"""Tudo o que o programa mostra ou pergunta passa por aqui.

As funções (ok, erro, escolher, perguntar, confirmar_envio...) repassam para o "motor" em
uso: a janela do app.py, que troca o motor com usar(). O motor Terminal (texto) fica só
para os comandos de teste dos desenvolvedores (python edocs_http.py ..., diario_http.py).
"""
import msvcrt
import sys

from rich.console import Console

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

console = Console(highlight=False)


class Terminal:
    """Motor de texto: mostra no console e lê do teclado."""

    CORES = {"ok": ("🟩", "green"), "erro": ("🟥", "red"), "aviso": ("🟨", "yellow"),
             "info": ("🔄️", "cyan")}

    def mensagem(self, tipo: str, texto: str) -> None:
        if tipo == "titulo":
            console.rule(f"[bold]{texto}[/]", style="dim")
        else:
            icone, cor = self.CORES[tipo]
            console.print(f"{icone} [{cor}]{texto}[/]")

    def processando(self, texto: str):
        return console.status(f"[cyan]{texto}[/]", spinner="dots")

    def progresso(self, n: int, total: int, descricao: str) -> None:
        self.mensagem("titulo", f"{n}/{total}")

    def escolher(self, pergunta: str, opcoes: list[str], multiplo: bool = False,
                 vazio: str | None = None, marcados: list[int] | None = None) -> list[int]:
        """Índices escolhidos (a partir de 0), na ordem digitada. Com multiplo=False devolve
        um só. vazio: o que significa não responder nada; None torna a resposta obrigatória.
        marcados: escolha anterior, repetida com Enter (aí 0 = vazio)."""
        console.print(f"\n[bold]{pergunta}[/]")
        for i, opcao in enumerate(opcoes, 1):
            console.print(f"  [bold cyan]{i}[/] {opcao}")
        if multiplo:
            dica = f"vazio = {vazio}" if vazio else "pelo menos um"
            if marcados:
                dica = ("Enter = os mesmos da última vez: " + ", ".join(opcoes[i] for i in marcados)
                        + (f" · 0 = {vazio}" if vazio else ""))
            console.print(f"[dim]Números separados por vírgula ({dica})[/]")
        while True:
            resposta = console.input("› ").strip()
            if marcados and not resposta:
                return list(marcados)
            if marcados and vazio and resposta == "0":
                return []
            numeros = resposta.replace(",", " ").split()
            validos = all(n.isdigit() and 1 <= int(n) <= len(opcoes) for n in numeros)
            if validos and (numeros or vazio) and (multiplo or len(numeros) <= 1):
                return [int(n) - 1 for n in dict.fromkeys(numeros)]
            aviso("Resposta inválida")

    def escolher_pessoas(self, pergunta: str, pesquisar, recentes: list[dict],
                         marcados: list[dict], vazio: str | None) -> list[dict]:
        """Pesquisa no site e monta a lista na ordem. Números escolhem dos últimos usados ou
        do resultado da pesquisa; texto pesquisa; Enter termina."""
        escolhidos = list(marcados)
        lista = list(recentes)
        console.print(f"\n[bold]{pergunta}[/]")
        console.print(f"[dim]{DICA_PESQUISA}[/]")
        while True:
            if escolhidos:
                console.print("Escolhidos: " + ", ".join(f"{i}º {p['nome']}"
                                                         for i, p in enumerate(escolhidos, 1)))
            for i, p in enumerate(lista, 1):
                console.print(f"  [bold cyan]{i}[/] {descrever_pessoa(p)}")
            fim = "termina" if escolhidos else (f"= {vazio}" if vazio else "")
            dica = "número = escolher · -número = tirar dos escolhidos · " if lista or escolhidos else ""
            resposta = console.input(f"[dim]{dica}texto = pesquisar"
                                     + (f" · Enter {fim}" if fim else "") + "[/]\n› ").strip()
            if not resposta:
                if escolhidos or vazio is not None:
                    return escolhidos
                aviso("Escolha pelo menos uma pessoa")
            elif resposta.lstrip("-").isdigit() and len(resposta.lstrip("-")) <= 2:
                n = int(resposta.lstrip("-"))
                if resposta.startswith("-") and 1 <= n <= len(escolhidos):
                    escolhidos.pop(n - 1)
                elif not resposta.startswith("-") and 1 <= n <= len(lista):
                    if all(p["cpf"] != lista[n - 1]["cpf"] for p in escolhidos):
                        escolhidos.append(lista[n - 1])
                else:
                    aviso("Número inválido")
            else:
                with processando("Pesquisando no Edocs"):
                    lista = pesquisar(resposta)
                if not lista:
                    aviso("Ninguém encontrado")

    def perguntar(self, pergunta: str, senha: bool = False) -> str:
        return ler_senha(f"{pergunta}: ") if senha else console.input(f"{pergunta}: ").strip()

    def login(self, titulo: str, explicacao: str) -> tuple[str, str]:
        console.print(f"[bold]{titulo}[/]")
        console.print(f"[dim]{explicacao}[/]")
        usuario = senha = ""
        while not usuario:
            usuario = console.input("Usuário: ").strip()
        while not senha:
            senha = ler_senha("Senha: ")
        return usuario, senha

    def confirmar(self, pergunta: str, sim: str = "Sim", nao: str = "Não") -> bool:
        return console.input(f"{pergunta} [dim](S = {sim} · N = {nao})[/]: ").strip().upper() == "S"

    def confirmar_envio(self, ato: str = "", texto: str = "", folha: str = "", word=None,
                        unico: bool = False) -> str:
        """S → 'enviar'; T → 'todos' (este e os próximos sem perguntar); N → 'pular';
        Esc/Ctrl+C → 'cancelar'. Enter não envia, para evitar acidente."""
        console.print("[bold]Está tudo ok para o envio?[/] [bold cyan]S/T/N[/]")
        console.print("[dim]S envia · T envia este e todos os próximos · N pula este ato · Esc/Ctrl+C cancela tudo[/]")
        while msvcrt.kbhit():
            msvcrt.getwch()
        while True:
            tecla = msvcrt.getwch()
            if tecla in ("s", "S"):
                return "enviar"
            if tecla in ("t", "T"):
                return "todos"
            if tecla in ("n", "N"):
                return "pular"
            if tecla in ("\x1b", "\x03"):  # Esc ou Ctrl+C
                return "cancelar"
            if tecla in ("\x00", "\xe0"):
                msvcrt.getwch()

    def anexos_dos_atos(self, lista: list[dict]) -> list[list] | str:
        """Para cada ato, os caminhos dos arquivos (separados por ;); vazio = sem anexo
        (pulado); Esc no primeiro cancela."""
        saida = []
        for item in lista:
            donos = item.get("pessoas") or [""]
            escolhas = []
            for dono in donos:
                quem = f" · {dono}" if dono else ""
                resposta = console.input(f"Anexo(s) de {item['ato']}{quem} [dim](caminhos separados por ;)[/]: ")
                caminhos = [c.strip().strip('"') for c in resposta.split(";") if c.strip()]
                escolhas += [[c, dono] for c in caminhos] if dono else caminhos
            saida.append(escolhas)
        return saida if any(saida) else "cancelar"


_motor = Terminal()


def usar(motor) -> None:
    """Troca o motor (a janela do app.py chama com o dela)."""
    global _motor
    _motor = motor


def ok(msg: str) -> None:
    _motor.mensagem("ok", msg)


def erro(msg: str) -> None:
    _motor.mensagem("erro", msg)


def aviso(msg: str) -> None:
    _motor.mensagem("aviso", msg)


def info(msg: str) -> None:
    _motor.mensagem("info", msg)


def titulo(texto: str) -> None:
    _motor.mensagem("titulo", texto)


def processando(msg: str):
    return _motor.processando(msg)


def progresso(n: int, total: int, descricao: str = "") -> None:
    """Ato n de total (começa em 1)."""
    _motor.progresso(n, total, descricao)


def escolher(pergunta: str, opcoes: list[str], multiplo: bool = False,
             vazio: str | None = None, marcados: list[int] | None = None) -> list[int]:
    """marcados: índices que já vêm escolhidos (ex.: a escolha da última vez)."""
    return _motor.escolher(pergunta, opcoes, multiplo, vazio, marcados)


DICA_PESQUISA = ("Pesquise por CPF, nome ou matrícula. A matrícula precisa do dígito "
                 "verificador no final (ex.: 123450).")


def descrever_pessoa(p: dict) -> str:
    """'NOME · CPF 012.345.678-90 · mat. 123450 · CARGO'."""
    cpf = p.get("cpf", "")
    partes = [p.get("nome") or "?", f"CPF {cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}" if cpf else "",
              f"mat. {p['matricula']}" if p.get("matricula") else "", p.get("cargo", "")]
    return " · ".join(x for x in partes if x)


def escolher_pessoas(pergunta: str, pesquisar, recentes: list[dict], marcados: list[dict],
                     vazio: str | None = None) -> list[dict]:
    """Pessoas na ordem escolhida. pesquisar(termo) → [pessoa] (no site); recentes: os
    últimos usados, para um clique; marcados: a escolha da última vez (já vem escolhida)."""
    return _motor.escolher_pessoas(pergunta, pesquisar, recentes, marcados, vazio)


def perguntar(pergunta: str, senha: bool = False) -> str:
    return _motor.perguntar(pergunta, senha)


def login(titulo: str, explicacao: str) -> tuple[str, str]:
    return _motor.login(titulo, explicacao)


def confirmar(pergunta: str, sim: str = "Sim", nao: str = "Não") -> bool:
    return _motor.confirmar(pergunta, sim, nao)


def confirmar_envio(ato: str = "", texto: str = "", folha: str = "", word=None, unico: bool = False) -> str:
    """'enviar', 'todos' (este e os próximos sem perguntar), 'pular' ou 'cancelar'.
    texto = o ato montado, para a janela mostrar; word() grava o ato num Word (para imprimir).
    unico: um ato só (Ato manual): só Enviar ou Cancelar."""
    return _motor.confirmar_envio(ato, texto, folha, word, unico)


def anexos_dos_atos(lista: list[dict]) -> list[list] | str:
    """Os anexos de todos os atos, perguntados de uma vez antes do envio. lista: um item por
    ato, {"ato": descrição, "pessoas": nomes (ato com várias) ou None}. Devolve, na mesma
    ordem, a lista de arquivos de cada ato (vazia = sem anexo, o ato é pulado), ou
    'cancelar'. Cada arquivo é o caminho ou, num ato com várias pessoas, [caminho, nome]."""
    return _motor.anexos_dos_atos(lista)


# --- Só no terminal -----------------------------------------------------------


def ler_senha(prompt: str = "Senha: ") -> str:
    """Lê a senha mostrando * a cada caractere (Windows)."""
    console.print(prompt, end="")
    sys.stdout.flush()
    senha = []
    while True:
        c = msvcrt.getwch()
        if c in ("\r", "\n"):
            print()
            return "".join(senha)
        if c == "\x03":
            raise KeyboardInterrupt
        if c in ("\x00", "\xe0"):  # setas, F1...
            msvcrt.getwch()
        elif c == "\b":
            if senha:
                senha.pop()
                sys.stdout.write("\b \b")
        elif c >= " ":  # ignora Esc, Tab e outros controles
            senha.append(c)
            sys.stdout.write("*")
        sys.stdout.flush()


