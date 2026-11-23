"""Testes offline da pesquisa de pessoas para vistos e assinantes (não acessa o site).

Usam as gravações reais (capturas/pesquisa nos vistos.har e pesquisa nos assinantes.har):
  1. lê matrícula, CPF, nome, cargo, lotação e o id interno das respostas da pesquisa;
  2. cada pesquisa preenche UM campo só: CPF (11 dígitos), matrícula (outros números) ou nome;
  3. a busca por CPF que não acha tenta como matrícula;
  4. a escolha guarda as 5 últimas pessoas e a escolha do processo (num config temporário),
     sem o próprio usuário e sem ids internos do site;
  5. com vistos e assinantes fixos no processo, o envio não pergunta (desligado, pergunta).

Rode com:  python testes/testar_pesquisa_pessoas.py
"""
import json
import sys
import tempfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import edocs_http  # noqa: E402
import main  # noqa: E402
import ui  # noqa: E402

HARS = [RAIZ / "capturas" / "pesquisa nos vistos.har", RAIZ / "capturas" / "pesquisa nos assinantes.har"]
F = edocs_http.F


class _RespostaFalsa:
    def __init__(self, texto: str, url: str = "https://x/f/n/servidoressel"):
        self.text = texto
        self.content = texto.encode("utf-8")
        self.url = url
        self.status_code = 200

    def raise_for_status(self) -> None:
        pass


def respostas_da_pesquisa(har: Path) -> list[tuple[dict, str]]:
    """[(campos enviados, html devolvido)] de cada pesquisa gravada."""
    saida = []
    for e in json.loads(har.read_text(encoding="utf-8"))["log"]["entries"]:
        r = e["request"]
        if r["method"] == "POST" and r["url"].endswith("servidoressel"):
            campos = {p["name"]: p["value"] for p in r["postData"]["params"]}
            saida.append((campos, e["response"]["content"]["text"]))
    return saida


def teste_1_le_resultados() -> None:
    total = 0
    for har in HARS:
        for campos, html in respostas_da_pesquisa(har):
            pessoas = edocs_http.servidores_da_tela(edocs_http.Tela(None, _RespostaFalsa(html)))
            assert len(pessoas) == 1, pessoas
            p = pessoas[0]
            assert p["oid"].isdigit() and len(p["cpf"]) == 11 and p["nome"], p
            assert p["matricula"].isdigit() and p["cargo"] and p["lotacao"] and p["lookup"], p
            # O resultado bate com o que foi pesquisado (na gravação, os campos das
            # pesquisas anteriores ficaram preenchidos)
            preenchidos = {k.split(":")[-1]: v for k, v in campos.items()
                           if k.split(":")[-1] in ("nrMatricula", "nrCpf", "nmPessoa") and v}
            assert preenchidos, campos
            for campo, valor in preenchidos.items():
                alvo = {"nrMatricula": p["matricula"], "nrCpf": p["cpf"], "nmPessoa": p["nome"]}[campo]
                assert valor.upper() in alvo.upper(), (campo, valor, p)
            total += 1
    assert total >= 6, total
    print(f"  [1] lê as {total} respostas reais (matrícula, CPF, nome...) OK")


class _TelaPesquisa:
    """Janela de seleção falsa: guarda o que foi pesquisado e devolve a resposta escolhida."""

    def __init__(self, enviados: list, resposta):
        self.enviados, self.resposta = enviados, resposta

    def enviar(self, botao, valores=None, parcial=False):
        self.enviados.append(valores)
        return self.resposta(valores)


def _pesquisar(termo: str, resposta) -> list[dict]:
    enviados: list = []
    guardado = edocs_http.abrir
    edocs_http.abrir = lambda sessao, caminho: _TelaPesquisa(enviados, resposta)
    try:
        edocs_http.pesquisar_servidores(None, termo)
    finally:
        edocs_http.abrir = guardado
    return enviados


def teste_2_um_campo_por_pesquisa() -> None:
    html = respostas_da_pesquisa(HARS[0])[0][1]
    com_resultado = lambda valores: edocs_http.Tela(None, _RespostaFalsa(html))  # noqa: E731
    assert _pesquisar("987.654.321-00", com_resultado) == [{F + "nrCpf": "987.654.321-00"}]
    assert _pesquisar("98765432100", com_resultado) == [{F + "nrCpf": "987.654.321-00"}]
    assert _pesquisar("12345-0", com_resultado) == [{F + "nrMatricula": "123450"}]
    assert _pesquisar("  gabriel   jalles ", com_resultado) == [{F + "nmPessoa": "gabriel jalles"}]
    print("  [2] CPF, matrícula ou nome: um campo por pesquisa ........ OK")


def teste_3_cpf_que_nao_acha_tenta_matricula() -> None:
    vazio = '<html><body><form id="corpo:formulario" action="servidoressel"></form></body></html>'
    sem_resultado = lambda valores: edocs_http.Tela(None, _RespostaFalsa(vazio))  # noqa: E731
    assert _pesquisar("98765432100", sem_resultado) == [
        {F + "nrCpf": "987.654.321-00"}, {F + "nrMatricula": "98765432100"}]
    print("  [3] 11 números sem resultado como CPF: tenta matrícula .. OK")


def teste_4_guarda_ultimos() -> None:
    pasta = Path(tempfile.mkdtemp())
    guardados = main.CONFIG, main.processo, ui._motor
    main.CONFIG, main.processo = pasta / "config.json", pasta / "Processo X"
    eu = {"cpf": "98765432100", "nome": "EU", "oid": "1"}
    gente = [{"cpf": f"{i:011d}", "nome": f"PESSOA {i}", "matricula": str(i), "cargo": "C",
              "lotacao": "L", "oid": str(i), "lookup": "x"} for i in range(1, 8)]
    vistos_pela_tela = []

    class Motor:
        def escolher_pessoas(self, pergunta, pesquisar, recentes, marcados, vazio):
            vistos_pela_tela.append((recentes, marcados, pesquisar("qualquer")))
            return self.resposta

    ui._motor = Motor()
    try:
        for i in range(0, 7, 2):  # escolhe 2 de cada vez: 1,2 · 3,4 · 5,6 · 7
            ui._motor.resposta = gente[i:i + 2]
            main.escolher_pessoas("?", lambda t: [eu] + gente, eu["cpf"], None, "vistos")
        cfg = json.loads(main.CONFIG.read_text(encoding="utf-8"))
        recentes, marcados, achados = vistos_pela_tela[-1]
        assert eu not in achados  # o próprio usuário nunca aparece
        assert [p["nome"] for p in recentes] == ["PESSOA 5", "PESSOA 6", "PESSOA 3", "PESSOA 4",
                                                 "PESSOA 1"], recentes
        assert [p["nome"] for p in marcados] == ["PESSOA 5", "PESSOA 6"]  # a escolha anterior
        assert [p["nome"] for p in cfg["ultimas_escolhas"]["Processo X"]["vistos"]] == ["PESSOA 7"]
        assert all("oid" not in p and "lookup" not in p for p in cfg["pessoas_recentes"])
    finally:
        main.CONFIG, main.processo, ui._motor = guardados
    print("  [4] guarda as 5 últimas pessoas e a escolha do processo .. OK")


def teste_5_fixos_do_processo() -> None:
    pasta = Path(tempfile.mkdtemp()) / "Processo Fixo"
    pasta.mkdir()
    guardados = main.CONFIG, main.processo, ui._motor
    main.CONFIG, main.processo = pasta.parent / "config.json", pasta
    perguntas = []

    class Motor:
        def mensagem(self, tipo, texto):
            pass

        def escolher_pessoas(self, pergunta, pesquisar, recentes, marcados, vazio):
            perguntas.append(pergunta)
            return [{"cpf": "33333333333", "nome": "C"}]

    ui._motor = Motor()
    eu = "98765432100"
    fixas = {"fixas": True, "vistos": [{"cpf": "11111111111", "nome": "A"}, {"cpf": eu, "nome": "EU"}],
             "assinantes": [{"cpf": "22222222222", "nome": "B"}]}
    try:
        # Fixos: não pergunta; quem envia está na lista e passa a vistar em 1º
        json.dump({"pessoas": fixas}, (pasta / "processo.json").open("w", encoding="utf-8"))
        assert main.pessoas_do_envio(eu, None) == ([eu, "11111111111"], ["22222222222"])
        assert perguntas == []
        # Quem envia fora da lista não vista (nada é colocado sozinho)
        sem_mim = fixas | {"vistos": [{"cpf": "11111111111", "nome": "A"}]}
        json.dump({"pessoas": sem_mim}, (pasta / "processo.json").open("w", encoding="utf-8"))
        assert main.pessoas_do_envio(eu, None) == (["11111111111"], ["22222222222"])
        # Desligado: pergunta vistos e assinantes; só vista quem foi escolhido
        json.dump({"pessoas": fixas | {"fixas": False}}, (pasta / "processo.json").open("w", encoding="utf-8"))
        assert main.pessoas_do_envio(eu, lambda t: []) == (["33333333333"], ["33333333333"])
        assert len(perguntas) == 2
        # Processo antigo (sem "pessoas" no processo.json): pergunta
        json.dump({}, (pasta / "processo.json").open("w", encoding="utf-8"))
        main.pessoas_do_envio(eu, lambda t: [])
        assert len(perguntas) == 4
    finally:
        main.CONFIG, main.processo, ui._motor = guardados
    print("  [5] vistos fixos dispensam a pergunta; quem envia só vista se estiver (em 1º) OK")


def main_() -> None:
    print("Testes da pesquisa de pessoas no Edocs (offline)\n")
    for teste in (teste_1_le_resultados, teste_2_um_campo_por_pesquisa,
                  teste_3_cpf_que_nao_acha_tenta_matricula, teste_4_guarda_ultimos,
                  teste_5_fixos_do_processo):
        teste()
    print("\n5 testes passaram.")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main_()
