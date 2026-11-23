"""Tabelas de consulta: uma tabela de fora (ex.: todos os servidores da Prefeitura) que
completa a linha da planilha do processo. A pessoa escreve o começo do nome (ou a matrícula)
na Nova linha, escolhe na lista e o resto vem da tabela.

A tabela é cadastrada uma vez (vale para todos os processos) e fica em configs/config.json:
    "tabelas_consulta": {"Servidores": {"arquivo": "C:/.../servidores.xlsx", "aba": "",
                                        "atualizar_dias": 7, "atualizada_em": "2026-10-20"}}
"atualizar_dias" (0 = nunca): passado esse prazo, a Nova linha pede a tabela mais nova.

Cada processo diz como usar a tabela (processo.json, "consulta"):
    {"tabela": "Servidores",
     "chaves": [{"minha": "NOME", "dela": "NOME COMPLETO"}, {"minha": "MATRÍCULA", "dela": "MATRIC"}],
     "puxar":  [{"minha": "SEXO", "dela": "SEXO", "traducao": {"FEMININO": "F", "MASCULINO": "M"}}, ...]}
"minha" é a coluna da planilha do processo; "dela", a da tabela. Um item pode juntar a
coluna dela com outras ("juntar": {"com": ["dv_matricula"], "separador": "-", "milhar": true}:
30244 e 9 viram 30.244-9). E pode só valer em certos casos ("se": [{"coluna": "cd_cargo_efetivo",
"op": "tem_valor"}]: o cargo efetivo só vem para quem tem um; ver condicoes_valem). As chaves acham a pessoa
(por qualquer uma delas); "puxar" diz o que copiar. O que vem da tabela é copiado: se a
tabela mudar depois, os atos já feitos não mudam.

As sugestões (sugerir_chaves, sugerir_puxar, sugerir_traducao) comparam os nomes das
colunas, para a pessoa só conferir.
"""
import csv
import difflib
import json
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import openpyxl

import atos
import de_para
import main

CHAVE_CONFIG = "tabelas_consulta"
NOME_PASTA = "tabelas de consulta"
PASTA = main.ENTRADAS / NOME_PASTA  # deste computador (as que vieram no pacote do build)


def pasta() -> Path:
    """Onde ficam as tabelas trazidas (importadas, ou levadas para a pasta compartilhada)."""
    return (main.extras_compartilhados() or main.ENTRADAS) / NOME_PASTA
MINIMO = 0.75  # parecença mínima entre nomes de colunas para sugerir a ligação
# Nomes diferentes para a mesma coisa (normalizados), comuns nas tabelas de pessoal
SINONIMOS = {
    "NOME": {"NOMECOMPLETO", "NOMEDOSERVIDOR", "NOMESERVIDOR", "SERVIDOR", "NOMEFUNCIONARIO", "FUNCIONARIO",
             "NMPESSOA", "NOMEPESSOA", "NMSERVIDOR"},
    "MATRICULA": {"MATRIC", "MATR", "MAT", "NMATRICULA", "NUMEROMATRICULA", "REGISTRO"},
    "SECRETARIA": {"LOTACAO", "ORGAO", "SECRETARIALOTACAO", "ORGAOLOTACAO"},
    # NÚCLEO é o nome usado aqui; UNIDADE é o mesmo (nome antigo, ainda comum nas tabelas)
    "NUCLEO": {"UNIDADE", "UNIDADELOTACAO", "NUCLEOLOTACAO", "SETOR", "SETORLOTACAO", "DEPARTAMENTO"},
    "UNIDADE": {"NUCLEO", "UNIDADELOTACAO", "NUCLEOLOTACAO", "SETOR", "SETORLOTACAO", "DEPARTAMENTO"},
    "SEXO": {"GENERO"},
    "NIVEL": {"NIVELQUALIFICACAO", "NIVELDEQUALIFICACAO", "QUALIFICACAO"},
    "CPF": {"NCPF", "NUMEROCPF"},
}


# --- Tabelas cadastradas ------------------------------------------------------------

def tabelas() -> dict[str, dict]:
    """{nome: {"arquivo", "aba", "atualizar_dias", "atualizada_em"}} das tabelas cadastradas."""
    return dict(main.ler_config().get(CHAVE_CONFIG) or {})


def salvar_tabela(nome: str, arquivo: Path | str, aba: str = "", atualizar_dias: int = 0) -> None:
    """Cadastra (ou troca o arquivo de) uma tabela: conta como atualizada hoje. Com a pasta
    compartilhada, o arquivo é copiado para lá (o de Downloads ninguém mais enxerga)."""
    raiz = main.pasta_compartilhada()
    arquivo = Path(arquivo)
    if raiz and raiz.resolve() not in arquivo.resolve().parents:
        import shutil
        pasta().mkdir(parents=True, exist_ok=True)
        shutil.copy2(arquivo, pasta() / arquivo.name)
        arquivo = pasta() / arquivo.name
    todas = tabelas()
    todas[nome] = {"arquivo": str(arquivo), "aba": aba, "atualizar_dias": max(0, int(atualizar_dias or 0)),
                   "atualizada_em": date.today().isoformat()}
    main.salvar_config(**{CHAVE_CONFIG: todas})


PACOTE = main.ENTRADAS / "veio_no_pacote.json"  # o build.py põe aqui as tabelas e os de-para


def receber_do_pacote() -> list[str]:
    """Primeira abertura de um Edodocs novo: cadastra as tabelas de consulta e junta os de-para
    que vieram no pacote (build.py), sem trocar o que já existe aqui. Roda uma vez: depois o
    arquivo some (se a pessoa remover uma tabela, ela não volta sozinha). Devolve o que entrou."""
    if not PACOTE.is_file():
        return []
    try:
        dados = json.loads(PACOTE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []
    entrou = []
    cadastradas = tabelas()
    for nome, tabela in (dados.get("tabelas") or {}).items():
        arquivo = PASTA / Path(str(tabela.get("arquivo", ""))).name
        atual = cadastradas.get(nome)
        if arquivo.is_file() and not (atual and Path(atual.get("arquivo") or "").is_file()):
            cadastrar(nome, dict(tabela, arquivo=str(arquivo)))
            entrou.append(f"tabela {nome}")
    for nome, pares in (dados.get("de_para") or {}).items():
        if isinstance(pares, dict):
            de_para.juntar(nome, {str(k): str(v) for k, v in pares.items()})
            entrou.append(f"de-para {nome}")
    PACOTE.unlink(missing_ok=True)
    return entrou


def cadastrar(nome: str, dados: dict) -> None:
    """Cadastra a tabela como veio (ex.: importada com os processos, com o prazo e a data)."""
    todas = tabelas()
    todas[nome] = dict(dados)
    main.salvar_config(**{CHAVE_CONFIG: todas})


def ajustar_tabela(nome: str, **mudancas) -> None:
    """Muda só alguns dados da tabela (ex.: atualizar_dias=7, atualizada_em="2026-10-20")."""
    todas = tabelas()
    if nome in todas:
        todas[nome] = {**todas[nome], **mudancas}
        main.salvar_config(**{CHAVE_CONFIG: todas})


def marcar_atualizada(nome: str) -> None:
    ajustar_tabela(nome, atualizada_em=date.today().isoformat())


def atualizada_em(dados: dict) -> date | None:
    """Quando os dados foram atualizados: o dia em que a tabela foi escolhida ou, se o arquivo
    foi trocado depois no mesmo lugar, o dia em que ele mudou."""
    dias = []
    try:
        dias.append(date.fromisoformat(dados.get("atualizada_em") or ""))
    except ValueError:
        pass
    try:
        dias.append(datetime.fromtimestamp(Path(dados["arquivo"]).stat().st_mtime).date())
    except (OSError, KeyError):
        pass
    return max(dias) if dias else None


def dias_sem_atualizar(nome: str) -> int | None:
    """Há quantos dias a tabela não é atualizada, se já passou do prazo que a pessoa escolheu
    (atualizar_dias; 0 = nunca pede). None: está em dia."""
    dados = tabelas().get(nome) or {}
    prazo = int(dados.get("atualizar_dias") or 0)
    quando = atualizada_em(dados)
    if not prazo or not quando:
        return None
    dias = (date.today() - quando).days
    return dias if dias >= prazo else None


def remover_tabela(nome: str) -> None:
    todas = tabelas()
    todas.pop(nome, None)
    main.salvar_config(**{CHAVE_CONFIG: todas})


# --- Leitura ------------------------------------------------------------------------

@dataclass
class Tabela:
    colunas: list[str]                               # como estão no cabeçalho
    linhas: list[dict[str, str]] = field(default_factory=list)  # {coluna: valor em texto}

    def valores(self, coluna: str, limite: int = 2000) -> list[str]:
        """Os valores diferentes de uma coluna (na ordem em que aparecem)."""
        vistos = dict.fromkeys(l.get(coluna, "") for l in self.linhas[:limite])
        return [v for v in vistos if v]

    def exemplo(self, coluna: str) -> str:
        return next((l[coluna] for l in self.linhas if l.get(coluna)), "")


_cache: dict[tuple, Tabela] = {}


def abas(arquivo: Path) -> list[str]:
    if Path(arquivo).suffix.lower() == ".csv":
        return []
    wb = openpyxl.load_workbook(arquivo, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def ler(arquivo: Path | str, aba: str = "") -> Tabela:
    """Lê a tabela (Excel ou CSV). O cabeçalho é a primeira linha com 2 ou mais células
    preenchidas. Guarda em memória até o arquivo mudar."""
    arquivo = Path(arquivo)
    chave = (str(arquivo), aba, arquivo.stat().st_mtime)
    if chave not in _cache:
        _cache.clear()  # uma tabela por vez em memória (podem ser grandes)
        _cache[chave] = _ler_csv(arquivo) if arquivo.suffix.lower() == ".csv" else _ler_excel(arquivo, aba)
    return _cache[chave]


def ler_cadastrada(nome: str) -> Tabela:
    dados = tabelas().get(nome)
    if not dados:
        raise KeyError(f"A tabela {nome} não está cadastrada")
    return ler(dados["arquivo"], dados.get("aba", ""))


def _montar(linhas) -> Tabela:
    cab: list[str] = []
    registros = []
    for linha in linhas:
        valores = [atos.formatar(v) for v in linha]
        if not cab:
            if sum(bool(v) for v in valores) >= 2:
                cab = [v or f"Coluna {i + 1}" for i, v in enumerate(valores)]
            continue
        if any(valores):
            registros.append({c: v for c, v in zip(cab, valores) if v})
    return Tabela(cab, registros)


def _ler_excel(arquivo: Path, aba: str) -> Tabela:
    wb = openpyxl.load_workbook(arquivo, read_only=True, data_only=True)
    try:
        ws = wb[aba] if aba and aba in wb.sheetnames else wb.worksheets[0]
        return _montar(ws.iter_rows(values_only=True))
    finally:
        wb.close()


def _ler_csv(arquivo: Path) -> Tabela:
    bruto = arquivo.read_bytes()
    for codificacao in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            texto = bruto.decode(codificacao)
            break
        except UnicodeDecodeError:
            continue
    try:
        dialeto = csv.Sniffer().sniff(texto[:5000], delimiters=";,\t|")
    except csv.Error:
        dialeto = csv.excel
    return _montar(csv.reader(texto.splitlines(), dialeto))


# --- Sugestões -----------------------------------------------------------------------

def _palavras(nome: str) -> list[str]:
    return [atos.normalizar(p) for p in re.split(r"[^0-9A-Za-zÀ-ÿº°]+", nome) if atos.normalizar(p)]


def parecenca(minha: str, dela: str) -> float:
    """0 a 1: o quanto o nome de uma coluna lembra o da outra ("MATRÍCULA" e "MATRIC": 0,8)."""
    a, b = atos.normalizar(minha), atos.normalizar(dela)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if b in SINONIMOS.get(a, ()) or a in SINONIMOS.get(b, ()):
        return 0.9
    curta, longa = sorted((a, b), key=len)
    if curta in longa and len(curta) >= 3:
        # dv_matricula é só o dígito: cd_matricula (o número) ganha dela
        digito = longa.startswith("DV") and not curta.startswith("DV")
        return 0.6 + 0.3 * len(curta) / len(longa) - (0.1 if digito else 0)
    pa, pb = set(_palavras(minha)), set(_palavras(dela))
    if pa and pb and (pa <= pb or pb <= pa):
        return 0.6 + 0.3 * len(curta) / len(longa)
    return difflib.SequenceMatcher(None, a, b).ratio() * 0.95


def _melhores(pares: list[tuple[float, str, str]]) -> dict[str, str]:
    """Liga cada coluna minha à mais parecida da tabela, sem usar a mesma duas vezes."""
    escolhidos: dict[str, str] = {}
    usadas: set[str] = set()
    for nota, minha, dela in sorted(pares, reverse=True):
        if nota >= MINIMO and minha not in escolhidos and dela not in usadas:
            escolhidos[minha] = dela
            usadas.add(dela)
    return escolhidos


def sugerir_chaves(minhas: list[str], tabela: Tabela) -> list[dict]:
    """Como achar a pessoa: nome, matrícula e CPF, os que existirem dos dois lados."""
    candidatas = [m for m in minhas if atos.normalizar(m) in ("NOME", "MATRICULA", "CPF")
                  or any(atos.normalizar(m) in s for s in (SINONIMOS["NOME"], SINONIMOS["MATRICULA"]))]
    pares = [(parecenca(m, d), m, d) for m in candidatas for d in tabela.colunas]
    ligados = _melhores(pares)
    return [_com_digito({"minha": m, "dela": ligados[m]}, tabela) for m in minhas if m in ligados]


def _com_digito(item: dict, tabela: Tabela) -> dict:
    """cd_matricula com uma dv_matricula ao lado: já vem juntada (30.244-9)."""
    resto = re.sub(r"^CD", "", atos.normalizar(item["dela"]))
    if resto != atos.normalizar(item["dela"]):
        digito = next((c for c in tabela.colunas if atos.normalizar(c) == "DV" + resto), None)
        if digito:
            item["juntar"] = {"com": [digito], "separador": "-", "milhar": True}
    return item


def sugerir_puxar(minhas: list[str], tabela: Tabela, chaves: list[dict],
                  opcoes: dict[str, list[str]] | None = None) -> list[dict]:
    """O que copiar da tabela: cada coluna minha (fora as chaves) com a mais parecida da
    tabela. Coluna minha com lista de opções (F/M) ganha a tradução dos valores."""
    opcoes = opcoes or {}
    usadas = {c["dela"] for c in chaves}
    livres = [m for m in minhas if m not in {c["minha"] for c in chaves}]
    pares = [(parecenca(m, d), m, d) for m in livres for d in tabela.colunas if d not in usadas]
    ligados = _melhores(pares)
    saida = []
    for minha in livres:
        if minha in ligados:
            saida.append(item_puxar(minha, ligados[minha], tabela, opcoes.get(minha)))
    return saida


def item_puxar(minha: str, dela: str, tabela: Tabela, lista: list[str] | None = None,
               nome_de_para: str | None = "sugerir", parte: dict | None = None) -> dict:
    """Uma coluna de "puxar": com o de-para que conhece os códigos dela (se houver; parte:
    só um pedaço do valor é o código) e a tradução para a lista de opções da minha coluna
    (sobre os valores já com o de-para)."""
    valores = tabela.valores(dela)
    if nome_de_para == "sugerir":
        nome_de_para = de_para.sugerir(valores)
    item = {"minha": minha, "dela": dela}
    if nome_de_para:
        item["de_para"] = nome_de_para
        if parte and int(parte.get("n") or 0) > 0:
            item["de_para_parte"] = {"de": parte.get("de", "inicio"), "n": int(parte["n"])}
        listas = de_para.todos()
        valores = list(dict.fromkeys(de_para.traduzir(nome_de_para, v, listas, item.get("de_para_parte"))
                                     for v in valores))
    item["traducao"] = sugerir_traducao(lista, valores) if lista else {}
    if not lista and (formato := sugerir_formato(minha, valores)):
        item["formato"] = formato  # coluna com lista de opções: a tradução é que decide
    return item


def sugerir_traducao(opcoes: list[str], valores: list[str]) -> dict[str, str]:
    """Valores da tabela que viram uma das opções da minha coluna: "FEMININO" → "F",
    "COMISSIONADO" → "COMISSÃO". Os iguais (sem acento/maiúscula) não precisam de tradução."""
    alvo = {atos.normalizar(o): o for o in opcoes if o}
    traducao = {}
    for valor in valores[:60]:
        n = atos.normalizar(valor)
        if not n or n in alvo:
            continue
        melhor, tamanho = None, 0
        for no, original in alvo.items():
            comum = len(_prefixo(n, no))
            # "F" e "FEMININO": a opção inteira no começo; "COMISSAO" e "COMISSIONADO": 6 letras
            if (comum == len(no) or comum >= 5) and comum > tamanho:
                melhor, tamanho = original, comum
        if melhor:
            traducao[valor] = melhor
    return traducao


def _prefixo(a: str, b: str) -> str:
    n = 0
    while n < min(len(a), len(b)) and a[n] == b[n]:
        n += 1
    return a[:n]


# --- Uso: achar a pessoa e completar a linha -----------------------------------------

def _digitos(texto: str) -> str:
    return re.sub(r"\D", "", str(texto))


def buscar(consulta: dict, tabela: Tabela, termo: str, limite: int = 8) -> list[dict]:
    """Linhas da tabela em que alguma chave combina com o que foi digitado: palavras pelo
    começo ("paulo ces" acha "PEDRO MODELO COSTA") ou números ("38830" acha "38.830-0")."""
    termo = str(termo or "").strip()
    palavras = [atos.normalizar(p) for p in termo.split() if atos.normalizar(p)]
    numero = _digitos(termo) if re.fullmatch(r"[\d.\-/\s]+", termo) else ""
    if not palavras or (not numero and sum(map(len, palavras)) < 2):
        return []
    chaves = consulta.get("chaves") or []
    achados = []
    for linha in tabela.linhas:
        nota = 0
        for chave in chaves:
            valor = valor_juntado(chave, linha)
            if not valor:
                continue
            if numero:
                # Com o dígito juntado: "30244" é a matrícula 30.244-x (não a 3.024-4); "30.244-9", a inteira
                d, bruto = _digitos(valor), _digitos(linha.get(chave.get("dela", ""), ""))
                nota = max(nota, (2 if "-" in termo else 3) if bruto == numero else 0,
                           (3 if "-" in termo or bruto == d else 2.5) if d == numero else 0,
                           2 if d.startswith(numero) or bruto.startswith(numero) else 0)
            else:
                nomes = [atos.normalizar(p) for p in valor.split()]
                if all(any(n.startswith(p) for n in nomes) for p in palavras):
                    nota = max(nota, 3 if atos.normalizar(valor) == "".join(palavras)
                               else 2 if nomes[:1] and nomes[0].startswith(palavras[0]) else 1)
        if nota:
            achados.append((-nota, linha))
            if len(achados) >= limite * 20:
                break
    achados.sort(key=lambda x: x[0])
    return [linha for _, linha in achados[:limite]]


def completar(consulta: dict, linha: dict) -> dict[str, str]:
    """{coluna minha: valor} com o que vem da tabela: as chaves e as colunas de "puxar",
    com o código trocado pela descrição (de-para) e traduzidas (FEMININO → F)."""
    saida = {}
    listas = de_para.todos()
    normalizada = None
    for item in (consulta.get("chaves") or []) + (consulta.get("puxar") or []):
        if item.get("se"):
            normalizada = normalizada or {atos.normalizar(k): v for k, v in linha.items()}
            if not condicoes_valem(item["se"], normalizada):
                continue  # ex.: CARGO EFETIVO de quem não tem cargo efetivo (cd_cargo_efetivo = 0)
        valor = _sem_traducao(item, linha, listas)
        if valor:
            traducao = {atos.normalizar(k): v for k, v in (item.get("traducao") or {}).items()}
            traduzido = traducao.get(atos.normalizar(valor))
            # O formato (ASSESSOR → Assessor) vale para o que vem da tabela; o que a tradução
            # escolheu da lista (F, COMISSÃO) fica como está na lista
            saida[item["minha"]] = traduzido if traduzido else formatar(valor, item.get("formato"))
    return saida


# Comparações do "Só preencher se" (a da tabela de servidores: código 0 = não tem)
OPERADORES_SE = {"tem_valor": "tem valor (não está vazio e não é 0)", "preenchido": "está preenchido",
                 "vazio": "está vazio", "igual": "é", "diferente": "não é", "contem": "contém"}
SEM_VALOR_SE = {"tem_valor", "preenchido", "vazio"}


def condicoes_valem(condicoes: list[dict], linha_normalizada: dict) -> bool:
    """Todas as condições do "Só preencher se" valem para esta pessoa da tabela?"""
    import processos  # as comparações são as mesmas das regras do processo
    for cond in condicoes:
        if cond.get("op") == "tem_valor":
            valor = str(linha_normalizada.get(atos.normalizar(cond.get("coluna", "")), "")).strip()
            if not valor or re.fullmatch(r"0+([.,]0+)?", valor):
                return False
        elif not processos.condicao_vale({"campo": cond.get("coluna", ""), "op": cond.get("op", "igual"),
                                          "valor": cond.get("valor", "")}, linha_normalizada):
            return False
    return True


def descrever_se(condicoes: list[dict]) -> str:
    """'cd_cargo_efetivo tem valor (não está vazio e não é 0)', para mostrar na tela."""
    return " e ".join(f"{c.get('coluna', '?')} {OPERADORES_SE.get(c.get('op', 'igual'), c.get('op'))}"
                      + ("" if c.get("op") in SEM_VALOR_SE else f" {c.get('valor', '')}") for c in condicoes)


FORMATOS = {"": "Como está na tabela", "maiusculas": "MAIÚSCULAS", "titulo": "Primeiras Maiúsculas",
            "minusculas": "minúsculas"}
# Ficam minúsculas no meio do texto, em "Primeiras Maiúsculas"
PEQUENAS = {"de", "da", "do", "das", "dos", "e", "em", "na", "no", "nas", "nos", "a", "o", "as", "os", "ao",
            "aos", "à", "às", "para", "por", "com"}


def formatar(texto: str, formato: str | None) -> str:
    """Maiúsculas/minúsculas do valor. "titulo": "ASSESSOR TÉCNICO DAM-15" → "Assessor Técnico
    DAM-15"; de, da, do... ficam minúsculos; siglas com número (DAM-15) e romanos (II) ficam."""
    if not formato or not texto:
        return texto
    if formato == "maiusculas":
        return texto.upper()
    if formato == "minusculas":
        return texto.lower()
    palavras = texto.split(" ")
    saida = []
    def pedaco(texto: str, antes_de_numero: bool) -> str:
        if re.search(r"\d", texto) or antes_de_numero or (re.fullmatch(r"[IVX]+", texto) and len(texto) <= 4):
            return texto  # DAM-15, CC-6, II: sigla e número ficam
        return re.sub(r"[^\W\d_]+", lambda m: m.group(0)[0].upper() + m.group(0)[1:].lower(), texto)

    for i, palavra in enumerate(palavras):
        if i and palavra.lower() in PEQUENAS:
            saida.append(palavra.lower())
            continue
        partes = palavra.split("-")  # ESPECIAIS-CC-6: cada pedaço do seu jeito → Especiais-CC-6
        saida.append("-".join(pedaco(p, j + 1 < len(partes) and bool(re.search(r"\d", partes[j + 1])))
                              for j, p in enumerate(partes)))
    return " ".join(saida)


def sugerir_formato(minha: str, valores: list[str]) -> str:
    """Texto todo em maiúsculas na tabela (ASSESSOR TÉCNICO): sugere Primeiras Maiúsculas. O
    nome da pessoa fica como está (nos atos, ele vai em maiúsculas)."""
    if "NOME" in atos.normalizar(minha):
        return ""
    textos = [v for v in valores[:200] if re.search(r"[^\W\d_]{3,}", v)]
    if not textos:
        return ""
    maiusculos = sum(v == v.upper() and " " in v.strip() for v in textos)
    return "titulo" if maiusculos / len(textos) >= 0.7 else ""


def descricao(consulta: dict, linha: dict, campos: int = 4) -> str:
    """Uma linha para a lista de resultados: as chaves e o cargo/secretaria (ajudam a separar
    pessoas com o mesmo nome)."""
    def ordem(item: dict) -> int:
        nome = atos.normalizar(item["minha"])
        return 0 if re.search(r"SECRET|LOTAC|ORGAO", nome) else 1 if re.search(r"CARGO|NUCLEO|UNIDADE", nome) else 2

    puxar = sorted(consulta.get("puxar") or [], key=ordem)
    itens = list({i["dela"]: i for i in reversed((consulta.get("chaves") or []) + puxar)}.values())
    itens = list(reversed(itens))[:campos]  # um por coluna da tabela, na ordem: chaves, secretaria, cargo
    listas = de_para.todos()
    valores = (_sem_traducao(i, linha, listas) for i in itens)  # a matrícula inteira, a descrição
    return " · ".join(v for v in valores if v)[:160]


def _milhar(numero: str) -> str:
    """"30244" → "30.244" (só número; o resto fica como está)."""
    if not numero.isdigit() or len(numero) <= 3:
        return numero
    grupos = []
    while numero:
        grupos.insert(0, numero[-3:])
        numero = numero[:-3]
    return ".".join(grupos)


def valor_juntado(item: dict, linha: dict) -> str:
    """O valor da coluna da tabela, juntado com as outras do item (ex.: cd_matricula e
    dv_matricula → 30.244-9)."""
    valor = linha.get(item.get("dela", ""), "")
    juntar = item.get("juntar")
    if not juntar or not valor:
        return valor
    if juntar.get("milhar"):
        valor = _milhar(valor)
    partes = [linha.get(c, "") for c in juntar.get("com") or []]
    return str(juntar.get("separador", "-")).join([valor] + [p for p in partes if p])


def _sem_traducao(item: dict, linha: dict, listas: dict) -> str:
    """Juntado e com o de-para (a descrição no lugar do código); falta só a tradução."""
    valor = valor_juntado(item, linha)
    if valor and item.get("de_para"):
        valor = de_para.traduzir(item["de_para"], valor, listas, item.get("de_para_parte"))
    return valor
