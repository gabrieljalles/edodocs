import copy
import html
from html.entities import codepoint2name
import re
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run

W14_ATTRS = (
    "{http://schemas.microsoft.com/office/word/2010/wordml}paraId",
    "{http://schemas.microsoft.com/office/word/2010/wordml}textId",
)


def normalizar(txt) -> str:
    txt = unicodedata.normalize("NFKD", str(txt or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Z0-9]", "", txt.upper())


# Aba usada pelo modelo do Word (consulta original da mala direta)
ABA = "processosPendentes"

UNIDADES = ["zero", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove",
            "dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete",
            "dezoito", "dezenove"]
DEZENAS = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta",
           "oitenta", "noventa"]
CENTENAS = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos",
            "setecentos", "oitocentos", "novecentos"]


def por_extenso(n: int) -> str:
    if n < 20:
        return UNIDADES[n]
    if n < 100:
        d, u = divmod(n, 10)
        return DEZENAS[d] + (f" e {UNIDADES[u]}" if u else "")
    if n == 100:
        return "cem"
    if n < 1000:
        c, r = divmod(n, 100)
        return CENTENAS[c] + (f" e {por_extenso(r)}" if r else "")
    return str(n)


# Erros de fórmula do Excel são tratados como célula vazia
ERROS_EXCEL = {"#NAME?", "#VALUE!", "#REF!", "#N/A", "#DIV/0!", "#NUM!", "#NULL!"}


def formatar(v) -> str:
    if v is None or (isinstance(v, str) and v.strip() in ERROS_EXCEL):
        return ""
    if isinstance(v, datetime):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def ler_planilha(caminho: Path, aba: str | None = None, obrigatoria: str = "NOME") -> list[dict]:
    """Lê a planilha; cada registro é {coluna normalizada: valor}. Ignora as linhas em que a
    coluna obrigatória (normalizada, ex.: 'NOME') está vazia. Sem aba, usa a 'processosPendentes'
    (a da mala direta antiga) ou a primeira."""
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    aba = aba or ABA
    ws = wb[aba] if aba in wb.sheetnames else wb.active
    linhas = ws.iter_rows(values_only=True)
    cab = []
    for c in next(linhas):
        c = normalizar(c)
        # Coluna repetida ganha número, como no Word: a 2ª "ANO" vira "ANO1"
        if c and c in cab:
            n = 1
            while f"{c}{n}" in cab:
                n += 1
            c = f"{c}{n}"
        cab.append(c)

    registros = []
    for n, linha in enumerate(linhas, 2):
        dados = {c: formatar(v) for c, v in zip(cab, linha) if c}
        if dados.get(obrigatoria):
            dados["_LINHA"] = n  # número da linha no Excel (para gravar o status)
            registros.append(dados)
    wb.close()
    return registros


def cabecalho_planilha(caminho: Path, aba: str | None = None) -> tuple[list[str], list[str]]:
    """(abas da planilha, colunas da aba) com os nomes como estão no Excel. Lê só o
    cabeçalho direto do arquivo (rápido mesmo em planilha grande)."""
    import planilha_status

    abas, ativa = planilha_status.abas(caminho)
    aba = aba or ABA
    cabecalho = planilha_status.cabecalho(caminho, aba if aba in abas else ativa)
    colunas, vistas = [], set()
    for _, nome in sorted(cabecalho.items()):
        nome = nome.strip()
        if not nome:
            continue
        # Coluna repetida ganha número, como no Word e em ler_planilha: a 2ª "ANO" vira "ANO1"
        if normalizar(nome) in vistas:
            n = 1
            while normalizar(f"{nome}{n}") in vistas:
                n += 1
            nome = f"{nome}{n}"
        vistas.add(normalizar(nome))
        colunas.append(nome)
    return abas, colunas


def proximo_dia_util(dia: date) -> date:
    """Dia útil seguinte (pula sábado e domingo). Feriados não são considerados."""
    dia += timedelta(days=1)
    while dia.weekday() >= 5:
        dia += timedelta(days=1)
    return dia


def ler_data(texto: str) -> date | None:
    """Lê 'dd/mm/aaaa' (com ou sem hora) ou 'aaaa-mm-dd'."""
    texto = (texto or "").strip()
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(texto, fmt).date()
        except ValueError:
            pass
    return None


def _nome_campo(instr: str) -> tuple[str, str]:
    """(nome normalizado, nome para mostrar): 'MERGEFIELD "Nº_PROC"' -> ('NOPROC', 'Nº PROC')."""
    m = re.search(r'MERGEFIELD\s+"?([^"]+?)"?\s*(\\|$)', instr.strip())
    if not m:
        return "", ""
    return normalizar(m.group(1)), " ".join(m.group(1).replace("_", " ").split())


# Campo digitado no Word: {{NOME}}, {{ Nº DO PROCESSO }}...
CHAVES = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")

# Marcadores do trecho do Word que se repete para cada pessoa do ato: os parágrafos de
# {{REPETIR}} até {{FIM REPETIR}} (os marcadores somem do ato montado)
REPETIR, FIM_REPETIR = "REPETIR", "FIMREPETIR"

# Campos que o programa preenche sozinho: {nome normalizado: (nome para mostrar, exemplo)}
QUANTIDADE = normalizar("QUANTIDADE DE PESSOAS")
TITULO_NO_ATO = "_TITULO_NO_ATO"  # vai no lugar da 1ª linha do Word (Processo.valores)
AUTOMATICOS = {normalizar(rotulo): (rotulo, exemplo) for rotulo, exemplo in (
    ("QUANTIDADE DE PESSOAS", "quantas pessoas o ato tem: 1, 2, 3..."),
    ("ITEM", "I, II, III... (no trecho que repete)"),
    ("Nº DO ITEM", "1, 2, 3... (no trecho que repete)"),
    ("LETRA DO ITEM", "a, b, c... (no trecho que repete)"),
    ("FIM DO ITEM", "; nas pessoas do meio e . na última"),
)}


def romano(n: int) -> str:
    saida = ""
    for valor, letras in ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
                          (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")):
        while n >= valor:
            saida += letras
            n -= valor
    return saida


def campos_do_item(n: int, total: int) -> dict:
    """Campos automáticos da n-ésima pessoa (1, 2...) no trecho que repete."""
    return {"ITEM": romano(n), normalizar("Nº DO ITEM"): str(n),
            "LETRADOITEM": chr(ord("a") + (n - 1) % 26), "FIMDOITEM": "." if n == total else ";"}


def itens(dados: dict) -> list[dict]:
    """As linhas da planilha que formam o ato (uma só, se o processo não junta pessoas)."""
    return dados.get("_ITENS") or [dados]


def _rotulo(nome: str) -> str:
    return " ".join(nome.replace("_", " ").split())


def campos_do_modelo(modelo: Path) -> list[tuple[str, str]]:
    """Campos do Word, na ordem em que aparecem e sem repetir: os de mala direta («NOME») e
    os digitados entre chaves duplas ({{NOME}}).
    [(nome normalizado, nome para mostrar)], ex.: [('NOPROC', 'Nº PROC'), ...]."""
    doc = Document(modelo)
    campos, instr, dentro = {}, "", False
    for p in doc.element.body.iter(qn("w:p")):
        texto = "".join(t.text or "" for t in p.iter(qn("w:t")))
        for m in CHAVES.finditer(texto):
            if normalizar(m.group(1)) not in (REPETIR, FIM_REPETIR):
                campos.setdefault(normalizar(m.group(1)), _rotulo(m.group(1)))
    for r in doc.element.body.iter(qn("w:r")):
        fc = r.find(qn("w:fldChar"))
        tipo = fc.get(qn("w:fldCharType")) if fc is not None else None
        if tipo == "begin":
            instr, dentro = "", True
        elif tipo in ("separate", "end") and dentro:
            nome, rotulo = _nome_campo(instr)
            if nome:
                campos.setdefault(nome, rotulo)
            dentro = False
        elif dentro:
            instr += "".join(t.text or "" for t in r.iter(qn("w:instrText")))
    return list(campos.items())


def _preencher_paragrafo(p, campos: dict) -> set[str]:
    """Substitui os MERGEFIELD do parágrafo. Retorna os campos sem dado."""
    faltando = set()
    runs = list(p.iter(qn("w:r")))
    i = 0
    while i < len(runs):
        fc = runs[i].find(qn("w:fldChar"))
        if fc is None or fc.get(qn("w:fldCharType")) != "begin":
            i += 1
            continue

        # Coleta begin ... end
        grupo, instr, resultado, fase = [runs[i]], "", [], "instr"
        j = i + 1
        while j < len(runs):
            r = runs[j]
            grupo.append(r)
            f = r.find(qn("w:fldChar"))
            tipo = f.get(qn("w:fldCharType")) if f is not None else None
            if tipo == "separate":
                fase = "result"
            elif tipo == "end":
                break
            elif fase == "instr":
                instr += "".join(t.text or "" for t in r.iter(qn("w:instrText")))
            else:
                resultado.append(r)
            j += 1

        nome, rotulo = _nome_campo(instr)
        if not nome:
            i = j + 1
            continue

        valor = campos.get(nome, "")
        falta = not valor and nome not in campos.get("_VAZIOS", ())
        base = resultado[0] if resultado else grupo[0]
        novo = copy.deepcopy(base)
        for filho in list(novo):
            if filho.tag != qn("w:rPr"):
                novo.remove(filho)
        t = novo.makeelement(qn("w:t"), {})
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = f"[{rotulo}]" if falta else valor
        novo.append(t)
        if falta:
            faltando.add(rotulo)
            rpr = novo.find(qn("w:rPr"))
            if rpr is None:
                rpr = novo.makeelement(qn("w:rPr"), {})
                novo.insert(0, rpr)
            rpr.append(rpr.makeelement(qn("w:highlight"), {qn("w:val"): "yellow"}))

        grupo[0].addprevious(novo)
        for r in grupo:
            r.getparent().remove(r)
        i = j + 1
    return faltando


def _preencher_chaves(p, campos: dict) -> set[str]:
    """Substitui os {{CAMPO}} digitados no parágrafo. O Word costuma quebrar o texto em
    vários pedaços (runs), então o campo pode começar num pedaço e terminar em outro: o valor
    fica no pedaço onde o campo começa (com a formatação dele) e o resto do campo é apagado."""
    nos = [t for r in p.iter(qn("w:r")) for t in r.findall(qn("w:t"))]
    texto = "".join(t.text or "" for t in nos)
    achados = list(CHAVES.finditer(texto))
    if not achados:
        return set()
    # Posição de cada caractere: (nó, índice dentro do nó)
    posicoes = [(no, i) for no in nos for i in range(len(no.text or ""))]
    novos = {id(no): list(no.text or "") for no in nos}
    faltando = set()
    for m in reversed(achados):
        nome = normalizar(m.group(1))
        valor = campos.get(nome, "")
        if not valor and nome not in campos.get("_VAZIOS", ()):
            faltando.add(_rotulo(m.group(1)))
            valor = f"[{_rotulo(m.group(1))}]"
        if not valor:  # vazio de propósito leva junto a vírgula ou a barra de antes
            virgula = re.search(r"\s*[,/]\s*$", texto[:m.start()])  # ", {{NÚCLEO}}", "HRB/{{INICIAIS}}"
            for k in range(virgula.start() if virgula else m.start(), m.start()):
                no, i = posicoes[k]
                novos[id(no)][i] = ""
        primeiro, inicio = posicoes[m.start()]
        for k in range(m.start(), m.end()):
            no, i = posicoes[k]
            novos[id(no)][i] = ""
        novos[id(primeiro)][inicio] = valor
    for no in nos:
        no.text = "".join(novos[id(no)])
        no.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return faltando


def _texto(el) -> str:
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def _tem_campo(p) -> bool:
    return bool(CHAVES.search(_texto(p))) or any(
        "MERGEFIELD" in (i.text or "") for i in p.iter(qn("w:instrText")))


def trecho_que_repete(bloco: list) -> tuple[int, int] | None:
    """(primeiro, último) elemento entre {{REPETIR}} e {{FIM REPETIR}}, contando os dos
    marcadores; None se não há trecho marcado. Serve para os elementos do corpo (parágrafos)
    e para as linhas de uma tabela; no corpo, os marcadores de dentro de uma tabela não contam
    (são da linha da tabela que repete)."""
    inicio = fim = None
    for i, el in enumerate(bloco):
        if el.tag == qn("w:tbl"):
            continue
        nomes = {normalizar(m.group(1)) for m in CHAVES.finditer(_texto(el))}
        if REPETIR in nomes and inicio is None:
            inicio = i
        if FIM_REPETIR in nomes and inicio is not None:
            fim = i
            break
    return (inicio, fim) if fim is not None else None


def _so_marcadores(el) -> bool:
    """Parágrafo / linha de tabela que só tem {{REPETIR}} / {{FIM REPETIR}}."""
    texto = _texto(el)
    nomes = {normalizar(m.group(1)) for m in CHAVES.finditer(texto)}
    resto = CHAVES.sub(lambda m: "" if normalizar(m.group(1)) in (REPETIR, FIM_REPETIR) else m.group(0), texto)
    return bool(nomes & {REPETIR, FIM_REPETIR}) and not resto.strip()


def _repetir_linhas(tabela, por_item) -> set[str]:
    """Tabela com uma linha entre {{REPETIR}} e {{FIM REPETIR}} (ex.: a linha de um candidato
    nomeado): a linha sai uma vez para cada pessoa do ato, preenchida com os dados dela.
    por_item() devolve os campos de cada pessoa. Retorna os campos sem dado."""
    linhas = tabela.findall(qn("w:tr"))
    trecho = trecho_que_repete(linhas)
    if trecho is None:
        return set()
    inicio, fim = trecho
    molde = [r for r in linhas[inicio:fim + 1] if not _so_marcadores(r)]
    faltando = set()
    for campos in por_item():
        for original in molde:
            linha = copy.deepcopy(original)
            for x in linha.iter():
                for a in W14_ATTRS:
                    x.attrib.pop(a, None)
            for p in linha.iter(qn("w:p")):
                faltando |= _preencher_paragrafo(p, campos)
                faltando |= _preencher_chaves(p, campos)
            linhas[inicio].addprevious(linha)
    for r in linhas[inicio:fim + 1]:
        tabela.remove(r)
    return faltando


def _preencher_trecho(bloco: list, campos: dict, por_item=None) -> tuple[list, set[str]]:
    """Cópias preenchidas dos elementos: [(elemento, sumiu)]. Um parágrafo que só tinha
    campos e ficou vazio (ex.: um {{CONSIDERANDO}} que não se aplica) some do ato. Numa tabela,
    a linha entre {{REPETIR}} e {{FIM REPETIR}} sai uma vez por pessoa (por_item)."""
    faltando = set()
    saida = []
    for original in bloco:
        el = copy.deepcopy(original)
        for x in el.iter():
            for a in W14_ATTRS:
                x.attrib.pop(a, None)
        tinha = el.tag == qn("w:p") and _tem_campo(el) and _texto(el).strip()
        if por_item is not None:
            for tabela in list(el.iter(qn("w:tbl"))):
                faltando |= _repetir_linhas(tabela, por_item)
        for p in el.iter(qn("w:p")):
            faltando |= _preencher_paragrafo(p, campos)
            faltando |= _preencher_chaves(p, campos)
        saida.append((el, bool(tinha) and not _texto(el).strip()))
    return saida, faltando


def _vazio(el) -> bool:
    return el.tag == qn("w:p") and not _texto(el).strip()


def _preencher(bloco: list, dados: dict, valores=dict) -> tuple[list, set[str]]:
    """Copia o conteúdo do modelo e preenche com um ato. valores(dados) devolve os valores de
    cada campo (as ligações do processo). O trecho entre {{REPETIR}} e {{FIM REPETIR}} sai uma
    vez para cada pessoa do ato (dados["_ITENS"]), com os dados dela e os campos de item
    (ITEM, FIM DO ITEM...). Retorna (elementos, faltando)."""
    marcadores = {REPETIR: "", FIM_REPETIR: ""}

    def campos_de(d: dict) -> dict:
        campos = valores(d) | marcadores
        campos["_VAZIOS"] = set(campos.get("_VAZIOS", ())) | set(marcadores)
        return campos

    pessoas = itens(dados)
    guardados = []

    def por_item() -> list[dict]:
        """Os campos de cada pessoa do ato (com ITEM, FIM DO ITEM...), calculados uma vez."""
        if not guardados:
            guardados.extend(campos_de(dict(p, _ITENS=pessoas) | campos_do_item(n, len(pessoas)))
                             for n, p in enumerate(pessoas, 1))
        return guardados

    topo = campos_de(dados)
    trecho = trecho_que_repete(bloco)
    if trecho is None:
        partes, faltando = _preencher_trecho(bloco, topo, por_item)
    else:
        inicio, fim = trecho
        partes, faltando = _preencher_trecho(bloco[:inicio], topo, por_item)
        for campos in por_item():
            feitas, falta = _preencher_trecho(bloco[inicio:fim + 1], campos)
            partes += feitas
            faltando |= falta
        depois, falta = _preencher_trecho(bloco[fim + 1:], topo, por_item)
        partes += depois
        faltando |= falta

    # Tira os parágrafos que sumiram; se isso deixar duas linhas em branco seguidas (a de
    # antes e a de depois do parágrafo que sumiu), fica uma só
    copia = []
    pular_branco = False
    for el, sumiu in partes:
        if sumiu:
            pular_branco = bool(copia) and _vazio(copia[-1])
            continue
        if pular_branco and _vazio(el):
            pular_branco = False
            continue
        pular_branco = False
        copia.append(el)
    if topo.get(TITULO_NO_ATO):
        _trocar_primeira_linha(copia, topo[TITULO_NO_ATO])
    return copia, faltando


def _trocar_primeira_linha(copia: list, texto: str) -> None:
    """O título do processo no lugar da 1ª linha do Word ("DECRETO S/Nº" → "DECRETO S/Nº -
    EXONERA FULANA"), com a formatação do começo dela."""
    for el in copia:
        if el.tag == qn("w:p") and _texto(el).strip():
            nos = [t for t in el.iter(qn("w:t")) if t.text]
            nos[0].text = texto
            nos[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            for no in nos[1:]:
                no.text = ""
            return


def paragrafos_do_modelo(modelo: Path) -> list[str]:
    """Texto de cada elemento do corpo do Word (parágrafo ou tabela), na ordem: o que o
    assistente mostra para a pessoa escolher o trecho que repete."""
    _, _, bloco = _abrir_modelo(modelo)
    return [_texto(el) for el in bloco]


def marcar_repeticao(modelo: Path, primeiro: int | None, ultimo: int | None = None) -> None:
    """Grava no Word o trecho que repete: um parágrafo {{REPETIR}} antes do elemento
    'primeiro' e um {{FIM REPETIR}} depois do 'ultimo' (índices de paragrafos_do_modelo).
    Os marcadores antigos saem antes. primeiro=None só tira os marcadores."""
    doc = Document(modelo)
    body = doc.element.body
    _tirar_marcadores_das_tabelas(body)
    elementos = [el for el in body if el.tag != qn("w:sectPr")]
    alvos = (elementos[primeiro], elementos[ultimo]) if primeiro is not None else None
    for el in elementos:
        achados = list(CHAVES.finditer(_texto(el)))
        if el.tag != qn("w:p") or not {normalizar(m.group(1)) for m in achados} & {REPETIR, FIM_REPETIR}:
            continue
        sozinho = not CHAVES.sub("", _texto(el)).strip()
        if sozinho and (alvos is None or el not in alvos):
            body.remove(el)
        else:  # marcador no meio do texto: tira só ele (os outros campos ficam como estão)
            manter = {normalizar(m.group(1)): m.group(0) for m in achados}
            manter.update({REPETIR: "", FIM_REPETIR: "", "_VAZIOS": {REPETIR, FIM_REPETIR}})
            _preencher_chaves(el, manter)
    if alvos:
        for texto, vizinho, antes in (("{{REPETIR}}", alvos[0], True), ("{{FIM REPETIR}}", alvos[1], False)):
            p = body.makeelement(qn("w:p"), {})
            ppr = vizinho.find(qn("w:pPr")) if vizinho.tag == qn("w:p") else None
            if ppr is not None:
                p.append(copy.deepcopy(ppr))
            r = p.makeelement(qn("w:r"), {})
            t = r.makeelement(qn("w:t"), {})
            t.text = texto
            r.append(t)
            p.append(r)
            if antes:
                vizinho.addprevious(p)
            else:
                vizinho.addnext(p)
    _guardar_versao(modelo)
    doc.save(modelo)


def _tirar_marcadores_das_tabelas(body) -> None:
    """Tira {{REPETIR}} / {{FIM REPETIR}} de dentro das tabelas (os outros campos ficam)."""
    for tabela in body.iter(qn("w:tbl")):
        for p in tabela.iter(qn("w:p")):
            achados = list(CHAVES.finditer(_texto(p)))
            if {normalizar(m.group(1)) for m in achados} & {REPETIR, FIM_REPETIR}:
                manter = {normalizar(m.group(1)): m.group(0) for m in achados}
                manter.update({REPETIR: "", FIM_REPETIR: "", "_VAZIOS": {REPETIR, FIM_REPETIR}})
                _preencher_chaves(p, manter)


def _escrever_na_ponta(p, texto: str, no_inicio: bool) -> None:
    """Põe texto no começo (ou no fim) do parágrafo, com a formatação do pedaço que já está lá."""
    nos = list(p.iter(qn("w:t")))
    if nos:
        no = nos[0] if no_inicio else nos[-1]
        no.text = texto + (no.text or "") if no_inicio else (no.text or "") + texto
        no.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        return
    r = p.makeelement(qn("w:r"), {})
    t = r.makeelement(qn("w:t"), {})
    t.text = texto
    r.append(t)
    p.append(r)


def linhas_das_tabelas(modelo: Path) -> dict[int, list[list[str]]]:
    """{índice da tabela em paragrafos_do_modelo: [[texto de cada célula] de cada linha]}."""
    _, _, bloco = _abrir_modelo(modelo)
    return {i: [[_texto(c) for c in tr.findall(qn("w:tc"))] for tr in el.findall(qn("w:tr"))]
            for i, el in enumerate(bloco) if el.tag == qn("w:tbl")}


def preparar_linha_pessoa(modelo: Path, indice: int, linha: int) -> dict:
    """A pessoa clicou na linha de uma das pessoas numa tabela (ex.: a do 1º candidato
    nomeado). Deixa o Word pronto para repetir essa linha:
      - as linhas logo abaixo, do mesmo formato (os outros candidatos), saem do Word;
      - a linha fica entre {{REPETIR}} (no começo da 1ª célula) e {{FIM REPETIR}} (no fim
        da última), e sai uma vez para cada pessoa do ato.
    indice é o da tabela em paragrafos_do_modelo; linha, a posição dela na tabela.
    Devolve {"removidos"}."""
    doc = Document(modelo)
    body = doc.element.body
    tabela = [el for el in body if el.tag != qn("w:sectPr")][indice]
    # Um ato só tem um trecho que repete: tira os marcadores antigos (parágrafos e tabelas)
    _tirar_marcadores_das_tabelas(body)
    for el in [el for el in body if el.tag == qn("w:p") and _so_marcadores(el)]:
        body.remove(el)
    linhas = tabela.findall(qn("w:tr"))
    alvo = linhas[linha]
    celulas = len(alvo.findall(qn("w:tc")))
    removidos = 0
    for outra in linhas[linha + 1:]:
        if len(outra.findall(qn("w:tc"))) != celulas:
            break
        tabela.remove(outra)
        removidos += 1
    paragrafos = list(alvo.iter(qn("w:p")))
    _escrever_na_ponta(paragrafos[0], "{{REPETIR}}", True)
    _escrever_na_ponta(paragrafos[-1], "{{FIM REPETIR}}", False)
    _guardar_versao(modelo)
    doc.save(modelo)
    return {"removidos": removidos}


def _guardar_versao(modelo: Path) -> None:
    """Pasta compartilhada: cópia do processo antes de o app mudar o Word dele."""
    try:
        import main
        main.guardar_versao(Path(modelo).parent)
    except Exception:
        pass


def campos_que_repetem(modelo: Path) -> set[str]:
    """Campos (normalizados) do trecho que se repete para cada pessoa: os parágrafos entre
    {{REPETIR}} e {{FIM REPETIR}} ou a linha da tabela. São os dados de cada pessoa; os outros
    campos do Word são iguais para todo o ato."""
    _, _, bloco = _abrir_modelo(modelo)
    trecho = trecho_que_repete(bloco)
    elementos = bloco[trecho[0]:trecho[1] + 1] if trecho else []
    for tabela in (el for el in bloco if el.tag == qn("w:tbl")):
        linhas = tabela.findall(qn("w:tr"))
        if linha := trecho_que_repete(linhas):
            elementos += linhas[linha[0]:linha[1] + 1]
    nomes = set()
    for el in elementos:
        nomes |= {normalizar(m.group(1)) for m in CHAVES.finditer(_texto(el))}
        for instr in el.iter(qn("w:instrText")):
            nome, _ = _nome_campo(instr.text or "")
            if nome:
                nomes.add(nome)
    return nomes - {REPETIR, FIM_REPETIR}


def resumo_da_repeticao(modelo: Path) -> str | None:
    """O que se repete para cada pessoa, para mostrar: o começo do parágrafo, ou as células
    da linha da tabela ("{{ORDEM}} | {{LISTA}} | ... | {{NOME}}"). None se nada repete."""
    _, _, bloco = _abrir_modelo(modelo)

    def sem_marcadores(texto: str) -> str:
        return CHAVES.sub(lambda m: "" if normalizar(m.group(1)) in (REPETIR, FIM_REPETIR) else m.group(0),
                          texto).strip()

    textos = []
    trecho = trecho_que_repete(bloco)
    if trecho:
        textos = [sem_marcadores(_texto(el)) for el in bloco[trecho[0]:trecho[1] + 1]]
    else:
        for tabela in (el for el in bloco if el.tag == qn("w:tbl")):
            linhas = tabela.findall(qn("w:tr"))
            trecho = trecho_que_repete(linhas)
            if trecho:
                textos = [" | ".join(sem_marcadores(_texto(c)) for c in tr.findall(qn("w:tc")))
                          for tr in linhas[trecho[0]:trecho[1] + 1]]
                break
    textos = [t for t in textos if t.strip(" |")]
    if not textos:
        return None
    return (textos[0][:110] + "…") if len(textos[0]) > 110 else textos[0]


# Começo do parágrafo de uma pessoa numa lista: "I –", "II -", "1)", "a)"...
INICIO_ITEM = re.compile(r"^\s*([IVXLCDM]+|\d{1,3}|[a-z])\s*(?:[–—-]|\))\s*")
# Fim do parágrafo de uma pessoa: ";", "; e", "." (o que vira {{FIM DO ITEM}})
FIM_ITEM = re.compile(r"\s*(?:;\s*e|;|\.)\s*$")


def preparar_trecho_pessoa(modelo: Path, indice: int) -> dict:
    """A pessoa clicou no parágrafo de uma das pessoas do ato (ex.: "I – PAULO..., da
    Secretaria de Governo;"). Deixa o Word pronto para repetir esse parágrafo:
      - o "I" do começo vira {{ITEM}} e o ";" / "." do fim vira {{FIM DO ITEM}};
      - os parágrafos seguidos das outras pessoas ("II – LUIZ...") saem do Word;
      - o parágrafo (e a linha em branco depois dele, se as pessoas eram separadas por uma)
        fica entre {{REPETIR}} e {{FIM REPETIR}}.
    indice é o de paragrafos_do_modelo. Devolve {"removidos", "numerou", "fim"}."""
    from modelo_automatico import _trocar  # troca texto sem perder a formatação

    doc = Document(modelo)
    body = doc.element.body
    elementos = [el for el in body if el.tag != qn("w:sectPr")]
    alvo = elementos[indice]
    texto = _texto(alvo)
    trocas = []
    inicio = INICIO_ITEM.match(texto)
    if inicio and "{{" not in inicio.group(1):
        trocas.append((inicio.start(1), inicio.end(1), "{{ITEM}}"))
    fim = FIM_ITEM.search(texto)
    if fim and "FIM DO ITEM" not in texto.upper():
        trocas.append((fim.start(), fim.end(), "{{FIM DO ITEM}}"))
    if trocas:
        _trocar(alvo, trocas)

    # Os parágrafos das outras pessoas logo abaixo, numerados do mesmo jeito (II, III... /
    # 2, 3... / b, c...), saem junto com as linhas em branco entre eles
    def estilo(numero: str) -> str:
        return "numero" if numero.isdigit() else "letra" if numero.islower() else "romano"

    remover, brancos = [], []
    for el in elementos[indice + 1:]:
        if el.tag == qn("w:p") and not _texto(el).strip():
            brancos.append(el)
            continue
        outro = INICIO_ITEM.match(_texto(el)) if el.tag == qn("w:p") else None
        if inicio and outro and estilo(outro.group(1)) == estilo(inicio.group(1)):
            remover += brancos + [el]
            brancos = []
            continue
        break
    separado_por_branco = any(not _texto(r).strip() for r in remover)
    for el in remover:
        body.remove(el)
    _guardar_versao(modelo)
    doc.save(modelo)
    # Trecho: o parágrafo e, se as pessoas eram separadas por linha em branco, uma em branco
    elementos = paragrafos_do_modelo(modelo)
    ultimo = indice
    if separado_por_branco and indice + 1 < len(elementos) and not elementos[indice + 1].strip():
        ultimo = indice + 1
    marcar_repeticao(modelo, indice, ultimo)
    return {"removidos": sum(1 for r in remover if _texto(r).strip()), "numerou": bool(inicio),
            "fim": bool(fim)}


def _abrir_modelo(modelo: Path):
    """Abre o modelo e separa o conteúdo do corpo (sem a seção final)."""
    doc = Document(modelo)
    body = doc.element.body
    sect = body.find(qn("w:sectPr"))
    bloco = [el for el in body if el is not sect]
    for el in bloco:
        body.remove(el)
    return doc, sect, bloco


def blocos(copia: list, sem_titulo: bool) -> list:
    """Parágrafos e tabelas do ato preenchido, na ordem. sem_titulo tira a primeira linha
    com texto, que é o título (o Diário recebe o título num campo próprio), e as linhas em
    branco logo depois dela."""
    saida = []
    for el in copia:
        saida += [el] if el.tag in (qn("w:p"), qn("w:tbl")) else list(el.iter(qn("w:p")))
    if sem_titulo:
        for i, el in enumerate(saida):
            if el.tag == qn("w:p") and _texto(el).strip():
                saida = saida[i + 1:]
                while saida and saida[0].tag == qn("w:p") and not _texto(saida[0]).strip():
                    saida = saida[1:]
                break
    return saida


def _celulas(tabela) -> list[list[tuple]]:
    """As células de cada linha que aparecem: [(célula, colspan, rowspan)]. Células mescladas
    na vertical (vMerge) viram uma só, com rowspan; na horizontal (gridSpan), colspan."""
    grade = []  # por linha: [(coluna da grade, célula, colspan, mescla)]
    for tr in tabela.findall(qn("w:tr")):
        coluna, linha = 0, []
        for tc in tr.findall(qn("w:tc")):
            tcpr = tc.find(qn("w:tcPr"))
            span = tcpr.find(qn("w:gridSpan")) if tcpr is not None else None
            colspan = int(span.get(qn("w:val"))) if span is not None else 1
            vmerge = tcpr.find(qn("w:vMerge")) if tcpr is not None else None
            mescla = None if vmerge is None else (vmerge.get(qn("w:val")) or "continue")
            linha.append((coluna, tc, colspan, mescla))
            coluna += colspan
        grade.append(linha)
    saida = []
    for n, linha in enumerate(grade):
        celulas = []
        for coluna, tc, colspan, mescla in linha:
            if mescla == "continue":
                continue
            rowspan = 1
            if mescla == "restart":
                for abaixo in grade[n + 1:]:
                    if not any(c == coluna and m == "continue" for c, _, _, m in abaixo):
                        break
                    rowspan += 1
            celulas.append((tc, colspan, rowspan))
        saida.append(celulas)
    return saida


def tabela_html(tabela, paragrafo_html) -> str:
    """Tabela do Word em <table> com bordas; paragrafo_html(p) desenha cada parágrafo das
    células (o mesmo usado fora da tabela)."""
    linhas = []
    for celulas in _celulas(tabela):
        tds = []
        for tc, colspan, rowspan in celulas:
            attrs = (f' colspan="{colspan}"' if colspan > 1 else "") + (f' rowspan="{rowspan}"' if rowspan > 1 else "")
            estilo = "border:1px solid #000;padding:1pt 4pt;vertical-align:middle"
            tcpr = tc.find(qn("w:tcPr"))
            fundo = tcpr.find(qn("w:shd")) if tcpr is not None else None
            if fundo is not None and (fundo.get(qn("w:fill")) or "auto").lower() not in ("auto", "ffffff"):
                estilo += f";background-color:#{fundo.get(qn('w:fill'))}"
            conteudo = "".join(paragrafo_html(p) for p in tc.iter(qn("w:p")))
            tds.append(f'<td{attrs} style="{estilo}">{conteudo}</td>')
        linhas.append(f"<tr>{''.join(tds)}</tr>")
    return ('<table style="border-collapse:collapse;width:100%;margin:6pt 0" border="1" cellpadding="2">'
            f"<tbody>{''.join(linhas)}</tbody></table>")


def _texto_bloco(el) -> str:
    if el.tag == qn("w:tbl"):  # uma linha por linha da tabela, células separadas por " | "
        def celula(tc) -> str:
            return " ".join(" ".join(Paragraph(p, None).text.split()) for p in tc.iter(qn("w:p"))).strip()

        return "\n".join(" | ".join(celula(tc) for tc, _, _ in celulas) for celulas in _celulas(el))
    return Paragraph(el, None).text


def textos(modelo: Path, registros: list[dict], sem_titulo: bool = False, valores=dict) -> list[str]:
    """Texto de cada ato preenchido (um por registro), sem gerar arquivo."""
    _, _, bloco = _abrir_modelo(modelo)
    saida = []
    for dados in registros:
        copia, _ = _preencher(bloco, dados, valores)
        saida.append("\n".join(_texto_bloco(el) for el in blocos(copia, sem_titulo)).strip())
    return saida


def _entidades(texto: str) -> str:
    """Escapa o texto como o editor (TinyMCE) dos sites faz: <, >, & e todo caractere fora
    do ASCII viram entidade (ç → &ccedil;, º → &ordm;). O gerador de PDF do Edocs acusa
    "Arquivo ... Corrompido!" com acento cru no HTML enviado direto por requisição."""
    texto = html.escape(texto, quote=False)
    return "".join(c if ord(c) < 128 else
                   f"&{codepoint2name[ord(c)]};" if ord(c) in codepoint2name else f"&#{ord(c)};"
                   for c in texto)


def _html_paragrafo(p, part) -> str:
    """Converte um parágrafo do Word em <p> com alinhamento, recuo, espaçamento, fonte,
    negrito, itálico e sublinhado. Sai em XHTML, como o editor dos sites gera (<br />,
    <strong>, <em>, acentos como entidade), para poder ir direto na requisição."""
    par = Paragraph(p, part)
    pf = par.paragraph_format
    estilo = []
    alinh = {0: "left", 1: "center", 2: "right", 3: "justify"}.get(
        int(pf.alignment) if pf.alignment is not None else None)
    if alinh:
        estilo.append(f"text-align:{alinh}")
    if pf.first_line_indent:
        estilo.append(f"text-indent:{pf.first_line_indent.pt:.1f}pt")
    if pf.left_indent:
        estilo.append(f"margin-left:{pf.left_indent.pt:.1f}pt")
    antes = pf.space_before.pt if pf.space_before else 0
    depois = pf.space_after.pt if pf.space_after else 0
    estilo.append(f"margin-top:{antes:.1f}pt;margin-bottom:{depois:.1f}pt")
    if isinstance(pf.line_spacing, float):
        estilo.append(f"line-height:{pf.line_spacing}")

    partes = []
    for el in p.iter(qn("w:r")):
        run = Run(el, par)
        conteudo = ""
        for filho in el:
            if filho.tag == qn("w:t"):
                conteudo += _entidades(filho.text or "")
            elif filho.tag == qn("w:tab"):
                conteudo += "&emsp;"
            elif filho.tag in (qn("w:br"), qn("w:cr")):
                conteudo += "<br />"
        if not conteudo:
            continue
        css = []
        if run.font.name:
            css.append(f"font-family:{run.font.name}")
        if run.font.size:
            css.append(f"font-size:{run.font.size.pt:g}pt")
        if run.font.all_caps:
            css.append("text-transform:uppercase")
        if css:
            conteudo = f'<span style="{";".join(css)}">{conteudo}</span>'
        if run.underline:
            conteudo = f'<span style="text-decoration:underline">{conteudo}</span>'
        if run.italic:
            conteudo = f"<em>{conteudo}</em>"
        if run.bold:
            conteudo = f"<strong>{conteudo}</strong>"
        partes.append(conteudo)

    corpo = "".join(partes) or "&nbsp;"
    return f'<p style="{";".join(estilo)}">{corpo}</p>'


def salvar_ato(modelo: Path, dados: dict, destino: Path, valores=dict) -> Path:
    """Grava o ato preenchido num Word novo (o modelo com cabeçalho, rodapé e formatação),
    para abrir e imprimir. O modelo não é mexido."""
    doc, sect, bloco = _abrir_modelo(modelo)
    copia, _ = _preencher(bloco, dados, valores)
    body = doc.element.body
    for el in copia:
        if sect is not None:
            sect.addprevious(el)
        else:
            body.append(el)
    destino.parent.mkdir(parents=True, exist_ok=True)
    doc.save(destino)
    return destino


def htmls(modelo: Path, registros: list[dict], sem_titulo: bool = False, valores=dict) -> list[str]:
    """HTML formatado de cada ato (mesma formatação do Word), para colar nos sites."""
    doc, _, bloco = _abrir_modelo(modelo)
    saida = []
    for dados in registros:
        copia, _ = _preencher(bloco, dados, valores)

        def paragrafo(p) -> str:
            return _html_paragrafo(p, doc.part)

        partes = [tabela_html(el, paragrafo) if el.tag == qn("w:tbl") else paragrafo(el)
                  for el in blocos(copia, sem_titulo)]
        saida.append("\n".join(partes))
    return saida
