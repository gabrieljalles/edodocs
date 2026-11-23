"""Grava a coluna "status" na planilha de dados mexendo só nessas células.

Edita o XML do .xlsx diretamente (em vez de salvar pelo openpyxl), porque a planilha tem
milhares de fórmulas e o openpyxl descartaria os valores já calculados delas.
"""
import os
import re
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
N = {"m": NS}

ENVIADO = "Enviado com sucesso"
ERRO = "Houve um erro"
# (cor da fonte, cor do fundo) no padrão "Bom"/"Ruim" do Excel
CORES = {ENVIADO: ("FF006100", "FFC6EFCE"), ERRO: ("FF9C0006", "FFFFC7CE")}
AMARELO = ("FF9C5700", "FFFFEB9C")  # padrão "Neutro" do Excel

# Conferência da publicação no Diário Oficial (coluna status_publicado)
PUBLICADO = "Publicado com sucesso no diário oficial"
NAO_ENCONTRADO = "Não encontrado"
NAO_ENCONTRADO_FINAL = "Não encontrado após 3 tentativas"
ERRO_CONFERENCIA = "Houve um erro na conferência"


def em_uso(caminho: Path) -> str | None:
    """A planilha está aberta no Excel (neste computador ou por outra pessoa na rede)?
    None = livre para gravar. Senão, o nome de quem abriu, lido do arquivo-trava "~$..." que
    o Excel cria ao lado da planilha ("" se não der para saber).

    O que vale é o arquivo travado: o Excel segura a planilha aberta sem deixar ninguém
    gravar (também na rede). Gravar por cima seria perder o que a outra pessoa salvar depois."""
    try:
        with open(caminho, "r+b"):
            return None
    except PermissionError:
        pass
    except OSError:
        return None  # não existe / caminho inválido: quem chamou trata
    for trava in (caminho.with_name("~$" + caminho.name), caminho.with_name("~$" + caminho.name[2:])):
        try:
            dados = trava.read_bytes()
        except OSError:
            continue
        if dados:
            return dados[1:1 + dados[0]].decode("latin-1", "ignore").strip()
    return ""


def eh_erro(texto: str) -> bool:
    return str(texto or "").startswith(ERRO)


def _cores(texto: str):
    texto = str(texto or "")
    if texto.startswith(PUBLICADO):
        return CORES[ENVIADO]
    if texto == NAO_ENCONTRADO_FINAL or eh_erro(texto):
        return CORES[ERRO]
    if texto.startswith(NAO_ENCONTRADO):
        return AMARELO
    return CORES.get(texto, (None, None))


def proxima_tentativa(atual: str) -> str:
    """'' -> 'Não encontrado 1' -> 'Não encontrado 2' -> 'Não encontrado após 3 tentativas'."""
    m = re.fullmatch(rf"{NAO_ENCONTRADO} (\d)", str(atual or "").strip())
    n = int(m.group(1)) + 1 if m else 1
    return NAO_ENCONTRADO_FINAL if n >= 3 else f"{NAO_ENCONTRADO} {n}"


def tentativas_feitas(atual: str) -> int:
    """Quantas conferências já deram 'não encontrado' (0, 1 ou 2)."""
    m = re.fullmatch(rf"{NAO_ENCONTRADO} (\d)", str(atual or "").strip())
    return int(m.group(1)) if m else 0


def conferencia_finalizada(atual: str) -> bool:
    """Publicado ou esgotou as 3 tentativas: não confere mais."""
    atual = str(atual or "")
    return atual.startswith(PUBLICADO) or atual == NAO_ENCONTRADO_FINAL


def _letra(col: int) -> str:
    s = ""
    while col:
        col, r = divmod(col - 1, 26)
        s = chr(65 + r) + s
    return s


def _numero_coluna(ref: str) -> int:
    n = 0
    for c in re.match(r"[A-Z]+", ref).group():
        n = n * 26 + ord(c) - 64
    return n


def colunas(caminho: Path, aba: str, nomes: list[str],
            apelidos: dict[str, list[str]] | None = None) -> dict[str, int]:
    """Número de cada coluna pelo cabeçalho (ou por um nome antigo em 'apelidos'). As que
    não existem ganham as próximas colunas livres depois do cabeçalho, na ordem pedida."""
    existentes, ultima = {}, 0
    for i, v in sorted(cabecalho(caminho, aba).items()):
        if v.strip():
            existentes[v.strip().lower()] = i
            ultima = i
    resultado = {}
    for nome in nomes:
        achados = [n.lower() for n in [nome, *(apelidos or {}).get(nome, [])] if n.lower() in existentes]
        if achados:
            resultado[nome] = existentes[achados[0]]
        else:
            ultima += 1
            resultado[nome] = ultima
    return resultado


def abas(caminho: Path) -> tuple[list[str], str]:
    """(nomes das abas, aba ativa) lendo só o workbook.xml."""
    with zipfile.ZipFile(caminho) as z:
        wb = etree.fromstring(z.read("xl/workbook.xml"))
    nomes = [f.get("name") for f in wb.findall("m:sheets/m:sheet", N)]
    vista = wb.find("m:bookViews/m:workbookView", N)
    ativa = int(vista.get("activeTab", 0)) if vista is not None else 0
    return nomes, nomes[min(ativa, len(nomes) - 1)]


def cabecalho(caminho: Path, aba: str) -> dict[int, str]:
    """{número da coluna: texto} da 1ª linha, lendo só o começo do XML da aba (o openpyxl
    levava ~10 s numa planilha grande só para isso)."""
    return linha_da_planilha(caminho, aba, 1)


def linha_da_planilha(caminho: Path, aba: str, numero: int) -> dict[int, str]:
    """{número da coluna: texto} de uma linha, direto do arquivo (para até chegar nela)."""
    return linhas_da_planilha(caminho, aba, [numero])[numero]


def linhas_da_planilha(caminho: Path, aba: str, numeros: list[int]) -> dict[int, dict[int, str]]:
    """{linha: {número da coluna: texto}} de várias linhas, lendo o arquivo uma vez só (para
    na maior delas). Linha que não existe vem vazia."""
    procuradas = set(numeros)
    with zipfile.ZipFile(caminho) as z:
        celulas: dict[int, dict[int, tuple[str | None, str]]] = {n: {} for n in procuradas}
        with z.open(_caminho_aba(z, aba)) as f:
            for i, (_, row) in enumerate(etree.iterparse(f, tag=f"{{{NS}}}row"), 1):
                atual = int(row.get("r") or i)
                if atual in procuradas:
                    for c in row.findall("m:c", N):
                        if c.get("t") == "inlineStr":
                            texto = "".join(c.itertext())
                        else:
                            v = c.find("m:v", N)
                            texto = v.text or "" if v is not None else ""
                        celulas[atual][_numero_coluna(c.get("r"))] = (c.get("t"), texto)
                if atual >= max(procuradas):
                    break
                row.clear()
        # Textos compartilhados (t="s"): lê só até o maior índice usado nessas linhas
        indices = {int(t) for linha in celulas.values() for tipo, t in linha.values()
                   if tipo == "s" and t.isdigit()}
        compartilhados: dict[int, str] = {}
        if indices and "xl/sharedStrings.xml" in z.namelist():
            with z.open("xl/sharedStrings.xml") as f:
                for i, (_, si) in enumerate(etree.iterparse(f, tag=f"{{{NS}}}si")):
                    if i in indices:
                        compartilhados[i] = "".join(t.text or "" for t in si.iter(f"{{{NS}}}t"))
                    si.clear()
                    if i >= max(indices):
                        break
    return {n: {col: compartilhados.get(int(t), "") if tipo == "s" else t for col, (tipo, t) in linha.items()}
            for n, linha in celulas.items()}


def _caminho_aba(z: zipfile.ZipFile, aba: str) -> str:
    wb = etree.fromstring(z.read("xl/workbook.xml"))
    folhas = wb.findall("m:sheets/m:sheet", N)
    folha = next((f for f in folhas if f.get("name") == aba), folhas[0])
    rid = folha.get(f"{{{NS_R}}}id")
    rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    alvo = next(r.get("Target") for r in rels if r.get("Id") == rid)
    return alvo.lstrip("/") if alvo.startswith("/") else "xl/" + alvo


def _estilo(styles, fonte: str, fundo: str) -> int:
    """Índice de um estilo de célula com essa fonte/fundo (reaproveita se já existir)."""
    fonts = styles.find("m:fonts", N)
    fills = styles.find("m:fills", N)
    xfs = styles.find("m:cellXfs", N)

    def indice(lista, teste):
        for i, el in enumerate(lista):
            if teste(el):
                return i
        return None

    id_fonte = indice(fonts, lambda f: f.find("m:color", N) is not None
                      and f.find("m:color", N).get("rgb") == fonte and f.find("m:b", N) is not None)
    if id_fonte is None:
        f = etree.SubElement(fonts, f"{{{NS}}}font")
        etree.SubElement(f, f"{{{NS}}}b")
        etree.SubElement(f, f"{{{NS}}}color", rgb=fonte)
        etree.SubElement(f, f"{{{NS}}}sz", val="11")
        etree.SubElement(f, f"{{{NS}}}name", val="Calibri")
        id_fonte = len(fonts) - 1
        fonts.set("count", str(len(fonts)))

    id_fundo = indice(fills, lambda f: f.find("m:patternFill/m:fgColor", N) is not None
                      and f.find("m:patternFill/m:fgColor", N).get("rgb") == fundo)
    if id_fundo is None:
        f = etree.SubElement(fills, f"{{{NS}}}fill")
        p = etree.SubElement(f, f"{{{NS}}}patternFill", patternType="solid")
        etree.SubElement(p, f"{{{NS}}}fgColor", rgb=fundo)
        etree.SubElement(p, f"{{{NS}}}bgColor", indexed="64")
        id_fundo = len(fills) - 1
        fills.set("count", str(len(fills)))

    id_xf = indice(xfs, lambda x: x.get("fontId") == str(id_fonte) and x.get("fillId") == str(id_fundo))
    if id_xf is None:
        etree.SubElement(xfs, f"{{{NS}}}xf", numFmtId="0", fontId=str(id_fonte),
                         fillId=str(id_fundo), borderId="0", xfId="0",
                         applyFont="1", applyFill="1")
        id_xf = len(xfs) - 1
        xfs.set("count", str(len(xfs)))
    return id_xf


def _escrever_celula(sheet_data, linha: int, col: int, texto: str, estilo: int | None) -> None:
    ref = f"{_letra(col)}{linha}"
    row = sheet_data.find(f"m:row[@r='{linha}']", N)
    if row is None:
        row = etree.Element(f"{{{NS}}}row", r=str(linha))
        depois = next((r for r in sheet_data if int(r.get("r")) > linha), None)
        if depois is None:
            sheet_data.append(row)
        else:
            depois.addprevious(row)
        row.text = None

    celula = row.find(f"m:c[@r='{ref}']", N)
    if texto == "":
        if celula is not None:
            row.remove(celula)  # apagar: a célula some (valor e cor)
        return
    if celula is None:
        celula = etree.Element(f"{{{NS}}}c", r=ref)
        depois = next((c for c in row if _numero_coluna(c.get("r")) > col), None)
        if depois is None:
            row.append(celula)
        else:
            depois.addprevious(celula)
    for filho in list(celula):
        celula.remove(filho)
    celula.set("t", "inlineStr")
    if estilo is not None:
        celula.set("s", str(estilo))
    elif "s" in celula.attrib:
        del celula.attrib["s"]
    is_ = etree.SubElement(celula, f"{{{NS}}}is")
    etree.SubElement(is_, f"{{{NS}}}t").text = texto

    # A linha pode declarar o intervalo de colunas usado: amplia se preciso
    spans = row.get("spans")
    if spans:
        ini, fim = (int(x) for x in spans.split(":"))
        if col > fim:
            row.set("spans", f"{ini}:{col}")


def gravar(caminho: Path, aba: str, cols: dict[str, int], valores: dict[str, dict[int, str]]) -> None:
    """valores = {nome da coluna: {número da linha no Excel: texto}}; cols vem de colunas().
    Cria os cabeçalhos se não existirem. Lança PermissionError se o arquivo estiver aberto
    (ou se outro computador ficar gravando nele por muito tempo)."""
    import trava
    try:
        with trava.trava(Path(caminho)):
            _gravar(caminho, aba, cols, valores)
    except trava.Ocupado as e:
        raise PermissionError(str(e)) from None


def _gravar(caminho: Path, aba: str, cols: dict[str, int], valores: dict[str, dict[int, str]]) -> None:
    if em_uso(caminho) is not None:
        raise PermissionError("planilha aberta no Excel")
    with zipfile.ZipFile(caminho) as z:
        nome_aba = _caminho_aba(z, aba)
        arquivos = {i.filename: (i, z.read(i.filename)) for i in z.infolist()}

    sheet = etree.fromstring(arquivos[nome_aba][1])
    styles = etree.fromstring(arquivos["xl/styles.xml"][1])
    dados = sheet.find("m:sheetData", N)

    for nome, por_linha in valores.items():
        col = cols[nome]
        _escrever_celula(dados, 1, col, nome, None)
        for linha, texto in por_linha.items():
            fonte, fundo = _cores(texto)
            estilo = _estilo(styles, fonte, fundo) if fonte else None
            _escrever_celula(dados, linha, col, texto, estilo)
    col = max(cols[nome] for nome in valores)

    dim = sheet.find("m:dimension", N)
    if dim is not None:
        m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", dim.get("ref", ""))
        if m and _numero_coluna(m.group(3)) < col:
            dim.set("ref", f"{m.group(1)}{m.group(2)}:{_letra(col)}{m.group(4)}")

    novos = {
        nome_aba: etree.tostring(sheet, xml_declaration=True, encoding="UTF-8", standalone=True),
        "xl/styles.xml": etree.tostring(styles, xml_declaration=True, encoding="UTF-8", standalone=True),
    }
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=caminho.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as saida:
            for nome, (info, conteudo) in arquivos.items():
                saida.writestr(info, novos.get(nome, conteudo))
        os.replace(tmp, caminho)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
