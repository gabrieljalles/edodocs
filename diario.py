import re
from datetime import date
from pathlib import Path

import openpyxl
import pypdfium2 as pdfium

# "ATO ADMINISTRATIVO SMA Nº 2599/2026" (tolera N°, Nº, No, espaços e quebras de linha)
PADRAO_ATO = re.compile(
    r"ATO\s+ADMINISTRATIVO\s+SMA\s+N\s*[º°oO.]*\s*([\d.]+)\s*/\s*(\d{4})", re.IGNORECASE
)


def chave(numero: str, ano: str) -> str:
    return f"{int(numero.replace('.', ''))}/{ano}"


PADRAO_EDICAO = re.compile(r"EDI[ÇC][ÃA]O\s*N\s*[º°oO.]*\s*(\d+)", re.IGNORECASE)
PADRAO_DATA = re.compile(r"(\d{1,2})\s+DE\s+([A-ZÇÃÉÊÍÓÔÚ]+)\s+DE\s+(\d{4})", re.IGNORECASE)
MESES = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO",
         "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]


def _texto_pdf(caminho: Path) -> str:
    pdf = pdfium.PdfDocument(caminho)
    try:
        return " ".join(p.get_textpage().get_text_range() for p in pdf)
    finally:
        pdf.close()


def info_edicao(texto: str) -> tuple[str | None, date | None]:
    """(número da edição, data do Diário) lidos do cabeçalho da 1ª página."""
    inicio = texto[:3000]
    m = PADRAO_EDICAO.search(inicio)
    numero = m.group(1) if m else None
    dia = None
    d = PADRAO_DATA.search(inicio)
    if d and d.group(2).upper() in MESES:
        try:
            dia = date(int(d.group(3)), MESES.index(d.group(2).upper()) + 1, int(d.group(1)))
        except ValueError:
            pass
    return numero, dia


def edicao(texto: str, caminho: Path) -> str:
    """Identificação do Diário: 'Nº da edição (data)', ex.: '7450 (01/10/2026)'.
    Sem esses dados no PDF, usa o nome do arquivo."""
    numero, dia = info_edicao(texto)
    if not numero:
        return caminho.stem
    return numero + (f" ({dia:%d/%m/%Y})" if dia else "")


SITE = "https://diario.uberlandia.mg.gov.br"
_ROTULO_EDICAO = re.compile(
    r"Edição nº ([\w-]+)/Ano [IVXLC]+(?:\s|<[^>]+>)*[A-Za-zçáÁ-]+, (\d{1,2}) de (\w+) de (\d{4})")
_LINK_PDF = re.compile(r"/diario-oficial/versao-pdf/(\d+)")


def _baixar(url: str) -> bytes:
    import urllib.request
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read()


def edicoes_no_site() -> list[tuple[date, str, str]]:
    """Edições listadas em 'Todas Edições' (as mais recentes): [(data, número, id do PDF)].
    Cada rótulo 'Edição nº ... <data>' aparece logo antes do link do PDF dela."""
    import html as _html
    pagina = _html.unescape(re.sub(r"\s+", " ", _baixar(SITE + "/diario-oficial").decode("utf8", "ignore")))
    rotulos = list(_ROTULO_EDICAO.finditer(pagina))
    links = list(_LINK_PDF.finditer(pagina))
    edicoes = []
    for i, r in enumerate(rotulos):
        fim = rotulos[i + 1].start() if i + 1 < len(rotulos) else len(pagina)
        link = next((l for l in links if r.end() <= l.start() < fim), None)
        mes = r.group(3).upper()
        if link and mes in MESES:
            dia = date(int(r.group(4)), MESES.index(mes) + 1, int(r.group(2)))
            edicoes.append((dia, r.group(1), link.group(1)))
    return edicoes


def baixar_diarios(datas: set[date], pasta: Path, avisar=lambda msg: None) -> set[date]:
    """Baixa do site os Diários das datas pedidas que ainda não estão na pasta (todas as
    edições do dia, inclusive as extras '-A'). Retorna as datas que não estavam no site."""
    ja_tem = {d for d in datas if any(pasta.glob(f"Diario_Oficial_{d:%Y-%m-%d}*.pdf"))}
    faltam = datas - ja_tem
    if not faltam:
        return set()
    pasta.mkdir(parents=True, exist_ok=True)
    disponiveis = edicoes_no_site()
    for dia, numero, id_pdf in disponiveis:
        if dia in faltam:
            destino = pasta / f"Diario_Oficial_{dia:%Y-%m-%d}_ed{numero}.pdf"
            if not destino.exists():
                destino.write_bytes(_baixar(f"{SITE}/diario-oficial/versao-pdf/{id_pdf}"))
                avisar(f"Baixado: Diário {numero} de {dia:%d/%m/%Y}")
    return faltam - {d for d, _, _ in disponiveis}


def indexar_pdfs(pasta: Path, avisar=lambda msg: None,
                 padrao: re.Pattern = PADRAO_ATO) -> dict[date, list[tuple[str, set[str]]]]:
    """Lê todos os PDFs da pasta e agrupa por data do Diário. Renomeia cada arquivo para
    'Diario_Oficial_AAAA-MM-DD_edNNNN.pdf', para ficar fácil saber de que dia é cada um.
    padrao: como o título do ato aparece no PDF (vem do título do Diário do processo).
    Retorna {data: [(identificação da edição, atos), ...]}."""
    por_data: dict[date, list[tuple[str, set[str]]]] = {}
    for pdf in sorted(pasta.glob("*.pdf")):
        texto = _texto_pdf(pdf)
        numero, dia = info_edicao(texto)
        if dia is None:
            avisar(f"Sem data no cabeçalho, ignorado: {pdf.name}")
            continue
        nome = f"Diario_Oficial_{dia:%Y-%m-%d}" + (f"_ed{numero}" if numero else "") + ".pdf"
        novo = pasta / nome
        if pdf != novo and not novo.exists():
            try:
                pdf.rename(novo)
                pdf = novo
            except OSError:
                pass  # arquivo aberto em outro programa: segue com o nome antigo
        rotulo = edicao(texto, pdf)
        # Edições extras ("7450-A"): o cabeçalho só traz o número; o nome do arquivo tem o sufixo
        m = re.search(r"_ed([\w-]+)\.pdf$", pdf.name)
        if m and numero and m.group(1).startswith(numero):
            rotulo = rotulo.replace(numero, m.group(1), 1)
        atos = {chave(n, a) for n, a in padrao.findall(texto)}
        por_data.setdefault(dia, []).append((rotulo, atos))
    return por_data


def salvar_divergencias(erros: list[str], destino: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "divergencias"
    ws.append(["ds_erro"])
    for e in erros:
        ws.append([e])
    ws.column_dimensions["A"].width = 45
    destino.parent.mkdir(parents=True, exist_ok=True)
    wb.save(destino)
