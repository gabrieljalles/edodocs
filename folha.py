"""Prévia do ato como o PAPEL do Word: a página com as margens do documento, o cabeçalho
(com a logo) e o rodapé, e cada parágrafo com o recuo, o alinhamento, o espaçamento e a
fonte do Word. Usada em toda tela que mostra um ato (tabela, envio, assistente).

O texto puro (atos.textos) jogava tudo grudado na esquerda e a pessoa achava que o ato
estava errado. Aqui a formatação vem do parágrafo e, se ele não tiver, do estilo dele
(e dos estilos de que este herda), como o Word faz.
"""
import base64
import html
from pathlib import Path

from docx.oxml.ns import qn
from docx.shared import Length
from docx.text.paragraph import Paragraph
from docx.text.run import Run

import atos

EMU_POR_PT = 12700
ALINHAMENTOS = {0: "left", 1: "center", 2: "right", 3: "justify"}


def _do_estilo(par: Paragraph, atributo: str):
    """Formato do parágrafo: o direto; senão o do estilo (e dos estilos de que ele herda)."""
    valor = getattr(par.paragraph_format, atributo)
    estilo = par.style
    while valor is None and estilo is not None:
        valor = getattr(estilo.paragraph_format, atributo)
        estilo = estilo.base_style
    return valor


def _fonte_padrao(doc) -> tuple[str, float]:
    """Fonte e tamanho que o documento usa quando o texto não diz nada (Normal ou padrões)."""
    normal = doc.styles["Normal"].font
    nome, tamanho = normal.name, normal.size.pt if normal.size else None
    padroes = doc.styles.element.find(qn("w:docDefaults"))
    if padroes is not None:
        rpr = padroes.find(".//" + qn("w:rPr"))
        if rpr is not None:
            fontes, sz = rpr.find(qn("w:rFonts")), rpr.find(qn("w:sz"))
            if not nome and fontes is not None:
                nome = fontes.get(qn("w:ascii")) or fontes.get(qn("w:hAnsi"))
            if not tamanho and sz is not None:
                tamanho = int(sz.get(qn("w:val"))) / 2
    return nome or "Calibri", tamanho or 11


def _imagem(desenho, part) -> str:
    """<img> de uma figura do Word (logo do cabeçalho, imagem no texto), embutida."""
    blip = desenho.find(".//" + qn("a:blip"))
    if blip is None:
        return ""
    parte = part.related_parts.get(blip.get(qn("r:embed")))
    if parte is None:
        return ""
    tipo = getattr(parte, "content_type", "image/png")
    dados = base64.b64encode(parte.blob).decode()
    tamanho = desenho.find(".//" + qn("wp:extent"))
    estilo = ""
    if tamanho is not None:
        largura, altura = int(tamanho.get("cx")) / EMU_POR_PT, int(tamanho.get("cy")) / EMU_POR_PT
        estilo = f' style="width:{largura:.1f}pt;height:{altura:.1f}pt"'
    return f'<img src="data:{tipo};base64,{dados}"{estilo} alt="">'


def _paragrafo(p, part) -> str:
    """Um parágrafo do Word em <p>, com recuos, alinhamento, espaçamento e fonte."""
    par = Paragraph(p, part)
    estilo = []
    alinhamento = _do_estilo(par, "alignment")
    if alinhamento is not None:
        estilo.append(f"text-align:{ALINHAMENTOS.get(int(alinhamento), 'left')}")
    for atributo, css in (("first_line_indent", "text-indent"), ("left_indent", "margin-left"),
                          ("right_indent", "margin-right"), ("space_before", "margin-top"),
                          ("space_after", "margin-bottom")):
        valor = _do_estilo(par, atributo)
        if valor:
            estilo.append(f"{css}:{Length(valor).pt:.1f}pt")
    entrelinha = _do_estilo(par, "line_spacing")
    if isinstance(entrelinha, float):
        estilo.append(f"line-height:{entrelinha * 1.15:.2f}")  # "simples" no Word ≈ 1.15
    elif entrelinha is not None:
        estilo.append(f"line-height:{Length(entrelinha).pt:.1f}pt")

    partes = []
    for el in p.iter(qn("w:r")):
        run = Run(el, par)
        conteudo = ""
        for filho in el:
            if filho.tag == qn("w:t"):
                conteudo += html.escape(filho.text or "")
            elif filho.tag == qn("w:tab"):
                conteudo += '<span class="tab"></span>'
            elif filho.tag in (qn("w:br"), qn("w:cr")):
                conteudo += "<br>"
            elif filho.tag in (qn("w:drawing"), qn("w:pict")):
                conteudo += _imagem(filho, part)
        if not conteudo:
            continue
        css = []
        if run.font.name:
            css.append(f"font-family:'{run.font.name}'")
        if run.font.size:
            css.append(f"font-size:{run.font.size.pt:g}pt")
        if run.bold:
            css.append("font-weight:bold")
        if run.italic:
            css.append("font-style:italic")
        if run.underline:
            css.append("text-decoration:underline")
        if run.font.all_caps:
            css.append("text-transform:uppercase")
        partes.append(f'<span style="{";".join(css)}">{conteudo}</span>' if css else conteudo)
    return f'<p style="{";".join(estilo)}">{"".join(partes) or "&nbsp;"}</p>'


def _bloco(elemento, part) -> str:
    """Parágrafos de um pedaço do documento (corpo, cabeçalho, rodapé; tabelas por linha)."""
    return "".join(_paragrafo(p, part) for p in elemento.iter(qn("w:p")))


def folhas(modelo: Path, registros: list[dict], valores=dict, papel: bool = True,
           sem_titulo: bool = False) -> list[str]:
    """HTML da folha de cada ato preenchido (um por registro). papel=False: só o texto (sem
    página, cabeçalho e rodapé), como fica no editor dos sites; sem_titulo tira a 1ª linha
    (o Diário recebe o título num campo próprio)."""
    doc, _, bloco = atos._abrir_modelo(modelo)
    secao = doc.sections[0]
    fonte, tamanho = _fonte_padrao(doc)

    def margem(valor, padrao: float) -> float:
        return valor.pt if valor is not None else padrao

    largura = margem(secao.page_width, 595)
    cima, baixo = margem(secao.top_margin, 72), margem(secao.bottom_margin, 72)
    esquerda, direita = margem(secao.left_margin, 72), margem(secao.right_margin, 72)
    dist_cabecalho = margem(secao.header_distance, 36)
    cabecalho = _bloco(secao.header._element, secao.header.part) if secao.header is not None else ""
    rodape = _bloco(secao.footer._element, secao.footer.part) if secao.footer is not None else ""

    saida = []
    for dados in registros:
        copia, _ = atos._preencher(bloco, dados, valores)
        corpo = "".join(atos.tabela_html(el, lambda p: _paragrafo(p, doc.part)) if el.tag == qn("w:tbl")
                        else _paragrafo(el, doc.part) for el in atos.blocos(copia, sem_titulo))
        if not papel:
            saida.append(f'<div class="folha-texto" style="font-family:\'{fonte}\';font-size:{tamanho:g}pt">'
                         f'{corpo}</div>')
            continue
        saida.append(
            f'<div class="folha" style="width:{largura:.0f}pt;font-family:\'{fonte}\';'
            f'font-size:{tamanho:g}pt;padding:{dist_cabecalho:.0f}pt {direita:.0f}pt {baixo / 2:.0f}pt '
            f'{esquerda:.0f}pt">'
            f'<div class="folha-cabecalho">{cabecalho}</div>'
            f'<div class="folha-corpo" style="margin-top:{max(cima - dist_cabecalho - 40, 12):.0f}pt">'
            f'{corpo}</div>'
            f'<div class="folha-rodape">{rodape}</div></div>')
    return saida


# Aparência do papel (cada tela inclui com gui.add_css, uma vez)
CSS = """
.folha-area { background: #e4e7eb; border-radius: 8px; padding: 14px; overflow: auto; }
body.body--dark .folha-area { background: #0f1a2c; }
.folha { background: #fff; color: #000; margin: 0 auto; box-shadow: 0 2px 10px rgba(0,0,0,.25);
         zoom: var(--zoom-folha, .8); line-height: 1.15; box-sizing: border-box; }
.folha p { margin: 0; }
.folha .tab { display: inline-block; width: 36pt; }
.folha img { max-width: 100%; }
.folha-rodape { margin-top: 24pt; }
/* Só o texto, como no editor dos sites (e-docs, Diário) */
.folha-texto { background: #fff; color: #000; padding: 14px 18px; line-height: 1.15; zoom: .85; }
.folha-texto p { margin: 0; }
.folha-texto .tab { display: inline-block; width: 36pt; }
"""
