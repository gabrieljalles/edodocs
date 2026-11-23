"""O HTML do Ato manual vai ao e-docs no jeito do editor do site: sem acento cru (o PDF do
e-docs acusa "arquivo corrompido"), tags vazias fechadas e tabela com border-collapse.

Rode com:  python testes/testar_ato_manual_html.py
"""
import re
import sys
from html.entities import name2codepoint
from pathlib import Path

import lxml.etree as et

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import ato_manual  # noqa: E402

TABELA_EXCEL = ('<div>Relação</div><table style="width: 100%;"><colgroup><col style="width: 50%;">'
                '<col style="width: 50%;"></colgroup><tbody><tr><td colspan="2" rowspan="2">ÓRGÃO<br> Nº 1</td>'
                '</tr><tr></tr><tr><td>MATRÍCULA:</td><td>27810-6</td></tr></tbody></table>')


def _como_xml(html: str) -> None:
    """Lê como o gerador de PDF (XHTML): toda tag fechada."""
    texto = re.sub(r"&(\w+);", lambda m: f"&#{name2codepoint.get(m.group(1), 63)};", html)
    et.fromstring(f"<div>{texto}</div>")


def teste_1_sem_acento_cru() -> None:
    h = ato_manual.html_para_edocs(TABELA_EXCEL)
    assert all(ord(c) < 128 for c in h), h
    assert "&Oacute;RG&Atilde;O" in h and "N&ordm;" in h and "MATR&Iacute;CULA" in h
    print("  [1] acentos como entidade (ç → &ccedil;) ................ OK")


def teste_2_tags_fechadas() -> None:
    h = ato_manual.html_para_edocs(TABELA_EXCEL)
    assert "<br />" in h and "<br>" not in h
    assert h.count("<col ") == 2 and len(re.findall(r"<col [^>]*/>", h)) == 2
    assert "<colgroup>" in h and "</colgroup>" in h
    _como_xml(h)
    print("  [2] <br />, <col /> e XHTML válido ...................... OK")


def teste_3_tabela_como_no_site() -> None:
    h = ato_manual.html_para_edocs(TABELA_EXCEL)
    assert '<table style="border-collapse: collapse; width: 100%;">' in h, h[:120]
    h = ato_manual.html_para_edocs('<table border="1"><tr><td>a</td></tr></table>')
    assert 'style="border-collapse: collapse;"' in h
    print("  [3] tabela com border-collapse .......................... OK")


def teste_4_texto_simples() -> None:
    h = ato_manual.html_para_edocs("<div>Ação <b>já</b></div><div>linha<br>2</div>")
    assert h == "<p>A&ccedil;&atilde;o <b>j&aacute;</b></p><p>linha<br />2</p>", h
    print("  [4] parágrafos e negrito como antes ..................... OK")


if __name__ == "__main__":
    testes = [teste_1_sem_acento_cru, teste_2_tags_fechadas, teste_3_tabela_como_no_site, teste_4_texto_simples]
    for teste in testes:
        teste()
    print(f"\n{len(testes)} testes passaram.")
