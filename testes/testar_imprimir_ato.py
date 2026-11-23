"""Abrir o ato no Word para imprimir (main.word_do_ato): o ato de uma pessoa e o de várias
pessoas (sempre o ato inteiro, qualquer que seja a linha escolhida), com o cabeçalho e a
formatação do modelo, sem mexer no modelo.

    python testes/testar_imprimir_ato.py
"""
import shutil
import sys
import tempfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

from docx import Document  # noqa: E402

import main  # noqa: E402
import processos  # noqa: E402

passou = 0


def ok(texto: str) -> None:
    global passou
    passou += 1
    print(f"  [{passou}] {texto} " + "." * max(2, 55 - len(texto)) + " OK")


def textos(caminho: Path) -> list[str]:
    return [p.text for p in Document(caminho).paragraphs if p.text.strip()]


TMP = Path(tempfile.mkdtemp(prefix="edodocs_imprimir_"))
original = main.PASTA_IMPRESSAO
main.PASTA_IMPRESSAO = TMP / "impressao"
try:
    pasta = TMP / "Estabilidade"
    pasta.mkdir()
    doc = Document()
    doc.sections[0].header.paragraphs[0].text = "PREFEITURA DE UBERLÂNDIA"
    doc.add_paragraph("DECRETO S/Nº")
    p = doc.add_paragraph()
    p.add_run("{{NOME}}").bold = True
    p.add_run(", matrícula nº {{MATRÍCULA}}, fica {{DECLARADO}} estável.")
    doc.save(pasta / "uma.docx")

    doc = Document()
    doc.sections[0].header.paragraphs[0].text = "PREFEITURA DE UBERLÂNDIA"
    for texto in ("DECRETO S/Nº", "Ficam declarados estáveis:", "{{REPETIR}}",
                  "{{ITEM}} – {{NOME}}, matrícula nº {{MATRÍCULA}}{{FIM DO ITEM}}", "{{FIM REPETIR}}", "Publique-se."):
        doc.add_paragraph(texto)
    doc.save(pasta / "varias.docx")
    antes = (pasta / "uma.docx").read_bytes()

    proc = processos.Processo(pasta, {
        "modelo": "uma.docx", "coluna_nome": "NOME", "coluna_numero": "Nº DO DECRETO",
        "juntar": {"coluna": "Nº DO DECRETO"},
        "outros_modelos": [{"modelo": "varias.docx",
                            "se": [{"campo": "QUANTIDADE DE PESSOAS", "op": "maior", "valor": "1"}]}],
        "ligacoes": {"DECLARADO": {"tipo": "genero", "coluna": "SEXO", "masculino": "declarado",
                                   "feminino": "declarada"}}})
    linhas = [
        {"NODODECRETO": "1", "NOME": "MARIA SOUZA", "MATRICULA": "2.222-2", "SEXO": "F", "_LINHA": 2},
        {"NODODECRETO": "2", "NOME": "ANA LIMA", "MATRICULA": "1.111-1", "SEXO": "F", "_LINHA": 3},
        {"NODODECRETO": "2", "NOME": "BETO REIS", "MATRICULA": "3.333-3", "SEXO": "M", "_LINHA": 4},
        {"NODODECRETO": "2", "NOME": "CAIO DIAS", "MATRICULA": "4.444-4", "SEXO": "M", "_LINHA": 5},
    ]

    # Uma pessoa
    sozinho = main.word_do_ato(proc.ato_da_linha(linhas, 2), proc)
    assert textos(sozinho) == ["DECRETO S/Nº", "MARIA SOUZA, matrícula nº 2.222-2, fica declarada estável."], textos(sozinho)
    word = Document(sozinho)
    assert word.paragraphs[1].runs[0].bold and word.sections[0].header.paragraphs[0].text == "PREFEITURA DE UBERLÂNDIA"
    assert (pasta / "uma.docx").read_bytes() == antes
    ok("ato de uma pessoa, com cabeçalho e negrito")

    # Várias pessoas: o ato inteiro, igual qualquer que seja a linha
    esperado = ["DECRETO S/Nº", "Ficam declarados estáveis:", "I – ANA LIMA, matrícula nº 1.111-1;",
                "II – BETO REIS, matrícula nº 3.333-3;", "III – CAIO DIAS, matrícula nº 4.444-4.", "Publique-se."]
    for linha in (3, 4, 5):
        caminho = main.word_do_ato(proc.ato_da_linha(linhas, linha), proc)
        assert textos(caminho) == esperado, (linha, textos(caminho))
        assert caminho.name == "Estabilidade - 2 · ANA LIMA e mais 2.docx", caminho.name
    ok("várias pessoas: o ato inteiro, de qualquer linha")

    # Word aberto (arquivo travado): grava com outro nome
    travado = main.PASTA_IMPRESSAO / "Estabilidade - 1 · MARIA SOUZA.docx"
    trava = open(travado, "rb")
    try:
        import os
        if os.name == "nt":
            outro = main.word_do_ato(proc.ato_da_linha(linhas, 2), proc)
            assert outro.name == "Estabilidade - 1 · MARIA SOUZA (2).docx", outro.name
    finally:
        trava.close()
    ok("com o Word aberto, grava com outro nome")
finally:
    main.PASTA_IMPRESSAO = original
    shutil.rmtree(TMP, ignore_errors=True)

print(f"\n{passou} testes passaram.")
