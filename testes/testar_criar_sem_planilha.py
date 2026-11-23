"""Testes offline de criar um processo só com o Word (sem planilha).

  1. sugestões do que parece dado: nome em maiúsculas, matrícula, data por extenso;
  2. marcar campos (texto, data, opções) gera o modelo, as ligações e a linha de exemplo;
  3. campo "Escolha entre opções": SIM escreve o texto, NÃO não escreve (e o parágrafo some);
  4. a planilha criada tem as colunas, a lista SIM/NÃO e o formato de data;
  5. ato de várias pessoas: clicar no parágrafo de uma pessoa prepara a repetição
     (I → {{ITEM}}, ; → {{FIM DO ITEM}}, os outros parágrafos saem);
  6. do começo ao fim: Word marcado + planilha criada + duas pessoas = o decreto montado.

Rode com:  python testes/testar_criar_sem_planilha.py
"""
import sys
import tempfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import openpyxl  # noqa: E402
from docx import Document  # noqa: E402

import atos  # noqa: E402
import marcar_ui  # noqa: E402
import planilha_nova  # noqa: E402
import processos  # noqa: E402

TEMP = Path(tempfile.mkdtemp())
TEXTO_UMA = [
    "DECRETO S/Nº",
    "O PREFEITO DE UBERLÂNDIA, no uso de suas atribuições legais,",
    "Considerando o Formulário de Desligamento,",
    "DECRETA:",
    "Art. 1º Fica exonerado a pedido, ANTONIO EXEMPLO DOS SANTOS, matrícula nº 31.147-2, "
    "retroativo a 3 de agosto de 2026.",
]
TEXTO_VARIAS = [
    "DECRETO S/Nº",
    "Art. 1º Ficam exonerados os servidores abaixo mencionados, sendo:",
    "",
    "I – PEDRO MODELO COSTA MARTINS, matrícula nº 38.830-0, da Secretaria Municipal de Governo;",
    "",
    "II – LUIZ CLAUDIO DE ANDRADE FREIRE, matrícula nº 38.786-0, da Secretaria de Desenvolvimento Social.",
    "",
    "Art. 2º Este Decreto entra em vigor na data de sua publicação.",
]


def word(nome: str, paragrafos: list[str]) -> Path:
    doc = Document()
    for texto in paragrafos:
        doc.add_paragraph(texto)
    caminho = TEMP / nome
    doc.save(caminho)
    return caminho


def marca(paragrafos, texto, campo, tipo="texto", opcoes=None):
    ocorrencias = marcar_ui._ocorrencias(paragrafos, texto)
    assert ocorrencias, texto
    return marcar_ui.Marca(campo, tipo, texto, ocorrencias, opcoes or [])


def teste_1_sugestoes() -> None:
    achados = {(TEXTO_UMA[n][i:f], nome) for n, i, f, nome in marcar_ui.sugestoes(TEXTO_UMA)}
    assert ("ANTONIO EXEMPLO DOS SANTOS", "NOME") in achados, achados
    assert ("31.147-2", "MATRÍCULA") in achados
    assert ("3 de agosto de 2026", "DATA") in achados
    assert not any(t.startswith("PREFEITO") or t.startswith("O PREFEITO") for t, _ in achados), achados
    assert marcar_ui.data_do_extenso("1º de setembro de 2026") == "01/09/2026"
    print("  [1] sugestões do que parece dado ......................... OK")


def criar_modelo_uma() -> tuple[Path, dict, dict]:
    origem = word("uma.docx", TEXTO_UMA)
    marcas = [
        marca(TEXTO_UMA, "ANTONIO EXEMPLO DOS SANTOS", "NOME"),
        marca(TEXTO_UMA, "31.147-2", "MATRÍCULA"),
        marca(TEXTO_UMA, "3 de agosto de 2026", "DATA", "data"),
        marca(TEXTO_UMA, "a pedido, ", "A PEDIDO", "opcoes",
              [{"valor": "SIM", "texto": "a pedido, "}, {"valor": "NÃO", "texto": ""}]),
        marca(TEXTO_UMA, "Considerando o Formulário de Desligamento,", "CONSIDERANDO", "opcoes",
              [{"valor": "SIM", "texto": "Considerando o Formulário de Desligamento,"},
               {"valor": "NÃO", "texto": ""}]),
    ]
    destino = TEMP / "modelo_uma.docx"
    ligacoes, exemplo = marcar_ui.criar_modelo(origem, destino, marcas)
    return destino, ligacoes, exemplo


def teste_2_marcar_campos() -> None:
    modelo, ligacoes, exemplo = criar_modelo_uma()
    assert [r for _, r in atos.campos_do_modelo(modelo)] == ["CONSIDERANDO", "A PEDIDO", "NOME", "MATRÍCULA", "DATA"]
    assert ligacoes["DATA"] == {"tipo": "data_extenso", "coluna": "DATA"}
    assert ligacoes["A PEDIDO"]["tipo"] == "opcoes"
    assert exemplo["NOME"] == "ANTONIO EXEMPLO DOS SANTOS" and exemplo["DATA"] == "03/08/2026"
    assert exemplo["APEDIDO"] == "SIM"
    proc = processos.Processo(TEMP, {"modelo": str(modelo), "ligacoes": ligacoes})
    texto = proc.textos([dict(exemplo, _LINHA=2)])[0]
    assert texto == "\n".join(TEXTO_UMA), texto  # o exemplo monta o ato igual ao original
    print("  [2] marcar campos gera o modelo e o exemplo .............. OK")


def teste_3_opcoes() -> None:
    modelo, ligacoes, exemplo = criar_modelo_uma()
    proc = processos.Processo(TEMP, {"modelo": str(modelo), "ligacoes": ligacoes})
    nao = dict(exemplo, APEDIDO="NÃO", CONSIDERANDO="Não", _LINHA=3)
    texto = proc.textos([nao])[0]
    assert "Fica exonerado ANTONIO" in texto and "Considerando" not in texto, texto
    assert "[" not in texto
    assert processos.valor_da_regra(ligacoes["A PEDIDO"], {"APEDIDO": "sim"}) == "a pedido, "
    print("  [3] escolha entre opções ................................. OK")


def teste_4_planilha_criada() -> None:
    modelo, ligacoes, _ = criar_modelo_uma()
    proc = processos.Processo(TEMP, {"modelo": str(modelo), "ligacoes": ligacoes, "coluna_numero": "Nº DO ATO",
                                     "caixas": ["CONFERIDO"]})
    colunas = proc.colunas_planilha_nova()
    assert [c["nome"] for c in colunas] == ["Nº DO ATO", "CONSIDERANDO", "A PEDIDO", "NOME", "MATRÍCULA",
                                            "DATA", "CONFERIDO"], colunas
    caminho = planilha_nova.criar(TEMP / "nova.xlsx", colunas)
    wb = openpyxl.load_workbook(caminho)
    ws = wb[planilha_nova.ABA]
    assert [c.value for c in ws[1]] == [c["nome"] for c in colunas]
    listas = {str(v.sqref).split(":")[0]: v.formula1 for v in ws.data_validations.dataValidation}
    assert listas["C2"] == '"SIM,NÃO"', listas
    assert ws["F2"].number_format == "dd/mm/yyyy"
    assert atos.ler_planilha(caminho, planilha_nova.ABA, "NOME") == []  # vazia, pronta para preencher
    print("  [4] planilha criada com listas e datas ................... OK")


def teste_5_paragrafo_de_cada_pessoa() -> None:
    modelo = word("varias.docx", TEXTO_VARIAS)
    indice = atos.paragrafos_do_modelo(modelo).index(TEXTO_VARIAS[3])
    feito = atos.preparar_trecho_pessoa(modelo, indice)
    assert feito == {"removidos": 1, "numerou": True, "fim": True}, feito
    assert atos.paragrafos_do_modelo(modelo) == [
        "DECRETO S/Nº",
        "Art. 1º Ficam exonerados os servidores abaixo mencionados, sendo:",
        "",
        "{{REPETIR}}",
        "{{ITEM}} – PEDRO MODELO COSTA MARTINS, matrícula nº 38.830-0, da Secretaria Municipal de Governo"
        "{{FIM DO ITEM}}",
        "",
        "{{FIM REPETIR}}",
        "Art. 2º Este Decreto entra em vigor na data de sua publicação.",
    ], atos.paragrafos_do_modelo(modelo)
    print("  [5] parágrafo de cada pessoa preparado ................... OK")


def teste_6_do_comeco_ao_fim() -> None:
    origem = word("varias2.docx", TEXTO_VARIAS)
    paragrafos = TEXTO_VARIAS
    marcas = [marca(paragrafos, "PEDRO MODELO COSTA MARTINS", "NOME"),
              marca(paragrafos, "38.830-0", "MATRÍCULA"),
              marca(paragrafos, "Secretaria Municipal de Governo", "SECRETARIA")]
    modelo = TEMP / "modelo_varias.docx"
    ligacoes, exemplo = marcar_ui.criar_modelo(origem, modelo, marcas)
    atos.preparar_trecho_pessoa(modelo, atos.paragrafos_do_modelo(modelo).index(
        "I – {{NOME}}, matrícula nº {{MATRÍCULA}}, da {{SECRETARIA}};"))
    config = {"modelo": str(modelo), "ligacoes": ligacoes, "coluna_numero": "Nº DO ATO",
              "juntar": {"coluna": "Nº DO ATO"}}
    proc = processos.Processo(TEMP, config)
    colunas = proc.colunas_planilha_nova()
    assert [c["nome"] for c in colunas] == ["Nº DO ATO", "NOME", "MATRÍCULA", "SECRETARIA"]
    planilha = planilha_nova.criar(TEMP / "varias.xlsx", colunas)
    wb = openpyxl.load_workbook(planilha)
    ws = wb[planilha_nova.ABA]
    ws.append(["1", "ANA", "1.111-1", "Secretaria de Saúde"])
    ws.append(["1", "BIA", "2.222-2", "Secretaria de Obras"])
    wb.save(planilha)
    atos_ = proc.agrupar(atos.ler_planilha(planilha, planilha_nova.ABA, "NOME"))
    texto = proc.textos(atos_)[0]
    assert "I – ANA, matrícula nº 1.111-1, da Secretaria de Saúde;\n\nII – BIA, matrícula nº 2.222-2, " \
           "da Secretaria de Obras." in texto, texto
    print("  [6] do Word marcado ao decreto montado ................... OK")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    testes = [teste_1_sugestoes, teste_2_marcar_campos, teste_3_opcoes, teste_4_planilha_criada,
              teste_5_paragrafo_de_cada_pessoa, teste_6_do_comeco_ao_fim]
    for teste in testes:
        teste()
    print(f"\n{len(testes)} testes passaram.")
