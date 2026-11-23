"""Testes offline do "ato pronto vira modelo" (modelo_automatico.py).

  1. descobre de qual linha da planilha é o ato pronto;
  2. acha os dados da planilha no texto, inclusive número por extenso e palavras que mudam
     com o sexo (à servidora / ao servidor, lotada / lotado);
  3. o modelo criado gera o ato certo para as OUTRAS pessoas da planilha;
  4. número de lei ("Lei Complementar nº 340") vem desmarcado e ano de data não vira campo;
  5. o título sai da primeira linha, com o ano automático.

Rode com:  python testes/testar_modelo_automatico.py
"""
import sys
import tempfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

from docx import Document  # noqa: E402

import assistente  # noqa: E402
import atos  # noqa: E402
import modelo_automatico as ma  # noqa: E402
import processos  # noqa: E402

TEMP = Path(tempfile.mkdtemp())


def ato_pronto_da_maria() -> tuple[Path, Path, list[dict], list[str]]:
    """O Word de exemplo preenchido para a MARIA (sem campo nenhum), como alguém guardaria."""
    modelo, planilha = assistente.gerar_exemplos(TEMP)
    proc = processos.Processo(TEMP, {"modelo": str(modelo), "ligacoes": {
        "DIAS POR EXTENSO": {"tipo": "extenso", "coluna": "DIAS"},
        "AO A SERVIDOR": {"tipo": "genero", "coluna": "SEXO", "feminino": "à servidora", "masculino": "ao servidor"},
        "LOTADO": {"tipo": "genero", "coluna": "SEXO", "feminino": "lotada", "masculino": "lotado"}}})
    registros = atos.ler_planilha(planilha, "atos", "NOME")
    doc, sect, bloco = atos._abrir_modelo(modelo)
    copia, _ = atos._preencher(bloco, registros[0], proc.valores)
    for el in copia:
        sect.addprevious(el)
    pronto = TEMP / "Ato da Maria.docx"
    doc.save(pronto)
    return pronto, planilha, registros, atos.cabecalho_planilha(planilha, "atos")[1]


def teste_1_2_encontra() -> None:
    pronto, _, registros, colunas = ato_pronto_da_maria()
    assert atos.campos_do_modelo(pronto) == []
    achados, linha = ma.encontrar(pronto, registros, colunas)
    assert linha["NOME"] == "MARIA DA SILVA", linha
    por_coluna = {a.coluna: a.valor for a in achados}
    assert por_coluna["NOME"] == "MARIA DA SILVA"
    assert por_coluna["MATRÍCULA"] == "12.345-6"
    assert por_coluna["DIAS (por extenso)"] == "quinze"
    assert por_coluna["Muda com SEXO: à servidora / ao servidor"] == "à servidora"
    assert por_coluna["Muda com SEXO: lotada / lotado"] == "lotada"
    print("  [1] descobre a linha do ato pronto ...................... OK")
    print("  [2] acha dados, por extenso e palavras de gênero ......... OK")


def teste_3_modelo_serve_para_outros() -> None:
    pronto, _, registros, colunas = ato_pronto_da_maria()
    achados, _ = ma.encontrar(pronto, registros, colunas)
    modelo = ma.aplicar(pronto, TEMP / "modelo.docx", achados)
    proc = processos.Processo(TEMP, {"modelo": str(modelo), "ligacoes": ma.ligacoes(achados)})
    joao = proc.textos([registros[1]])[0]
    assert "30 (trinta) dia(s) de licença ao servidor JOÃO SOUZA, matrícula nº 23.456-7" in joao, joao
    assert "lotado na Secretaria Municipal de Obras" in joao, joao
    assert proc.textos([registros[0]])[0] == atos.textos(pronto, [{}])[0]  # a Maria fica igual
    print("  [3] o modelo criado serve para as outras pessoas ........ OK")


def teste_4_referencias_e_datas() -> None:
    doc = Document()
    doc.add_paragraph("ATO Nº 340/2026")
    doc.add_paragraph("Com base na Lei Complementar nº 340, de 23 de janeiro de 2004, "
                      "concede a FULANO DE TAL a partir de 05/08/2026.")
    pronto = TEMP / "referencias.docx"
    doc.save(pronto)
    registros = [{"NUMERO": "340", "ANO": "2026", "NOME": "FULANO DE TAL", "_LINHA": 2}]
    achados, _ = ma.encontrar(pronto, registros, ["NUMERO", "ANO", "NOME"])
    trecho = {(a.valor, a.ocorrencias[0][0]): a.usar for a in achados}
    assert trecho[("340", 0)] is True and trecho[("340", 1)] is False, trecho  # lei vem desmarcada
    assert sum(a.valor == "2026" for a in achados) == 1, achados  # o 2026 da data não conta
    print("  [4] referência de lei desmarcada; ano de data ignorado .. OK")


def teste_5_titulo() -> None:
    doc = Document()
    doc.add_paragraph("ATO ADMINISTRATIVO SMA Nº {{Nº NO DRIVE}}/2026")
    modelo = TEMP / "titulo.docx"
    doc.save(modelo)
    assert ma.titulo_do_modelo(modelo, "2026") == "ATO ADMINISTRATIVO SMA Nº {Nº NO DRIVE}/{ano do ato}"
    print("  [5] título tirado da primeira linha, com ano automático .. OK")


def teste_6_exemplo_do_treino() -> None:
    planilha, ato = assistente.gerar_ato_pronto_exemplo(TEMP / "treino")
    assert atos.campos_do_modelo(ato) == []  # é um ato preenchido, como o de quem treina
    registros = atos.ler_planilha(planilha, "atos", "NOME")
    achados, linha = ma.encontrar(ato, registros, atos.cabecalho_planilha(planilha, "atos")[1])
    assert linha["NOME"] == "MARIA DA SILVA" and all(a.usar for a in achados) and len(achados) >= 8, achados
    print("  [6] o exemplo do treino é reconhecido por inteiro ...... OK")


def main() -> None:
    print("Testes do ato pronto que vira modelo (offline)\n")
    for teste in (teste_1_2_encontra, teste_3_modelo_serve_para_outros, teste_4_referencias_e_datas,
                  teste_5_titulo, teste_6_exemplo_do_treino):
        teste()
    print("\n6 testes passaram.")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
