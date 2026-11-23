"""Testes offline das colunas da tabela dos atos (ordem, visíveis e caixinhas).

  1. padrão: número do ato, datas das etapas, campos do Word e caixinhas, todas visíveis;
  2. a ordem e o "ver" salvos valem; coluna nova entra no fim; coluna que sumiu sai;
  3. marcar a caixinha grava SIM numa coluna nova da planilha (sem mexer no resto).

Rode com:  python testes/testar_colunas_tabela.py
"""
import sys
import tempfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import openpyxl  # noqa: E402

import atos  # noqa: E402
import planilha_status  # noqa: E402
import processos  # noqa: E402

TEMP = Path(tempfile.mkdtemp())
CAMPOS = [("NOME", "NOME"), ("CARGO", "CARGO")]


def teste_1_padrao() -> None:
    proc = processos.Processo(TEMP, {"caixas": ["Conferido"]})
    colunas = proc.colunas_tabela(CAMPOS)
    assert [c["nome"] for c in colunas] == ["PORTARIASMA", "DATAEDOCS", "DATADIARIO", "_DATAPUBLICACAO",
                                            "NOME", "CARGO", "CONFERIDO"], colunas
    assert all(c["ver"] for c in colunas)
    assert [c["caixa"] for c in colunas][-1] is True and colunas[-1]["rotulo"] == "Conferido"
    print("  [1] colunas padrão ...................................... OK")


def teste_2_ordem_salva() -> None:
    proc = processos.Processo(TEMP, {"caixas": ["CONFERIDO"], "colunas_tabela": [
        {"nome": "CONFERIDO", "ver": True}, {"nome": "NOME", "ver": True},
        {"nome": "DATAEDOCS", "ver": False}, {"nome": "SUMIU", "ver": True}]})
    colunas = proc.colunas_tabela(CAMPOS)
    assert [c["nome"] for c in colunas][:3] == ["CONFERIDO", "NOME", "DATAEDOCS"]
    assert colunas[2]["ver"] is False
    assert "SUMIU" not in [c["nome"] for c in colunas]
    assert [c["nome"] for c in colunas][-1] == "CARGO"  # não estava na ordem salva: vai pro fim
    print("  [2] ordem e visíveis salvos .............................. OK")


def teste_3_marcar_grava_na_planilha() -> None:
    caminho = TEMP / "dados.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "atos"
    ws.append(["NOME", "DIAS"])
    ws.append(["MARIA", 15])
    ws.append(["JOÃO", 30])
    wb.save(caminho)
    cols = planilha_status.colunas(caminho, "atos", ["CONFERIDO"])
    assert cols == {"CONFERIDO": 3}
    planilha_status.gravar(caminho, "atos", cols, {"CONFERIDO": {3: "SIM"}})
    registros = atos.ler_planilha(caminho, "atos", "NOME")
    assert [r.get("CONFERIDO", "") for r in registros] == ["", "SIM"], registros
    assert registros[0]["DIAS"] == "15"
    # Desmarcar apaga o SIM; a coluna (já com cabeçalho) é reaproveitada
    planilha_status.gravar(caminho, "atos", planilha_status.colunas(caminho, "atos", ["CONFERIDO"]),
                           {"CONFERIDO": {3: ""}})
    assert all(not r.get("CONFERIDO") for r in atos.ler_planilha(caminho, "atos", "NOME"))
    assert atos.normalizar("Sim") in processos.MARCADO and atos.normalizar("x") in processos.MARCADO
    print("  [3] marcar grava SIM na planilha ......................... OK")


if __name__ == "__main__":
    teste_1_padrao()
    teste_2_ordem_salva()
    teste_3_marcar_grava_na_planilha()
    print("\n3 testes passaram.")
