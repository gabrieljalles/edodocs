"""Testes offline da edição de linhas da planilha (não precisam do Excel).

  1. o formato da planilha é descoberto: colunas digitadas (texto/número/data),
     calculadas (fórmula) e de controle (status_...);
  2. a fórmula copiada para a linha nova tem as referências ajustadas (A3 → A4, $A$1 fica);
  3. incluir linha: vai logo depois da última com dados (ignorando #N/A colado), copia a
     fórmula e grava número como número e data como data;
  4. editar linha: troca só os valores digitados, mantém fórmula e formato;
  5. excluir sem o Excel: limpa os valores digitados e os de controle, mantém a fórmula;
  6. incluir várias linhas de uma vez: uma depois da outra, cada uma com a fórmula ajustada.

Rode com:  python testes/testar_planilha_edicao.py
"""
import sys
import tempfile
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import openpyxl  # noqa: E402

import planilha_edicao as edicao  # noqa: E402


def planilha_exemplo() -> Path:
    caminho = Path(tempfile.mkdtemp()) / "teste.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "processosPendentes"
    ws.append(["MATRIC", "DIAS", "INICIO", "NOME", "status_edocs"])
    ws.append([111, 15, datetime(2026, 7, 6), '=IF(A2<>"","SERV "&A2&" "&$B$1,"")', "Enviado com sucesso"])
    ws.append([222, 30, datetime(2026, 8, 1), '=IF(A3<>"","SERV "&A3&" "&$B$1,"")', ""])
    ws["E10"] = "#N/A"  # resto de erro colado como valor, longe dos dados
    for linha in (2, 3):
        ws.cell(linha, 3).number_format = "dd/mm/yyyy"
    wb.save(caminho)
    return caminho


def ler(caminho: Path, formulas: bool = False):
    wb = openpyxl.load_workbook(caminho, data_only=not formulas)
    return wb["processosPendentes"]


def teste_1_formato() -> None:
    colunas = edicao.esquema(planilha_exemplo(), None)
    resumo = {c.nome: (c.tipo, c.formula, c.controle) for c in colunas}
    assert resumo["MATRIC"] == ("numero", False, False), resumo
    assert resumo["INICIO"] == ("data", False, False), resumo
    assert resumo["NOME"][1] is True, resumo
    assert resumo["status_edocs"][2] is True, resumo
    assert [c.nome for c in colunas if c.editavel] == ["MATRIC", "DIAS", "INICIO"]
    print("  [1] formato da planilha descoberto ................... OK")


def teste_2_deslocar_formula() -> None:
    f = edicao.deslocar_formula
    assert f('IF(A3<>"","SERV "&A3&" "&$B$1,"")', 1) == 'IF(A4<>"","SERV "&A4&" "&$B$1,"")'
    assert f("VLOOKUP(A2,'[1]geral REL 14'!$A:$AI,2,0)", 5) == "VLOOKUP(A7,'[1]geral REL 14'!$A:$AI,2,0)"
    assert f("SUM($A2:A$9)+LOG10(B2)", 2) == "SUM($A4:A$9)+LOG10(B4)"
    assert f('TEXT(F2,"dd/mm/aaaa")', 1) == 'TEXT(F3,"dd/mm/aaaa")'  # texto entre aspas intacto
    print("  [2] fórmula ajustada para a linha nova ................ OK")


def teste_3_incluir() -> None:
    caminho = planilha_exemplo()
    colunas = edicao.esquema(caminho, None)
    nova = edicao.incluir_linha(caminho, None, colunas, {1: "333", 2: "20", 3: "01/09/2026"})
    assert nova == 4, nova
    ws = ler(caminho, formulas=True)
    assert ws["A4"].value == 333 and ws["B4"].value == 20, (ws["A4"].value, ws["B4"].value)
    assert ws["C4"].value == datetime(2026, 9, 1), ws["C4"].value
    assert ws["C4"].number_format == "dd/mm/yyyy", ws["C4"].number_format
    assert ws["D4"].value == '=IF(A4<>"","SERV "&A4&" "&$B$1,"")', ws["D4"].value
    print("  [3] incluir: linha certa, fórmula e tipos ............ OK")


def teste_4_editar() -> None:
    caminho = planilha_exemplo()
    colunas = edicao.esquema(caminho, None)
    edicao.editar_linha(caminho, None, 3, colunas, {1: "999", 2: "", 3: "15/08/2026"})
    ws = ler(caminho, formulas=True)
    assert ws["A3"].value == 999 and ws["B3"].value is None, (ws["A3"].value, ws["B3"].value)
    assert ws["C3"].value == datetime(2026, 8, 15), ws["C3"].value
    assert ws["D3"].value.startswith("=IF(A3"), ws["D3"].value
    assert ler(caminho, formulas=True)["A2"].value == 111  # outra linha intacta
    print("  [4] editar: só os valores digitados mudam ............ OK")


def teste_5_limpar() -> None:
    caminho = planilha_exemplo()
    colunas = edicao.esquema(caminho, None)
    edicao.limpar_linha(caminho, None, 2, colunas)
    ws = ler(caminho, formulas=True)
    assert ws["A2"].value is None and ws["E2"].value is None, (ws["A2"].value, ws["E2"].value)
    assert ws["D2"].value.startswith("=IF(A2"), ws["D2"].value
    assert ws["A3"].value == 222
    print("  [5] excluir sem Excel: limpa e mantém a fórmula ...... OK")


def teste_6_incluir_varias() -> None:
    caminho = planilha_exemplo()
    colunas = edicao.esquema(caminho, None)
    novas = edicao.incluir_linhas(caminho, None, colunas, [{1: "333", 2: "20"}, {1: "444", 2: "10"},
                                                           {1: "555", 3: "02/09/2026"}])
    assert novas == [4, 5, 6], novas
    ws = ler(caminho, formulas=True)
    assert [ws[f"A{n}"].value for n in novas] == [333, 444, 555]
    assert ws["D6"].value == '=IF(A6<>"","SERV "&A6&" "&$B$1,"")', ws["D6"].value
    assert ws["C6"].value == datetime(2026, 9, 2) and ws["C6"].number_format == "dd/mm/yyyy"
    assert ws["B6"].value is None  # campo não preenchido fica vazio
    print("  [6] incluir várias linhas de uma vez ................. OK")


def main() -> None:
    print("Testes da edição de linhas da planilha (offline)\n")
    testes = [teste_1_formato, teste_2_deslocar_formula, teste_3_incluir, teste_4_editar, teste_5_limpar,
              teste_6_incluir_varias]
    for teste in testes:
        teste()
    print(f"\n{len(testes)} testes passaram.")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
