"""Cria a planilha de um processo para quem não tem uma (assistente, passo 3 → "Criar a
planilha para mim"): uma coluna por campo do Word, com lista suspensa no Excel para as
colunas de opções (SIM/NÃO, F/M...) e formato de data nas datas. A pessoa só preenche.

    criar(destino, processo.colunas_planilha_nova(campos))
"""
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import atos

ABA = "dados"
LINHAS_PRONTAS = 2000  # até onde vão as listas e o formato de data
SIM_NAO = {"SIM", "NAO"}
_FINA = Side(style="thin", color="000000")
BORDA = Border(left=_FINA, right=_FINA, top=_FINA, bottom=_FINA)


def _opcoes(opcoes: list[str]) -> list[str]:
    """SIM/NÃO sempre com as duas opções (a regra pode só ter falado do SIM)."""
    if opcoes and {atos.normalizar(o) for o in opcoes} <= SIM_NAO:
        return ["SIM", "NÃO"]
    return opcoes


def estilizar(ws, colunas: int, ate_linha: int = LINHAS_PRONTAS + 1) -> None:
    """Sem cores: bordas pretas finas em todas as células e o cabeçalho em negrito preto. As
    linhas de baixo já vão com borda (a linha nova copia o formato da de cima)."""
    for i in range(1, colunas + 1):
        ws.cell(1, i).font = Font(bold=True, color="000000")
        for linha in range(1, ate_linha + 1):
            ws.cell(linha, i).border = BORDA


def criar(destino: Path, colunas: list[dict]) -> Path:
    """Grava a planilha vazia (só o cabeçalho) com as listas e formatos de cada coluna."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = ABA
    ws.freeze_panes = "A2"
    for i, coluna in enumerate(colunas, 1):
        ws.cell(1, i, coluna["nome"]).alignment = Alignment(vertical="center")
        letra = get_column_letter(i)
        intervalo = f"{letra}2:{letra}{LINHAS_PRONTAS + 1}"
        opcoes = _opcoes(coluna.get("opcoes") or [])
        if coluna["tipo"] == "lista" and opcoes:
            lista = DataValidation(type="list", formula1='"' + ",".join(o.replace('"', "") for o in opcoes) + '"',
                                   allow_blank=True)
            # A lista ajuda, mas não trava: um valor novo ainda pode ser digitado
            lista.showErrorMessage = False
            lista.add(intervalo)
            ws.add_data_validation(lista)
        elif coluna["tipo"] == "data":
            for linha in range(2, LINHAS_PRONTAS + 2):
                ws.cell(linha, i).number_format = "dd/mm/yyyy"
        largura = max(len(coluna["nome"]), *(len(o) for o in opcoes), 10) + 4
        ws.column_dimensions[letra].width = min(largura, 45)
    ws.row_dimensions[1].height = 22
    estilizar(ws, len(colunas))
    destino.parent.mkdir(parents=True, exist_ok=True)
    wb.save(destino)
    return destino
