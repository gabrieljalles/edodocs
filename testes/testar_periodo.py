"""Regra "periodo": de uma data a outra, como nas portarias de designação.

    python testes/testar_periodo.py
"""
import sys
sys.stdout.reconfigure(encoding="utf-8")
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import processos  # noqa: E402

regra = {"tipo": "periodo", "coluna": "INÍCIO", "ate": "FIM"}
casos = [
    ("09/11/2026", "28/11/2026", "09 a 28 de novembro de 2026"),
    ("23/11/2026", "07/12/2026", "23 de novembro a 07 de dezembro de 2026"),
    ("21/12/2026", "04/01/2027", "21 de dezembro de 2026 a 04 de janeiro de 2027"),
    ("01/12/2026", "", "01 de dezembro de 2026"),
    ("", "18/12/2026", ""),
]
for inicio, fim, esperado in casos:
    saiu = processos.valor_da_regra(regra, {"INICIO": inicio, "FIM": fim})
    assert saiu == esperado, (inicio, fim, saiu)
assert processos.periodo(date(2026, 5, 4), date(2026, 5, 4)) == "04 de maio de 2026"
print("  [1] período: mês e ano só quando mudam, dia com 2 algarismos ... OK")

proc = processos.Processo(Path("."), {"ligacoes": {"PERÍODO": regra}})
colunas = {c["nome"]: c["tipo"] for c in proc.colunas_planilha_nova([("PERIODO", "PERÍODO")])}
assert colunas.get("INÍCIO") == "data" and colunas.get("FIM") == "data", colunas
print("  [2] a planilha ganha as duas colunas de data .................. OK")
print("\n2 testes passaram.")
