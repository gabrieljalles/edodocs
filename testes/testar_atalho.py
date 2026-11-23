"""Testes do atalho na área de trabalho (com um .exe falso e uma área de trabalho de teste:
a área de trabalho de verdade não é tocada).

  1. na primeira vez, cria Edodocs.lnk apontando para o .exe, com o ícone dele;
  2. na segunda vez (mesmo caminho), não cria de novo, nem se a pessoa tiver apagado;
  3. se o programa mudou de pasta, cria o atalho para o lugar novo;
  4. aberto de dentro do .zip (pasta temporária do Windows) ou fora do .exe, não cria.

Rode com:  python testes/testar_atalho.py
"""
import shutil
import sys
import tempfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import atalho  # noqa: E402
# Carrega o pywin32 antes de fingir ser o .exe (ele também olha sys.frozen ao ser importado)
import pythoncom  # noqa: E402,F401
import win32com.client  # noqa: E402,F401
from win32com.shell import shell  # noqa: E402,F401

# Fora da pasta temporária do Windows (lá o atalho não é criado de propósito)
BASE = RAIZ / "build" / "teste_atalho"


def ler(lnk: Path) -> tuple[str, str]:
    from win32com.client import Dispatch
    pythoncom.CoInitialize()
    try:
        a = Dispatch("WScript.Shell").CreateShortcut(str(lnk))
        return a.TargetPath, a.IconLocation
    finally:
        pythoncom.CoUninitialize()


def preparar(pasta: str) -> Path:
    exe = BASE / pasta / "Edodocs.exe"
    exe.parent.mkdir(parents=True, exist_ok=True)
    exe.write_bytes(b"")
    return exe


def main() -> None:
    print("Testes do atalho na área de trabalho\n")
    shutil.rmtree(BASE, ignore_errors=True)
    mesa = BASE / "area de trabalho"
    mesa.mkdir(parents=True)
    guardado = (getattr(sys, "frozen", None), sys.executable, atalho.area_de_trabalho)
    atalho.area_de_trabalho = lambda: mesa
    try:
        sys.frozen = True
        sys.executable = str(preparar("Edodocs"))
        feito = atalho.criar_se_preciso(None)
        alvo, icone = ler(mesa / "Edodocs.lnk")
        assert feito == sys.executable and Path(alvo) == Path(sys.executable), (feito, alvo)
        assert icone.startswith(sys.executable), icone
        print("  [1] primeira vez: cria o atalho com o ícone ............ OK")

        (mesa / "Edodocs.lnk").unlink()  # a pessoa apagou
        assert atalho.criar_se_preciso(feito) is None and not (mesa / "Edodocs.lnk").exists()
        print("  [2] mesma pasta: não recria (nem se a pessoa apagou) ... OK")

        sys.executable = str(preparar("Outra pasta"))
        novo = atalho.criar_se_preciso(feito)
        assert novo == sys.executable and Path(ler(mesa / "Edodocs.lnk")[0]) == Path(novo)
        print("  [3] programa mudou de pasta: atalho para o lugar novo .. OK")

        (mesa / "Edodocs.lnk").unlink()
        sys.executable = str(Path(tempfile.gettempdir()) / "Temp1_Edodocs.zip" / "Edodocs.exe")
        assert atalho.criar_se_preciso(None) is None and not (mesa / "Edodocs.lnk").exists()
        del sys.frozen
        assert atalho.criar_se_preciso(None) is None
        print("  [4] de dentro do .zip ou fora do .exe: não cria ......... OK")
    finally:
        congelado, sys.executable, atalho.area_de_trabalho = guardado
        if congelado is not None:
            sys.frozen = congelado
        elif hasattr(sys, "frozen"):
            del sys.frozen
        shutil.rmtree(BASE, ignore_errors=True)
    print("\n4 testes passaram.")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
