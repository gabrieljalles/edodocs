"""Atalho do Edodocs na área de trabalho, criado na primeira vez que o Edodocs.exe abre.

Não recria se a pessoa apagar (só se o programa for movido para outra pasta). Não cria quando
o .exe roda de dentro do .zip (o Windows abre de uma pasta temporária que some depois) nem
fora do executável (python app.py).
"""
import os
import sys
import tempfile
from pathlib import Path

NOME = "Edodocs.lnk"


def area_de_trabalho() -> Path:
    """A pasta da área de trabalho de verdade (pode estar dentro do OneDrive)."""
    try:
        from win32com.shell import shell, shellcon
        return Path(shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOPDIRECTORY, None, 0))
    except Exception:
        return Path(os.environ.get("USERPROFILE", Path.home())) / "Desktop"


def _de_dentro_do_zip(exe: Path) -> bool:
    temporaria = Path(tempfile.gettempdir()).resolve()
    return temporaria in exe.resolve().parents


def criar_se_preciso(ja_criado_para: str | None) -> str | None:
    """Cria o atalho se ainda não foi criado para ESTE caminho do programa. Devolve o caminho
    do .exe para guardar na configuração (None se nada foi feito)."""
    if not getattr(sys, "frozen", False):
        return None
    exe = Path(sys.executable)
    if ja_criado_para == str(exe) or _de_dentro_do_zip(exe):
        return None
    import pythoncom
    from win32com.client import Dispatch

    pythoncom.CoInitialize()
    try:
        atalho = Dispatch("WScript.Shell").CreateShortcut(str(area_de_trabalho() / NOME))
        atalho.TargetPath = str(exe)
        atalho.WorkingDirectory = str(exe.parent)
        atalho.IconLocation = f"{exe},0"
        atalho.Description = "Edodocs: atos no e-docs e no Diário Oficial"
        atalho.Save()
    finally:
        pythoncom.CoUninitialize()
    return str(exe)
