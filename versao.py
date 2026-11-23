"""Versão do Edodocs: a data do último commit e o código curto dele (ex.: "2026.10.16 · 2ce046f").

No executável, vem do assets/versao.txt que o build.py grava na hora de gerar o .exe; rodando
pelo código (python app.py), é lida do git. Sem nenhum dos dois: "desenvolvimento".
"""
import subprocess
import sys
from functools import lru_cache
from pathlib import Path

RAIZ = Path(__file__).parent
ARQUIVO = Path(getattr(sys, "_MEIPASS", RAIZ)) / "assets" / "versao.txt"


def do_git(pasta: Path = RAIZ) -> str | None:
    """A versão pelo último commit; com mudanças ainda sem commit, ganha um "+" no fim."""
    try:
        def git(*args) -> str:
            return subprocess.run(["git", *args], cwd=pasta, capture_output=True, text=True,
                                  timeout=5, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                                  check=True).stdout.strip()

        data, codigo = git("log", "-1", "--format=%cd %h", "--date=format:%Y.%m.%d").split()
        sujo = bool(git("status", "--porcelain", "--untracked-files=no"))
        return f"{data} · {codigo}{'+' if sujo else ''}"
    except Exception:
        return None


@lru_cache(maxsize=1)
def atual() -> str:
    try:
        texto = ARQUIVO.read_text(encoding="utf-8").strip()
        if texto:
            return texto
    except OSError:
        pass
    return do_git() or "desenvolvimento"
