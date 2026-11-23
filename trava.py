"""Só um computador grava um arquivo por vez (planilha ou configuração na pasta da rede).

Quem vai gravar cria, ao lado do arquivo, o "~edodocs <nome>.trava" com o nome de quem está
gravando; os outros esperam ele sumir. Gravar é rápido (segundos): trava mais velha que
VELHA é de um programa que fechou no meio, e é ignorada.
"""
import getpass
import os
import socket
import threading
import time
from contextlib import contextmanager
from pathlib import Path

ESPERA = 30   # segundos esperando a vez antes de desistir
VELHA = 120   # segundos: trava esquecida
_minhas = threading.local()  # travas que esta linha de execução já tem (pode pedir de novo)


class Ocupado(Exception):
    """Outra pessoa está gravando o arquivo há mais tempo que ESPERA."""

    def __init__(self, quem: str):
        super().__init__(f"{quem or 'outra pessoa'} está salvando neste arquivo agora; tente de novo em instantes")
        self.quem = quem


def arquivo_da_trava(arquivo: Path) -> Path:
    return arquivo.with_name(f"~edodocs {arquivo.name}.trava")


def quem_esta(arquivo: Path) -> str | None:
    """Quem está gravando agora (None = ninguém)."""
    trava = arquivo_da_trava(arquivo)
    try:
        if time.time() - trava.stat().st_mtime > VELHA:
            return None
        return trava.read_text(encoding="utf-8").strip()
    except OSError:
        return None


def _eu() -> str:
    try:
        return f"{getpass.getuser()} ({socket.gethostname()})"
    except Exception:
        return "outro computador"


@contextmanager
def trava(arquivo: Path, espera: float = ESPERA):
    """with trava(planilha): grava... Espera a vez; Ocupado se demorar demais. Quem já tem a
    trava pode pedir de novo (ex.: incluir linhas, que lê e depois grava)."""
    caminho = arquivo_da_trava(arquivo)
    tenho = getattr(_minhas, "travas", None)
    if tenho is None:
        tenho = _minhas.travas = {}
    chave = str(caminho).lower()
    if tenho.get(chave):
        tenho[chave] += 1
        try:
            yield
        finally:
            tenho[chave] -= 1
        return
    limite = time.monotonic() + espera
    while True:
        try:
            fd = os.open(caminho, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            break
        except FileExistsError:
            try:
                if time.time() - caminho.stat().st_mtime > VELHA:
                    caminho.unlink(missing_ok=True)  # de um programa que fechou no meio
                    continue
            except OSError:
                continue
            if time.monotonic() > limite:
                raise Ocupado(quem_esta(arquivo) or "") from None
            time.sleep(0.3)
        except PermissionError:
            # No Windows, a trava que o outro está apagando agora dá "acesso negado": é esperar
            if time.monotonic() > limite:
                raise Ocupado(quem_esta(arquivo) or "") from None
            time.sleep(0.1)
        except OSError:
            fd = None  # pasta onde não dá para criar a trava: grava sem ela (como antes)
            break
    try:
        if fd is not None:
            os.write(fd, _eu().encode("utf-8"))
            os.close(fd)
        tenho[chave] = 1
        yield
    finally:
        tenho.pop(chave, None)
        if fd is not None:
            try:
                caminho.unlink(missing_ok=True)
            except OSError:
                pass
