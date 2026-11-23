"""Credenciais salvas em arquivo local. A senha é criptografada com o DPAPI do Windows
(só o mesmo usuário do Windows, no mesmo computador, consegue ler)."""
import base64
import ctypes
import json
from ctypes import wintypes
from pathlib import Path


class _Blob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _dpapi(dados: bytes, cifrar: bool) -> bytes:
    entrada = _Blob(len(dados), ctypes.create_string_buffer(dados, len(dados)))
    saida = _Blob()
    crypt32 = ctypes.windll.crypt32
    fn = crypt32.CryptProtectData if cifrar else crypt32.CryptUnprotectData
    if not fn(ctypes.byref(entrada), None, None, None, None, 0, ctypes.byref(saida)):
        raise ctypes.WinError()
    try:
        return ctypes.string_at(saida.pbData, saida.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(saida.pbData)


def _ler(arquivo: Path) -> dict:
    """Credenciais de todos os sites: {site: {"usuario", "senha"}}.
    O formato antigo (um só usuário, sem site) e a chave antiga "webedocs" viram Edocs."""
    try:
        d = json.loads(arquivo.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    if "usuario" in d:
        return {"edocs": d}
    if "webedocs" in d:
        d.setdefault("edocs", d.pop("webedocs"))
    return d


def carregar(arquivo: Path, site: str = "edocs") -> tuple[str, str] | None:
    try:
        d = _ler(arquivo)[site]
        senha = _dpapi(base64.b64decode(d["senha"]), cifrar=False).decode()
        return d["usuario"], senha
    except (OSError, ValueError, KeyError):
        return None


def salvar(arquivo: Path, usuario: str, senha: str, site: str = "edocs") -> None:
    """Salva o usuário do site, mantendo os dos outros sites."""
    arquivo.parent.mkdir(parents=True, exist_ok=True)
    todos = _ler(arquivo)
    cifrada = base64.b64encode(_dpapi(senha.encode(), cifrar=True)).decode()
    todos[site] = {"usuario": usuario, "senha": cifrada}
    arquivo.write_text(json.dumps(todos, indent=2), encoding="utf-8")


def remover(arquivo: Path, site: str) -> None:
    """Apaga só o usuário do site (ex.: login recusado), mantendo os outros."""
    todos = _ler(arquivo)
    if todos.pop(site, None) is not None:
        arquivo.write_text(json.dumps(todos, indent=2), encoding="utf-8")
