"""Abre pastas e arquivos (PDF, Word, manual) em PRIMEIRO PLANO, na frente do Edodocs.

Com os.startfile sozinho, o Windows abria a pasta atrás da janela do Edodocs (um programa em
segundo plano não pode jogar outra janela para a frente), e a pessoa clicava de novo achando
que nada tinha aberto. Aqui, depois de abrir, a nova janela é procurada e trazida para a frente.
"""
import ctypes
import os
import subprocess
import threading
import time
from ctypes import wintypes
from pathlib import Path

_user32 = ctypes.windll.user32
_ENUM = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


def _janelas() -> dict[int, str]:
    """{identificador: título} das janelas visíveis com título."""
    janelas: dict[int, str] = {}

    def cada(janela, _):
        if _user32.IsWindowVisible(janela):
            tamanho = _user32.GetWindowTextLengthW(janela)
            if tamanho:
                texto = ctypes.create_unicode_buffer(tamanho + 1)
                _user32.GetWindowTextW(janela, texto, tamanho + 1)
                janelas[janela] = texto.value
        return True

    _user32.EnumWindows(_ENUM(cada), 0)
    return janelas


def _para_frente(janela: int) -> None:
    # Apertar e soltar o Alt libera o Windows a trocar a janela da frente
    _user32.keybd_event(0x12, 0, 0, 0)
    _user32.keybd_event(0x12, 0, 2, 0)
    _user32.ShowWindow(janela, 9 if _user32.IsIconic(janela) else 5)  # restaurar / mostrar
    _user32.SetForegroundWindow(janela)


def _trazer(antes: dict[int, str], nome: str) -> None:
    """Espera a janela do que foi aberto aparecer (até 6 s) e a traz para a frente."""
    alvo = nome.lower()
    for _ in range(40):
        time.sleep(0.15)
        agora = _janelas()
        novas = [j for j in agora if j not in antes]
        escolhida = (next((j for j in novas if alvo and alvo in agora[j].lower()), None)
                     or (novas[0] if novas else None))
        if escolhida:
            _para_frente(escolhida)
            return
    # Nenhuma janela nova (ex.: a pasta já estava aberta): traz a que tem o nome dela
    existente = next((j for j, t in _janelas().items() if alvo and alvo in t.lower()), None)
    if existente:
        _para_frente(existente)


def mostrar_na_pasta(arquivo) -> None:
    """Abre a pasta do arquivo, com ele já selecionado, e traz a janela para a frente."""
    arquivo = Path(arquivo)
    antes = _janelas()
    subprocess.Popen(["explorer", "/select,", str(arquivo)])
    threading.Thread(target=_trazer, args=(antes, arquivo.parent.name), daemon=True).start()


def abrir(caminho) -> None:
    """Abre a pasta ou o arquivo e traz a janela dele para a frente."""
    caminho = Path(caminho)
    antes = _janelas()
    os.startfile(caminho)
    nome = caminho.name if caminho.is_dir() else caminho.stem
    threading.Thread(target=_trazer, args=(antes, nome), daemon=True).start()
