"""Editor das colunas da tabela dos atos: a ordem (setas), quais aparecem (olho) e as
caixinhas de marcar criadas pela pessoa (ex.: CONFERIDO). Usado no fim do assistente
(antes de salvar o processo) e no botão Colunas da página Atos.

Mexe só no config do processo: "colunas_tabela" (ordem e visibilidade) e "caixas" (ver
processos.py). Quem chama decide quando salvar.
"""
from nicegui import ui as gui

import atos


class EditorColunas:
    def __init__(self, config: dict, colunas, ao_mudar=None):
        """colunas() devolve a lista atual (Processo.colunas_tabela com este config)."""
        self.config = config
        self.colunas = colunas
        self.ao_mudar = ao_mudar or (lambda: None)
        with gui.column().classes("w-full gap-2"):
            with gui.row().classes("w-full items-center gap-2 no-wrap"):
                self.nova = gui.input("Nova caixinha (nome da coluna)", placeholder="Ex.: CONFERIDO") \
                    .props("outlined dense").classes("col")
                self.nova.on("keydown.enter", self._adicionar)
                gui.button("Adicionar", icon="check_box", on_click=self._adicionar).props("outline dense no-caps")
            gui.label("A caixinha vira uma coluna na tabela dos atos, para marcar à mão (fica gravada na "
                      "planilha como SIM). Use as setas para mudar a ordem e o olho para mostrar ou "
                      "esconder.").classes("text-caption t-suave")
            self.lista = gui.column().classes("w-full gap-0")
        self.desenhar()

    def desenhar(self) -> None:
        self.lista.clear()
        colunas = self.colunas()
        with self.lista:
            for i, c in enumerate(colunas):
                with gui.row().classes("w-full items-center gap-1 no-wrap q-py-none"):
                    gui.button(icon="arrow_upward", on_click=lambda i=i: self._mover(i, -1)) \
                        .props(f"flat round dense size=sm {'disable' if i == 0 else ''}")
                    gui.button(icon="arrow_downward", on_click=lambda i=i: self._mover(i, 1)) \
                        .props(f"flat round dense size=sm {'disable' if i == len(colunas) - 1 else ''}")
                    gui.button(icon="visibility" if c["ver"] else "visibility_off",
                               on_click=lambda i=i: self._alternar(i)) \
                        .props(f"flat round dense size=sm color={'primary' if c['ver'] else 'grey-6'}") \
                        .tooltip("Mostrar / esconder na tabela")
                    if c["caixa"]:
                        gui.icon("check_box", color="primary", size="18px")
                    gui.label(c["rotulo"]).classes("text-body2 col " + ("t-forte" if c["ver"] else "t-fraco"))
                    if c["caixa"]:
                        gui.button(icon="delete", on_click=lambda c=c: self._remover(c["rotulo"])) \
                            .props("flat round dense size=sm color=negative") \
                            .tooltip("Tirar a caixinha da tabela (o que já foi marcado fica na planilha)")

    def _gravar(self, colunas: list[dict]) -> None:
        self.config["colunas_tabela"] = [{"nome": c["nome"], "ver": c["ver"]} for c in colunas]
        self.desenhar()
        self.ao_mudar()

    def _mover(self, i: int, passo: int) -> None:
        colunas = self.colunas()
        j = i + passo
        if 0 <= j < len(colunas):
            colunas[i], colunas[j] = colunas[j], colunas[i]
            self._gravar(colunas)

    def _alternar(self, i: int) -> None:
        colunas = self.colunas()
        colunas[i]["ver"] = not colunas[i]["ver"]
        self._gravar(colunas)

    def _adicionar(self) -> None:
        nome = " ".join((self.nova.value or "").split()).upper()
        if not nome:
            gui.notify("Escreva o nome da caixinha (ex.: CONFERIDO)", type="warning")
            return
        if any(atos.normalizar(c["nome"]) == atos.normalizar(nome) for c in self.colunas()):
            gui.notify(f"Já existe uma coluna chamada {nome}", type="warning")
            return
        self.config["caixas"] = list(self.config.get("caixas") or []) + [nome]
        self.nova.value = ""
        self._gravar(self.colunas())  # a nova entra no fim, visível

    def _remover(self, rotulo: str) -> None:
        self.config["caixas"] = [c for c in self.config.get("caixas") or [] if c != rotulo]
        self._gravar(self.colunas())
