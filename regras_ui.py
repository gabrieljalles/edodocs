"""Janelas do assistente para montar regras sem fórmula (usadas pelo assistente.py):

    editar_texto_condicional   um campo do Word cujo texto muda conforme a planilha:
                               "Se A PEDIDO é SIM, escrever 'artigo 48'; senão, ..."
    editar_condicoes           só a lista de condições (ex.: quando usar outro Word)
    escolher_trecho            os parágrafos do Word que se repetem para cada pessoa

Tudo é escolhido em listas: a coluna, a comparação ("é", "não é", "está vazio"...) e o valor,
que já vem com os valores que existem na planilha. As regras ficam no processo.json (ver
processos.py: ligação "condicao", "outros_modelos").
"""
import html
from pathlib import Path

from nicegui import ui as gui

import atos
import processos

ROTULO_QUANTIDADE = atos.AUTOMATICOS[atos.QUANTIDADE][0]


def _valores_da_coluna(registros: list[dict], coluna: str, limite: int = 40) -> list[str]:
    """Valores diferentes que a coluna tem na planilha (para a pessoa só escolher)."""
    nome = atos.normalizar(coluna)
    if nome == atos.QUANTIDADE:
        return ["1", "2", "3"]
    vistos: dict[str, str] = {}
    for r in registros:
        valor = str(r.get(nome, "")).strip()
        if valor:
            vistos.setdefault(atos.normalizar(valor), valor)
        if len(vistos) >= limite:
            break
    return sorted(vistos.values(), key=str.lower)


class _Condicoes:
    """As linhas "COLUNA · comparação · valor" de uma lista de condições (todas precisam
    valer). Mexe direto na lista recebida."""

    def __init__(self, condicoes: list[dict], colunas: list[str], registros: list[dict], ao_mudar=None):
        self.condicoes = condicoes
        self.colunas = list(dict.fromkeys(colunas + [ROTULO_QUANTIDADE]))
        self.registros = registros
        self.ao_mudar = ao_mudar or (lambda: None)
        self.area = gui.column().classes("w-full gap-1")
        self.desenhar()

    def desenhar(self) -> None:
        self.area.clear()
        with self.area:
            for i, cond in enumerate(self.condicoes):
                self._linha(i, cond)
            gui.button("e também...", icon="add", on_click=self._adicionar) \
                .props("flat dense no-caps size=sm color=primary") \
                .tooltip("Acrescenta outra condição: todas precisam valer ao mesmo tempo")

    def _adicionar(self) -> None:
        self.condicoes.append({"campo": self.colunas[0] if self.colunas else "", "op": "igual", "valor": ""})
        self.desenhar()
        self.ao_mudar()

    def _linha(self, i: int, cond: dict) -> None:
        with gui.row().classes("w-full items-center gap-2 no-wrap"):
            gui.label("Se" if i == 0 else "e").classes("text-weight-medium t-forte w-6")
            gui.select(self.colunas, value=cond.get("campo") if cond.get("campo") in self.colunas else None,
                       label="Coluna", on_change=lambda e: self._mudar(cond, "campo", e.value, redesenhar=True)) \
                .props("dense outlined use-input input-debounce=0 options-dense").classes("w-56")
            gui.select(processos.OPERADORES, value=cond.get("op", "igual"),
                       on_change=lambda e: self._mudar(cond, "op", e.value, redesenhar=True)) \
                .props("dense outlined options-dense").classes("w-56")
            if cond.get("op", "igual") not in processos.SEM_VALOR:
                opcoes = _valores_da_coluna(self.registros, cond.get("campo", ""))
                atuais = [v.strip() for v in str(cond.get("valor", "")).split(";") if v.strip()]
                opcoes = list(dict.fromkeys(opcoes + atuais))
                multiplo = cond.get("op", "igual") in ("igual", "diferente")
                gui.select(opcoes, value=atuais if multiplo else (atuais[0] if atuais else None),
                           label="Valor (escolha ou digite)", multiple=multiplo, with_input=True,
                           new_value_mode="add-unique",
                           on_change=lambda e: self._mudar(
                               cond, "valor", ";".join(e.value) if isinstance(e.value, list) else (e.value or ""))) \
                    .props("dense outlined options-dense use-chips").classes("col")
            else:
                gui.space()
            gui.button(icon="close", on_click=lambda: self._remover(i)) \
                .props("flat round dense size=sm color=grey-7").tooltip("Tirar esta condição")

    def _mudar(self, cond: dict, chave: str, valor, redesenhar: bool = False) -> None:
        cond[chave] = valor or ""
        if redesenhar:
            self.desenhar()
        self.ao_mudar()

    def _remover(self, i: int) -> None:
        del self.condicoes[i]
        self.desenhar()
        self.ao_mudar()


def _cabecalho(titulo: str, explicacao: str, icone: str) -> None:
    with gui.row().classes("items-center gap-2 no-wrap"):
        gui.icon(icone, color="primary", size="28px")
        gui.label(titulo).classes("text-h6 t-forte")
    gui.label(explicacao).classes("text-body2 t-suave")


def editar_texto_condicional(camada, rotulo: str, regra: dict, colunas: list[str], registros: list[dict],
                             exemplo, ao_fechar) -> None:
    """Janela das regras de um campo: lista de "Se ... escrever ..." e o "senão".
    exemplo() devolve (descrição do ato da prévia, valor do campo) para mostrar ao vivo."""
    regra.setdefault("casos", [])
    regra.setdefault("senao", "")
    with camada, gui.dialog().props("persistent") as janela, \
            gui.card().classes("cartao w-[980px] max-w-full q-pa-lg gap-3"):
        _cabecalho("{{" + rotulo + "}} muda conforme a planilha",
                   "O Edodocs olha as regras de cima para baixo e usa o texto da primeira que servir. "
                   "Se nenhuma servir, usa o texto do fim.", "alt_route")
        area = gui.column().classes("w-full gap-3").style("max-height: 56vh; overflow-y: auto")
        with gui.card().classes("w-full fundo-cartao q-pa-sm"):
            with gui.row().classes("w-full items-center gap-2 no-wrap"):
                gui.icon("visibility", color="primary")
                ao_vivo = gui.label("").classes("text-body2 t-forte col")

        def atualizar() -> None:
            ato, valor = exemplo()
            ao_vivo.text = (f"Para {ato}: " + (f"“{valor}”" if valor else "(nada: o campo fica vazio)")) \
                if ato else "Escolha a planilha para ver um exemplo."

        def desenhar() -> None:
            area.clear()
            with area:
                for i, caso in enumerate(regra["casos"]):
                    with gui.card().classes("cartao w-full q-pa-sm gap-1"):
                        with gui.row().classes("w-full items-center"):
                            gui.label(f"Regra {i + 1}").classes("text-subtitle2 text-primary")
                            gui.space()
                            if i:
                                gui.button(icon="arrow_upward", on_click=lambda i=i: subir(i)) \
                                    .props("flat round dense size=sm").tooltip("Olhar esta regra antes")
                            gui.button(icon="delete", on_click=lambda i=i: remover(i)) \
                                .props("flat round dense size=sm color=negative").tooltip("Apagar a regra")
                        caso.setdefault("se", [])
                        _Condicoes(caso["se"], colunas, registros, atualizar)
                        with gui.row().classes("w-full items-center gap-2 no-wrap"):
                            gui.label("escrever:").classes("text-weight-medium t-forte")
                            gui.input(value=caso.get("texto", ""), placeholder="(nada)",
                                      on_change=lambda e, c=caso: (c.update(texto=e.value or ""), atualizar())) \
                                .props("dense outlined").classes("col")
                gui.button("Adicionar regra", icon="add", on_click=adicionar).props("outline no-caps")
                with gui.card().classes("cartao w-full q-pa-sm"):
                    with gui.row().classes("w-full items-center gap-2 no-wrap"):
                        gui.label("Se nenhuma regra servir, escrever:" if regra["casos"] else "Sempre escrever:") \
                            .classes("text-weight-medium t-forte")
                        gui.input(value=regra.get("senao", ""), placeholder="(nada)",
                                  on_change=lambda e: (regra.update(senao=e.value or ""), atualizar())) \
                            .props("dense outlined").classes("col")
            atualizar()

        def adicionar() -> None:
            regra["casos"].append({"se": [{"campo": colunas[0] if colunas else "", "op": "igual", "valor": ""}],
                                   "texto": ""})
            desenhar()

        def remover(i: int) -> None:
            del regra["casos"][i]
            desenhar()

        def subir(i: int) -> None:
            casos = regra["casos"]
            casos[i - 1], casos[i] = casos[i], casos[i - 1]
            desenhar()

        desenhar()
        gui.label("Dica: no texto, {{COLUNA}} põe um dado da planilha (ex.: retroativo a {{DATA}}). "
                  "Deixar vazio não escreve nada; se o parágrafo do Word ficar vazio, ele some do ato.") \
            .classes("text-caption t-suave")
        with gui.row().classes("w-full justify-end"):
            gui.button("Pronto", icon="check", on_click=lambda: (janela.close(), ao_fechar())) \
                .props("unelevated no-caps")
    janela.open()


def editar_condicoes(camada, titulo: str, explicacao: str, condicoes: list[dict], colunas: list[str],
                     registros: list[dict], ao_fechar) -> None:
    """Janela só com uma lista de condições (todas precisam valer)."""
    with camada, gui.dialog().props("persistent") as janela, \
            gui.card().classes("cartao w-[900px] max-w-full q-pa-lg gap-3"):
        _cabecalho(titulo, explicacao, "rule")
        if not condicoes:
            condicoes.append({"campo": ROTULO_QUANTIDADE, "op": "maior", "valor": "1"})
        _Condicoes(condicoes, colunas, registros)
        with gui.row().classes("w-full justify-end"):
            gui.button("Pronto", icon="check", on_click=lambda: (janela.close(), ao_fechar())) \
                .props("unelevated no-caps")
    janela.open()


def descrever_condicoes(condicoes: list[dict]) -> str:
    return " e ".join(processos.descrever_condicao(c) for c in condicoes) or "(sem condição: nunca é usado)"


def escolher_trecho(camada, modelo: Path, ao_marcar) -> None:
    """Janela com os parágrafos do Word: a pessoa marca os que se repetem para cada pessoa e
    o Edodocs grava {{REPETIR}} / {{FIM REPETIR}} em volta deles no Word."""
    paragrafos = atos.paragrafos_do_modelo(modelo)
    _, _, bloco = atos._abrir_modelo(modelo)
    atual = atos.trecho_que_repete(bloco)
    marcadores = {atos.REPETIR, atos.FIM_REPETIR}
    # Os parágrafos dos marcadores não aparecem: a pessoa só vê o texto do ato
    visiveis = [(i, t) for i, t in enumerate(paragrafos)
                if not ({atos.normalizar(m.group(1)) for m in atos.CHAVES.finditer(t)} & marcadores
                        and not atos.CHAVES.sub("", t).strip())]
    escolhidos = {i for i, _ in visiveis if atual and atual[0] <= i <= atual[1]}
    with camada, gui.dialog().props("persistent") as janela, \
            gui.card().classes("cartao w-[900px] max-w-full q-pa-lg gap-3"):
        _cabecalho("Qual parte se repete para cada pessoa?",
                   f"Marque os parágrafos de {modelo.name} que saem uma vez para cada pessoa do ato "
                   "(normalmente a linha com {{NOME}}, {{MATRÍCULA}}...). Pode marcar uma linha em branco "
                   "junto, para separar as pessoas.", "repeat")
        caixas = {}
        with gui.column().classes("w-full gap-0").style("max-height: 58vh; overflow-y: auto"):
            for i, texto in visiveis:
                with gui.row().classes("w-full items-start gap-2 no-wrap q-py-xs"):
                    caixas[i] = gui.checkbox(value=i in escolhidos)
                    gui.html(html.escape(texto) if texto.strip() else '<span class="t-suave">(linha em branco)</span>') \
                        .classes("text-body2 t-forte q-pt-xs")

        def gravar() -> None:
            marcados = [i for i, c in caixas.items() if c.value]
            try:
                atos.marcar_repeticao(modelo, min(marcados) if marcados else None,
                                      max(marcados) if marcados else None)
            except PermissionError:
                gui.notify("O Word está aberto: feche o arquivo e tente de novo", type="negative")
                return
            janela.close()
            if marcados and max(marcados) - min(marcados) + 1 > len(marcados):
                gui.notify("Marquei do primeiro ao último parágrafo escolhido (o trecho precisa ser seguido)",
                           type="info")
            ao_marcar()

        with gui.row().classes("w-full justify-end gap-2"):
            gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
            gui.button("Nada se repete", on_click=lambda: ([c.set_value(False) for c in caixas.values()], gravar())) \
                .props("flat no-caps")
            gui.button("Marcar no Word", icon="check", on_click=gravar).props("unelevated no-caps")
    janela.open()


def resumo_trecho(modelo: Path) -> str | None:
    """O começo do trecho que repete (parágrafo ou linha de tabela), ou None se o Word não tem."""
    return atos.resumo_da_repeticao(modelo)
