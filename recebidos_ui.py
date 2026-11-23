"""Página Caixa de Saída (a Caixa de Entrada do e-docs) e o cartão da IA em
Configurações.

Os recebidos aparecem em 3 colunas: cada quadradinho com o PDF para abrir, o resumo da IA (só
com a IA conectada: sem ela, o resumo nem é mencionado) e os
botões Ciente / Vistar dele. "Dar ciente em todos" dá ciente em tudo o que tem o botão Ciente;
"Vistar tudo" dá o ciente que faltar e depois vista os que são de visto. Despachos não têm
ciente nem visto: ficam para o site (botão do globo). Os de visto também podem ser
cancelados (jogados fora, sem ciente e sem visto): um a um ou marcando vários.
A IA não resume sozinha: o ícone de IA do quadradinho resume aquele, "Resumir tudo" resume
todos. Um resumo já feito antes (guardado) aparece direto.
"""
import asyncio
import html
import re

from nicegui import background_tasks, run
from nicegui import ui as gui

import credenciais
import edocs_http
import ia
import main
from abrir import abrir

PASTA = main.SAIDAS / "Recebidos"


INVALIDOS = re.compile(r'[<>:"/\\|?*]')


def _nome_arquivo(doc: str, nome: str) -> str:
    return f"{doc} - {INVALIDOS.sub('_', nome)}"


FALHOU = "A IA não conseguiu resumir"


def _sem_resumo(item: dict) -> bool:
    """Ainda dá para pedir o resumo: nunca pedido ou a IA falhou (tenta de novo)."""
    return item["resumo"] == "" or str(item["resumo"]).startswith(FALHOU)


# Caixa de Saída em Cards (quadradinhos, 3 por linha) ou em Lista (uma linha por documento)
CARDS, LISTA = "cards", "lista"
CHAVE_MODO = "caixa_saida_modo"


class PaginaRecebidos:
    def __init__(self, tela):
        self.tela = tela
        self.trava = asyncio.Lock()       # uma ação no site por vez
        # um pedido ao site por vez (a sessão é uma só): a Tramitação entra entre os pedidos
        # de uma ação longa (baixar os PDFs), sem esperar ela acabar
        self.trava_pedido = asyncio.Lock()
        self.sessao = None
        self.itens: dict[str, dict] = {}  # doc → dados do recebido (+ "pdf", "resumo", "feito")
        self.carregado = False
        self.marcados: set[str] = set()  # documentos marcados para cancelar
        self.com_ia = ia.conectada()
        self.resumir_todos = False  # "Resumir tudo": os PDFs que ainda estão baixando também
        with gui.row().classes("w-full items-center gap-2"):
            self.explicacao = gui.label("").classes("text-body1 t-forte")
            self._explicar()
            gui.space()
            # Cards (como era) ou Lista (um embaixo do outro, mais documentos por vez); fica lembrado
            self.modo = main.ler_config().get(CHAVE_MODO) or CARDS
            gui.toggle({CARDS: "Cards", LISTA: "Lista"}, value=self.modo, on_change=self._trocar_modo) \
                .props("dense no-caps unelevated toggle-color=primary").tooltip("Como mostrar os documentos")
            gui.button(icon="public", on_click=lambda: self.tela._abrir_site("edocs")) \
                .props("flat round dense").classes("t-suave").tooltip("Abrir o e-docs já logado")
            gui.button("Atualizar", icon="refresh", on_click=self.carregar).props("flat no-caps")
            gui.button("Abrir pasta", icon="folder_open", on_click=lambda: abrir(self._pasta())) \
                .props("flat no-caps")
            self.botao_cancelar = gui.button("Cancelar marcados", icon="delete_sweep", color="negative",
                                             on_click=lambda: self._confirmar_cancelar(sorted(self.marcados))) \
                .props("outline no-caps")
            self.botao_cancelar.set_visibility(False)
            self.botao_resumir = gui.button("Resumir tudo", icon="auto_awesome", on_click=self._resumir_tudo) \
                .props("outline no-caps").tooltip("A IA resume todos os documentos (um por um)")
            self.botao_resumir.set_visibility(self.com_ia)
            self.botao_ciente = gui.button("Dar ciente em todos", icon="done_all",
                                           on_click=self._ciente_em_todos).props("outline no-caps")
            self.botao_vistar = gui.button("Vistar tudo", icon="fact_check", color="positive",
                                           on_click=self._vistar_tudo).props("unelevated no-caps")
        self.situacao = gui.label("").classes("text-caption t-suave")
        self.grade = gui.column().classes("w-full q-mt-sm")

    @staticmethod
    def _pasta():
        PASTA.mkdir(parents=True, exist_ok=True)
        return PASTA

    # --- Sessão no e-docs -------------------------------------------------------------------
    async def _no_site(self, funcao, *args):
        """Roda funcao(sessao, *args) fora da tela; entra de novo se a sessão tiver caído."""
        for tentativa in range(2):
            async with self.trava_pedido:
                if self.sessao is None:
                    dados = credenciais.carregar(main.CREDENCIAIS, "edocs")
                    if not dados:
                        raise RuntimeError("cadastre seu usuário do e-docs em Configurações")
                    self.sessao = await run.io_bound(edocs_http.entrar, *dados)
                try:
                    return await run.io_bound(funcao, self.sessao, *args)
                except edocs_http.SessaoExpirada:
                    self.sessao = None
        raise RuntimeError("a sessão do e-docs caiu; tente de novo")

    def _avisar(self, texto: str, type: str = "info") -> None:
        """Aviso na tela pela própria página (o botão clicado pode ter sido redesenhado)."""
        with self.situacao:
            gui.notify(texto, type=type)

    def _ocupado(self, sim: bool) -> None:
        for botao in (self.botao_ciente, self.botao_vistar, self.botao_cancelar):
            botao.set_enabled(not sim)

    def _marcar(self, doc: str, valor: bool) -> None:
        if valor:
            self.marcados.add(doc)
        else:
            self.marcados.discard(doc)
        self._mostrar_marcados()

    def _mostrar_marcados(self) -> None:
        self.marcados &= {d for d, i in self.itens.items() if i.get("cancelavel") and not i["feito"]}
        self.botao_cancelar.set_visibility(bool(self.marcados))
        self.botao_cancelar.text = f"Cancelar marcados ({len(self.marcados)})"

    def _explicar(self) -> None:
        self.explicacao.text = ("Documentos recebidos por você no e-docs"
                                + (". Clique no ícone de IA para resumir um documento." if self.com_ia else "."))

    # --- Carregar: lista, PDFs e resumos -----------------------------------------------------
    async def carregar(self) -> None:
        if self.trava.locked():
            return
        async with self.trava:
            self.carregado = True
            self.com_ia = ia.conectada()  # a chave pode ter sido configurada depois de abrir
            self.resumir_todos = False
            self.botao_resumir.set_visibility(self.com_ia)
            self._explicar()
            self._ocupado(True)
            self.situacao.text = "Buscando os recebidos no e-docs..."
            try:
                lista = await self._no_site(edocs_http.recebidos)
            except Exception as e:
                self.situacao.text = f"Não deu para buscar os recebidos: {e}"
                self._ocupado(False)
                return
            self.itens = {r["doc"]: dict(r, pdf=None, resumo=None, feito="") for r in lista}
            self.marcados.clear()
            self._mostrar_marcados()
            self._desenhar()
            self._ocupado(False)
            # Um por um: baixa o PDF (o resumo só quando pedido)
            for n, (doc, item) in enumerate(list(self.itens.items()), 1):
                self.situacao.text = f"Baixando {n} de {len(self.itens)}..."
                await self._baixar(doc, item)
            self.situacao.text = (f"{len(self.itens)} recebido(s). Os PDFs ficam em saidas\\Recebidos."
                                  if self.itens else "Nenhum documento na sua Caixa de Entrada.")

    async def _baixar(self, doc: str, item: dict) -> None:
        try:
            nome, conteudo = await self._no_site(edocs_http.baixar_recebido, doc)
            arquivo = self._pasta() / _nome_arquivo(doc, nome)
            arquivo.write_bytes(conteudo)
            item["pdf"] = arquivo
        except Exception as e:
            item["resumo"] = f"Não deu para baixar o PDF: {e}"
            self._atualizar(doc)
            return
        # "" = ainda sem resumo (o ícone de IA resume); o já feito antes aparece direto
        item["resumo"] = (ia.resumo_guardado(doc) or "") if self.com_ia else ""
        self._atualizar(doc)
        if self.resumir_todos:
            await self._resumir(doc)

    async def _resumir(self, doc: str) -> None:
        """A IA resume um documento (o PDF já baixado)."""
        item = self.itens.get(doc)
        if not item or not item["pdf"] or item["resumo"] == "…":
            return
        item["resumo"] = "…"
        self._atualizar(doc)
        try:
            item["resumo"] = await run.io_bound(ia.resumir_pdf, doc, item["pdf"].read_bytes(),
                                                item.get("titulo", ""))
        except Exception as e:
            item["resumo"] = f"{FALHOU}: {type(e).__name__}. Abra o PDF para ler."
        self._atualizar(doc)

    async def _resumir_tudo(self) -> None:
        """Resume todos os que ainda não têm resumo; os que ainda estão baixando são resumidos
        assim que o PDF chegar."""
        self.resumir_todos = True
        faltam = [d for d, i in self.itens.items() if i["pdf"] and _sem_resumo(i)]
        if not faltam and all(i["pdf"] for i in self.itens.values()):
            self._avisar("Todos os documentos já estão resumidos")
            return
        self.botao_resumir.set_enabled(False)
        try:
            for doc in faltam:
                await self._resumir(doc)
        finally:
            self.botao_resumir.set_enabled(True)

    # --- Quadradinhos ------------------------------------------------------------------------
    def _trocar_modo(self, e) -> None:
        self.modo = e.value or CARDS
        main.salvar_config(**{CHAVE_MODO: self.modo})
        self._desenhar()

    def _desenhar(self) -> None:
        self.grade.clear()
        self.cartoes = {}
        lista = self.modo == LISTA
        with self.grade:
            with (gui.column().classes("w-full gap-0 cartao fundo-cartao").style("overflow: hidden")
                  if lista else gui.grid(columns=3).classes("w-full gap-4")):
                for doc in self.itens:
                    self.cartoes[doc] = (gui.row().classes("w-full items-center gap-2 no-wrap q-px-md q-py-xs linha-caixa")
                                         if lista else gui.card().classes("cartao q-pa-md w-full gap-2"))
                    self._atualizar(doc)

    def _atualizar(self, doc: str) -> None:
        cartao = getattr(self, "cartoes", {}).get(doc)
        item = self.itens.get(doc)
        if cartao is None or item is None:
            return
        try:
            cartao.clear()
        except RuntimeError:
            return  # a janela foi fechada enquanto baixava os PDFs: não há mais tela para atualizar
        if self.modo == LISTA:
            self._linha(cartao, doc, item)
            return
        papel = " / ".join(dict.fromkeys(p for p in item["papeis"] if p)) or "Recebido"
        with cartao:
            with gui.row().classes("w-full items-center gap-2 no-wrap"):
                if item.get("cancelavel") and not item["feito"]:
                    gui.checkbox(value=doc in self.marcados, on_change=lambda e, d=doc: self._marcar(d, e.value)) \
                        .props("dense").tooltip("Marcar para cancelar")
                if item.get("tramitacao"):  # o caminho do documento, ao lado do status
                    gui.button(icon="route", on_click=lambda d=doc: self._mostrar_tramitacao(d)) \
                        .props("flat round dense size=sm color=blue-grey-8") \
                        .tooltip("Tramitação: por onde o documento já passou e quem ainda falta")
                # Cancelado ou vistado: o documento saiu da caixa, então fica só o resultado (sem o
                # papel "Visto" junto, que daria a entender as duas coisas)
                if item["feito"] not in ("Cancelado", "Vistado"):
                    gui.chip(papel, color=None).props("dense square").classes("chip-campo")
                if item.get("andamento"):  # é a vez dele no "Vistar tudo"
                    with gui.row().classes("items-center gap-1 no-wrap"):
                        gui.spinner(size="16px", color="primary")
                        gui.label(item["andamento"]).classes("text-caption text-weight-medium text-primary")
                elif item.get("na_fila"):
                    gui.chip("Na fila", icon="schedule", color=None).props("dense square outline") \
                        .classes("t-suave")
                if item["feito"]:
                    cancelado = item["feito"] == "Cancelado"
                    gui.chip(item["feito"], icon="delete" if cancelado else "check",
                             color="negative" if cancelado else "positive", text_color="white") \
                        .props("dense square")
                gui.space()
                gui.label(item.get("data", "")[:10]).classes("text-caption t-suave")
            gui.label(item.get("titulo") or item.get("identificacao") or f"Documento {doc}") \
                .classes("text-subtitle1 text-weight-medium t-forte")
            gui.label(" · ".join(x for x in (item.get("tipo"), item.get("remetente")) if x)) \
                .classes("text-caption t-suave")
            if not self.com_ia:
                # Sem IA: nada de resumo; só o aviso enquanto baixa (ou o erro do download)
                if item["resumo"] is None:
                    with gui.row().classes("items-center gap-2"):
                        gui.spinner(size="sm")
                        gui.label("Baixando o documento...").classes("text-body2 t-suave")
                elif item["resumo"]:
                    gui.label(item["resumo"]).classes("text-body2 text-negative")
            else:
                with gui.element("div").classes("w-full fundo-cartao rounded q-pa-sm") \
                        .style("min-height: 110px"):
                    if item["resumo"] in (None, "…"):
                        with gui.row().classes("items-center gap-2"):
                            gui.spinner(size="sm")
                            gui.label("Baixando o documento..." if item["resumo"] is None
                                      else "A IA está resumindo...").classes("text-body2 t-suave")
                    elif not item["resumo"]:
                        with gui.row().classes("items-center gap-2 no-wrap"):
                            gui.icon("auto_awesome", color="grey-6", size="18px")
                            gui.label("Ainda sem resumo. Clique no ícone de IA para resumir.") \
                                .classes("text-body2 t-suave")
                    else:
                        with gui.row().classes("items-start gap-2 no-wrap"):
                            gui.icon("auto_awesome", color="primary", size="18px").classes("q-mt-xs")
                            gui.label(item["resumo"]).classes("text-body2 t-forte")
            with gui.row().classes("w-full items-center gap-2"):
                gui.button("Abrir PDF", icon="picture_as_pdf",
                           on_click=lambda d=doc: abrir(self.itens[d]["pdf"])) \
                    .props("flat dense no-caps").set_enabled(bool(item["pdf"]))
                if self.com_ia:
                    ia_botao = gui.button(icon="auto_awesome", on_click=lambda d=doc: self._resumir(d)) \
                        .props("flat round dense color=primary") \
                        .tooltip("Resumir este documento com a IA")
                    # Já resumido (ou resumindo): o resumo fica guardado, não precisa de novo
                    ia_botao.set_enabled(bool(item["pdf"]) and _sem_resumo(item))
                gui.space()
                if "Despacho" in item["papeis"] and not (item["ciente"] or item["vistavel"]):
                    gui.label("Despachar ou encaminhar no site").classes("text-caption t-suave")
                    gui.button(icon="public", on_click=lambda: self.tela._abrir_site("edocs")) \
                        .props("flat round dense size=sm").tooltip("Abrir o e-docs já logado")
            # As ações numa linha só, com a mesma largura e o mesmo desenho (fundo claro, texto e
            # ícone fortes da mesma cor); só a cor muda: vermelho cancela, azul ciente, verde vista
            acoes = self._acoes(doc, item)
            if acoes:
                with gui.row().classes("w-full gap-2 no-wrap"):
                    for texto, icone, cor, ao_clicar, dica in acoes:
                        self._botao_acao(texto, icone, cor, ao_clicar).classes("col").tooltip(dica)

    def _acoes(self, doc: str, item: dict) -> list:
        """As ações do documento: (texto, ícone, cor, ao clicar, dica)."""
        acoes = []
        if item.get("cancelavel") and not item["feito"]:
            acoes.append(("Cancelar", "delete", "red", lambda d=doc: self._confirmar_cancelar([d]),
                          "Joga fora o documento no e-docs, sem ciente e sem visto"))
        if item["ciente"]:
            acoes.append(("Ciente", "done", "blue", lambda d=doc: self._um(d, "ciente"), "Dá ciente no documento"))
        if item["vistavel"]:
            acoes.append(("Vistar", "fact_check", "green", lambda d=doc: self._um(d, "vistar"),
                          "Dá ciente (se faltar) e vista o documento"))
        return acoes

    def _linha(self, linha, doc: str, item: dict) -> None:
        """O documento numa linha só (modo Lista): marcar, situação, título, de quem, data e ações.
        O resumo da IA aparece passando o mouse no título."""
        papel = " / ".join(dict.fromkeys(p for p in item["papeis"] if p)) or "Recebido"
        with linha:
            if item.get("cancelavel") and not item["feito"]:
                gui.checkbox(value=doc in self.marcados, on_change=lambda e, d=doc: self._marcar(d, e.value)) \
                    .props("dense").tooltip("Marcar para cancelar")
            else:
                gui.element("div").style("width: 27px; flex-shrink: 0")
            # Situação numa largura fixa, para os títulos ficarem alinhados
            with gui.row().classes("items-center gap-1 no-wrap").style("width: 150px; flex-shrink: 0"):
                if item.get("andamento"):
                    gui.spinner(size="16px", color="primary")
                    gui.label(item["andamento"]).classes("text-caption text-weight-medium text-primary ellipsis")
                elif item["feito"]:
                    cancelado = item["feito"] == "Cancelado"
                    gui.chip(item["feito"], icon="delete" if cancelado else "check",
                             color="negative" if cancelado else "positive", text_color="white") \
                        .props("dense square")
                elif item.get("na_fila"):
                    gui.chip("Na fila", icon="schedule", color=None).props("dense square outline").classes("t-suave")
                else:
                    gui.chip(papel, color=None).props("dense square").classes("chip-campo")
            with gui.column().classes("col gap-0").style("min-width: 0"):
                titulo = gui.label(item.get("titulo") or item.get("identificacao") or f"Documento {doc}") \
                    .classes("text-body2 text-weight-medium t-forte ellipsis w-full")
                if self.com_ia and item["resumo"] and item["resumo"] != "…":
                    titulo.tooltip(item["resumo"])
                gui.label(" · ".join(x for x in (item.get("tipo"), item.get("remetente")) if x)) \
                    .classes("text-caption t-suave ellipsis w-full")
            gui.label(item.get("data", "")[:10]).classes("text-caption t-suave").style("flex-shrink: 0")
            with gui.row().classes("items-center gap-1 no-wrap").style("flex-shrink: 0"):
                if item.get("tramitacao"):
                    gui.button(icon="route", on_click=lambda d=doc: self._mostrar_tramitacao(d)) \
                        .props("flat round dense size=sm color=blue-grey-8").tooltip("Tramitação")
                if item["pdf"]:
                    gui.button(icon="picture_as_pdf", on_click=lambda d=doc: abrir(self.itens[d]["pdf"])) \
                        .props("flat round dense size=sm color=negative").tooltip("Abrir PDF")
                else:
                    gui.spinner(size="16px").tooltip("Baixando o documento...")
                if self.com_ia:
                    if item["resumo"] == "…":
                        gui.spinner(size="16px", color="primary").tooltip("A IA está resumindo...")
                    else:
                        ia_botao = gui.button(icon="auto_awesome", on_click=lambda d=doc: self._resumir(d)) \
                            .props("flat round dense size=sm color=primary") \
                            .tooltip(item["resumo"] or "Resumir este documento com a IA")
                        ia_botao.set_enabled(bool(item["pdf"]) and _sem_resumo(item) or bool(item["resumo"]))
            # Ações num bloco de largura fixa: datas e ícones alinhados de uma linha para outra
            with gui.row().classes("items-center gap-1 no-wrap justify-end").style("width: 290px; flex-shrink: 0"):
                for texto, icone, cor, ao_clicar, dica in self._acoes(doc, item):
                    self._botao_acao(texto, icone, cor, ao_clicar).props("size=sm").tooltip(dica)

    @staticmethod
    def _botao_acao(texto: str, icone: str, cor: str, ao_clicar):
        """Botão de ação do quadradinho: todos iguais, só a cor muda (red, blue, green)."""
        return gui.button(texto, icon=icone, color=f"{cor}-1", on_click=ao_clicar) \
            .props(f"unelevated dense no-caps text-color={cor}-9 padding='4px 12px'") \
            .classes("rounded-borders text-weight-medium")

    # --- Ações -----------------------------------------------------------------------------
    async def _fazer(self, doc: str, acao: str) -> None:
        """Ciente ou visto em um documento (dentro da trava). O quadradinho mostra que é a
        vez dele enquanto o site responde."""
        item = self.itens[doc]
        item["andamento"] = "Dando ciente..." if acao == "ciente" else "Vistando..."
        item["na_fila"] = False
        self._atualizar(doc)
        try:
            if acao == "ciente":
                await self._no_site(edocs_http.dar_ciente, doc)
                item["ciente"], item["feito"] = False, "Ciente"
            else:
                await self._no_site(edocs_http.vistar_documento, doc)
                item["ciente"] = item["vistavel"] = False
                item["feito"] = "Vistado"
        finally:
            item["andamento"] = ""
            self._atualizar(doc)

    def _titulo(self, doc: str) -> str:
        item = self.itens.get(doc, {})
        return item.get("titulo") or item.get("identificacao") or f"Documento {doc}"

    async def _um(self, doc: str, acao: str) -> None:
        if self.trava.locked():
            self._avisar("Aguarde terminar o que está em andamento", type="warning")
            return
        async with self.trava:
            self._ocupado(True)
            self.tela.andamento(f"{'Dando ciente' if acao == 'ciente' else 'Vistando'}: {self._titulo(doc)}")
            try:
                await self._fazer(doc, acao)
                self._avisar(("Ciente dado" if acao == "ciente" else "Documento vistado") +
                             f": {self._titulo(doc)}", type="positive")
            except Exception as e:
                self._avisar(f"Não deu certo: {e}", type="negative")
            finally:
                self._ocupado(False)
                self.tela.andamento(None)

    def _confirmar(self, titulo: str, texto: str, depois) -> None:
        with self.tela.camada, gui.dialog() as janela, gui.card().classes("cartao w-[520px] q-pa-lg"):
            gui.label(titulo).classes("text-h6")
            gui.label(texto).classes("text-body2").style("white-space: pre-line")
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Confirmar", icon="check",
                           on_click=lambda: (janela.close(), background_tasks.create(depois()))) \
                    .props("unelevated no-caps")
        janela.open()

    def _ciente_em_todos(self) -> None:
        alvos = [d for d, i in self.itens.items() if i["ciente"]]
        if not alvos:
            self._avisar("Nenhum recebido esperando ciente", type="info")
            return
        self._confirmar("Dar ciente em todos?", f"Vai dar ciente em {len(alvos)} documento(s) no e-docs.",
                        lambda: self._em_lote(alvos, []))

    def _vistar_tudo(self) -> None:
        cientes = [d for d, i in self.itens.items() if i["ciente"] and not i["vistavel"]]
        vistos = [d for d, i in self.itens.items() if i["vistavel"]]
        if not (cientes or vistos):
            self._avisar("Nenhum recebido esperando ciente ou visto", type="info")
            return
        self._confirmar("Vistar tudo?",
                        f"Primeiro dá ciente no que faltar ({len(cientes)} documento(s)) e depois vista "
                        f"{len(vistos)} documento(s) no e-docs.\nO visto não pode ser desfeito pelo Edodocs.",
                        lambda: self._em_lote(cientes, vistos))

    # --- Tramitação (o caminho do documento) ---
    def _mostrar_tramitacao(self, doc: str) -> None:
        """Janela com o caminho do documento: verde onde já passou, destaque onde está parado
        e cinza o que falta."""
        with self.tela.camada, gui.dialog() as janela, gui.card().classes("cartao w-[560px] max-w-full q-pa-lg gap-3"):
            with gui.row().classes("w-full items-start no-wrap gap-2"):
                gui.icon("route", color="primary", size="26px")
                with gui.column().classes("col gap-0"):
                    gui.label("Tramitação").classes("text-h6 t-forte")
                    gui.label(self._titulo(doc)).classes("text-caption t-suave")
                gui.button(icon="close", on_click=janela.close).props("flat round dense").tooltip("Fechar")
            area = gui.column().classes("w-full items-center")
            with area, gui.row().classes("items-center gap-2 q-pa-md"):
                gui.spinner()
                gui.label("Buscando no e-docs...").classes("t-suave")
        janela.open()
        background_tasks.create(self._carregar_tramitacao(doc, area))

    async def _carregar_tramitacao(self, doc: str, area) -> None:
        try:
            passos = await self._no_site(edocs_http.tramitacao, self.itens[doc]["tramitacao"])
        except Exception as e:
            area.clear()
            with area:
                gui.label(f"Não deu para buscar a tramitação: {e}").classes("text-body2 text-negative")
            return
        area.clear()
        feitos = sum(p["feito"] for p in passos)
        atual = next((p for p in passos if not p["feito"]), None)
        with area:
            gui.html(_caminho_html(passos), sanitize=False).classes("w-full")
            resumo = f"{feitos} de {len(passos)} etapas" + (
                f" · agora: {atual['situacao'].lower()} ({atual['nome']})" if atual else " · concluído")
            gui.label(resumo).classes("text-caption t-suave q-mt-sm")

    # --- Cancelar (jogar fora, sem ciente e sem visto) ---
    def _confirmar_cancelar(self, docs: list[str]) -> None:
        docs = [d for d in docs if d in self.itens]
        if not docs:
            return
        with self.tela.camada, gui.dialog() as janela, gui.card().classes("cartao w-[560px] max-w-full q-pa-lg gap-2"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("delete_forever", color="negative", size="28px")
                gui.label(f"Cancelar {len(docs)} documento(s)?").classes("text-h6")
            gui.label("O documento é cancelado no e-docs: sai da sua Caixa de Entrada sem ciente e sem visto. "
                      "Não dá para desfazer pelo Edodocs.").classes("text-body2")
            with gui.element("div").classes("w-full fundo-previa rounded q-pa-sm") \
                    .style("max-height: 160px; overflow-y: auto"):
                for d in docs:
                    gui.label(self._titulo(d)).classes("text-caption t-forte")
            # O e-docs exige um motivo com pelo menos 5 palavras: o campo fica em destaque, conta
            # as palavras e só libera o botão quando chega lá
            with gui.column().classes("w-full gap-1 q-pa-sm rounded").style("border: 2px solid #ef9a9a; background: #fff5f5"):
                with gui.row().classes("items-center gap-1"):
                    gui.icon("edit_note", color="negative", size="22px")
                    gui.label("Escreva o motivo do cancelamento").classes("text-subtitle2 text-negative")
                motivo = gui.textarea(placeholder=f"Ex.: {edocs_http.MOTIVO_CANCELAR_RECEBIDO}") \
                    .props("outlined dense autogrow autofocus bg-color=white input-class=text-grey-10") \
                    .classes("w-full")
                with gui.row().classes("w-full items-center no-wrap"):
                    contador = gui.label("").classes("text-caption")
                    gui.space()
                    gui.button("Usar o texto padrão", icon="auto_fix_high",
                               on_click=lambda: motivo.set_value(edocs_http.MOTIVO_CANCELAR_RECEBIDO)) \
                        .props("flat dense no-caps size=sm color=negative")

            def palavras() -> int:
                return len((motivo.value or "").split())

            def contar() -> None:
                n = palavras()
                ok = n >= edocs_http.PALAVRAS_MOTIVO
                contador.text = (f"{n} palavra(s) · ok" if ok else
                                 f"{n} de {edocs_http.PALAVRAS_MOTIVO} palavras (o e-docs exige no mínimo "
                                 f"{edocs_http.PALAVRAS_MOTIVO})")
                contador.classes(replace="text-caption " + ("text-positive" if ok else "text-negative"))
                botao.set_enabled(ok)

            def confirmar() -> None:
                if palavras() < edocs_http.PALAVRAS_MOTIVO:
                    return
                janela.close()
                background_tasks.create(self._cancelar(docs, motivo.value))

            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Voltar", on_click=janela.close).props("flat no-caps")
                botao = gui.button("Cancelar documento(s)", icon="delete", color="negative", on_click=confirmar) \
                    .props("unelevated no-caps")
            motivo.on_value_change(lambda e: contar())
            contar()
        janela.open()

    async def _cancelar(self, docs: list[str], motivo: str) -> None:
        if self.trava.locked():
            self._avisar("Aguarde terminar o que está em andamento", type="warning")
            return
        async with self.trava:
            self._ocupado(True)
            erros = []
            try:
                for n, doc in enumerate(docs):
                    item = self.itens[doc]
                    self.tela.andamento(f"Cancelando: {self._titulo(doc)}", n, len(docs))
                    item["andamento"] = "Cancelando..."
                    self._atualizar(doc)
                    try:
                        await self._no_site(edocs_http.cancelar_recebido, doc, motivo)
                        item["feito"] = "Cancelado"
                        item["ciente"] = item["vistavel"] = False
                    except Exception as e:
                        erros.append(f"{self._titulo(doc)}: {e}")
                    finally:
                        item["andamento"] = ""
                        self.marcados.discard(doc)
                        self._atualizar(doc)
                self.tela.andamento("Pronto" if not erros else f"Pronto, com {len(erros)} problema(s)",
                                    len(docs), len(docs))
                await asyncio.sleep(1.2)
            finally:
                self._ocupado(False)
                self.tela.andamento(None)
                self._mostrar_marcados()
            feitos = len(docs) - len(erros)
            self.situacao.text = (f"{feitos} documento(s) cancelado(s)." if not erros else
                                  f"Cancelados {feitos}; problema(s): " + "; ".join(erros))
            self._avisar(f"{feitos} documento(s) cancelado(s)" if not erros else
                         f"{len(erros)} documento(s) não foram cancelados",
                         type="positive" if not erros else "warning")

    async def _em_lote(self, cientes: list[str], vistos: list[str]) -> None:
        if self.trava.locked():
            self._avisar("Aguarde terminar o que está em andamento", type="warning")
            return
        async with self.trava:
            self._ocupado(True)
            erros = []
            # Vistar tudo: primeiro os cientes que faltam, depois os vistos. A barra abaixo do
            # cabeçalho mostra a etapa, o documento da vez e quantos já foram
            tarefas = [(d, "ciente") for d in cientes] + [(d, "vistar") for d in vistos]
            for doc, _ in tarefas:
                self.itens[doc]["na_fila"] = True
                self._atualizar(doc)
            try:
                for n, (doc, acao) in enumerate(tarefas):
                    etapa = "Dando ciente" if acao == "ciente" else "Vistando"
                    if cientes and vistos:  # as duas etapas: diz em qual está
                        etapa = (f"1/2 · Dando ciente ({n + 1} de {len(cientes)})" if acao == "ciente" else
                                 f"2/2 · Vistando ({n - len(cientes) + 1} de {len(vistos)})")
                    self.tela.andamento(f"{etapa}: {self._titulo(doc)}", n, len(tarefas))
                    self.situacao.text = f"{etapa}: {n + 1} de {len(tarefas)}..."
                    try:
                        await self._fazer(doc, acao)
                    except Exception as e:
                        erros.append(f"{self._titulo(doc)}: {e}")
                self.tela.andamento("Pronto" if not erros else f"Pronto, com {len(erros)} problema(s)",
                                    len(tarefas), len(tarefas))
                await asyncio.sleep(1.2)  # deixa ver a barra cheia
            finally:
                for doc, _ in tarefas:
                    self.itens[doc]["na_fila"] = False
                    self._atualizar(doc)
                self._ocupado(False)
                self.tela.andamento(None)
            self.situacao.text = ("Pronto." if not erros else
                                  f"Pronto, com {len(erros)} problema(s): " + "; ".join(erros))
            self._avisar("Pronto" if not erros else f"{len(erros)} documento(s) não deram certo",
                         type="positive" if not erros else "warning")


def _caminho_html(passos: list[dict]) -> str:
    """O caminho do documento: uma linha no meio e as etapas alternando dos dois lados. Verde
    onde já passou, a etapa da vez em destaque (anel azul), cinza o que falta."""
    VERDE, CINZA, AZUL = "#16a34a", "#cfd6e0", "#2459a8"
    atual = next((i for i, p in enumerate(passos) if not p["feito"]), None)
    linhas = []
    for i, p in enumerate(passos):
        feito, da_vez = p["feito"], i == atual
        cor = VERDE if feito else AZUL if da_vez else CINZA
        # a linha de cima é verde se chegou até aqui; a de baixo, se a próxima etapa já foi
        cima = VERDE if feito or da_vez else CINZA
        baixo = VERDE if i + 1 < len(passos) and passos[i + 1]["feito"] else CINZA
        ponto = (f"background:{cor};border:2px solid {cor}" if feito else
                 f"background:#fff;border:3px solid {cor};box-shadow:0 0 0 4px rgba(36,89,168,.15)" if da_vez else
                 f"background:#fff;border:2px solid {cor}")
        texto = (f'<div style="font-size:10px;letter-spacing:.04em;color:#8a94a3">{html.escape(p["tipo"])}</div>'
                 f'<div class="t-forte" style="font-size:13px;font-weight:600">{html.escape(p["nome"].title())}</div>'
                 f'<div style="font-size:12px;color:{VERDE if feito else AZUL if da_vez else "#8a94a3"}">'
                 f'{html.escape(p["situacao"].capitalize())}</div>'
                 + (f'<div style="font-size:11px;color:#8a94a3">{html.escape(p["data"])}</div>' if p["data"] else ""))
        esquerda = i % 2 == 0
        lado = f'<div style="padding:6px 12px;text-align:{"right" if esquerda else "left"}">{texto}</div>'
        centro = (f'<div style="position:relative;display:flex;justify-content:center;align-items:center">'
                  f'<div style="position:absolute;left:50%;top:0;height:50%;width:2px;margin-left:-1px;'
                  f'background:{cima if i else "transparent"}"></div>'
                  f'<div style="position:absolute;left:50%;bottom:0;height:50%;width:2px;margin-left:-1px;'
                  f'background:{baixo if i + 1 < len(passos) else "transparent"}"></div>'
                  f'<div style="position:relative;width:14px;height:14px;border-radius:50%;{ponto}"></div></div>')
        vazio = "<div></div>"
        linhas.append((lado + centro + vazio) if esquerda else (vazio + centro + lado))
    return ('<div style="display:grid;grid-template-columns:1fr 34px 1fr;align-items:stretch">'
            + "".join(linhas) + "</div>")


# --- Cartão "Inteligência artificial" em Configurações ---------------------------------------------
class PainelIA:
    """Cartão da IA em Configurações: os provedores cadastrados (um em uso, cada um com a sua
    chave e o seu modelo) e o formulário para adicionar outro."""

    def __init__(self):
        self.disponiveis: dict[str, list[tuple[str, str]]] = {}  # provedor → modelos que a chave vê
        self.novo = ia.ANTHROPIC if not ia.provedores() else "DeepSeek"
        with gui.card().classes("cartao q-pa-lg w-full"):
            with gui.row().classes("items-center gap-2"):
                gui.icon("auto_awesome", color="primary", size="26px")
                gui.label("Inteligência artificial (resumos)").classes("text-h6")
            gui.label("Os documentos são enviados à IA em uso para resumir. Você pode cadastrar mais de "
                      "um provedor e escolher qual usar. As chaves ficam criptografadas neste computador, "
                      "como as senhas.").classes("text-caption t-suave")
            self.area = gui.column().classes("w-full gap-3")
        self._desenhar()

    def _desenhar(self) -> None:
        self.area.clear()
        lista = ia.provedores()
        em_uso = (ia.provedor() or {}).get("nome")
        with self.area:
            if not lista:
                with gui.row().classes("items-center gap-2"):
                    gui.element("span").classes("bolinha b-branco")
                    gui.label("Nenhuma IA configurada").classes("text-body2")
            for p in lista:
                self._provedor(p, p["nome"] == em_uso)
            gui.separator().classes("q-my-xs")
            self._formulario()

    def _provedor(self, p: dict, ativo: bool) -> None:
        nome = p["nome"]
        modelos = ia.modelos_instalados(nome)
        opcoes = {m: m for m in modelos}
        opcoes.update({m: (f"{rotulo} ({m})" if rotulo != m else m) for m, rotulo in self.disponiveis.get(nome, [])})
        with gui.element("div").classes("w-full rounded-borders q-pa-sm fundo-cartao") \
                .style("border: 2px solid #2459a8" if ativo else "border: 1px solid rgba(128,128,128,.3)"):
            with gui.row().classes("w-full items-center gap-2 no-wrap"):
                gui.button(icon="radio_button_checked" if ativo else "radio_button_unchecked",
                           on_click=lambda n=nome: self._usar(n)) \
                    .props(f"flat round dense color={'primary' if ativo else 'grey-6'}") \
                    .tooltip("Em uso" if ativo else "Usar esta IA")
                with gui.column().classes("gap-0 col"):
                    with gui.row().classes("items-center gap-2"):
                        gui.label(nome).classes("text-subtitle2 text-weight-medium t-forte")
                        if ativo:
                            gui.badge("em uso", color="primary").props("outline")
                    gui.label("Lê o PDF inteiro" if p.get("tipo") == "anthropic"
                              else f"Lê só o texto do PDF · {p.get('url', '')}") \
                        .classes("text-caption t-suave ellipsis")
                gui.button(icon="refresh", on_click=lambda n=nome: self._buscar_modelos(n)) \
                    .props("flat round dense").tooltip("Buscar os modelos que a chave pode usar")
                gui.button(icon="close", on_click=lambda n=nome: self._remover(n)) \
                    .props("flat round dense color=negative").tooltip("Remover esta IA (e a chave)")
            gui.select(opcoes, value=ia.modelo_escolhido(nome) or None, label="Modelo",
                       with_input=True, new_value_mode="add-unique",
                       on_change=lambda e, n=nome: self._escolher(n, e.value)) \
                .props('outlined dense options-dense hint="Escolha da lista (atualize com ↻) ou escreva o nome do modelo"') \
                .classes("w-full q-mt-xs")

    def _formulario(self) -> None:
        pronto = ia.MODELOS_PRONTOS.get(self.novo, ia.MODELOS_PRONTOS[ia.OUTRO])
        gui.label("Adicionar uma IA").classes("text-subtitle2 t-forte")
        with gui.row().classes("w-full items-start gap-2 no-wrap"):
            gui.select(list(ia.MODELOS_PRONTOS), value=self.novo, label="Provedor",
                       on_change=lambda e: self._trocar_novo(e.value)) \
                .props("outlined dense").classes("col")
            nome = gui.input("Nome (para você reconhecer)", value="" if self.novo == ia.OUTRO else self.novo) \
                .props("outlined dense").classes("col")
        url = None
        if pronto["tipo"] != "anthropic":
            url = gui.input("Endereço da API", value=pronto["url"],
                            placeholder="https://.../v1 (o endereço no formato da OpenAI)") \
                .props("outlined dense").classes("w-full")
            gui.label("Esta IA recebe só o texto do PDF: um PDF digitalizado (só imagem) não dá para "
                      "resumir.").classes("text-caption t-suave")
        with gui.row().classes("w-full items-start gap-2 no-wrap"):
            chave = gui.input("Chave da API" + (" (não precisa)" if pronto.get("sem_chave") else ""),
                              password=True, password_toggle_button=True) \
                .props("outlined dense").classes("col")
            botao = gui.button("Salvar e testar", icon="key").props("unelevated no-caps")
            botao.on_click(lambda: self._salvar(nome.value, pronto, url.value if url else "", chave.value, botao))

    def _trocar_novo(self, valor: str) -> None:
        self.novo = valor
        self._desenhar()

    async def _salvar(self, nome: str, pronto: dict, url: str, valor: str, botao) -> None:
        nome, valor, url = (nome or "").strip(), (valor or "").strip(), (url or "").strip()
        if not nome:
            gui.notify("Dê um nome para esta IA", type="warning")
            return
        if pronto["tipo"] != "anthropic" and not url:
            gui.notify("Escreva o endereço da API", type="warning")
            return
        if not valor and not pronto.get("sem_chave") and self.novo != ia.OUTRO:
            gui.notify("Cole a chave da API", type="warning")
            return
        botao.props("loading")
        try:
            modelos = await run.io_bound(ia.modelos_disponiveis, pronto["tipo"], url, valor)
        except Exception as e:
            botao.props(remove="loading")
            gui.notify(f"Não funcionou: {e or type(e).__name__}", type="negative", multi_line=True)
            return
        ia.salvar_provedor(nome, pronto["tipo"], url, valor)
        self.disponiveis[nome] = modelos
        falta = "" if ia.modelo_escolhido(nome) else " Escolha o modelo na lista."
        gui.notify(f"{nome} salvo e testado: {len(modelos)} modelo(s) disponíveis. Ela é a IA em uso agora."
                   + falta, type="positive", multi_line=True)
        self._desenhar()

    async def _buscar_modelos(self, nome: str) -> None:
        p = ia.provedor(nome)
        try:
            self.disponiveis[nome] = await run.io_bound(ia.modelos_disponiveis, p["tipo"], p.get("url", ""),
                                                        ia.chave(nome))
        except Exception as e:
            gui.notify(f"{nome}: {e or type(e).__name__}", type="negative", multi_line=True)
            return
        gui.notify(f"{nome}: {len(self.disponiveis[nome])} modelo(s) na lista", type="positive")
        self._desenhar()

    def _usar(self, nome: str) -> None:
        ia.usar_provedor(nome)
        gui.notify(f"IA em uso: {nome}", type="positive")
        self._desenhar()

    def _remover(self, nome: str) -> None:
        ia.remover_provedor(nome)
        self.disponiveis.pop(nome, None)
        gui.notify(f"{nome} removida", type="info")
        self._desenhar()

    def _escolher(self, nome: str, modelo) -> None:
        modelo = (modelo or "").strip()
        if not modelo:
            return
        ia.adicionar_modelo(nome, modelo)
        ia.escolher_modelo(nome, modelo)
        gui.notify(f"{nome}: modelo {modelo}", type="positive")
