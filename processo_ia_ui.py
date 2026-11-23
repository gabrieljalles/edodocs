"""Janela "Montar com IA" (Novo processo → Montar com IA).

1. A pessoa junta os atos que já fez (Word ou PDF) e, se quiser, explica o processo.
2. A IA monta (processo_ia.montar_com_ia), com o andamento na tela.
3. Se a IA ficou em dúvida, ela pergunta, uma pergunta por vez (um clique por resposta). Se
   alguma resposta não for a sugestão dela, a IA ajusta (processo_ia.ajustar_com_respostas).
4. O resultado abre no assistente, em "Como vai ficar", para a pessoa conferir e salvar.
"""
import tempfile
import time
from pathlib import Path

from nicegui import run
from nicegui import ui as gui

import assistente
import ia
import main
import processo_ia
from assistente import Assistente


def nome_livre(nome: str) -> str:
    """O nome sugerido pela IA, sem repetir o de um processo que já existe."""
    nome = " ".join(processo_ia.INVALIDOS.sub(" ", nome or "").split()).rstrip(".") or "Processo novo"
    candidato, n = nome, 2
    while main.problema_no_nome(candidato):
        candidato, n = f"{nome} {n}", n + 1
    return candidato


class JanelaIA:
    def __init__(self, tela):
        self.tela = tela
        self.arquivos: list[Path] = []
        self.trabalhando = False
        self.cancelado = False
        self.andamento = ""
        self.explicacao = ""
        self.desde = 0.0
        self.exemplos: list = []
        self.pasta: Path | None = None
        self.resultado = None
        self.perguntas: list[dict] = []
        self.respostas: list[dict] = []
        prov = ia.provedor() or {}
        self.provedor = prov.get("nome", "a IA")
        with tela.camada, gui.dialog().props("persistent") as self.janela, \
                gui.card().classes("cartao w-[760px] max-w-full q-pa-lg gap-3"):
            with gui.row().classes("items-center gap-2 no-wrap w-full"):
                gui.icon("auto_awesome", color="primary", size="30px")
                gui.label("Montar o processo com IA").classes("text-h6 t-forte")
                gui.space()
                gui.button(icon="close", on_click=self._fechar).props("flat round dense")
            self.area = gui.column().classes("w-full gap-3")
        self._desenhar_escolha()
        gui.timer(1.0, self._atualizar_andamento)
        self.janela.open()

    # --- 1. Os arquivos ---------------------------------------------------------------------
    def _desenhar_escolha(self) -> None:
        self.area.clear()
        with self.area:
            gui.label("Junte atos que você já fez deste processo. A IA estuda o que muda de um para o outro "
                      "e monta o Word, a planilha e as regras. No fim você confere tudo antes de salvar.") \
                .classes("text-body1")
            with gui.column().classes("gap-0 q-pl-sm"):
                for dica in ("Quanto mais variações, melhor: um de homem e um de mulher, um de uma pessoa e um "
                             "de várias, um com e um sem considerando...",
                             "De 2 a 10 atos costuma bastar. Precisa de pelo menos um Word (.docx ou .doc); "
                             "PDFs ajudam a mostrar outras variações."):
                    gui.label("• " + dica).classes("text-body2 t-suave")
            with gui.row().classes("items-center gap-3"):
                gui.button("Juntar arquivos...", icon="attach_file", on_click=self._juntar) \
                    .props("unelevated no-caps")
                gui.label(f"{len(self.arquivos)} arquivo(s)" if self.arquivos else "Nenhum arquivo ainda") \
                    .classes("text-body2 t-suave")
            if self.arquivos:
                with gui.column().classes("w-full gap-0 fundo-cartao rounded q-pa-sm"):
                    for caminho in self.arquivos:
                        with gui.row().classes("w-full items-center gap-2 no-wrap"):
                            gui.icon("picture_as_pdf" if caminho.suffix.lower() == ".pdf" else "description",
                                     color="primary", size="20px")
                            gui.label(caminho.name).classes("text-body2 t-forte col ellipsis")
                            gui.button(icon="close", on_click=lambda c=caminho: self._tirar(c)) \
                                .props("flat round dense size=sm").tooltip("Tirar da lista")
            gui.textarea(
                "Quer explicar alguma coisa? (opcional)", value=self.explicacao,
                on_change=lambda e: setattr(self, "explicacao", e.value or ""),
                placeholder="Ex.: quando a pessoa tem processo administrativo, entra o considerando do processo. "
                            "Com mais de uma pessoa, vai a tabela do anexo.") \
                .props("outlined autogrow").classes("w-full")
            with gui.row().classes("items-start gap-2 no-wrap fundo-cartao rounded q-pa-sm"):
                gui.icon("privacy_tip", color="orange-8", size="22px")
                gui.label(f"O texto destes atos (nomes, matrículas, CPF...) vai para {self.provedor}, a IA "
                          "escolhida em Configurações. Use só se a sua equipe permite enviar esses dados.") \
                    .classes("text-body2 t-forte")
            rascunho = assistente.ler_rascunho()
            if rascunho:
                with gui.row().classes("items-start gap-2 no-wrap"):
                    gui.icon("warning", color="orange-8", size="22px")
                    gui.label(f"Você tem o processo “{rascunho.get('nome') or 'sem nome'}” pela metade. Ele será "
                              "substituído pelo que a IA montar.").classes("text-body2 t-suave")
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=self._fechar).props("flat no-caps")
                gui.button("Montar com IA", icon="auto_awesome", on_click=self._montar) \
                    .props("unelevated no-caps").set_enabled(bool(self.arquivos))

    async def _juntar(self) -> None:
        novos = await self.tela._escolher_arquivos_do_computador("Word ou PDF", processo_ia.EXTENSOES)
        for caminho in novos:
            if caminho.suffix.lower() not in processo_ia.EXTENSOES:
                gui.notify(f"{caminho.name}: só Word (.docx, .doc) ou PDF", type="warning")
            elif caminho.name.startswith("~$"):
                continue  # arquivo temporário do Word aberto
            elif caminho not in self.arquivos:
                self.arquivos.append(caminho)
        if len(self.arquivos) > processo_ia.MAX_ARQUIVOS:
            gui.notify(f"Use até {processo_ia.MAX_ARQUIVOS} arquivos: fiquei com os primeiros", type="warning")
            self.arquivos = self.arquivos[:processo_ia.MAX_ARQUIVOS]
        self._desenhar_escolha()

    def _tirar(self, caminho: Path) -> None:
        self.arquivos.remove(caminho)
        self._desenhar_escolha()

    # --- 2. A IA trabalhando ------------------------------------------------------------------
    def _avisar(self, texto: str) -> None:
        self.andamento = texto  # vem de outra thread: o timer mostra

    def _atualizar_andamento(self) -> None:
        if self.trabalhando and hasattr(self, "lbl_andamento"):
            self.lbl_andamento.text = self.andamento
            segundos = int(time.monotonic() - self.desde)
            self.lbl_tempo.text = f"{segundos // 60} min {segundos % 60:02d} s"

    def _desenhar_andamento(self, texto: str) -> None:
        self.trabalhando, self.cancelado, self.desde = True, False, time.monotonic()
        self.andamento = texto
        self.area.clear()
        with self.area:
            with gui.row().classes("items-center gap-3 no-wrap w-full"):
                gui.spinner("dots", size="lg", color="primary")
                with gui.column().classes("gap-0 col"):
                    self.lbl_andamento = gui.label(self.andamento).classes("text-subtitle1 t-forte")
                    self.lbl_tempo = gui.label("").classes("text-caption t-suave")
            gui.linear_progress(show_value=False).props("indeterminate rounded color=primary").classes("w-full")
            gui.label("Costuma levar de 2 a 10 minutos: a IA estuda cada ato e o Edodocs confere o que ela "
                      "montou refazendo cada exemplo. Pode deixar esta janela aberta e fazer outra coisa.") \
                .classes("text-body2 t-suave")
            with gui.row().classes("w-full justify-end"):
                gui.button("Cancelar", on_click=self._fechar).props("flat no-caps")

    async def _montar(self) -> None:
        self._desenhar_andamento("Lendo os arquivos...")
        self.pasta = Path(tempfile.mkdtemp(prefix="edodocs_ia_"))
        try:
            self.exemplos = await run.io_bound(processo_ia.ler_arquivos, self.arquivos, self.pasta)
            resultado = await run.io_bound(processo_ia.montar_com_ia, self.exemplos, self.pasta / "processo",
                                           self.explicacao, self._avisar)
        except Exception as e:
            self.trabalhando = False
            if not self.cancelado:
                self._desenhar_erro(e)
            return
        self.trabalhando = False
        if self.cancelado:
            return
        self.resultado, self.respostas = resultado, []
        self.perguntas = processo_ia.perguntas_do_plano(resultado.plano)
        if self.perguntas:
            self._desenhar_pergunta(0)
        else:
            self._abrir(resultado)

    def _abrir(self, resultado, respostas: list[dict] | None = None) -> None:
        self.janela.close()
        dados = processo_ia.rascunho_do_resultado(resultado, nome_livre(resultado.plano.get("nome", "")),
                                                  respostas)
        assistente.apagar_rascunho()
        Assistente(self.tela, rascunho=dados)

    # --- 3. As perguntas da IA (uma por vez) ---------------------------------------------------
    def _desenhar_pergunta(self, i: int) -> None:
        total = len(self.perguntas)
        p = self.perguntas[i]
        self.area.clear()
        with self.area:
            with gui.row().classes("items-center gap-2 no-wrap w-full"):
                gui.icon("quiz", color="primary", size="26px")
                gui.label("A IA montou o processo. Antes de terminar, ela quer confirmar algumas coisas com você"
                          if i == 0 else "Mais uma pergunta").classes("text-subtitle1 t-forte col")
                gui.label(f"{i + 1} de {total}").classes("text-body2 t-suave")
            gui.linear_progress(value=i / total, show_value=False).props("rounded color=primary") \
                .classes("w-full")
            if i == 0:
                gui.label("Responda com um clique. Não tem resposta errada: se não souber, clique em "
                          "“Não sei” e fica como a IA sugeriu.").classes("text-body2 t-suave")
            gui.label(p["pergunta"]).classes("text-h6 t-forte")
            if p["por_que"]:
                with gui.row().classes("w-full items-start gap-2 no-wrap"):
                    gui.icon("lightbulb", color="orange-8", size="20px").classes("q-mt-xs")
                    gui.label(p["por_que"]).classes("text-body1 t-forte")
            if p["contexto"]:
                with gui.column().classes("w-full gap-0 fundo-cartao rounded q-pa-sm"):
                    gui.label("No ato que você mandou:").classes("text-caption text-weight-bold t-suave")
                    gui.label(p["contexto"]).classes("text-body2 t-forte").style("font-style: italic")
            with gui.column().classes("w-full gap-2"):
                for j, opcao in enumerate(p["opcoes"]):
                    with gui.button(on_click=lambda o=opcao["texto"], s=(j == 0): self._responder(i, o, s)) \
                            .props("outline no-caps align=left").classes("w-full q-py-sm"):
                        with gui.row().classes("items-start gap-2 no-wrap w-full"):
                            gui.icon("auto_awesome" if j == 0 else "radio_button_unchecked", size="20px") \
                                .classes("q-mt-xs")
                            with gui.column().classes("gap-0 col items-start"):
                                gui.label(opcao["texto"]).classes("text-body1 text-left")
                                if opcao["como_fica"]:
                                    gui.label("O ato fica: " + opcao["como_fica"]) \
                                        .classes("text-body2 text-left t-suave").style("font-style: italic")
                            if j == 0:
                                gui.badge("sugestão da IA").props("outline color=primary")
            with gui.row().classes("w-full items-center gap-2"):
                outra = gui.input(placeholder="Ou escreva outra resposta...").props("outlined dense") \
                    .classes("col")
                gui.button("Responder", on_click=lambda: self._responder(i, outra.value or "", False)) \
                    .props("flat no-caps").bind_enabled_from(outra, "value", backward=lambda v: bool((v or "").strip()))
            with gui.row().classes("w-full items-center gap-2"):
                if i > 0:
                    gui.button("Voltar", icon="arrow_back", on_click=lambda: self._desenhar_pergunta(i - 1)) \
                        .props("flat no-caps")
                gui.space()
                gui.button("Não sei: deixe como a IA sugeriu",
                           on_click=lambda: self._responder(i, p["opcoes"][0]["texto"], True, nao_sei=True)) \
                    .props("flat no-caps")
                if total - i > 1:
                    gui.button("Pular as perguntas", on_click=self._pular).props("flat no-caps") \
                        .tooltip("Fica tudo como a IA sugeriu")

    async def _responder(self, i: int, texto: str, sugerida: bool, nao_sei: bool = False) -> None:
        texto = texto.strip()
        if not texto:
            return
        p = self.perguntas[i]
        resposta = f"Não sei (ficou: {texto})" if nao_sei else texto
        self.respostas = self.respostas[:i] + [{"pergunta": p["pergunta"], "resposta": resposta,
                                                "sugerida": sugerida}]
        if i + 1 < len(self.perguntas):
            self._desenhar_pergunta(i + 1)
        else:
            await self._ajustar()

    async def _pular(self) -> None:
        for p in self.perguntas[len(self.respostas):]:
            self.respostas.append({"pergunta": p["pergunta"], "resposta": f"Não sei (ficou: {p['opcoes'][0]['texto']})",
                                   "sugerida": True})
        await self._ajustar()

    async def _ajustar(self) -> None:
        if not processo_ia.precisa_ajustar(self.respostas):  # o plano já é assim: nada a pedir
            self._abrir(self.resultado, self.respostas)
            return
        self._desenhar_andamento("A IA está ajustando o processo com as suas respostas...")
        try:
            novo = await run.io_bound(processo_ia.ajustar_com_respostas, self.resultado, self.exemplos,
                                      self.respostas, self.pasta / "ajustado", self.explicacao, self._avisar)
        except Exception as e:
            self.trabalhando = False
            if not self.cancelado:
                self._desenhar_erro(e, ajuste=True)
            return
        self.trabalhando = False
        if not self.cancelado:
            self._abrir(novo, self.respostas)

    def _desenhar_erro(self, erro: Exception, ajuste: bool = False) -> None:
        if isinstance(erro, ia.SemChave):
            texto = "Configure a IA em Configurações → Inteligência artificial."
        elif isinstance(erro, (ia.ErroIA, RuntimeError)):
            texto = str(erro)
        else:
            texto = f"{type(erro).__name__}: {erro}"
        self.area.clear()
        with self.area:
            with gui.row().classes("items-start gap-2 no-wrap"):
                gui.icon("error", color="negative", size="28px")
                with gui.column().classes("gap-1"):
                    gui.label("Não deu para montar o processo").classes("text-subtitle1 t-forte")
                    gui.label(texto[:1500]).classes("text-body2 t-forte break-all")
                    gui.label("Tente de novo (às vezes a IA acerta na segunda vez), junte outros exemplos ou "
                              "monte você mesmo pelo passo a passo.").classes("text-body2 t-suave")
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Fechar", on_click=self._fechar).props("flat no-caps")
                if ajuste:  # a 1ª versão continua boa: dá para seguir com ela
                    gui.button("Ficar com a 1ª versão", icon="undo",
                               on_click=lambda: self._abrir(self.resultado, None)).props("outline no-caps")
                    gui.button("Tentar de novo", icon="refresh", on_click=self._ajustar) \
                        .props("unelevated no-caps")
                    return
                gui.button("Voltar aos arquivos", icon="arrow_back", on_click=self._desenhar_escolha) \
                    .props("outline no-caps")
                gui.button("Tentar de novo", icon="refresh", on_click=self._montar).props("unelevated no-caps")

    def _fechar(self) -> None:
        if self.trabalhando:
            self.cancelado = True  # a resposta que chegar depois é descartada
            gui.notify("Cancelado. Se a IA já estava respondendo, o pedido pode ser cobrado.", type="info")
        self.janela.close()
