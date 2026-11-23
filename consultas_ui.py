"""Telas das tabelas de consulta (consultas.py):

  - PassoAPassoConsulta: liga um processo pronto a uma tabela (ex.: Servidores), em 4 telas:
    qual tabela, como achar a pessoa, o que puxar e um teste. Quase tudo vem sugerido.
  - BuscaNaTabela: na janela da linha, a caixa "Procurar..." que preenche o resto da linha
    com os dados da pessoa escolhida (em verde: veio da tabela).
"""
import asyncio
import tempfile
from datetime import date
from pathlib import Path

from nicegui import app as napp
from nicegui import run
from nicegui import ui as gui

import atos
import consultas
import de_para
import main
import planilha_edicao
import processos

VERDE = "veio-da-tabela"  # classe do app.CSS: verde que se ajusta ao tema (dia/noite)


def planilha_de(pasta: Path) -> Path | None:
    """A planilha que o processo usa (a "pessoal" escolhida no app, ou a do processo)."""
    pessoal = main.ler_config().get("planilhas", {}).get(pasta.name)
    if pessoal:
        return Path(pessoal) if Path(pessoal).is_file() else None
    planilha = processos.carregar(pasta).planilha_padrao()
    return planilha if planilha and planilha.is_file() else None


def resumo(consulta: dict | None) -> str:
    """"Servidores · acha por NOME ou MATRÍCULA · preenche 8 colunas" (para o cartão do processo)."""
    if not consulta or not consulta.get("tabela"):
        return ""
    chaves = " ou ".join(c["minha"] for c in consulta.get("chaves") or [])
    return (f"{consulta['tabela']} · acha por {chaves or '?'} · preenche "
            f"{len(consulta.get('puxar') or [])} coluna(s)")


async def escolher_arquivo(camada, rotulo: str, extensoes: tuple[str, ...]) -> Path | None:
    """Janela de abrir arquivo do Windows; no navegador, envio do arquivo (fica uma cópia)."""
    if napp.native.main_window is not None:
        import webview
        tipo = webview.FileDialog.OPEN if hasattr(webview, "FileDialog") else webview.OPEN_DIALOG
        escolhidos = await napp.native.main_window.create_file_dialog(
            tipo, file_types=(f"{rotulo} ({';'.join('*' + e for e in extensoes)})",))
        return Path(escolhidos[0]) if escolhidos else None
    futuro = asyncio.get_running_loop().create_future()
    with camada, gui.dialog() as janela, gui.card().classes("cartao q-pa-lg"):
        gui.label(f"Envie o arquivo ({', '.join(extensoes)})").classes("text-h6")

        async def recebido(e) -> None:
            destino = Path(tempfile.mkdtemp()) / e.file.name
            await e.file.save(destino)
            janela.close()
            if not futuro.done():
                futuro.set_result(destino)

        gui.upload(on_upload=recebido, auto_upload=True, max_files=1).props(f"accept={','.join(extensoes)}")
        gui.button("Cancelar", on_click=lambda: (janela.close(), futuro.done() or futuro.set_result(None))) \
            .props("flat")
    janela.open()
    return await futuro


def campo_prazo(valor: int):
    """A pergunta "de quanto em quanto tempo atualizar" (em dias; 0 = nunca pedir)."""
    with gui.column().classes("w-full gap-0"):
        prazo = gui.number("Pedir para atualizar os dados a cada", value=valor, min=0, max=365, step=1,
                           format="%d", suffix="dias") \
            .props('outlined stack-label bg-color=white input-class=text-grey-10 label-color=grey-8') \
            .classes("w-full")
        with gui.row().classes("gap-1 items-center"):
            gui.label("Atalhos:").classes("text-caption t-suave")
            for dias, rotulo in ((7, "toda semana"), (15, "a cada 15 dias"), (30, "todo mês"), (0, "nunca pedir")):
                gui.chip(rotulo, on_click=lambda d=dias: prazo.set_value(d)) \
                    .props("dense outline clickable color=primary")
        gui.label("Passado esse prazo, ao incluir uma linha, o Edodocs pede a versão mais nova da tabela "
                  "(ex.: exportada de novo do sistema).").classes("text-caption t-suave")
    return prazo


def texto_prazo(dados: dict) -> tuple[str, bool]:
    """("Atualizar a cada 7 dias · atualizada em 20/10/2026", vencida?) para mostrar na tela."""
    prazo = int(dados.get("atualizar_dias") or 0)
    quando = consultas.atualizada_em(dados)
    data = f"atualizada em {quando:%d/%m/%Y}" if quando else "sem data de atualização"
    if not prazo:
        return f"Não pede atualização · {data}", False
    vencida = bool(quando) and (date.today() - quando).days >= prazo
    return f"Atualizar a cada {prazo} dia(s) · {data}" + (" · está na hora de atualizar" if vencida else ""), \
        vencida


async def trocar_arquivo(camada, nome: str) -> bool:
    """Escolhe o arquivo novo da tabela (mesmas colunas, dados mais novos). True: trocou."""
    dados = consultas.tabelas().get(nome)
    caminho = await escolher_arquivo(camada, "Tabela", (".xlsx", ".xlsm", ".csv"))
    if not caminho or not dados:
        return False
    try:
        nomes_abas = await run.io_bound(consultas.abas, caminho)
        aba = dados.get("aba", "") if dados.get("aba") in nomes_abas else (nomes_abas[0] if nomes_abas else "")
        nova = await run.io_bound(consultas.ler, caminho, aba)
    except Exception as e:
        gui.notify(f"Não consegui abrir a tabela ({type(e).__name__}). Ela está aberta no Excel?", type="negative")
        return False
    if dados.get("arquivo") and Path(dados["arquivo"]).is_file():
        try:
            antiga = await run.io_bound(consultas.ler, dados["arquivo"], dados.get("aba", ""))
            faltam = [c for c in antiga.colunas if c not in nova.colunas]
            if faltam:
                gui.notify(f"Atenção: a tabela nova não tem as colunas {', '.join(faltam[:5])}. Confira em "
                           "Processos → Preencher sozinho.", type="warning", timeout=8000)
        except Exception:
            pass  # a antiga não abre mais: fica a nova
    consultas.salvar_tabela(nome, caminho, aba, dados.get("atualizar_dias") or 0)
    gui.notify(f"Tabela {nome} atualizada: {len(nova.linhas)} pessoa(s)", type="positive")
    return True


async def achar_tabela(camada, nome: str) -> bool:
    """A tabela que um processo usa não está neste computador (veio de outro, ou o arquivo
    mudou de lugar): a pessoa escolhe o Excel e ela fica cadastrada aqui. True: achou."""
    caminho = await escolher_arquivo(camada, "Tabela", (".xlsx", ".xlsm", ".csv"))
    if not caminho:
        return False
    dados = consultas.tabelas().get(nome) or {}
    try:
        nomes_abas = await run.io_bound(consultas.abas, caminho)
        aba = dados.get("aba", "") if dados.get("aba") in nomes_abas else (nomes_abas[0] if nomes_abas else "")
        tabela = await run.io_bound(consultas.ler, caminho, aba)
    except Exception as e:
        gui.notify(f"Não consegui abrir a tabela ({type(e).__name__}). Ela está aberta no Excel?", type="negative")
        return False
    consultas.salvar_tabela(nome, caminho, aba, dados.get("atualizar_dias") or 0)
    gui.notify(f"Tabela {nome} achada: {len(tabela.linhas)} pessoa(s)", type="positive")
    return True


def editar_tabela(camada, nome: str, ao_mudar=None) -> None:
    """Prazo de atualização e troca do arquivo de uma tabela já cadastrada."""
    dados = consultas.tabelas().get(nome) or {}
    ao_mudar = ao_mudar or (lambda: None)
    with camada, gui.dialog() as janela, gui.card().classes("cartao w-[520px] q-pa-lg gap-3"):
        gui.label(f"Tabela {nome}").classes("text-h6 t-forte")
        gui.label(texto_prazo(dados)[0]).classes("text-body2 t-suave")
        prazo = campo_prazo(int(dados.get("atualizar_dias") or 0))

        async def atualizar_agora() -> None:
            if await trocar_arquivo(camada, nome):
                consultas.ajustar_tabela(nome, atualizar_dias=int(prazo.value or 0))
                janela.close()
                ao_mudar()

        def guardar() -> None:
            consultas.ajustar_tabela(nome, atualizar_dias=max(0, int(prazo.value or 0)))
            janela.close()
            ao_mudar()

        gui.button("Atualizar agora (escolher o arquivo novo)...", icon="update", on_click=atualizar_agora) \
            .props("outline no-caps").classes("w-full")
        with gui.row().classes("w-full justify-end gap-2"):
            gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
            gui.button("Guardar", icon="check", on_click=guardar).props("unelevated no-caps")
    janela.open()


async def pedir_atualizacao(camada, nome: str) -> None:
    """Na Nova linha: se a tabela passou do prazo, pede a versão mais nova antes de procurar."""
    dias = consultas.dias_sem_atualizar(nome)
    if dias is None:
        return
    prazo = int((consultas.tabelas().get(nome) or {}).get("atualizar_dias") or 0)
    futuro = asyncio.get_running_loop().create_future()

    def fechar() -> None:
        janela.close()
        if not futuro.done():
            futuro.set_result(None)

    async def escolher() -> None:
        if await trocar_arquivo(camada, nome):
            fechar()

    def em_dia() -> None:
        consultas.marcar_atualizada(nome)
        gui.notify(f"Certo: pergunto de novo daqui a {prazo} dia(s)", type="info")
        fechar()

    with camada, gui.dialog().props("persistent") as janela, gui.card().classes("cartao w-[560px] q-pa-lg gap-3"):
        with gui.row().classes("items-center gap-2 no-wrap"):
            gui.icon("update", color="orange-8", size="32px")
            gui.label(f"Hora de atualizar a tabela {nome}").classes("text-h6 t-forte")
        gui.label(f"Os dados foram atualizados há {dias} dia(s), e o combinado é atualizar a cada {prazo} dia(s). "
                  "Pegue a versão mais nova (ex.: exporte de novo do sistema) e escolha o arquivo: assim a "
                  "linha vem com os dados certos.").classes("text-body1 t-forte")
        with gui.column().classes("w-full gap-2"):
            gui.button("Escolher a tabela nova...", icon="upload_file", on_click=escolher) \
                .props("unelevated no-caps").classes("w-full")
            gui.button("Os dados já estão em dia", icon="check", on_click=em_dia) \
                .props("outline no-caps").classes("w-full")
            gui.button("Agora não (pergunto na próxima linha)", on_click=fechar).props("flat no-caps").classes("w-full")
    janela.open()
    await futuro


def _titulo_passo(numero: int, titulo: str, explicacao: str) -> None:
    with gui.row().classes("items-start gap-3 no-wrap w-full"):
        gui.label(str(numero)).classes("text-white text-weight-bold text-center text-subtitle1") \
            .style("background:#2459a8;border-radius:50%;width:30px;height:30px;line-height:30px;flex-shrink:0")
        with gui.column().classes("gap-0"):
            gui.label(titulo).classes("text-h6 t-forte")
            gui.label(explicacao).classes("text-body2 t-suave")


class PassoAPassoConsulta:
    """Liga o processo (já com Word e planilha prontos) a uma tabela de consulta."""

    def __init__(self, camada, pasta: Path, ao_salvar=None):
        self.camada = camada
        self.pasta = pasta
        self.ao_salvar = ao_salvar or (lambda: None)
        self.proc = processos.carregar(pasta)
        atual = self.proc.config.get("consulta") or {}
        self.nome_tabela: str | None = atual.get("tabela") if atual.get("tabela") in consultas.tabelas() else None
        self.chaves: list[dict] = [dict(c) for c in atual.get("chaves") or []]
        self.puxar: dict[str, dict] = {p["minha"]: dict(p) for p in atual.get("puxar") or []}
        self.ja_tinha = bool(atual)
        self.tabela: consultas.Tabela | None = None
        self.minhas: list[str] = []
        self.opcoes: dict[str, list[str]] = {
            c["nome"]: c["opcoes"] for c in self.proc.colunas_planilha_nova() if c["tipo"] == "lista"}
        self._montar()

    # --- Janela ---
    def _montar(self) -> None:
        with self.camada, gui.dialog().props("persistent") as self.janela, \
                gui.card().classes("cartao w-[900px] max-w-full q-pa-none gap-0"):
            with gui.row().classes("w-full items-center q-px-lg q-py-md text-white no-wrap").style(
                    "background: linear-gradient(90deg,#1d4a8f,#2f6fcf);"):
                gui.icon("auto_fix_high", size="26px")
                gui.label(f"Preencher sozinho · {self.pasta.name}").classes("text-h6 ellipsis")
                gui.space()
                gui.button(icon="close", on_click=self.janela.close).props("flat round color=white")
            with gui.scroll_area().classes("w-full h-[560px]"):
                with gui.stepper().props("animated flat header-nav").classes("w-full bg-transparent") as self.passos:
                    with gui.step("tabela", title="A tabela", icon="table_view"):
                        self._passo_tabela()
                    with gui.step("chaves", title="Achar a pessoa", icon="person_search"):
                        self._passo_chaves()
                    with gui.step("puxar", title="O que vem junto", icon="input"):
                        self._passo_puxar()
                    with gui.step("teste", title="Testar", icon="science"):
                        self._passo_teste()
        self.janela.open()
        self._desenhar_tabelas()
        if self.nome_tabela:
            gui.timer(0.1, self._carregar_tabela, once=True)

    # --- 1. A tabela ---
    def _passo_tabela(self) -> None:
        _titulo_passo(1, "De onde vêm os dados?",
                      "Escolha a tabela com os dados das pessoas (ex.: a lista de todos os servidores). "
                      "Ela fica guardada e serve para os outros processos também.")
        self.area_tabelas = gui.column().classes("w-full gap-2 q-mt-md")
        with gui.row().classes("items-center gap-2 q-mt-sm"):
            gui.button("Adicionar uma tabela do computador...", icon="upload_file",
                       on_click=self._adicionar_tabela).props("outline no-caps")
            gui.label("Excel (.xlsx) ou .csv").classes("text-caption t-suave")
        self.area_previa = gui.column().classes("w-full gap-1 q-mt-md")
        with gui.stepper_navigation():
            self.botao_tabela = gui.button("Próximo", icon="arrow_forward", on_click=self._ir_chaves) \
                .props("unelevated no-caps")

    def _desenhar_tabelas(self) -> None:
        self.area_tabelas.clear()
        todas = consultas.tabelas()
        with self.area_tabelas:
            if not todas:
                with gui.row().classes("items-center gap-2 fundo-cartao rounded q-pa-md w-full"):
                    gui.icon("info", color="primary")
                    gui.label("Nenhuma tabela ainda. Clique em Adicionar e escolha o Excel com as pessoas.") \
                        .classes("text-body2 t-forte")
                return
            for nome, dados in todas.items():
                escolhida = nome == self.nome_tabela
                with gui.card().classes("cartao w-full q-pa-sm cursor-pointer").style(
                        "outline: 3px solid var(--q-primary)" if escolhida else "") \
                        .on("click", lambda n=nome: self._escolher_tabela(n)):
                    with gui.row().classes("w-full items-center gap-3 no-wrap"):
                        gui.icon("radio_button_checked" if escolhida else "radio_button_unchecked",
                                 color="primary", size="22px")
                        gui.icon("table_view", color="grey-7", size="26px")
                        with gui.column().classes("gap-0 col"):
                            gui.label(nome).classes("text-subtitle1 text-weight-medium t-forte")
                            existe = Path(dados["arquivo"]).is_file()
                            gui.label(dados["arquivo"] if existe else f"Arquivo não encontrado: {dados['arquivo']}") \
                                .classes("text-caption break-all " + ("t-suave" if existe else "text-negative"))
                            texto, vencida = texto_prazo(dados)
                            gui.label(texto).classes("text-caption " + ("text-orange-9 text-weight-bold" if vencida
                                                                         else "t-suave"))
                        gui.button(icon="update", on_click=lambda n=nome: editar_tabela(
                            self.camada, n, self._recarregar_tabelas)) \
                            .props("flat round dense color=primary") \
                            .tooltip("De quanto em quanto tempo atualizar · trocar pelo arquivo novo") \
                            .on("click.stop", lambda: None)
                        gui.button(icon="delete", on_click=lambda n=nome: self._remover_tabela(n)) \
                            .props("flat round dense color=negative").tooltip("Tirar esta tabela da lista") \
                            .on("click.stop", lambda: None)

    async def _adicionar_tabela(self) -> None:
        caminho = await escolher_arquivo(self.camada, "Tabela", (".xlsx", ".xlsm", ".csv"))
        if not caminho:
            return
        try:
            nomes_abas = await run.io_bound(consultas.abas, caminho)
        except Exception as e:
            gui.notify(f"Não consegui abrir a tabela ({type(e).__name__}). Ela está aberta no Excel?",
                       type="negative")
            return
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[480px] q-pa-lg gap-3"):
            gui.label("Nova tabela").classes("text-h6 t-forte")
            gui.label(caminho.name).classes("text-body2 t-suave break-all")
            nome = gui.input("Que nome ela vai ter?", value=caminho.stem.strip().title()[:40],
                             placeholder="Ex.: Servidores") \
                .props('outlined stack-label bg-color=white input-class=text-grey-10 label-color=grey-8 autofocus') \
                .classes("w-full")
            aba = gui.select(nomes_abas, value=nomes_abas[0], label="Qual folha da planilha?") \
                .props("outlined dense").classes("w-full") if len(nomes_abas) > 1 else None
            gui.label("De quanto em quanto tempo os dados dessa tabela mudam?").classes("text-body2 t-forte q-mt-sm")
            prazo = campo_prazo(7)

            def guardar() -> None:
                escolhido = " ".join(str(nome.value or "").split())
                if not escolhido:
                    gui.notify("Dê um nome para a tabela (ex.: Servidores)", type="warning")
                    return
                consultas.salvar_tabela(escolhido, caminho, aba.value if aba else "", int(prazo.value or 0))
                janela.close()
                self._escolher_tabela(escolhido)

            nome.on("keydown.enter", guardar)
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Guardar", icon="check", on_click=guardar).props("unelevated no-caps")
        janela.open()

    def _recarregar_tabelas(self) -> None:
        """Mudou o prazo ou o arquivo de uma tabela: redesenha (e relê a escolhida)."""
        self._desenhar_tabelas()
        if self.nome_tabela:
            gui.timer(0.05, self._carregar_tabela, once=True)

    def _remover_tabela(self, nome: str) -> None:
        consultas.remover_tabela(nome)
        if self.nome_tabela == nome:
            self.nome_tabela, self.tabela = None, None
            self.area_previa.clear()
        self._desenhar_tabelas()

    def _escolher_tabela(self, nome: str) -> None:
        if nome not in consultas.tabelas():  # o clique foi no lixo da tabela (acabou de sair)
            return
        self.nome_tabela = nome  # as ligações ficam: _carregar_tabela tira as de colunas que ela não tem
        self._desenhar_tabelas()
        gui.timer(0.05, self._carregar_tabela, once=True)

    async def _carregar_tabela(self) -> None:
        self.area_previa.clear()
        with self.area_previa, gui.row().classes("items-center gap-2"):
            gui.spinner()
            gui.label("Lendo a tabela e a planilha do processo...").classes("t-suave")
        planilha = planilha_de(self.pasta)
        self.botao_tabela.disable()  # Próximo só depois de ler (as colunas vêm daqui)
        try:
            self.tabela = await run.io_bound(consultas.ler_cadastrada, self.nome_tabela)
            colunas = await run.io_bound(planilha_edicao.esquema, planilha, self.proc.aba) if planilha else []
        except Exception as e:
            self.botao_tabela.enable()
            self.area_previa.clear()
            with self.area_previa:
                gui.label(f"Não consegui ler a tabela ({type(e).__name__}). Ela está aberta no Excel ou mudou "
                          "de lugar?").classes("text-negative")
            self.tabela = None
            return
        self.minhas = [c.nome for c in colunas if c.editavel]
        self.botao_tabela.enable()
        # A mesma tabela escolhida de novo (ou outra com as mesmas colunas) mantém o que já estava ligado
        self.chaves = [c for c in self.chaves if c.get("dela") in self.tabela.colunas]
        self.puxar = {m: p for m, p in self.puxar.items() if p.get("dela") in self.tabela.colunas}
        if not self.chaves:
            self.chaves = consultas.sugerir_chaves(self.minhas, self.tabela)
        if not self.puxar:
            self.puxar = {p["minha"]: p for p in
                          consultas.sugerir_puxar(self.minhas, self.tabela, self.chaves, self.opcoes)}
        self.area_previa.clear()
        with self.area_previa:
            gui.label(f"{len(self.tabela.linhas)} pessoa(s) · {len(self.tabela.colunas)} coluna(s). As primeiras:") \
                .classes("text-body2 t-forte")
            mostrar = self.tabela.colunas[:8]
            # Larga demais, rola só a prévia (os botões dos cartões das tabelas ficam à vista)
            with gui.element("div").classes("w-full").style("overflow-x: auto; max-width: 820px"):
                gui.table(rows=[{c: l.get(c, "") for c in mostrar} for l in self.tabela.linhas[:3]],
                          columns=[{"name": c, "label": c, "field": c, "align": "left"} for c in mostrar]) \
                    .props("dense flat bordered")
            if len(self.tabela.colunas) > 8:
                gui.label(f"e mais {len(self.tabela.colunas) - 8} coluna(s)").classes("text-caption t-suave")
        self._desenhar_chaves()
        self._desenhar_puxar()

    def _ir_chaves(self) -> None:
        if not self.tabela:
            gui.notify("Escolha (ou adicione) a tabela com as pessoas", type="warning")
            return
        if not self.minhas:
            gui.notify("Não achei a planilha do processo: confira em Editar", type="warning")
            return
        self._desenhar_chaves()
        self.passos.next()

    # --- 2. Como achar a pessoa ---
    def _passo_chaves(self) -> None:
        _titulo_passo(2, "Como achar a pessoa?",
                      "Na hora de incluir uma linha, você escreve um destes dados (o começo do nome, por "
                      "exemplo) e o Edodocs procura na tabela. Ligue cada coluna da sua planilha à coluna da "
                      "tabela que tem o mesmo dado.")
        self.area_chaves = gui.column().classes("w-full gap-2 q-mt-md")
        with gui.stepper_navigation():
            gui.button("Próximo", icon="arrow_forward", on_click=self._ir_puxar).props("unelevated no-caps")
            gui.button("Voltar", on_click=self.passos.previous).props("flat no-caps")

    def _desenhar_chaves(self) -> None:
        if not self.tabela:
            return
        self.area_chaves.clear()
        with self.area_chaves:
            with gui.row().classes("w-full items-center gap-2 no-wrap text-caption text-weight-bold t-suave"):
                gui.label("Na sua planilha").classes("w-60")
                gui.label("").classes("w-8")
                gui.label("Na tabela").classes("w-60")
                gui.label("Exemplo").classes("col")
            for chave in self.chaves:
                with gui.row().classes("w-full items-center gap-2 no-wrap"):
                    gui.select(self.minhas, value=chave["minha"] if chave["minha"] in self.minhas else None,
                               on_change=lambda e, c=chave: (c.update(minha=e.value), self._desenhar_chaves())) \
                        .props("outlined dense options-dense").classes("w-60")
                    gui.icon("arrow_forward", color="primary", size="24px").classes("w-8")
                    gui.select(self.tabela.colunas,
                               value=chave["dela"] if chave["dela"] in self.tabela.colunas else None,
                               on_change=lambda e, c=chave: (c.update(dela=e.value), self._desenhar_chaves())) \
                        .props("outlined dense options-dense use-input input-debounce=0").classes("w-60")
                    gui.label(consultas.valor_juntado(chave, self._linha_exemplo(chave["dela"]))
                              if chave["dela"] else "") \
                        .classes("text-body2 t-forte col ellipsis")
                    if chave["dela"]:
                        self._botao_juntar(chave, self._desenhar_chaves)
                    gui.button(icon="close", on_click=lambda c=chave: self._tirar_chave(c)) \
                        .props("flat round dense color=grey-7").tooltip("Tirar esta forma de achar")
            gui.button("Mais uma forma de achar (ex.: matrícula)", icon="add",
                       on_click=self._mais_chave).props("flat no-caps dense")
            if len(self.chaves) > 1:
                gui.label("Dá para achar a pessoa por qualquer uma delas.").classes("text-caption t-suave")

    def _linha_exemplo(self, coluna: str) -> dict:
        """Uma pessoa da tabela que tenha a coluna preenchida (para mostrar como fica)."""
        return next((l for l in self.tabela.linhas[:2000] if l.get(coluna)), {})

    def _botao_juntar(self, item: dict, redesenhar) -> None:
        """Juntar a coluna da tabela com outra (ex.: cd_matricula + dv_matricula → 30.244-9)."""
        juntar = item.get("juntar")
        rotulo = ("+ " + ", ".join(juntar.get("com") or [])) if juntar else ""
        gui.button(rotulo, icon="join_inner", on_click=lambda: self._editar_juntar(item, redesenhar)) \
            .props("flat dense no-caps no-wrap size=sm color=" + ("positive" if juntar else "grey-6")) \
            .tooltip(f"Junta com {', '.join(juntar.get('com') or [])}" if juntar else
                     "Juntar com outra coluna da tabela (ex.: matrícula + dígito)")

    def _editar_juntar(self, item: dict, redesenhar) -> None:
        atual = item.get("juntar") or {}
        outras = [c for c in self.tabela.colunas if c != item["dela"]]
        # Sugestão: a coluna do dígito (dv_...) que acompanha esta
        sugerida = next((c for c in outras if atos.normalizar(c).startswith("DV")
                         and atos.normalizar(c)[2:] in atos.normalizar(item["dela"])), None)
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[560px] max-w-full q-pa-lg gap-3"):
            gui.label(f"Juntar {item['dela']} com outra coluna").classes("text-h6 t-forte")
            gui.label("Ex.: a matrícula vem em duas colunas, o número (cd_matricula) e o dígito (dv_matricula): "
                      "juntando, fica 30.244-9.").classes("text-body2 t-suave")
            com = gui.select(outras, label="Juntar com", value=(atual.get("com") or [sugerida])[0],
                             with_input=True).props("outlined dense").classes("w-full")
            separador = gui.select({"-": "traço  (30244-9)", "": "nada  (302449)", " ": "espaço  (30244 9)",
                                    "/": "barra  (30244/9)", ".": "ponto  (30244.9)"},
                                   label="Entre as duas", value=atual.get("separador", "-")) \
                .props("outlined dense").classes("w-full")
            milhar = gui.checkbox("Ponto de milhar no número (30244 → 30.244)",
                                  value=atual.get("milhar", not atual))
            previa = gui.label("").classes("text-body1 t-forte")

            def regra() -> dict | None:
                return ({"com": [com.value], "separador": separador.value, "milhar": bool(milhar.value)}
                        if com.value else None)

            def atualizar(_=None) -> None:
                r = regra()
                exemplos = [l for l in self.tabela.linhas[:2000] if l.get(item["dela"])][:3]
                previa.set_text("Fica assim: " + " · ".join(
                    consultas.valor_juntado(dict(item, juntar=r), l) for l in exemplos) if exemplos else "")

            for campo in (com, separador, milhar):
                campo.on_value_change(atualizar)
            atualizar()

            def usar(r: dict | None) -> None:
                if r:
                    item["juntar"] = r
                else:
                    item.pop("juntar", None)
                janela.close()
                redesenhar()

            with gui.row().classes("w-full items-center gap-2"):
                if atual:
                    gui.button("Não juntar", icon="link_off", on_click=lambda: usar(None)) \
                        .props("flat no-caps color=negative")
                gui.space()
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Usar", icon="check", on_click=lambda: usar(regra())).props("unelevated no-caps")
        janela.open()

    def _mais_chave(self) -> None:
        self.chaves.append({"minha": None, "dela": None})
        self._desenhar_chaves()

    def _tirar_chave(self, chave: dict) -> None:
        self.chaves.remove(chave)
        self._desenhar_chaves()

    def _ir_puxar(self) -> None:
        self.chaves = [c for c in self.chaves if c.get("minha") or c.get("dela")]
        if not self.chaves or any(not (c.get("minha") and c.get("dela")) for c in self.chaves):
            gui.notify("Ligue pelo menos uma coluna (ex.: NOME → NOME COMPLETO), dos dois lados", type="warning")
            self._desenhar_chaves()
            return
        for c in self.chaves:  # uma coluna não é chave e "o que vem junto" ao mesmo tempo
            self.puxar.pop(c["minha"], None)
        self._desenhar_puxar()
        self.passos.next()

    # --- 3. O que vem junto ---
    def _passo_puxar(self) -> None:
        _titulo_passo(3, "O que vem junto?",
                      "Marque as colunas que o Edodocs preenche sozinho com os dados da pessoa. As que só "
                      "você sabe (A PEDIDO, a data do ato...) ficam desmarcadas: você preenche.")
        self.area_puxar = gui.column().classes("w-full gap-1 q-mt-md")
        with gui.stepper_navigation():
            gui.button("Próximo", icon="arrow_forward", on_click=self._ir_teste).props("unelevated no-caps")
            gui.button("Voltar", on_click=self.passos.previous).props("flat no-caps")

    def _desenhar_puxar(self) -> None:
        if not self.tabela:
            return
        self.area_puxar.clear()
        das_chaves = {c.get("minha") for c in self.chaves}
        exemplo = self.tabela.linhas[0] if self.tabela.linhas else {}
        # O exemplo é de uma pessoa que tenha o máximo de colunas ligadas preenchidas
        ligadas = [p["dela"] for p in self.puxar.values()]
        if ligadas and self.tabela.linhas:
            exemplo = max(self.tabela.linhas[:500], key=lambda l: sum(bool(l.get(d)) for d in ligadas))
        with self.area_puxar:
            with gui.row().classes("w-full items-center gap-2 no-wrap text-caption text-weight-bold t-suave"):
                gui.label("Na sua planilha").classes("w-56")
                gui.label("").classes("w-8")
                gui.label("Vem da coluna da tabela").classes("w-60")
                gui.label(f"Exemplo ({exemplo.get(self.chaves[0]['dela'], '') if self.chaves else ''})") \
                    .classes("col ellipsis")
            for minha in self.minhas:
                if minha in das_chaves:
                    continue
                item = self.puxar.get(minha)
                with gui.row().classes("w-full items-center gap-2 no-wrap"):
                    with gui.row().classes("w-56 items-center gap-1 no-wrap"):
                        gui.checkbox(value=bool(item), on_change=lambda e, m=minha: self._marcar(m, e.value)) \
                            .props("dense")
                        gui.label(minha).classes("text-body2 " + ("t-forte text-weight-medium" if item else "t-suave"))
                    gui.icon("arrow_back", color="primary" if item else "grey-5", size="22px").classes("w-8")
                    gui.select({c: c for c in self.tabela.colunas}, value=item["dela"] if item else None,
                               clearable=True, on_change=lambda e, m=minha: self._trocar_coluna(m, e.value)) \
                        .props("outlined dense options-dense use-input input-debounce=0").classes("w-60")
                    with gui.row().classes("col items-center gap-1 no-wrap").style("min-width: 0"):
                        if item:
                            valor = exemplo.get(item["dela"], "")
                            traduzido = consultas.completar({"puxar": [item]}, exemplo).get(minha, valor)
                            # O exemplo encolhe (com "..."), os botões ao lado ficam inteiros
                            gui.label(traduzido or "(vazio)").classes("text-body2 ellipsis "
                                                                     + ("t-forte" if traduzido else "t-suave")) \
                                .style("min-width: 0; max-width: 170px").tooltip(traduzido or "")
                            self._botao_juntar(item, self._desenhar_puxar)
                            self._botao_so_se(minha, item)
                            self._botao_de_para(minha, item)
                            if not self.opcoes.get(minha):  # coluna com lista: a tradução é que decide
                                self._botao_formato(item)
                            if item.get("traducao") or self.opcoes.get(minha):
                                gui.button(icon="translate", on_click=lambda m=minha: self._editar_traducao(m)) \
                                    .props("flat round dense size=sm color=primary") \
                                    .tooltip("Como os valores da tabela viram os da sua lista "
                                             + ", ".join(f"{k} → {v}" for k, v in (item.get("traducao") or {}).items()))
            n = len(self.puxar)
            gui.label(f"{n} coluna(s) marcada(s). Nada é apagado: o Edodocs só preenche o que estiver vazio.") \
                .classes("text-caption t-suave q-mt-sm")

    def _marcar(self, minha: str, ligado: bool) -> None:
        if not ligado:
            self.puxar.pop(minha, None)
        else:
            sugestao = consultas.sugerir_puxar([minha], self.tabela, self.chaves, self.opcoes)
            self.puxar[minha] = sugestao[0] if sugestao else {"minha": minha, "dela": None, "traducao": {}}
        self._desenhar_puxar()

    def _trocar_coluna(self, minha: str, dela: str | None) -> None:
        if not dela:
            self.puxar.pop(minha, None)
        else:
            self.puxar[minha] = consultas.item_puxar(minha, dela, self.tabela, self.opcoes.get(minha))
        self._desenhar_puxar()

    def _botao_de_para(self, minha: str, item: dict) -> None:
        """Código na tabela, descrição no ato: o de-para que troca um pelo outro (com o valor
        todo ou só uma parte dele, ex.: os 2 primeiros dígitos)."""
        import de_para_ui  # aqui: de_para_ui usa este módulo
        nome = item.get("de_para")
        regra = item.get("de_para_parte")
        valores = self.tabela.valores(item["dela"])
        so_codigos = not nome and de_para.parece_codigo(valores, len(self.tabela.linhas))
        # Com nome: o de-para ligado; só código: chama a atenção; o resto: só o ícone, discreto
        rotulo = (nome if len(nome) <= 16 else nome[:15] + "…") if nome else "Só código? De-para" if so_codigos else ""
        if nome and regra:  # curto: o tooltip explica
            rotulo += f" · {regra['n']} {'últ.' if regra.get('de') == 'fim' else 'prim.'}"
        with gui.button(rotulo, icon="swap_horiz").props(
                "flat dense no-caps no-wrap size=sm color=" + ("positive" if nome else "orange-9" if so_codigos
                                                               else "grey-6")) as botao:
            botao.tooltip(f"De-para: {nome}" + (f", com os {de_para.descrever_parte(regra)} caracteres" if regra else "")
                          if nome else "A tabela traz só o código? Ligue a um de-para (código → descrição)")
            with gui.menu():
                gui.label("Trocar o código pela descrição").classes("text-caption t-suave q-px-md q-pt-sm")
                gui.menu_item("Não usar de-para (entra como vem na tabela)",
                              on_click=lambda: self._usar_de_para(minha, None))
                listas = de_para.todos()
                for outro, pares in listas.items():
                    parte = round(100 * de_para.cobertura(pares, valores[:300], regra))
                    gui.menu_item(("✓ " if outro == nome else "") + f"{outro} (conhece {parte}% dos códigos)",
                                  on_click=lambda o=outro: de_para_ui.perguntar_parte(
                                      self.camada, valores, lambda r, o=o: self._usar_de_para(minha, o, r),
                                      regra, listas[o], f"Ligar ao de-para {o}: o código todo ou só uma parte?"))
                gui.separator()
                gui.menu_item("Criar um de-para com os códigos desta coluna...",
                              on_click=lambda: de_para_ui.perguntar_parte(
                                  self.camada, valores, lambda r: self._criar_de_para(minha, valores, r), regra))
                if nome:
                    gui.menu_item(f"Completar o de-para {nome}...",
                                  on_click=lambda: de_para_ui.EditorDePara(
                                      self.camada, nome, codigos=self._partes(valores, regra),
                                      ao_salvar=lambda novo: self._usar_de_para(minha, novo, regra)))

    @staticmethod
    def _partes(valores: list[str], regra: dict | None) -> list[str]:
        return list(dict.fromkeys(de_para.parte(v, regra) for v in valores if v))

    def _criar_de_para(self, minha: str, valores: list[str], regra: dict | None) -> None:
        import de_para_ui
        de_para_ui.EditorDePara(self.camada, codigos=self._partes(valores, regra), sugestao_nome=minha.title(),
                                ao_salvar=lambda novo: self._usar_de_para(minha, novo, regra))

    def _botao_so_se(self, minha: str, item: dict) -> None:
        """Só preencher em certos casos (ex.: CARGO EFETIVO só se cd_cargo_efetivo tem valor)."""
        tem = bool(item.get("se"))
        gui.button("Só se…" if tem else "", icon="filter_alt", on_click=lambda: self._editar_so_se(minha, item)) \
            .props("flat dense no-caps no-wrap size=sm color=" + ("positive" if tem else "grey-6")) \
            .tooltip(f"Só preenche se {consultas.descrever_se(item['se'])}" if tem else
                     "Só preencher em certos casos (ex.: o cargo efetivo só para quem tem um)")

    def _editar_so_se(self, minha: str, item: dict) -> None:
        condicoes = [dict(c) for c in item.get("se") or []]
        if not condicoes:  # sugestão: a coluna de código que acompanha (ds_cargo → cd_cargo_efetivo...)
            parecida = max(self.tabela.colunas, key=lambda c: consultas.parecenca(minha, c))
            condicoes.append({"coluna": parecida, "op": "tem_valor", "valor": ""})
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[760px] max-w-full q-pa-lg gap-3"):
            gui.label(f"{minha}: só preencher se…").classes("text-h6 t-forte")
            gui.label("O Edodocs só preenche este campo quando todas as condições valem para a pessoa escolhida "
                      "na tabela; senão, o campo fica vazio. Ex.: CARGO EFETIVO só se cd_cargo_efetivo tem valor "
                      "(quem não tem cargo efetivo vem com 0 ou vazio).").classes("text-body2 t-suave")
            area = gui.column().classes("w-full gap-2")
            previa = gui.label("").classes("text-body2 t-forte")

            def atualizar_previa() -> None:
                validas = [c for c in condicoes if c.get("coluna")]
                if not validas:
                    previa.set_text("")
                    return
                normal = [{atos.normalizar(k): v for k, v in l.items()} for l in self.tabela.linhas[:3000]]
                sim = sum(consultas.condicoes_valem(validas, l) for l in normal)
                previa.set_text(f"Nas primeiras {len(normal)} pessoas da tabela: preenche para {sim}, "
                                f"fica vazio para {len(normal) - sim}.")

            def desenhar() -> None:
                area.clear()
                with area:
                    for i, cond in enumerate(condicoes):
                        with gui.row().classes("w-full items-center gap-2 no-wrap"):
                            gui.label("Se" if i == 0 else "e").classes("text-weight-medium t-forte w-6")
                            gui.select(self.tabela.colunas, value=cond.get("coluna"), label="Coluna da tabela",
                                       with_input=True,
                                       on_change=lambda e, c=cond: (c.update(coluna=e.value), atualizar_previa())) \
                                .props("dense outlined options-dense").classes("w-60")
                            gui.select(consultas.OPERADORES_SE, value=cond.get("op", "tem_valor"),
                                       on_change=lambda e, c=cond: (c.update(op=e.value), desenhar())) \
                                .props("dense outlined options-dense").classes("w-72")
                            if cond.get("op", "tem_valor") not in consultas.SEM_VALOR_SE:
                                gui.input("Valor", value=cond.get("valor", ""),
                                          on_change=lambda e, c=cond: (c.update(valor=e.value or ""), atualizar_previa())) \
                                    .props("dense outlined").classes("col")
                            else:
                                gui.space()
                            gui.button(icon="close", on_click=lambda i=i: (condicoes.pop(i), desenhar())) \
                                .props("flat round dense size=sm color=grey-7").tooltip("Tirar esta condição")
                    gui.button("e também...", icon="add",
                               on_click=lambda: (condicoes.append({"coluna": None, "op": "tem_valor", "valor": ""}),
                                                 desenhar())).props("flat dense no-caps size=sm")
                atualizar_previa()

            def usar(lista: list[dict] | None) -> None:
                lista = [c for c in lista or [] if c.get("coluna")]
                if lista:
                    item["se"] = lista
                else:
                    item.pop("se", None)
                janela.close()
                self._desenhar_puxar()

            desenhar()
            with gui.row().classes("w-full items-center gap-2"):
                if item.get("se"):
                    gui.button("Preencher sempre", icon="filter_alt_off", on_click=lambda: usar(None)) \
                        .props("flat no-caps color=negative")
                gui.space()
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Usar", icon="check", on_click=lambda: usar(condicoes)).props("unelevated no-caps")
        janela.open()

    def _botao_formato(self, item: dict) -> None:
        """Maiúsculas/minúsculas do que vem da tabela (ASSESSOR TÉCNICO → Assessor Técnico)."""
        atual = item.get("formato", "")
        with gui.button(consultas.FORMATOS[atual] if atual else "", icon="text_fields").props(
                "flat dense no-caps no-wrap size=sm color=" + ("positive" if atual else "grey-6")) as botao:
            botao.tooltip("Maiúsculas e minúsculas: " + consultas.FORMATOS[atual])
            with gui.menu():
                exemplo = self.tabela.exemplo(item["dela"])
                for chave, rotulo in consultas.FORMATOS.items():
                    previa = consultas.formatar(exemplo, chave) if exemplo else ""
                    gui.menu_item(("✓ " if chave == atual else "") + rotulo + (f"  ({previa[:40]})" if previa else ""),
                                  on_click=lambda c=chave: self._usar_formato(item, c))

    def _usar_formato(self, item: dict, formato: str) -> None:
        if formato:
            item["formato"] = formato
        else:
            item.pop("formato", None)
        self._desenhar_puxar()

    def _usar_de_para(self, minha: str, nome: str | None, regra: dict | None = None) -> None:
        item = self.puxar.get(minha)
        if item:
            novo = consultas.item_puxar(minha, item["dela"], self.tabela, self.opcoes.get(minha), nome, regra)
            for chave in ("juntar", "formato", "se"):  # trocar o de-para não desfaz o juntar, o formato...
                if item.get(chave):
                    novo[chave] = item[chave]
            self.puxar[minha] = novo
        self._desenhar_puxar()

    def _editar_traducao(self, minha: str) -> None:
        """Os valores diferentes da coluna da tabela e o que cada um vira na planilha."""
        item = self.puxar[minha]
        todos_valores = self.tabela.valores(item["dela"])
        if item.get("de_para"):  # a tradução é sobre a descrição que o de-para já trocou
            listas = de_para.todos()
            todos_valores = list(dict.fromkeys(de_para.traduzir(item["de_para"], v, listas, item.get("de_para_parte"))
                                               for v in todos_valores))
        valores = todos_valores[:20]
        traducao = dict(item.get("traducao") or {})
        lista = self.opcoes.get(minha) or []
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[560px] q-pa-lg gap-2"):
            gui.label(f"{minha}: como os valores da tabela entram na planilha").classes("text-h6 t-forte")
            gui.label("Em branco: entra como está na tabela.").classes("text-caption t-suave")
            for valor in valores:
                with gui.row().classes("w-full items-center gap-2 no-wrap"):
                    gui.label(valor).classes("text-body2 t-forte w-56 ellipsis")
                    gui.icon("arrow_forward", color="primary")
                    gui.select(lista, value=traducao.get(valor), clearable=True, with_input=True,
                               new_value_mode="add-unique",
                               on_change=lambda e, v=valor: traducao.update({v: e.value})) \
                        .props("outlined dense options-dense").classes("col")
            if len(todos_valores) > 20:
                gui.label("Mostrando os 20 primeiros valores.").classes("text-caption t-suave")

            def guardar() -> None:
                item["traducao"] = {k: v for k, v in traducao.items() if v}
                janela.close()
                self._desenhar_puxar()

            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Guardar", icon="check", on_click=guardar).props("unelevated no-caps")
        janela.open()

    def _ir_teste(self) -> None:
        self.puxar = {m: p for m, p in self.puxar.items() if p.get("dela")}
        if not self.puxar:
            gui.notify("Marque pelo menos uma coluna para vir da tabela", type="warning")
            return
        self.passos.next()

    # --- 4. Testar e salvar ---
    def _passo_teste(self) -> None:
        _titulo_passo(4, "Teste antes de salvar",
                      "Escreva o começo de um nome (ou uma matrícula) e escolha a pessoa: veja como a linha "
                      "fica preenchida.")
        with gui.column().classes("w-full gap-2 q-mt-md"):
            self.area_teste = gui.column().classes("w-full gap-2")
            self.area_resultado = gui.column().classes("w-full gap-1")
        with gui.stepper_navigation():
            gui.button("Salvar", icon="save", color="positive", on_click=self._salvar).props("unelevated no-caps")
            gui.button("Voltar", on_click=self.passos.previous).props("flat no-caps")
            if self.ja_tinha:
                gui.button("Desligar o preencher sozinho", icon="link_off", on_click=self._desligar) \
                    .props("flat no-caps color=negative")
        self.passos.on_value_change(lambda e: e.value == "teste" and self._desenhar_teste())

    def _consulta(self) -> dict:
        return {"tabela": self.nome_tabela,
                "chaves": [{"minha": c["minha"], "dela": c["dela"],
                            **({"juntar": c["juntar"]} if c.get("juntar") else {})} for c in self.chaves],
                "puxar": [{"minha": m, "dela": p["dela"], "traducao": p.get("traducao") or {},
                           **({"juntar": p["juntar"]} if p.get("juntar") else {}),
                           **({"formato": p["formato"]} if p.get("formato") else {}),
                           **({"se": p["se"]} if p.get("se") else {}),
                           **({"de_para": p["de_para"]} if p.get("de_para") else {}),
                           **({"de_para_parte": p["de_para_parte"]} if p.get("de_para") and p.get("de_para_parte")
                              else {})}
                          for m, p in self.puxar.items() if m in self.minhas]}

    def _desenhar_teste(self) -> None:
        if not self.tabela:
            return
        self.area_teste.clear()
        self.area_resultado.clear()
        with self.area_teste:
            BuscaNaTabela(self._consulta(), self.tabela, ao_escolher=self._mostrar_teste)

    def _mostrar_teste(self, linha: dict) -> None:
        valores = consultas.completar(self._consulta(), linha)
        self.area_resultado.clear()
        with self.area_resultado:
            gui.label("A linha ficaria assim:").classes("text-body2 t-forte")
            with gui.grid(columns=2).classes("w-full gap-x-6 gap-y-0"):
                for minha in self.minhas:
                    with gui.row().classes("items-center gap-2 no-wrap"):
                        valor = valores.get(minha, "")
                        gui.icon("check_circle" if valor else "edit", color="positive" if valor else "grey-5",
                                 size="18px")
                        gui.label(f"{minha}:").classes("text-body2 t-suave")
                        gui.label(valor or "você preenche").classes(
                            "text-body2 ellipsis " + ("t-forte text-weight-medium" if valor else "t-suave"))

    def _salvar(self) -> None:
        proc = processos.carregar(self.pasta)  # o processo pode ter mudado enquanto a janela estava aberta
        proc.config["consulta"] = self._consulta()
        proc.salvar()
        import de_para_ui
        de_para_ui._redesenhar_cartoes()  # "usado em" do de-para, em Configurações
        self.janela.close()
        gui.notify(f"Pronto: na Nova linha de {self.pasta.name}, procure a pessoa e o resto vem da tabela "
                   f"{self.nome_tabela}", type="positive", timeout=6000)
        self.ao_salvar()

    def _desligar(self) -> None:
        proc = processos.carregar(self.pasta)
        proc.config.pop("consulta", None)
        proc.salvar()
        self.janela.close()
        gui.notify("Preencher sozinho desligado neste processo", type="info")
        self.ao_salvar()


class BuscaNaTabela:
    """Caixa de busca na tabela de consulta. Escolher uma pessoa chama ao_escolher(linha) ou,
    com campos ({coluna normalizada: campo da tela}), preenche os campos vazios (em verde)."""

    def __init__(self, consulta: dict, tabela: consultas.Tabela, campos: dict | None = None,
                 por_valor=None, ao_escolher=None):
        self.consulta = consulta
        self.tabela = tabela
        self.campos = campos or {}
        self.por_valor = por_valor or (lambda campo, valor: setattr(campo, "value", valor))
        self.ao_escolher = ao_escolher
        self.marcados: dict = {}  # campo → valor que a tabela pôs (dá para trocar de pessoa)
        chaves = " ou ".join(c["minha"].lower() for c in consulta.get("chaves") or [])
        with gui.column().classes("w-full gap-0"):
            with gui.row().classes("w-full items-center gap-2 no-wrap fundo-cartao rounded q-pa-sm"):
                gui.icon("person_search", color="primary", size="26px")
                self.entrada = gui.input(placeholder=f"Procurar em {consulta.get('tabela', 'tabela')}: "
                                                     f"escreva o {chaves or 'nome'}",
                                         on_change=lambda e: self._procurar(e.value)) \
                    .props('outlined dense clearable bg-color=white input-class=text-grey-10 debounce=250') \
                    .classes("col")
            self.resultados = gui.column().classes("w-full gap-0 q-px-sm")

    def _procurar(self, termo: str | None) -> None:
        self.resultados.clear()
        termo = (termo or "").strip()
        if not termo:
            return
        achados = consultas.buscar(self.consulta, self.tabela, termo)
        with self.resultados:
            if not achados:
                gui.label("Ninguém com isso na tabela. Se a pessoa for nova, preencha os campos abaixo.") \
                    .classes("text-caption t-suave q-pa-xs")
                return
            for linha in achados:
                with gui.row().classes("w-full items-center gap-2 no-wrap q-px-sm q-py-xs rounded cursor-pointer "
                                       "resultado-busca") \
                        .on("click", lambda l=linha: self._escolher(l)):
                    gui.icon("person", color="primary", size="20px")
                    gui.label(consultas.descricao(self.consulta, linha)).classes("text-body2 t-forte ellipsis")

    def _escolher(self, linha: dict) -> None:
        self.resultados.clear()
        self.entrada.value = ""
        if self.ao_escolher:
            self.ao_escolher(linha)
        if not self.campos:
            return
        preenchidos, mantidos = 0, []
        novos = consultas.completar(self.consulta, linha)
        # Trocou de pessoa e a nova não tem esse dado (ex.: sem cargo efetivo): o que a tabela
        # tinha posto da pessoa anterior sai
        usados = {self.campos.get(atos.normalizar(m)) for m in novos}
        for campo, posto in list(self.marcados.items()):
            if campo not in usados and str(campo.value or "").strip() == posto:
                self.por_valor(campo, "")
                campo.classes(remove=VERDE)
                del self.marcados[campo]
        for minha, valor in novos.items():
            campo = self.campos.get(atos.normalizar(minha))
            if campo is None:
                continue
            atual = str(campo.value or "").strip()
            # Vazio, ou posto pela tabela numa escolha anterior: troca. Escrito pela pessoa: fica.
            if atual and self.marcados.get(campo) != atual:
                if atual != valor:
                    mantidos.append(minha)
                continue
            self.por_valor(campo, valor)
            campo.classes(add=VERDE)
            self.marcados[campo] = valor
            preenchidos += 1
        nome = linha.get((self.consulta.get("chaves") or [{}])[0].get("dela", ""), "")
        gui.notify(f"Preenchi {preenchidos} campo(s) com os dados de {nome}"
                   + (f". Não mexi no que você já tinha escrito: {', '.join(mantidos)}" if mantidos else ""),
                   type="positive")

    def limpar_marcas(self) -> None:
        """O formulário foi limpo (ou preenchido de outro jeito): tira o verde."""
        for campo in self.marcados:
            campo.classes(remove=VERDE)
        self.marcados.clear()
