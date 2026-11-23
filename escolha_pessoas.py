"""Escolha de quem vista e quem assina, pesquisando as pessoas no próprio e-docs.

    montar_escolha_pessoas   a janela de pesquisa (usada no envio e para deixar fixos)
    PesquisaEdocs            pesquisa com o usuário do e-docs salvo (entra só na 1ª vez)
    PainelPessoas            cartão "Vistos e assinantes" do processo (fixos ou perguntados)
"""
import asyncio
import re

from nicegui import run
from nicegui import ui as gui

import ui


def mesma_pessoa(a: dict, b: dict) -> bool:
    """Mesmo CPF ou mesmo nome (sem acento, maiúscula ou espaço a mais)."""
    import atos
    return (bool(a.get("cpf")) and a.get("cpf") == b.get("cpf")) or         (bool(atos.normalizar(a.get("nome"))) and atos.normalizar(a.get("nome")) == atos.normalizar(b.get("nome")))


def sem_repetir(pessoas: list[dict]) -> list[dict]:
    """A lista sem a mesma pessoa duas vezes (fica a primeira)."""
    saida: list[dict] = []
    for p in pessoas:
        if not any(mesma_pessoa(p, q) for q in saida):
            saida.append(p)
    return saida


def montar_escolha_pessoas(cartao, pergunta: str, pesquisar, recentes: list[dict],
                           marcados: list[dict], vazio: str | None, responder, cancelar) -> None:
    """Pesquisa no Edocs (CPF, nome ou matrícula) e monta a lista em ordem, dentro do cartão
    de uma janela. Clicar numa pessoa da pesquisa põe ou tira ela da lista; as setas mudam a
    ordem (1º, 2º...). Os últimos escolhidos aparecem em cima, para um clique só.
    pesquisar(termo) → [pessoa] (roda fora da tela); responder(lista) e cancelar() fecham."""
    escolhidos: list[dict] = list(marcados)

    def tem(p: dict) -> bool:
        return any(mesma_pessoa(e, p) for e in escolhidos)

    def incluir(p: dict) -> None:
        # O mesmo nome duas vezes (ex.: a pessoa com dois cadastros no e-docs) dá erro de
        # "Seq. Assinatura em duplicidade" no envio: não entra de novo
        if tem(p):
            gui.notify(f"{p['nome']} já está na lista", type="warning")
            return
        escolhidos.append(p)
        desenhar()

    def tirar(p: dict) -> None:
        escolhidos[:] = [e for e in escolhidos if not mesma_pessoa(e, p)]
        desenhar()

    def alternar(p: dict) -> None:
        """Clique na pessoa da pesquisa: entra se não estava; sai se já estava."""
        tirar(p) if tem(p) else incluir(p)

    def mover(i: int, passo: int) -> None:
        """Sobe (-1) ou desce (+1) a pessoa na ordem."""
        j = i + passo
        if 0 <= j < len(escolhidos):
            escolhidos[i], escolhidos[j] = escolhidos[j], escolhidos[i]
            desenhar()

    def desenhar() -> None:
        area_escolhidos.clear()
        with area_escolhidos:
            if not escolhidos:
                gui.label("Ninguém escolhido ainda" + (f" ({vazio})" if vazio else "")) \
                    .classes("text-caption t-fraco q-py-xs")
            for i, p in enumerate(escolhidos):
                with gui.row().classes("w-full items-center gap-2 no-wrap fundo-cartao rounded q-px-sm"):
                    gui.label(f"{i + 1}º").classes("text-weight-bold text-primary").style("min-width: 26px")
                    gui.label(p["nome"]).classes("col ellipsis t-forte").tooltip(ui.descrever_pessoa(p))
                    subir = gui.button(icon="arrow_upward", on_click=lambda i=i: mover(i, -1)) \
                        .props("flat round dense size=sm").tooltip("Subir na ordem")
                    descer = gui.button(icon="arrow_downward", on_click=lambda i=i: mover(i, 1)) \
                        .props("flat round dense size=sm").tooltip("Descer na ordem")
                    subir.set_enabled(i > 0)
                    descer.set_enabled(i < len(escolhidos) - 1)
                    gui.button(icon="close", on_click=lambda p=p: tirar(p)) \
                        .props("flat round dense size=sm color=negative").tooltip("Tirar da lista")
        area_recentes.clear()
        livres = [p for p in recentes if not tem(p)]
        area_recentes.set_visibility(bool(livres))
        with area_recentes:
            gui.label("Últimos escolhidos").classes("text-caption t-suave w-full")
            for p in livres:
                gui.chip(p["nome"], icon="history", on_click=lambda p=p: incluir(p)) \
                    .props("clickable outline dense color=primary").tooltip(ui.descrever_pessoa(p))
        desenhar_resultados()
        confirmar.set_enabled(bool(escolhidos) or vazio is not None)

    resultados: list[dict] = []

    def desenhar_resultados() -> None:
        area_resultados.clear()
        with area_resultados:
            for p in resultados:
                detalhes = " · ".join(
                    x for x in (ui.descrever_pessoa(p | {"nome": "-"})[4:], p.get("lotacao")) if x)
                escolhido = tem(p)
                with gui.item(on_click=lambda p=p: alternar(p)).props("clickable dense") \
                    .classes("w-full rounded-borders") \
                    .tooltip("Clique para tirar da lista" if escolhido else "Clique para adicionar"):
                    with gui.item_section():
                        gui.item_label(p["nome"])
                        gui.item_label(detalhes).props("caption lines=2")
                    with gui.item_section().props("side"):
                        with gui.row().classes("items-center gap-1 no-wrap"):
                            gui.label("Escolhido" if escolhido else "Adicionar") \
                                .classes("text-caption " + ("text-primary" if escolhido else "t-suave"))
                            gui.icon("check_circle" if escolhido else "add_circle_outline",
                                     color="primary").classes("text-xl")

    async def buscar() -> None:
        termo = (campo.value or "").strip()
        if not termo:
            return
        if len(re.sub(r"\d", "", termo).strip()) and len(termo) < 3:
            aviso.text = "Digite pelo menos 3 letras do nome"
            return
        botao.props("loading")
        aviso.text = ""
        try:
            achados = await run.io_bound(pesquisar, termo)
        except Exception as e:  # sessão caiu, site fora...
            aviso.text = f"Não deu para pesquisar: {type(e).__name__} {e}".strip()
            achados = None
        finally:
            botao.props(remove="loading")
        if achados is None:
            return
        resultados[:] = achados
        if not achados:
            digitos = re.sub(r"\D", "", termo)
            aviso.text = ("Ninguém encontrado. Na matrícula, confira o dígito verificador no final."
                          if digitos and len(digitos) != 11 and digitos == re.sub(r"[\s.\-/]", "", termo)
                          else "Ninguém encontrado.")
        desenhar_resultados()  # a pessoa só entra na lista quando você clica nela

    with cartao:
        gui.label(pergunta).classes("text-h6")
        gui.label("Pesquise e clique na pessoa para adicionar (clique de novo para tirar). Use as setas "
                  "para mudar a ordem do visto/assinatura (1º, 2º...)."
                  + (" Já vêm marcados os escolhidos antes." if marcados else "")) \
            .classes("text-caption t-suave")
        area_escolhidos = gui.column().classes("w-full gap-1 min-h-[32px]")
        gui.separator()
        area_recentes = gui.row().classes("w-full gap-1 items-center")
        with gui.row().classes("w-full items-start no-wrap gap-2"):
            campo = gui.input("CPF, nome ou matrícula").props("outlined dense autofocus clearable") \
                .classes("grow")
            campo.on("keydown.enter", buscar)
            botao = gui.button("Pesquisar", icon="search", on_click=buscar).props("no-caps unelevated")
        gui.label(ui.DICA_PESQUISA).classes("text-caption t-fraco -mt-2")
        aviso = gui.label("").classes("text-caption text-orange-8")
        with gui.scroll_area().classes("max-h-[300px] w-full"):
            area_resultados = gui.list().classes("w-full")
        with gui.row().classes("w-full justify-end gap-2"):
            gui.button("Cancelar", on_click=cancelar).props("flat no-caps")
            confirmar = gui.button("Confirmar", icon="check",
                                   on_click=lambda: responder(list(escolhidos)))
    desenhar()


class PesquisaEdocs:
    """pesquisar(termo) com o usuário do e-docs salvo em Configurações. Entra no site só na
    primeira pesquisa (e de novo se a sessão cair). Roda fora da tela (run.io_bound)."""

    def __init__(self):
        self.sessao = None

    def __call__(self, termo: str) -> list[dict]:
        import credenciais
        import edocs_http
        import main

        dados = credenciais.carregar(main.CREDENCIAIS, "edocs")
        if not dados:
            raise RuntimeError("cadastre seu usuário do e-docs em Configurações")
        for tentativa in range(2):
            if self.sessao is None:
                try:
                    self.sessao = edocs_http.entrar(*dados)
                except edocs_http.LoginRecusado:
                    raise RuntimeError("o e-docs recusou o usuário salvo em Configurações")
            try:
                return edocs_http.pesquisar_servidores(self.sessao, termo)
            except edocs_http.SessaoExpirada:
                self.sessao = None
        raise RuntimeError("a sessão do e-docs caiu; tente de novo")


class PainelPessoas:
    """Cartão "Vistos e assinantes" de um processo. Ligado (fixas), o envio usa as pessoas
    escolhidas aqui sem perguntar; desligado, pergunta a cada envio. Edita o dicionário
    config["pessoas"] do processo (quem chama salva)."""

    def __init__(self, pessoas: dict, pesquisar=None, ao_mudar=None):
        self.pessoas = pessoas
        self.ao_mudar = ao_mudar or (lambda: None)
        self.pessoas.setdefault("fixas", False)
        self.pessoas.setdefault("vistos", [])
        self.pessoas.setdefault("assinantes", [])
        self.pesquisar = pesquisar or PesquisaEdocs()
        with gui.column().classes("w-full gap-2") as self.raiz:
            self.chave = gui.switch("Sempre os mesmos vistos e assinantes neste processo",
                                    value=self.pessoas["fixas"], on_change=self._ligar)
            self.explicacao = gui.label("").classes("text-caption t-suave")
            self.area = gui.column().classes("w-full gap-2")
        self._desenhar()

    def valido(self) -> bool:
        """Fixos precisam de pelo menos um assinante."""
        return not self.pessoas["fixas"] or bool(self.pessoas["assinantes"])

    async def _ligar(self, e) -> None:
        self.pessoas["fixas"] = e.value
        self._desenhar()
        self.ao_mudar()
        # Ao ligar sem ninguém escolhido, já pergunta quem vista e quem assina
        if e.value and not self.pessoas["assinantes"]:
            if await self.escolher("vistos"):
                await self.escolher("assinantes")

    def _desenhar(self) -> None:
        fixas = self.pessoas["fixas"]
        self.explicacao.text = (
            "Ligado: o envio usa as pessoas abaixo sem perguntar. Para vistar você mesmo, inclua o seu "
            "nome nos vistos."
            if fixas else "Desligado: a cada envio o Edodocs pergunta quem vista e quem assina.")
        self.area.clear()
        self.area.set_visibility(fixas)
        if not fixas:
            return
        with self.area:
            for qual, rotulo, vazio in (("vistos", "Vistam", "ninguém"),
                                        ("assinantes", "Assinam", "ninguém ainda: escolha quem assina")):
                lista = self.pessoas[qual]
                with gui.row().classes("w-full items-center gap-1 no-wrap"):
                    with gui.column().classes("gap-1 col"):
                        gui.label(rotulo).classes("text-caption t-suave")
                        with gui.row().classes("gap-1 items-center"):
                            if not lista:
                                gui.label(vazio).classes("text-body2 " + (
                                    "text-negative" if qual == "assinantes" else "t-fraco"))
                            for i, p in enumerate(lista, 1):
                                gui.chip(f"{i}º · {p['nome']}", color=None).classes("chip-campo") \
                                    .props("dense").tooltip(ui.descrever_pessoa(p))
                    gui.button("Escolher", icon="person_search",
                               on_click=lambda q=qual: self.escolher(q)).props("flat dense no-caps")

    async def escolher(self, qual: str) -> bool:
        """Abre a pesquisa para 'vistos' ou 'assinantes'. False se a pessoa cancelar."""
        import main

        pergunta = ("Quem vista neste processo? (para vistar você mesmo, inclua o seu nome)" if qual == "vistos"
                    else "Quem assina neste processo?")
        resposta: asyncio.Future = asyncio.get_running_loop().create_future()
        with self.raiz, gui.dialog().props("persistent") as dialogo, \
                gui.card().classes("cartao w-[640px] max-w-full q-pa-lg") as cartao:
            montar_escolha_pessoas(
                cartao, pergunta, self.pesquisar, main.recentes_sem(), list(self.pessoas[qual]),
                "ninguém" if qual == "vistos" else None,
                lambda lista: resposta.done() or resposta.set_result(lista),
                lambda: resposta.done() or resposta.set_result(None))
        dialogo.open()
        lista = await resposta
        dialogo.close()
        dialogo.delete()
        if lista is None:
            return False
        self.pessoas[qual] = main.guardar_pessoas(lista)
        main.lembrar_recentes(lista)
        self._desenhar()
        self.ao_mudar()
        return True
