"""Telas do de-para (de_para.py):

  - CartaoDePara: o cartão em Configurações com a lista dos de-para e o botão de criar.
  - EditorDePara: a janela com os códigos e as descrições. Os códigos podem vir de uma
    tabela de consulta (só falta escrever as descrições), de um Excel que já tem os dois, ou
    colados do Excel.
"""
from nicegui import run
from nicegui import ui as gui

import atos
import consultas
import consultas_ui
import de_para

MOSTRAR = 150  # linhas na janela de uma vez (o resto aparece pela busca)
_cartoes: list = []  # os cartões abertos (salvar um de-para em outra tela redesenha eles)


def _redesenhar_cartoes() -> None:
    for cartao in list(_cartoes):
        try:
            cartao.desenhar()
        except Exception:  # cartão de uma página que já fechou
            _cartoes.remove(cartao)


def campos_parte(valores: list[str], regra: dict | None = None, pares: dict | None = None,
                 ao_mudar=None):
    """Os campos "o valor todo ou só uma parte?" com a prévia (0312 → 03). Devolve uma função
    que dá a regra escolhida (None: o valor todo). pares: o de-para, para mostrar quanto ele
    conhece com essa escolha."""
    n_atual = int((regra or {}).get("n") or 0)
    with gui.column().classes("w-full gap-1"):
        modo = gui.radio({"todo": "O valor todo", "inicio": "Só o começo", "fim": "Só o fim"},
                         value=(regra or {}).get("de", "inicio") if n_atual else "todo").props("inline")
        with gui.row().classes("items-center gap-2") as linha_n:
            gui.label("Quantos caracteres:").classes("text-body2 t-forte")
            n = gui.number(value=n_atual or 2, min=1, max=30, step=1, format="%d") \
                .props("outlined dense").classes("w-24")
            gui.label("(pontos, traços e espaços não contam)").classes("text-caption t-suave")
        previa = gui.label("").classes("text-body2 t-forte")
        conhece = gui.label("").classes("text-caption t-suave")

    def regra_escolhida() -> dict | None:
        if modo.value == "todo" or not n.value:
            return None
        return {"de": modo.value, "n": max(1, int(n.value))}

    def atualizar(_=None) -> None:
        linha_n.set_visibility(modo.value != "todo")
        r = regra_escolhida()
        exemplos = [v for v in valores if v][:4]
        previa.set_text("Fica assim: " + " · ".join(f"{v} → {de_para.parte(v, r)}" for v in exemplos)
                        if exemplos else "")
        if pares is not None:
            conhece.set_text(f"O de-para conhece {round(100 * de_para.cobertura(pares, valores[:300], r))}% "
                             "dos códigos desta coluna assim.")
        if ao_mudar:
            ao_mudar(r)

    modo.on_value_change(atualizar)
    n.on_value_change(atualizar)
    atualizar()
    return regra_escolhida


def perguntar_parte(camada, valores: list[str], continuar, regra: dict | None = None,
                    pares: dict | None = None, titulo: str = "") -> None:
    """Janela: para achar no de-para, usar o valor todo ou só uma parte dele? continuar(regra)."""
    with camada, gui.dialog() as janela, gui.card().classes("cartao w-[560px] max-w-full q-pa-lg gap-3"):
        gui.label(titulo or "Usar o código todo ou só uma parte?").classes("text-h6 t-forte")
        gui.label("Às vezes só um pedaço do valor é o código (ex.: os 2 primeiros dígitos da lotação dizem "
                  "a secretaria).").classes("text-body2 t-suave")
        escolha = campos_parte(valores, regra, pares)

        def seguir() -> None:
            janela.close()
            continuar(escolha())

        with gui.row().classes("w-full justify-end gap-2"):
            gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
            gui.button("Continuar", icon="arrow_forward", on_click=seguir).props("unelevated no-caps")
    janela.open()


class CartaoDePara:
    def __init__(self, camada):
        self.camada = camada
        _cartoes.append(self)
        with gui.card().classes("cartao q-pa-lg w-full"):
            with gui.row().classes("items-center gap-2"):
                gui.icon("swap_horiz", color="primary", size="26px")
                gui.label("De-para (código → descrição)").classes("text-h6")
            gui.label("Para as tabelas de consulta que só trazem o código (ex.: lotação 0312): diga uma vez "
                      "o que cada código quer dizer (Secretaria Municipal de Governo) e o ato já vem com a "
                      "descrição.").classes("text-body2 t-suave")
            self.lista = gui.column().classes("w-full gap-2 q-mt-sm")
            with gui.row().classes("items-center gap-2"):
                gui.button("Novo de-para", icon="add",
                           on_click=lambda: EditorDePara(self.camada, ao_salvar=lambda _: self.desenhar())) \
                    .props("outline no-caps")
                gui.button("Exportar", icon="file_upload", on_click=self._exportar).props("flat no-caps") \
                    .tooltip("Grava os de-para num Excel (na pasta Downloads), para levar a outro computador")
                gui.button("Importar", icon="file_download", on_click=self._importar).props("flat no-caps") \
                    .tooltip("Traz de-para de um Excel exportado (ou feito à mão: uma folha por de-para, "
                             "código na coluna A e descrição na B)")
        self.desenhar()

    # --- Exportar / importar (Excel) ---
    def _exportar(self) -> None:
        todos = de_para.todos()
        if not todos:
            gui.notify("Nenhum de-para para exportar ainda", type="info")
            return
        marcados = set(todos)
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[480px] q-pa-lg gap-2"):
            gui.label("Exportar de-para").classes("text-h6 t-forte")
            gui.label("Vão para um Excel na sua pasta Downloads, uma folha para cada.").classes("text-body2 t-suave")
            for nome, pares in todos.items():
                gui.checkbox(f"{nome} ({len(pares)} códigos)", value=True,
                             on_change=lambda e, n=nome: (marcados.add if e.value else marcados.discard)(n))

            async def exportar() -> None:
                if not marcados:
                    gui.notify("Marque pelo menos um", type="warning")
                    return
                from datetime import datetime
                import processos_backup
                destino = processos_backup.pasta_downloads() / f"Edodocs - de-para {datetime.now():%Y-%m-%d %Hh%M}.xlsx"
                try:
                    await run.io_bound(de_para.exportar_excel, [n for n in todos if n in marcados], destino)
                except Exception as e:
                    gui.notify(f"Não consegui exportar: {type(e).__name__} {e}", type="negative")
                    return
                janela.close()
                gui.notify(f"{len(marcados)} de-para exportado(s): {destino.name}", type="positive")
                from abrir import mostrar_na_pasta
                mostrar_na_pasta(destino)

            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Exportar", icon="file_upload", on_click=exportar).props("unelevated no-caps")
        janela.open()

    async def _importar(self) -> None:
        caminho = await consultas_ui.escolher_arquivo(self.camada, "Excel de de-para", (".xlsx", ".xlsm"))
        if not caminho:
            return
        try:
            vindos = await run.io_bound(de_para.ler_excel, caminho)
        except Exception as e:
            gui.notify(f"Não consegui abrir ({type(e).__name__}). Ele está aberto no Excel?", type="negative")
            return
        if not vindos:
            gui.notify("Não achei de-para nesse arquivo (código na coluna A, descrição na B)", type="warning")
            return
        daqui = de_para.todos()
        # Novo: trazer. Já existe: juntar (só acrescenta o que falta), trocar ou deixar como está
        escolhas = {nome: "juntar" if nome in daqui else "trazer" for nome in vindos}
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[600px] max-w-full q-pa-lg gap-2"):
            gui.label("Importar de-para").classes("text-h6 t-forte")
            gui.label(caminho.name).classes("text-caption t-suave break-all")
            for nome, pares in vindos.items():
                with gui.row().classes("w-full items-center gap-2 no-wrap"):
                    with gui.column().classes("gap-0 col"):
                        gui.label(nome).classes("text-body1 text-weight-medium t-forte")
                        gui.label(f"{len(pares)} código(s)" + (f" · aqui já existe, com {len(daqui[nome])}"
                                                                if nome in daqui else " · novo")) \
                            .classes("text-caption t-suave")
                    opcoes = ({"juntar": "Juntar (só acrescenta)", "trocar": "Trocar pelo do arquivo",
                               "nao": "Não trazer"} if nome in daqui else {"trazer": "Trazer", "nao": "Não trazer"})
                    gui.select(opcoes, value=escolhas[nome],
                               on_change=lambda e, n=nome: escolhas.update({n: e.value})) \
                        .props("outlined dense").classes("w-56")

            def importar() -> None:
                feitos = []
                for nome, escolha in escolhas.items():
                    if escolha in ("trazer", "trocar"):
                        de_para.salvar(nome, vindos[nome])
                        feitos.append(nome)
                    elif escolha == "juntar":
                        de_para.juntar(nome, vindos[nome])
                        feitos.append(nome)
                janela.close()
                gui.notify(f"De-para importado(s): {', '.join(feitos)}" if feitos else "Nada importado",
                           type="positive" if feitos else "info")
                _redesenhar_cartoes()

            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Importar", icon="file_download", on_click=importar).props("unelevated no-caps")
        janela.open()

    def desenhar(self) -> None:
        self.lista.clear()
        todos = de_para.todos()
        usados = de_para.usados()
        with self.lista:
            if not todos:
                gui.label("Nenhum ainda.").classes("text-caption t-suave")
            for nome, pares in todos.items():
                faltam = sum(1 for v in pares.values() if not v)
                with gui.row().classes("w-full items-center gap-2 no-wrap fundo-cartao rounded q-pa-sm"):
                    gui.icon("swap_horiz", color="primary", size="22px")
                    with gui.column().classes("gap-0 col"):
                        gui.label(nome).classes("text-subtitle1 text-weight-medium t-forte")
                        partes = [f"{len(pares)} código(s)"]
                        if faltam:
                            partes.append(f"{faltam} sem descrição")
                        partes.append(f"usado em: {', '.join(usados[nome])}" if usados.get(nome)
                                      else "ainda não usado em nenhum processo")
                        gui.label(" · ".join(partes)).classes(
                            "text-caption " + ("text-orange-9" if faltam else "t-suave"))
                    gui.button(icon="edit", on_click=lambda n=nome: EditorDePara(
                        self.camada, n, ao_salvar=lambda _: self.desenhar())) \
                        .props("flat round dense color=primary").tooltip("Editar")
                    gui.button(icon="delete", on_click=lambda n=nome: self._remover(n)) \
                        .props("flat round dense color=negative").tooltip("Excluir")

    def _remover(self, nome: str) -> None:
        usados = de_para.usados().get(nome)
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[460px] q-pa-lg gap-3"):
            gui.label(f"Excluir o de-para {nome}?").classes("text-h6 t-forte")
            if usados:
                gui.label(f"Ele é usado em: {', '.join(usados)}. Sem ele, esses processos recebem o código "
                          "da tabela em vez da descrição.").classes("text-body2 text-orange-9")

            def excluir() -> None:
                de_para.remover(nome)
                janela.close()
                self.desenhar()

            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Excluir", icon="delete", color="negative", on_click=excluir).props("unelevated no-caps")
        janela.open()


class EditorDePara:
    """Janela de criar/editar um de-para. codigos: já vem com estes códigos (ex.: os valores
    de uma coluna da tabela de consulta). ao_salvar(nome) é chamado ao salvar."""

    def __init__(self, camada, nome: str | None = None, codigos: list[str] | None = None,
                 ao_salvar=None, sugestao_nome: str = ""):
        self.camada = camada
        self.antigo = nome
        self.ao_salvar = ao_salvar or (lambda nome: None)
        pares = dict(de_para.todos().get(nome) or {}) if nome else {}
        self.linhas: list[dict] = [{"codigo": k, "descricao": v} for k, v in pares.items()]
        self._acrescentar({c: "" for c in codigos or []})
        self.filtro = ""
        self.so_vazios = False
        with camada, gui.dialog().props("persistent") as self.janela, \
                gui.card().classes("cartao w-[860px] max-w-full q-pa-none gap-0"):
            with gui.row().classes("w-full items-center q-px-lg q-py-md text-white no-wrap").style(
                    "background: linear-gradient(90deg,#1d4a8f,#2f6fcf);"):
                gui.icon("swap_horiz", size="26px")
                gui.label("De-para: código → descrição").classes("text-h6")
                gui.space()
                gui.button(icon="close", on_click=self.janela.close).props("flat round color=white")
            with gui.column().classes("w-full q-px-lg q-pt-md gap-2"):
                self.nome = gui.input("Nome (ex.: Secretarias, Cargos)", value=nome or sugestao_nome) \
                    .props('outlined dense stack-label bg-color=white input-class=text-grey-10 label-color=grey-8') \
                    .classes("w-full")
                with gui.row().classes("items-center gap-2"):
                    gui.label("Trazer os códigos:").classes("text-body2 t-forte")
                    gui.button("De uma tabela de consulta", icon="table_view", on_click=self._da_tabela) \
                        .props("outline dense no-caps")
                    gui.button("De um Excel com código e descrição", icon="upload_file", on_click=self._do_excel) \
                        .props("outline dense no-caps")
                    gui.button("Colar do Excel", icon="content_paste", on_click=self._colar) \
                        .props("outline dense no-caps")
                with gui.row().classes("w-full items-center gap-3 no-wrap"):
                    gui.input(placeholder="Procurar código ou descrição", on_change=self._filtrar) \
                        .props("outlined dense clearable debounce=250").classes("col")
                    gui.switch("Só os sem descrição", on_change=self._so_vazios)
                self.resumo = gui.label("").classes("text-caption t-suave")
            with gui.scroll_area().classes("w-full h-[400px] q-px-lg"):
                self.area = gui.column().classes("w-full gap-1")
            with gui.row().classes("w-full items-center q-px-lg q-py-md gap-2"):
                gui.button("Mais um código", icon="add", on_click=self._mais).props("flat no-caps")
                gui.space()
                gui.button("Cancelar", on_click=self.janela.close).props("flat no-caps")
                gui.button("Salvar", icon="save", color="positive", on_click=self._salvar).props("unelevated no-caps")
        self.janela.open()
        self._desenhar()

    # --- Lista ---
    def _acrescentar(self, pares: dict[str, str]) -> int:
        """Códigos novos entram; os que já existem só ganham a descrição se estavam vazios."""
        por_chave = {de_para.chave(l["codigo"]): l for l in self.linhas}
        novos = 0
        for codigo, descricao in pares.items():
            linha = por_chave.get(de_para.chave(codigo))
            if linha is None:
                linha = {"codigo": codigo, "descricao": descricao}
                self.linhas.append(linha)
                por_chave[de_para.chave(codigo)] = linha
                novos += 1
            elif descricao and not linha["descricao"]:
                linha["descricao"] = descricao
                novos += 1
        return novos

    def _visiveis(self) -> list[dict]:
        termo = atos.normalizar(self.filtro)
        return [l for l in self.linhas
                if (not self.so_vazios or not l["descricao"].strip())
                and (not termo or termo in atos.normalizar(l["codigo"] + l["descricao"]))]

    def _desenhar(self) -> None:
        self.area.clear()
        visiveis = self._visiveis()
        self._atualizar_resumo()
        with self.area:
            if not self.linhas:
                with gui.column().classes("w-full items-center q-pa-lg gap-1"):
                    gui.icon("swap_horiz", color="grey-5", size="40px")
                    gui.label("Traga os códigos por um dos botões acima, ou clique em Mais um código.") \
                        .classes("text-body2 t-suave text-center")
                return
            with gui.row().classes("w-full items-center gap-2 no-wrap text-caption text-weight-bold t-suave"):
                gui.label("Código (como vem na tabela)").classes("w-48")
                gui.label("").classes("w-6")
                gui.label("Descrição (como vai no ato)").classes("col")
            for linha in visiveis[:MOSTRAR]:
                with gui.row().classes("w-full items-center gap-2 no-wrap"):
                    gui.input(value=linha["codigo"], on_change=lambda e, l=linha: l.update(codigo=e.value or "")) \
                        .props("outlined dense").classes("w-48")
                    gui.icon("arrow_forward", color="primary").classes("w-6")
                    gui.input(value=linha["descricao"], placeholder="escreva a descrição",
                              on_change=lambda e, l=linha: (l.update(descricao=e.value or ""),
                                                            self._atualizar_resumo())) \
                        .props("outlined dense").classes("col")
                    gui.button(icon="close", on_click=lambda l=linha: self._tirar(l)) \
                        .props("flat round dense color=grey-7").tooltip("Tirar este código")
            if len(visiveis) > MOSTRAR:
                gui.label(f"Mostrando {MOSTRAR} de {len(visiveis)}: use a busca para achar os outros.") \
                    .classes("text-caption t-suave q-pa-sm")

    def _atualizar_resumo(self) -> None:
        faltam = sum(1 for l in self.linhas if l["codigo"].strip() and not l["descricao"].strip())
        self.resumo.set_text(f"{len(self.linhas)} código(s)" + (f" · {faltam} sem descrição (entram como "
                                                                 "código até você escrever)" if faltam else ""))

    def _filtrar(self, e) -> None:
        self.filtro = e.value or ""
        self._desenhar()

    def _so_vazios(self, e) -> None:
        self.so_vazios = bool(e.value)
        self._desenhar()

    def _mais(self) -> None:
        self.linhas.insert(0, {"codigo": "", "descricao": ""})
        self.filtro, self.so_vazios = "", False
        self._desenhar()

    def _tirar(self, linha: dict) -> None:
        self.linhas.remove(linha)
        self._desenhar()

    def _trouxe(self, novos: int) -> None:
        gui.notify(f"{novos} código(s) novo(s) ou com descrição" if novos else "Nada novo: esses códigos já estavam aqui",
                   type="positive" if novos else "info")
        self._desenhar()

    # --- Trazer os códigos ---
    def _da_tabela(self) -> None:
        tabelas = consultas.tabelas()
        if not tabelas:
            gui.notify("Nenhuma tabela de consulta ainda: cadastre em Processos → Preencher sozinho", type="warning")
            return
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[520px] q-pa-lg gap-3"):
            gui.label("Pegar os códigos de uma tabela de consulta").classes("text-h6 t-forte")
            gui.label("Escolha a coluna que tem o código. Se outra coluna tiver a descrição, escolha também; "
                      "senão, você escreve.").classes("text-body2 t-suave")
            qual = gui.select(list(tabelas), label="Tabela", value=next(iter(tabelas))) \
                .props("outlined dense").classes("w-full")
            codigo = gui.select([], label="Coluna do código").props("outlined dense use-input").classes("w-full")
            descricao = gui.select([], label="Coluna da descrição (se tiver)", clearable=True) \
                .props("outlined dense use-input").classes("w-full")
            exemplo = gui.label("").classes("text-caption t-suave")
            area_parte = gui.column().classes("w-full gap-0")
            estado = {"tabela": None, "parte": lambda: None}

            async def ler(_=None) -> None:
                try:
                    estado["tabela"] = await run.io_bound(consultas.ler_cadastrada, qual.value)
                except Exception as e:
                    gui.notify(f"Não consegui ler a tabela ({type(e).__name__})", type="negative")
                    return
                codigo.set_options(estado["tabela"].colunas, value=None)
                descricao.set_options(estado["tabela"].colunas, value=None)

            def mostrar_exemplo(_=None) -> None:
                t = estado["tabela"]
                if t and codigo.value:
                    valores = t.valores(codigo.value)
                    exemplo.set_text(f"{len(valores)} código(s) diferentes, ex.: {', '.join(valores[:5])}")
                    area_parte.clear()
                    with area_parte:
                        gui.label("Usar o código todo ou só uma parte?").classes("text-body2 t-forte")
                        estado["parte"] = campos_parte(valores)

            def trazer() -> None:
                if not (estado["tabela"] and codigo.value):
                    gui.notify("Escolha a coluna do código", type="warning")
                    return
                novos = self._acrescentar(de_para.de_tabela(estado["tabela"], codigo.value, descricao.value,
                                                            estado["parte"]()))
                janela.close()
                self._trouxe(novos)

            qual.on_value_change(ler)
            codigo.on_value_change(mostrar_exemplo)
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Trazer", icon="check", on_click=trazer).props("unelevated no-caps")
        janela.open()
        gui.timer(0.05, ler, once=True)

    async def _do_excel(self) -> None:
        caminho = await consultas_ui.escolher_arquivo(self.camada, "Planilha", (".xlsx", ".xlsm", ".csv"))
        if not caminho:
            return
        try:
            tabela = await run.io_bound(consultas.ler, caminho)
        except Exception as e:
            gui.notify(f"Não consegui abrir ({type(e).__name__}). Ela está aberta no Excel?", type="negative")
            return
        if len(tabela.colunas) < 2:
            gui.notify("A planilha precisa de uma coluna com o código e outra com a descrição", type="warning")
            return
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[520px] q-pa-lg gap-3"):
            gui.label(caminho.name).classes("text-h6 t-forte break-all")
            codigo = gui.select(tabela.colunas, label="Coluna do código", value=tabela.colunas[0]) \
                .props("outlined dense").classes("w-full")
            descricao = gui.select(tabela.colunas, label="Coluna da descrição", value=tabela.colunas[1]) \
                .props("outlined dense").classes("w-full")
            primeira = tabela.linhas[0] if tabela.linhas else {}
            gui.label(f"{len(tabela.linhas)} linha(s). Ex.: {primeira.get(tabela.colunas[0], '')} → "
                      f"{primeira.get(tabela.colunas[1], '')}").classes("text-caption t-suave")

            def trazer() -> None:
                novos = self._acrescentar(de_para.de_tabela(tabela, codigo.value, descricao.value))
                janela.close()
                self._trouxe(novos)

            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Trazer", icon="check", on_click=trazer).props("unelevated no-caps")
        janela.open()

    def _colar(self) -> None:
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[560px] q-pa-lg gap-3"):
            gui.label("Colar do Excel").classes("text-h6 t-forte")
            gui.label("No Excel, selecione as duas colunas (código e descrição), copie (Ctrl+C) e cole aqui "
                      "(Ctrl+V).").classes("text-body2 t-suave")
            texto = gui.textarea(placeholder="0312\tSecretaria Municipal de Governo\n0401\tSecretaria Municipal "
                                             "de Saúde").props("outlined autogrow").classes("w-full")

            def trazer() -> None:
                pares = de_para.de_texto(texto.value)
                if not pares:
                    gui.notify("Não achei nada para trazer", type="warning")
                    return
                novos = self._acrescentar(pares)
                janela.close()
                self._trouxe(novos)

            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Trazer", icon="check", on_click=trazer).props("unelevated no-caps")
        janela.open()

    # --- Salvar ---
    def _salvar(self) -> None:
        nome = " ".join(str(self.nome.value or "").split())
        if not nome:
            gui.notify("Dê um nome ao de-para (ex.: Secretarias)", type="warning")
            return
        if nome != self.antigo and nome in de_para.todos():
            gui.notify(f"Já existe um de-para chamado {nome}", type="warning")
            return
        pares: dict[str, str] = {}
        for l in self.linhas:
            if l["codigo"].strip():
                pares.setdefault(l["codigo"].strip(), l["descricao"].strip())
        if not pares:
            gui.notify("Traga pelo menos um código", type="warning")
            return
        de_para.salvar(nome, pares, self.antigo)
        self.janela.close()
        gui.notify(f"De-para {nome} salvo: {len(pares)} código(s)", type="positive")
        self.ao_salvar(nome)
        _redesenhar_cartoes()
