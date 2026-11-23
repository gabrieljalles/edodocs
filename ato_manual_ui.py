"""Página "Ato manual": escrever um ato que não tem processo (cada um é diferente) numa tela
igual à do e-docs e enviar pelo Edodocs, sem abrir o site.

Duas abas: "Escrever ato" (a tela do e-docs: Setor, Tipo, Data, Nossa Identificação, Texto
do documento e, à direita, Visto(s), Assinante(s), Doc. Anexo(s) e Dest. Publicador(es)) e
"Meus atos manuais" (os rascunhos e os já enviados, para ver depois). Vistos e assinantes são
perguntados no fim, antes de enviar, como na execução padrão; os anexos são escolhidos na tela.
"""
import json

from nicegui import background_tasks, run
from nicegui import ui as gui

import ato_manual
import credenciais
import main

TIPO_PADRAO = "ATOS ADMINISTRATIVOS"
# Opções de Setor/Tipo/Grupo buscadas no e-docs, guardadas para abrir já preenchido
CHAVE_OPCOES = "ato_manual_opcoes"
# Último Setor/Tipo/Grupo usados num ato manual: o próximo começa com eles
CHAVE_ULTIMO = "ato_manual_ultimo"
SITUACOES = {ato_manual.RASCUNHO: ("Rascunho", "grey-7", "edit"),
             ato_manual.ENVIADO: ("Enviado", "positive", "check_circle"),
             ato_manual.ERRO: ("Com erro", "negative", "error")}
BARRA_EDITOR = [["undo", "redo"], ["bold", "italic", "underline", "strike", "hr"], ["tamanho"],
                ["left", "center", "right", "justify"], ["outdent", "indent"],
                ["unordered", "ordered", "tabela"], ["removeFormat"]]
# Tamanhos do texto (pt) da lista "Tamanho"
TAMANHOS = [6, 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36]
# Botão de tabela (como o do e-docs): uma grade para escolher linhas × colunas
GRADE_LINHAS, GRADE_COLUNAS = 8, 10
# Passar o mouse pinta a grade até a célula (no navegador, sem ir ao servidor)
PINTAR_GRADE = """(e) => {
  const cel = e.target.closest('.celula-grade'); if (!cel) return;
  const r = +cel.dataset.r, c = +cel.dataset.c, grade = cel.closest('.grade-tabela');
  grade.querySelectorAll('.celula-grade').forEach(x =>
    x.classList.toggle('marcada', +x.dataset.r <= r && +x.dataset.c <= c));
  grade.querySelector('.rotulo-grade').textContent = r + ' × ' + c;
}"""
# Colar do Excel, como no e-docs. A tabela do Excel vem com a formatação dela: células mescladas
# (colspan/rowspan), alinhamento, negrito/itálico, bordas, cor de fundo, quebra de linha dentro da
# célula e largura das colunas (os estilos do Excel vêm em classes .xl65... no <style> do que foi
# copiado; são lidos e postos em cada célula). Fora de uma tabela, ou numa tabela ainda vazia,
# entra a tabela do Excel inteira. Numa tabela já preenchida, só o texto, a partir da célula do
# cursor, e a tabela cresce até caber. Texto separado por tab (sem HTML) vira tabela simples.
COLAR_EXCEL = r"""(e) => {
  const cd = e.clipboardData; if (!cd) return;
  const html = cd.getData('text/html') || '', txt = cd.getData('text/plain') || '';
  const limpa = s => (s || '').replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
  const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const PROPS = ['text-align', 'vertical-align', 'font-weight', 'font-style', 'text-decoration',
                 'border-top', 'border-right', 'border-bottom', 'border-left', 'background-color', 'font-size'];
  let dados = null, tabelaHtml = '';
  if (/<table/i.test(html)) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const t = doc.querySelector('table');
    if (t) {
      // Estilos das classes do Excel (.xl65 {text-align:center; font-weight:700; border-...})
      const regras = {};
      let tamanhoGeral = '';
      for (const st of doc.querySelectorAll('style')) {
        try {
          const folha = new CSSStyleSheet();
          folha.replaceSync(st.textContent.replace(/<!--|-->/g, ''));
          for (const r of folha.cssRules) for (const s of (r.selectorText || '').split(',')) {
            const m = s.trim().match(/^(?:td|th)?\.([\w-]+)$/);
            if (m) regras[m[1]] = r.style;
            // A regra geral "td {font-size:11.0pt}": o tamanho de quem não tem o seu na classe
            else if (/^td$/i.test(s.trim()) && r.style.getPropertyValue('font-size'))
              tamanhoGeral = r.style.getPropertyValue('font-size');
          }
        } catch (_) {}
      }
      const estiloDa = c => {
        const est = {};
        const copiar = d => { if (d) for (const p of PROPS) { const v = d.getPropertyValue(p); if (v) est[p] = v; } };
        for (const k of c.classList) copiar(regras[k]);
        copiar(c.style);
        // Tamanho no texto da célula (<font> ou <span style> do Excel), se for um só
        const dentro = [...c.querySelectorAll('font[style], span[style], font[size]')]
          .map(f => f.style.fontSize).filter(Boolean);
        if (dentro.length && dentro.every(t => t === dentro[0])) est['font-size'] = dentro[0];
        // Em pt, como o e-docs grava ("8.0pt" → "8pt"); o que não for pt fica de fora
        if (est['font-size']) {
          const t = est['font-size'].match(/^([\d.]+)pt$/);
          if (t) est['font-size'] = `${+t[1]}pt`; else delete est['font-size'];
        }
        // O tamanho padrão da planilha não vai: o e-docs usa o dele (como quando cola no site);
        // só vai o tamanho que a célula tem de diferente
        const geral = (tamanhoGeral.match(/^([\d.]+)pt$/) || [])[1];
        if (geral && est['font-size'] === `${+geral}pt`) delete est['font-size'];
        const alinha = (c.getAttribute('align') || '').toLowerCase();
        if (['left', 'center', 'right', 'justify'].includes(alinha)) est['text-align'] = alinha;
        if (!['left', 'center', 'right', 'justify'].includes(est['text-align'])) delete est['text-align'];
        if (est['background-color'] && /^(transparent|white|#fff(fff)?)$/i.test(est['background-color']))
          delete est['background-color'];
        for (const p in est) est[p] = est[p].replace(/windowtext/gi, 'black').replace(/\bauto\b/g, 'black');
        if ((est['font-weight'] || '') === '400' || est['font-weight'] === 'normal') delete est['font-weight'];
        // Sem alinhamento no Excel ("Geral"): centralizado e no meio, como sai no PDF do e-docs
        if (!est['text-align']) est['text-align'] = 'center';
        if (!est['vertical-align'] || est['vertical-align'] === 'bottom') est['vertical-align'] = 'middle';
        return est;
      };
      // O texto da célula; <br> (Alt+Enter no Excel) vira quebra de linha
      const conteudo = no => [...no.childNodes].map(f => f.nodeType === 3 ? esc(f.textContent.replace(/\s+/g, ' '))
        : f.nodeName === 'BR' ? '<br>' : conteudo(f)).join('');
      // Largura das colunas pelas <col> do Excel
      const larguras = [];
      for (const col of t.querySelectorAll('col'))
        for (let i = 0; i < (+col.getAttribute('span') || 1); i++) larguras.push(parseFloat(col.getAttribute('width')) || 64);
      // Como o editor do e-docs faz ao colar do Excel: a altura da linha (em pt) na 1ª célula e a
      // largura (%) em cada célula, contando as colunas ocupadas por células mescladas de cima
      const total = larguras.reduce((a, b) => a + b, 0);
      const ocupadas = [];  // ocupadas[linha] = colunas já tomadas por um rowspan
      const linhas = [...t.rows].map((r, i) => {
        const tomadas = ocupadas[i] || new Set();
        const alturaPx = parseFloat(r.getAttribute('height')) || 0;
        const alturaPt = (r.style.height.match(/^([\d.]+)pt$/) || [])[1] || (alturaPx ? alturaPx * 0.75 : 0);
        let coluna = 0;
        return '<tr>' + [...r.cells].map((c, j) => {
          while (tomadas.has(coluna)) coluna++;
          const span = c.colSpan || 1;
          for (let k = 1; k < (c.rowSpan || 1); k++)
            for (let x = 0; x < span; x++) (ocupadas[i + k] = ocupadas[i + k] || new Set()).add(coluna + x);
          const est = estiloDa(c);
          if (j === 0 && alturaPt) est['height'] = `${+(+alturaPt).toFixed(2)}pt`;
          if (total && larguras.length >= coluna + span)
            est['width'] = `${(100 * larguras.slice(coluna, coluna + span).reduce((a, b) => a + b, 0) / total).toFixed(4)}%`;
          coluna += span;
          const css = Object.entries(est).map(([p, v]) => `${p}: ${v};`).join(' ');
          const texto = conteudo(c).replace(/^(\s|<br>)+|(\s|<br>)+$/g, '');
          return `<td${c.colSpan > 1 ? ` colspan="${c.colSpan}"` : ''}${c.rowSpan > 1 ? ` rowspan="${c.rowSpan}"` : ''}`
            + `${css ? ` style="${css}"` : ''}>${texto || '&nbsp;'}</td>`;
        }).join('') + '</tr>';
      }).join('');
      // Bordas finas em todas as células, como no e-docs (border="1"); as do Excel ficam por cima
      tabelaHtml = `<table style="border-collapse: collapse; width: 100%;" border="1"><tbody>${linhas}</tbody></table>`;
      dados = [...t.rows].map(r => [...r.cells].flatMap(c =>
        [limpa(c.textContent), ...Array(Math.max((c.colSpan || 1) - 1, 0)).fill('')]));
    }
  } else if (txt.includes('\t')) {
    dados = txt.replace(/(\r?\n)+$/, '').split(/\r?\n/).map(l => l.split('\t').map(limpa));
  }
  if (!dados || !dados.length) return;
  e.preventDefault();
  const ncol = Math.max(...dados.map(r => r.length));
  if (!tabelaHtml) {
    const w = (100 / ncol).toFixed(4);
    tabelaHtml = '<table style="border-collapse: collapse; width: 100%;" border="1"><tbody>' + dados.map(r => '<tr>'
      + Array.from({length: ncol}, (_, j) => `<td style="width: ${w}%;">${r[j] ? esc(r[j]) : '&nbsp;'}</td>`).join('')
      + '</tr>').join('') + '</tbody></table>';
  }
  const caixa = e.currentTarget.querySelector('.q-editor__content') || e.currentTarget;
  const sel = window.getSelection();
  let no = sel.rangeCount ? sel.getRangeAt(0).startContainer : null;
  if (no && no.nodeType === 3) no = no.parentNode;
  const celula = no && no.closest ? no.closest('td, th') : null;
  const tabela = celula && caixa.contains(celula) ? celula.closest('table') : null;
  if (tabela && limpa(tabela.textContent)) {
    // Tabela já preenchida: só o texto, a partir da célula do cursor; a tabela cresce até caber
    const corpo = tabela.tBodies[0] || tabela;
    const nova = () => { const c = document.createElement('td'); c.innerHTML = '&nbsp;'; return c; };
    const l0 = celula.parentElement.rowIndex, c0 = celula.cellIndex;
    while (tabela.rows.length < l0 + dados.length) corpo.appendChild(document.createElement('tr'));
    const largura = Math.max(...[...tabela.rows].map(r => r.cells.length), c0 + ncol);
    for (const r of tabela.rows) while (r.cells.length < largura) r.appendChild(nova());
    dados.forEach((linha, i) => linha.forEach((v, j) => {
      tabela.rows[l0 + i].cells[c0 + j].textContent = v || ' '; }));
  } else if (tabela) {
    // Tabela vazia (a da grade): vira a do Excel, do tamanho e com a formatação dela
    tabela.insertAdjacentHTML('afterend', tabelaHtml);
    tabela.remove();
  } else {
    document.execCommand('insertHTML', false, tabelaHtml + '<p>&nbsp;</p>');
  }
  caixa.dispatchEvent(new Event('input', {bubbles: true}));
}"""
CSS = """
<style>
.grade-tabela { display: grid; grid-template-columns: repeat(%d, 18px); gap: 3px; }
.celula-grade { width: 18px; height: 18px; border: 1px solid #94a3b8; border-radius: 2px; cursor: pointer; }
.celula-grade.marcada { background: var(--q-primary); border-color: var(--q-primary); }
.nicegui-editor table { border-collapse: collapse; }
.nicegui-editor td, .nicegui-editor th { padding: 2px 6px; min-width: 24px; }
.nicegui-editor table[border="1"] td, .nicegui-editor table[border="1"] th { border: 1px solid #64748b; }
/* Célula sem borda: linha tracejada só para guiar na tela (como no e-docs); não vai para o ato */
.nicegui-editor table:not([border="1"]) td:not([style*="border"]) { outline: 1px dashed #cbd5e1; outline-offset: -1px; }
</style>""" % GRADE_COLUNAS


def tabela_html(linhas: int, colunas: int) -> str:
    """Tabela vazia como a do editor do e-docs (bordas finas, largura toda), com um parágrafo
    depois para dar para continuar escrevendo abaixo dela."""
    largura = f"{100 / colunas:.4f}".rstrip("0").rstrip(".")
    linha = "<tr>" + "".join(f'<td style="width: {largura}%;">&nbsp;</td>' for _ in range(colunas)) + "</tr>"
    return ('<table style="border-collapse: collapse; width: 100%;" border="1"><tbody>'
            + linha * linhas + "</tbody></table><p>&nbsp;</p>")


def _titulo_da_tela(icone: str, texto: str):
    """A etiqueta azul dos títulos do e-docs. Devolve o rótulo (para trocar o texto)."""
    with gui.row().classes("items-center gap-2 no-wrap q-px-sm q-py-xs rounded-borders") \
            .style("background: var(--q-primary); color: white; width: fit-content"):
        gui.icon(icone, size="18px")
        return gui.label(texto).classes("text-caption text-weight-bold")


# A escolha antes de abrir a tela: cada conta do e-docs tem uma, a outra ou as duas
ESCOLHAS = [(ato_manual.INSTRUMENTO, "gavel", "Portarias, decretos e demais atos (com vistos, assinantes e "
             "destinatários publicadores)."),
            (ato_manual.GENERICO, "description", "Atestes, ofícios e demais documentos (com assinantes; "
             "vistos só se quiser).")]


class PaginaAtoManual:
    def __init__(self, tela):
        self.tela = tela  # a janela principal (Execução, diálogos, escolher arquivos)
        self.sessao = None
        self.opcoes = self._opcoes_guardadas()
        self.ato = self._ato_novo()
        self._sujo = False
        self._carregando = False
        self._desenhar()
        self._mostrar_ato()
        gui.timer(3.0, self._autosalvar)

    # --- Montagem da página ---
    def _desenhar(self) -> None:
        with gui.tabs().classes("w-full").props("align=left no-caps inline-label") as abas:
            gui.tab("escrever", "Escrever ato", icon="edit_note")
            self.aba_lista = gui.tab("lista", "Meus atos manuais", icon="history")
        self.abas = abas
        with gui.tab_panels(abas, value="escrever", on_change=self._trocou_aba) \
                .classes("w-full bg-transparent") as self.paineis:
            with gui.tab_panel("escrever").classes("q-pa-none q-pt-sm"):
                self.area_escolha = gui.column().classes("w-full gap-3")
                self._tela_edocs()
            with gui.tab_panel("lista").classes("q-pa-none q-pt-sm"):
                with gui.row().classes("w-full items-center gap-2"):
                    self.busca = gui.input(placeholder="Procurar pela Nossa Identificação ou pelo texto",
                                           on_change=lambda: self._desenhar_lista()) \
                        .props("outlined dense clearable").classes("col")
                    with self.busca.add_slot("prepend"):
                        gui.icon("search")
                    gui.button("Escrever outro ato", icon="add", on_click=self._novo).props("outline no-caps")
                # Vários rascunhos marcados vão juntos, com os mesmos vistos e assinantes
                with gui.row().classes("w-full items-center gap-2 q-mt-sm"):
                    self.marcar_todos = gui.checkbox("Marcar todos os rascunhos", on_change=self._marcar_todos)
                    gui.label("Marque os rascunhos que quer enviar juntos: quem vista e quem assina é "
                              "perguntado uma vez só, e vale para todos.").classes("text-caption t-suave col")
                    self.bt_enviar_varios = gui.button("Enviar atos (0)", icon="outgoing_mail",
                                                       on_click=self._enviar_marcados).props("unelevated no-caps")
                self.lista = gui.column().classes("w-full gap-2 q-mt-sm")
        self.marcados: set[str] = set()
        self._contar_marcados()

    def _desenhar_escolha(self) -> None:
        """Antes da tela: qual tela do e-docs (Instrumentos Jurídicos/Normativos ou Documentos
        Genéricos). A que a conta não tem fica apagada (depois de Buscar no e-docs)."""
        acesso = self.opcoes.get("acesso") or {}
        self.area_escolha.clear()
        with self.area_escolha:
            with gui.column().classes("w-full gap-0"):
                gui.label("Que documento você vai fazer?").classes("text-h6 t-forte")
                gui.label("Escolha a tela do e-docs. Cada conta do e-docs tem uma, a outra ou as duas.") \
                    .classes("text-body2 t-suave")
            with gui.row().classes("w-full gap-4 items-stretch"):
                for modelo, icone, explicacao in ESCOLHAS:
                    sem_acesso = acesso.get(modelo) is False
                    with gui.card().classes("cartao col q-pa-lg gap-2" + ("" if sem_acesso else " cursor-pointer")) \
                            .props("flat bordered").style("min-width: 260px" + ("; opacity: .5" if sem_acesso else "")) \
                            as cartao:
                        with gui.row().classes("items-center gap-3 no-wrap"):
                            gui.icon(icone, color="primary", size="34px")
                            gui.label(ato_manual.MODELOS[modelo]).classes("text-subtitle1 text-weight-bold t-forte")
                        gui.label(explicacao).classes("text-body2 t-suave")
                        if sem_acesso:
                            gui.label("A sua conta do e-docs não tem esta tela.").classes("text-caption text-negative")
                        else:
                            gui.button("Escrever", icon="arrow_forward",
                                       on_click=lambda m=modelo: self._escolher_modelo(m)) \
                                .props("unelevated no-caps").classes("self-start q-mt-sm")
                    if not sem_acesso:
                        cartao.on("click", lambda m=modelo: self._escolher_modelo(m))
            with gui.row().classes("w-full items-center gap-2"):
                self.status_escolha = gui.label(
                    "" if acesso else "Para saber quais telas a sua conta tem, clique em Buscar no e-docs.") \
                    .classes("text-caption t-suave col")
                gui.button("Buscar no e-docs", icon="sync", on_click=lambda: self._carregar_opcoes(todos=True)) \
                    .props("flat dense no-caps size=sm")

    def _escolher_modelo(self, modelo: str) -> None:
        if self.ato.get("modelo"):
            return  # clique no botão e no cartão juntos
        self.ato = self._ato_novo(modelo)
        self._mostrar_ato()
        if not (self.opcoes.get(modelo) or {}).get("setores") and credenciais.carregar(main.CREDENCIAIS, "edocs"):
            background_tasks.create(self._carregar_opcoes())

    def _tela_edocs(self) -> None:
        with gui.card().classes("cartao w-full q-pa-md gap-2") as self.cartao_form:
            # Botões no alto, como no e-docs
            with gui.row().classes("w-full items-center gap-2"):
                self.titulo_tela = _titulo_da_tela("send", "Instrumentos Jurídicos/Normativos")
                self.selo = gui.badge("").props("outline")
                gui.space()
                gui.button("Novo", icon="note_add", on_click=self._novo).props("flat no-caps") \
                    .tooltip("Começar outro ato em branco (este fica salvo em Meus atos manuais)")
                gui.button("Visualizar", icon="visibility", on_click=self._visualizar).props("flat no-caps")
                self.bt_copia = gui.button("Fazer uma cópia", icon="content_copy", on_click=self._copiar) \
                    .props("flat no-caps").tooltip("Um ato novo com o mesmo conteúdo, para enviar outro parecido")
                self.bt_salvar = gui.button("Salvar rascunho", icon="save", on_click=self._salvar) \
                    .props("outline no-caps")
                self.bt_enviar = gui.button("Enviar ao e-docs", icon="outgoing_mail", on_click=self._enviar) \
                    .props("unelevated no-caps")
            self.faixa = gui.row().classes("w-full items-center gap-2 rounded-borders q-pa-sm fundo-previa")
            self.faixa.set_visibility(False)

            with gui.row().classes("w-full gap-3 no-wrap items-start"):
                self.sel_setor = gui.select({}, label="Setor", on_change=self._trocou_setor) \
                    .props("outlined dense stack-label options-dense").classes("col").style("min-width: 0")
                self.sel_tipo = gui.select([], label="Tipo", on_change=lambda e: self._mudou("tipo", e.value)) \
                    .props("outlined dense stack-label options-dense").classes("col").style("min-width: 0")
                self.campo_data = gui.input("Data", on_change=lambda e: self._mudou("data", e.value)) \
                    .props('outlined dense stack-label placeholder="dd/mm/aaaa" mask="##/##/####" inputmode=numeric') \
                    .classes("w-44")
                with self.campo_data.add_slot("append"):
                    with gui.icon("event").classes("cursor-pointer"):
                        with gui.menu() as menu:
                            gui.date(mask="DD/MM/YYYY", on_change=lambda: menu.close()) \
                                .bind_value(self.campo_data).props("minimal")
            with gui.row().classes("w-full items-center gap-2 no-wrap -mt-1"):
                self.status_opcoes = gui.label("").classes("text-caption t-suave col")
                self.bt_opcoes = gui.button("Buscar no e-docs", icon="sync",
                                            on_click=lambda: self._carregar_opcoes()) \
                    .props("flat dense no-caps size=sm").tooltip("Atualizar as listas com as opções do seu e-docs")
            with gui.row().classes("w-full gap-3 no-wrap"):
                self.campo_titulo = gui.input("Nossa Identificação",
                                              on_change=lambda e: self._mudou("titulo", e.value)) \
                    .props("outlined dense stack-label").classes("col") \
                    .tooltip("O título do ato no e-docs (ex.: PORTARIA SMA Nº 123/2026)")
                self.campo_doc = gui.input("Identificação Sistema").props("outlined dense stack-label readonly") \
                    .classes("col")
            # Só nos Documentos Genéricos, como na tela deles
            with gui.row().classes("w-full gap-3 no-wrap items-center") as self.linha_generico:
                self.campo_iniciais = gui.input("Iniciais do Nome no Rodapé",
                                                on_change=lambda e: self._mudou("iniciais", e.value)) \
                    .props("outlined dense stack-label").classes("col").style("max-width: 320px")
                gui.label("Incluir Visto?").classes("text-caption text-weight-bold text-primary q-ml-md")
                self.radio_visto = gui.radio({True: "Sim", False: "Não"}, value=False,
                                             on_change=self._trocou_incluir_visto).props("inline dense")

            with gui.row().classes("w-full gap-4 no-wrap items-start"):
                with gui.column().classes("col gap-1").style("min-width: 0"):
                    gui.label("Texto do documento").classes("text-caption text-weight-bold text-primary")
                    self.editor = gui.editor(on_change=lambda e: self._mudou("texto", e.value)) \
                        .classes("w-full").props('min-height="380px" max-height="60vh"')
                    self.editor._props["toolbar"] = BARRA_EDITOR
                    self._botao_tabela()
                    self._botao_tamanho()
                    self.editor.on("paste", js_handler=COLAR_EXCEL)
                # À direita, na mesma ordem do e-docs
                with gui.column().classes("gap-4 q-pt-lg").style("width: 270px"):
                    self.rotulo_pessoas = {}
                    self.bloco_pessoas = {}
                    for chave, icone, rotulo in (("vistos", "fact_check", "Visto(s)"),
                                                 ("assinantes", "draw", "Assinante(s)")):
                        with gui.column().classes("gap-1 w-full") as self.bloco_pessoas[chave]:
                            _titulo_da_tela(icone, rotulo)
                            self.rotulo_pessoas[chave] = gui.label("").classes("text-caption t-suave") \
                                .style("white-space: pre-line")
                    with gui.column().classes("gap-1 w-full"):
                        _titulo_da_tela("attach_file", "Doc. Anexo(s)")
                        self.area_anexos = gui.column().classes("w-full gap-1")
                        self.bt_anexo = gui.button("Escolher PDF(s)", icon="folder_open", on_click=self._escolher_anexos) \
                            .props("flat dense no-caps")
                    with gui.column().classes("gap-1 w-full") as self.bloco_grupo:
                        _titulo_da_tela("groups", "Dest. Publicador(es)")
                        self.sel_grupo = gui.select({}, label="Grupo", on_change=lambda e: self._mudou("grupo", e.value)) \
                            .props("outlined dense stack-label options-dense").classes("w-full")

    def _botao_tamanho(self) -> None:
        """Tamanho do texto selecionado, em pontos (como no e-docs)."""
        with self.editor.add_slot("tamanho"):
            with gui.button("Tamanho", icon="format_size").props("flat dense no-caps size=sm padding='2px 6px'") \
                    .tooltip("Tamanho do texto: selecione o trecho e escolha") as botao:
                # Nem o botão nem a lista pegam o foco: o trecho selecionado continua selecionado
                botao.on("mousedown", js_handler="(e) => e.preventDefault()")
                with gui.menu().props("no-focus no-refocus") as menu:
                    # Clicar no menu não tira a seleção do texto
                    with gui.list().props("dense").on("mousedown", js_handler="(e) => e.preventDefault()"):
                        for pt in [None, *TAMANHOS]:
                            with gui.item(on_click=lambda pt=pt: self._mudar_tamanho(menu, pt)):
                                with gui.item_section():
                                    gui.item_label("Normal" if pt is None else f"{pt} pt") \
                                        .style("" if pt is None else f"font-size: {min(pt, 20)}pt")
        self.botao_tamanho = botao

    async def _mudar_tamanho(self, menu, pt: int | None) -> None:
        menu.close()
        if self.enviado:
            return
        # fontSize 7 marca o trecho (<font size="7">, ou <span> se o navegador usar CSS) e aí ele
        # vira <span style="font-size: Npt">, como o e-docs grava. "Normal" tira o tamanho.
        mudou = await gui.run_javascript(f"""
            const ed = getElement({self.editor.id}).$refs.qRef, caixa = ed.getContentEl();
            const sel = window.getSelection();
            if (!sel.rangeCount || sel.isCollapsed || !caixa.contains(sel.anchorNode)) return false;
            ed.runCmd('fontSize', '7');
            const marcados = [...caixa.querySelectorAll('font[size="7"]'),
                              ...[...caixa.querySelectorAll('span')].filter(s => /xxx-large/.test(s.style.fontSize))];
            for (const f of marcados) {{
              const s = document.createElement('span');
              {"" if pt is None else f"s.style.fontSize = '{pt}pt';"}
              s.innerHTML = f.innerHTML;
              // tamanhos de antes dentro do trecho saem (vale o novo)
              s.querySelectorAll('span').forEach(x => x.style.fontSize = '');
              if (s.getAttribute('style')) f.replaceWith(s); else f.replaceWith(...s.childNodes);
            }}
            caixa.dispatchEvent(new Event('input', {{bubbles: true}}));
            return true;""")
        if not mudou:
            gui.notify("Selecione o trecho do texto e depois escolha o tamanho", type="info")

    def _botao_tabela(self) -> None:
        """O botão de tabela na barra do editor: abre a grade; clicar insere a tabela no cursor."""
        gui.add_head_html(CSS)
        with self.editor.add_slot("tabela"):
            with gui.button(icon="grid_on").props("flat dense size=sm padding='4px'") \
                    .tooltip("Tabela: inserir ou apagar") as botao:
                with gui.menu().props("anchor='bottom left' self='top left'") as menu:
                    # Clicar no menu não tira o cursor do texto (é nele que a tabela entra ou sai)
                    with gui.column().classes("q-pa-sm gap-1") \
                            .on("mousedown", js_handler="(e) => e.preventDefault()"):
                        with gui.element("div").classes("grade-tabela").on("mouseover", js_handler=PINTAR_GRADE):
                            for r in range(1, GRADE_LINHAS + 1):
                                for c in range(1, GRADE_COLUNAS + 1):
                                    gui.element("div").classes("celula-grade") \
                                        .props(f'data-r="{r}" data-c="{c}"') \
                                        .on("click", lambda r=r, c=c: (menu.close(), self._inserir_tabela(r, c)))
                            gui.label("Escolha o tamanho").classes("rotulo-grade text-caption t-suave") \
                                .style(f"grid-column: 1 / span {GRADE_COLUNAS}")
                        gui.separator().classes("q-my-xs")
                        gui.button("Apagar a tabela onde está o cursor", icon="delete", color="negative",
                                   on_click=lambda: self._apagar_tabela(menu)) \
                            .props("flat dense no-caps size=sm")
        self.botao_tabela = botao

    async def _apagar_tabela(self, menu) -> None:
        """Tira a tabela em que o cursor está (clique numa célula dela antes)."""
        menu.close()
        if self.enviado:
            return
        apagou = await gui.run_javascript(f"""
            const ed = getElement({self.editor.id}).$refs.qRef, caixa = ed.getContentEl();
            const sel = window.getSelection();
            let no = sel.rangeCount ? sel.getRangeAt(0).startContainer : null;
            if (no && no.nodeType === 3) no = no.parentNode;
            const tabela = no && no.closest ? no.closest('table') : null;
            if (!tabela || !caixa.contains(tabela)) return false;
            tabela.remove();
            caixa.dispatchEvent(new Event('input', {{bubbles: true}}));
            return true;""")
        if not apagou:
            gui.notify("Clique dentro de uma célula da tabela que quer apagar e tente de novo", type="info")

    def _inserir_tabela(self, linhas: int, colunas: int) -> None:
        if self.enviado:
            return
        # O QEditor guarda onde estava o cursor e o devolve antes do comando
        gui.run_javascript(f"getElement({self.editor.id}).$refs.qRef.runCmd('insertHTML', "
                           f"{json.dumps(tabela_html(linhas, colunas))})")

    # --- Ato em tela ---
    def _ato_novo(self, modelo: str = "") -> dict:
        """Ato em branco. Sem modelo, a página mostra antes a escolha da tela."""
        ato = ato_manual.novo(modelo)
        if not modelo:
            return ato
        ultimo = self._ultimos().get(modelo) or {}
        ato.update({k: ultimo[k] for k in ("setor_posicao", "setor", "tipo", "grupo", "iniciais") if ultimo.get(k)})
        if not ato["tipo"] and modelo == ato_manual.INSTRUMENTO:
            ato["tipo"] = TIPO_PADRAO
        return ato

    @staticmethod
    def _ultimos() -> dict:
        """Último Setor/Tipo/Grupo de cada tela ({modelo: {...}}). O formato de antes (só
        Instrumentos) vale como Instrumentos."""
        ultimo = main.ler_config().get(CHAVE_ULTIMO) or {}
        if "tipo" in ultimo or "setor" in ultimo:
            ultimo = {ato_manual.INSTRUMENTO: ultimo}
        return ultimo

    @staticmethod
    def _opcoes_guardadas() -> dict:
        """As listas do e-docs guardadas, só se forem da conta cadastrada agora (cada conta
        tem as suas): {"cpf", "acesso": {modelo: bool}, <modelo>: {"setores", "tipos", "grupos"}}."""
        guardadas = main.ler_config().get(CHAVE_OPCOES) or {}
        dados = credenciais.carregar(main.CREDENCIAIS, "edocs")
        if not dados or guardadas.get("cpf") != main._cpf(dados[0]):
            return {}
        return guardadas

    @property
    def op(self) -> dict:
        """As listas do e-docs da tela do ato aberto."""
        return self.opcoes.get(self.ato.get("modelo") or ato_manual.INSTRUMENTO) or {}

    @property
    def enviado(self) -> bool:
        return self.ato["situacao"] == ato_manual.ENVIADO

    def _mostrar_ato(self) -> None:
        """Põe o ato (self.ato) nos campos. Ato novo sem tela escolhida: mostra a escolha."""
        escolher = not self.ato.get("modelo")
        self.area_escolha.set_visibility(escolher)
        self.cartao_form.set_visibility(not escolher)
        if escolher:
            self._desenhar_escolha()
            self._sujo = False
            return
        generico = self.ato["modelo"] == ato_manual.GENERICO
        self.titulo_tela.text = ato_manual.nome_do_modelo(self.ato)
        self.linha_generico.set_visibility(generico)
        self.bloco_grupo.set_visibility(not generico)
        self.bloco_pessoas["vistos"].set_visibility(ato_manual.pede_vistos(self.ato))
        self._carregando = True
        try:
            self._listas()
            self.campo_iniciais.value = self.ato.get("iniciais", "")
            self.radio_visto.value = bool(self.ato.get("incluir_visto"))
            self.campo_data.value = self.ato["data"]
            self.campo_titulo.value = self.ato["titulo"]
            self.campo_doc.value = self.ato["doc"] or "o e-docs cria ao enviar"
            self.editor.value = self.ato["texto"]
        finally:
            self._carregando = False
        self._sujo = False
        so_ver = self.enviado
        for campo in (self.sel_setor, self.sel_tipo, self.campo_data, self.campo_titulo, self.sel_grupo,
                      self.editor, self.campo_iniciais):
            campo.props("readonly") if so_ver else campo.props(remove="readonly")
        self.radio_visto.set_enabled(not so_ver)
        self.botao_tabela.set_enabled(not so_ver)
        self.botao_tamanho.set_enabled(not so_ver)
        self.bt_anexo.set_visibility(not so_ver)
        self.bt_salvar.set_visibility(not so_ver)
        self.bt_enviar.set_visibility(not so_ver)
        self.bt_copia.set_visibility(bool(self.ato["id"]))
        for chave, rotulo in self.rotulo_pessoas.items():
            pessoas = self.ato.get(chave) or []
            rotulo.text = ("\n".join(f"{i}º {p['nome']}" for i, p in enumerate(pessoas, 1)) or "ninguém") \
                if so_ver else "Perguntado no fim, antes de enviar"
        self._desenhar_anexos()
        self._desenhar_situacao()

    def _listas(self) -> None:
        """Setor, Tipo e Grupo com as opções do e-docs (ou só o valor do ato, sem elas)."""
        setores = self.op.get("setores") or []
        posicao = int(self.ato.get("setor_posicao") or 0)
        if setores:
            if not 1 <= posicao <= len(setores):
                posicao = 1
            self.sel_setor.set_options({i: f"{i}º · {t}" for i, t in enumerate(setores, 1)}, value=posicao)
            self.ato["setor_posicao"], self.ato["setor"] = posicao, setores[posicao - 1]
        else:
            self.sel_setor.set_options({posicao or 1: self.ato.get("setor") or "o 1º setor da sua lista"},
                                       value=posicao or 1)
        tipos = list((self.op.get("tipos") or {}).get(str(self.ato.get("setor_posicao") or 1)) or [])
        if not self.ato["tipo"] and tipos and not self.enviado:
            self.ato["tipo"] = tipos[0]
        if self.ato["tipo"] and self.ato["tipo"] not in tipos:
            tipos.append(self.ato["tipo"])
        self.sel_tipo.set_options(tipos, value=self.ato["tipo"] or None)
        grupos = {g: g for g in self.op.get("grupos") or []}
        if self.ato["grupo"] and self.ato["grupo"] not in grupos:
            grupos[self.ato["grupo"]] = self.ato["grupo"]
        if not self.ato["grupo"] and grupos:
            self.ato["grupo"] = next(iter(grupos))
        grupos = grupos or {"": "o 1º grupo da sua lista"}
        self.sel_grupo.set_options(grupos, value=self.ato["grupo"])
        generico = self.ato.get("modelo") == ato_manual.GENERICO
        if self.op:
            self.status_opcoes.text = (f"Listas do seu e-docs: {len(setores)} setor(es)"
                                       + ("." if generico else
                                          f", {len(self.op.get('grupos') or [])} grupo(s) de destinatários."))
        else:
            self.status_opcoes.text = (f"As listas de Setor, Tipo{'' if generico else ' e Grupo'} vêm do seu "
                                       "e-docs: clique em Buscar no e-docs.")

    def _desenhar_situacao(self) -> None:
        nome, cor, _ = SITUACOES.get(self.ato["situacao"], SITUACOES[ato_manual.RASCUNHO])
        self.selo.text = nome if self.ato["id"] else "Novo"
        self.selo.props(f"color={cor if self.ato['id'] else 'primary'}")
        self.faixa.clear()
        mensagem = ""
        if self.enviado:
            mensagem = (f"Enviado em {self.ato['enviado_em']} · documento {self.ato['doc']}. Para mandar outro "
                        "parecido, use “Fazer uma cópia”.")
        elif self.ato["situacao"] == ato_manual.ERRO:
            mensagem = f"Da última vez não deu certo: {self.ato['erro']}"
        self.faixa.set_visibility(bool(mensagem))
        if mensagem:
            with self.faixa:
                gui.icon("check_circle" if self.enviado else "error", color="positive" if self.enviado else "negative")
                gui.label(mensagem).classes("col text-body2")

    def _desenhar_anexos(self) -> None:
        self.area_anexos.clear()
        with self.area_anexos:
            if not self.ato["anexos"]:
                gui.label("Nenhum.") \
                    .classes("text-caption t-suave")
            for caminho in self.ato["anexos"]:
                with gui.row().classes("w-full items-center gap-1 no-wrap"):
                    gui.icon("picture_as_pdf", color="negative", size="18px")
                    gui.label(caminho.replace("\\", "/").rsplit("/", 1)[-1]).classes("text-body2 col ellipsis") \
                        .tooltip(caminho)
                    if not self.enviado:
                        gui.button(icon="close", on_click=lambda c=caminho: self._tirar_anexo(c)) \
                            .props("flat round dense size=sm").tooltip("Tirar este anexo")

    # --- Mudanças ---
    def _mudou(self, campo: str, valor) -> None:
        if self._carregando or self.enviado:
            return
        self.ato[campo] = valor if valor is not None else ""
        self._sujo = True

    def _trocou_setor(self, e) -> None:
        if self._carregando or self.enviado or not e.value:
            return
        setores = self.op.get("setores") or []
        self.ato["setor_posicao"] = int(e.value)
        if setores and int(e.value) <= len(setores):
            self.ato["setor"] = setores[int(e.value) - 1]
        self._sujo = True
        if self.op and str(e.value) not in (self.op.get("tipos") or {}):
            background_tasks.create(self._carregar_opcoes())  # os tipos podem mudar com o setor
        else:
            self._carregando = True
            try:
                self._listas()
            finally:
                self._carregando = False

    def _trocou_incluir_visto(self, e) -> None:
        self._mudou("incluir_visto", bool(e.value))
        self.bloco_pessoas["vistos"].set_visibility(ato_manual.pede_vistos(self.ato))

    async def _escolher_anexos(self) -> None:
        recusados = []
        for arquivo in await self.tela._escolher_arquivos_do_computador():
            if arquivo.suffix.lower() != ".pdf":  # o e-docs só recebe PDF como anexo
                recusados.append(arquivo.name)
            elif str(arquivo) not in self.ato["anexos"]:
                self.ato["anexos"].append(str(arquivo))
                self._sujo = True
        if recusados:
            gui.notify("Só aceita PDF. Não entrou: " + ", ".join(recusados), type="warning", multi_line=True)
        self._desenhar_anexos()

    def _tirar_anexo(self, caminho: str) -> None:
        self.ato["anexos"].remove(caminho)
        self._sujo = True
        self._desenhar_anexos()

    # --- Opções do e-docs ---
    def _avisar_opcoes(self, texto: str) -> None:
        self.status_opcoes.text = texto
        if not self.ato.get("modelo") and getattr(self, "status_escolha", None) is not None:
            self.status_escolha.text = texto

    async def _carregar_opcoes(self, todos: bool = False) -> None:
        """Busca no e-docs, com a conta cadastrada, as listas de Setor/Tipo (e Grupo, nos
        Instrumentos) da tela do ato. todos=True (ou na escolha da tela): das duas telas, e
        descobre qual delas a conta tem."""
        import edocs_http

        dados = credenciais.carregar(main.CREDENCIAIS, "edocs")
        if not dados:
            self._avisar_opcoes("Cadastre seu usuário do e-docs em Configurações para buscar as suas listas.")
            return
        atual = self.ato.get("modelo")
        modelos = [atual] if atual and not todos else list(ato_manual.MODELOS)
        posicao = int(self.ato.get("setor_posicao") or 1)
        self._avisar_opcoes("Buscando as suas opções no e-docs...")
        self.bt_opcoes.props("loading")
        achadas, acesso, grupos = {}, {}, None
        try:
            for _ in range(2):
                try:
                    if self.sessao is None:
                        self.sessao = await run.io_bound(edocs_http.entrar, *dados)
                    for modelo in modelos:
                        try:
                            achadas[modelo] = await run.io_bound(
                                edocs_http.opcoes_do_documento, self.sessao,
                                posicao if modelo == atual else 1, modelo)
                        except edocs_http.SessaoExpirada:
                            raise
                        except edocs_http.ErroEdocs:
                            achadas[modelo] = {"setores": [], "tipos": []}  # a tela não abriu: sem acesso
                        acesso[modelo] = bool(achadas[modelo]["setores"])
                    if acesso.get(ato_manual.INSTRUMENTO):
                        try:
                            grupos = await run.io_bound(edocs_http.grupos_destinatarios, self.sessao)
                        except edocs_http.ErroEdocs:
                            grupos = []
                    break
                except edocs_http.SessaoExpirada:
                    self.sessao = None
            else:
                raise RuntimeError("a sessão do e-docs caiu; tente de novo")
        except Exception as e:
            self.sessao = None
            self._avisar_opcoes(f"Não deu para buscar as opções no e-docs: {e}")
            return
        finally:
            self.bt_opcoes.props(remove="loading")
        cpf = main._cpf(dados[0])
        if self.opcoes.get("cpf") != cpf:
            self.opcoes = {"cpf": cpf}  # outra conta: as listas da anterior não valem
        self.opcoes["acesso"] = (self.opcoes.get("acesso") or {}) | acesso
        for modelo, opcoes in achadas.items():
            antes = self.opcoes.get(modelo) or {}
            pos = posicao if modelo == atual else 1
            self.opcoes[modelo] = {"setores": opcoes["setores"],
                                   "tipos": dict(antes.get("tipos") or {}) | {str(pos): opcoes["tipos"]},
                                   "grupos": (grupos or antes.get("grupos") or [])
                                   if modelo == ato_manual.INSTRUMENTO else []}
        main.salvar_config(**{CHAVE_OPCOES: self.opcoes})
        if not atual:  # ainda na escolha da tela: mostra qual a conta tem
            self._desenhar_escolha()
            sem = [ato_manual.MODELOS[m] for m, tem in acesso.items() if not tem]
            self.status_escolha.text = ("Pronto." if not sem else
                                        "A sua conta não tem: " + ", ".join(sem) + ".")
            return
        tipos = (achadas.get(atual) or {}).get("tipos") or []
        if self.ato["tipo"] not in tipos and tipos and not self.enviado:
            self.ato["tipo"] = TIPO_PADRAO if TIPO_PADRAO in tipos else tipos[0]
        self._carregando = True
        try:
            self._listas()
        finally:
            self._carregando = False
        if acesso.get(atual) is False:
            self.status_opcoes.text = (f"A sua conta do e-docs não abriu a tela {ato_manual.MODELOS[atual]}: "
                                       "confira se ela tem acesso a esta tela.")

    # --- Guardar ---
    def _tem_conteudo(self) -> bool:
        if not self.ato.get("modelo"):
            return False
        inicial = ato_manual.texto_inicial(modelo=self.ato["modelo"])
        return bool(str(self.ato["titulo"]).strip() or self.ato["anexos"]
                    or ato_manual.texto_puro(self.ato["texto"]) != ato_manual.texto_puro(inicial))

    def _guardar(self) -> None:
        ato_manual.salvar(self.ato)
        self._sujo = False
        ultimos = self._ultimos()
        ultimos[self.ato["modelo"]] = {k: self.ato.get(k) for k in ("setor_posicao", "setor", "tipo", "grupo", "iniciais")}
        main.salvar_config(**{CHAVE_ULTIMO: ultimos})
        self.bt_copia.set_visibility(True)
        self._desenhar_situacao()

    async def _autosalvar(self) -> None:
        """Rascunho salvo sozinho, alguns segundos depois de mexer (não se perde nada)."""
        if self._sujo and not self.enviado and self._tem_conteudo():
            await run.io_bound(self._guardar)

    def _salvar(self) -> None:
        self._guardar()
        gui.notify("Rascunho salvo em Meus atos manuais", type="positive")

    def _novo(self) -> None:
        if self._sujo and not self.enviado and self._tem_conteudo():
            self._guardar()
        self.ato = self._ato_novo()
        self._mostrar_ato()
        self.paineis.set_value("escrever")

    def _copiar(self) -> None:
        if self._sujo and not self.enviado and self._tem_conteudo():
            self._guardar()
        self.ato = ato_manual.copia(self.ato)
        self._mostrar_ato()
        gui.notify("Cópia aberta: ajuste e envie (o ato de antes continua como estava)", type="info")

    def _abrir(self, ato_id: str) -> None:
        if self._sujo and not self.enviado and self._tem_conteudo():
            self._guardar()
        if ato := ato_manual.ler(ato_id):
            self.ato = ato
            self._mostrar_ato()
            self.paineis.set_value("escrever")

    # --- Ver e enviar ---
    def _visualizar(self, ato: dict | None = None) -> None:
        ato = ato or self.ato
        with self.tela.camada, gui.dialog() as dialogo, gui.card().classes("cartao w-[860px] max-w-full q-pa-lg"):
            with gui.row().classes("w-full items-center gap-2"):
                gui.icon("visibility", color="primary", size="26px")
                gui.label(ato["titulo"] or "(sem Nossa Identificação)").classes("text-h6 col")
                gui.button(icon="close", on_click=dialogo.close).props("flat round dense")
            gui.html(f'<div class="folha-area" style="height:480px">{ato_manual.folha(ato)}</div>').classes("w-full")
            if ato["anexos"]:
                gui.label("Anexos: " + ", ".join(a.replace("\\", "/").rsplit("/", 1)[-1] for a in ato["anexos"])) \
                    .classes("text-caption t-suave")
        dialogo.open()

    def _enviar(self) -> None:
        if faltam := ato_manual.problemas(self.ato):
            gui.notify("Falta preencher: " + "; ".join(faltam), type="warning", multi_line=True)
            return
        if self.tela.rodando:
            gui.notify("Espere terminar o que está em execução", type="warning")
            return
        self._guardar()
        ato_id, titulo = self.ato["id"], self.ato["titulo"]
        self.tela.iniciar(f"Ato manual · {titulo}", lambda: ato_manual.enviar(ato_id),
                          subtitulo="Ato manual (sem processo)")

    def depois_da_execucao(self) -> None:
        """A Execução terminou: o ato pode ter virado enviado (ou erro)."""
        if self.ato.get("id") and (ato := ato_manual.ler(self.ato["id"])):
            self.ato = ato
            self._mostrar_ato()
        self._desenhar_lista()

    # --- Meus atos manuais ---
    def ao_abrir(self) -> None:
        """A página foi aberta pelo menu: lista atualizada e, na 1ª vez, as listas do e-docs."""
        self._desenhar_lista()
        if not self.opcoes.get("acesso") and not getattr(self, "_buscou", False) \
                and credenciais.carregar(main.CREDENCIAIS, "edocs"):
            self._buscou = True
            background_tasks.create(self._carregar_opcoes(todos=True))

    def _trocou_aba(self, e) -> None:
        if e.value == "lista":
            if self._sujo and not self.enviado and self._tem_conteudo():
                self._guardar()
            self._desenhar_lista()

    # --- Enviar vários de uma vez ---
    def _pode_marcar(self, ato: dict) -> bool:
        return ato["situacao"] != ato_manual.ENVIADO

    def _contar_marcados(self) -> None:
        self.bt_enviar_varios.text = f"Enviar atos ({len(self.marcados)})"
        self.bt_enviar_varios.set_enabled(bool(self.marcados))

    def _marcar(self, ato_id: str, valor: bool) -> None:
        (self.marcados.add if valor else self.marcados.discard)(ato_id)
        self._contar_marcados()

    def _marcar_todos(self, e) -> None:
        if getattr(self, "_redesenhando", False):
            return
        self.marcados = {a["id"] for a in self._visiveis if self._pode_marcar(a)} if e.value else set()
        self._desenhar_lista()

    def _enviar_marcados(self) -> None:
        if self.tela.rodando:
            gui.notify("Espere terminar o que está em execução", type="warning")
            return
        if self._sujo and not self.enviado and self._tem_conteudo():
            self._guardar()  # o ato aberto na tela pode estar entre os marcados
        atos = [a for a in ato_manual.listar() if a["id"] in self.marcados]
        com_falta = [f"{a['titulo'] or '(sem Nossa Identificação)'}: falta " + "; ".join(f)
                     for a in atos if (f := ato_manual.problemas(a))]
        if com_falta:
            gui.notify("Complete antes de enviar: " + " · ".join(com_falta), type="warning", multi_line=True,
                       timeout=8000)
            return
        ids = [a["id"] for a in sorted(atos, key=lambda a: a["id"])]  # na ordem em que foram criados
        self.marcados = set()
        self._contar_marcados()
        self.tela.iniciar(f"Atos manuais · {len(ids)} ato(s)", lambda: ato_manual.enviar_varios(ids),
                          subtitulo="Atos manuais (sem processo): os mesmos vistos e assinantes para todos")

    def _desenhar_lista(self) -> None:
        todos = ato_manual.listar()
        existentes = {a["id"] for a in todos if self._pode_marcar(a)}
        self.marcados &= existentes  # o que foi enviado ou apagado sai da marcação
        self._contar_marcados()
        self.aba_lista.props(f'label="Meus atos manuais ({len(todos)})"')
        termo = (self.busca.value or "").strip().lower()
        if termo:
            todos = [a for a in todos if termo in a["titulo"].lower()
                     or termo in ato_manual.texto_puro(a["texto"]).lower()]
        self._visiveis = todos
        marcaveis = [a["id"] for a in todos if self._pode_marcar(a)]
        self._redesenhando = True
        self.marcar_todos.value = bool(marcaveis) and all(i in self.marcados for i in marcaveis)
        self._redesenhando = False
        self.marcar_todos.set_visibility(bool(marcaveis))
        self.lista.clear()
        with self.lista:
            if not todos:
                with gui.column().classes("w-full items-center q-pa-xl gap-2 t-suave"):
                    gui.icon("edit_note", size="48px")
                    gui.label("Nenhum ato manual ainda." if not termo else "Nada encontrado.")
            for ato in todos:
                self._cartao(ato)

    def _cartao(self, ato: dict) -> None:
        nome, cor, icone = SITUACOES.get(ato["situacao"], SITUACOES[ato_manual.RASCUNHO])
        texto = " ".join(ato_manual.texto_puro(ato["texto"]).split())
        with gui.card().classes("cartao w-full q-pa-sm").props("flat bordered"):
            with gui.row().classes("w-full items-center gap-3 no-wrap"):
                if self._pode_marcar(ato):
                    gui.checkbox(value=ato["id"] in self.marcados,
                                 on_change=lambda e, i=ato["id"]: self._marcar(i, e.value)) \
                        .tooltip("Marcar para enviar junto com outros")
                else:
                    gui.element("div").style("width: 40px")  # alinhado com os que têm caixinha
                gui.icon(icone, color=cor, size="26px")
                # O texto ocupa só o que sobra e corta com "…"; os botões ficam num bloco fixo à direita
                with gui.column().classes("col gap-0").style("min-width: 0; overflow: hidden"):
                    with gui.row().classes("w-full items-center gap-2 no-wrap").style("min-width: 0"):
                        gui.label(ato["titulo"] or "(sem Nossa Identificação)") \
                             .classes("text-subtitle1 text-weight-medium t-forte ellipsis").style("min-width: 0")
                        gui.badge(nome, color=cor).props("outline").style("flex-shrink: 0")
                    gui.label(texto[:300]).classes("text-body2 t-suave ellipsis w-full").tooltip(texto[:600])
                    detalhes = [ato_manual.nome_do_modelo(ato), f"Data {ato['data']}", ato["tipo"]]
                    if ato["doc"]:
                        detalhes.append(f"documento {ato['doc']}")
                    if ato.get("criado_por"):
                        detalhes.append(f"criado por {ato['criado_por']}")
                    detalhes.append(f"enviado em {ato['enviado_em']}" if ato["enviado_em"]
                                    else f"alterado em {ato['alterado']}")
                    gui.label(" · ".join(d for d in detalhes if d)).classes("text-caption t-fraco ellipsis w-full")
                enviado = ato["situacao"] == ato_manual.ENVIADO
                with gui.row().classes("items-center gap-1 no-wrap").style("flex-shrink: 0"):
                    gui.button("Ver" if enviado else "Continuar", icon="open_in_new" if enviado else "edit",
                               on_click=lambda a=ato: self._abrir(a["id"])).props("flat dense no-caps") \
                         .tooltip("Abrir na tela do ato")
                    gui.button(icon="visibility", on_click=lambda a=ato: self._visualizar(a)) \
                         .props("flat round dense").tooltip("Ver como ficou o texto")
                    gui.button(icon="content_copy",
                               on_click=lambda a=ato: (self._abrir(a["id"]), self._copiar())) \
                         .props("flat round dense").tooltip("Fazer uma cópia (um ato novo com o mesmo conteúdo)")
                    gui.button(icon="delete", color="negative", on_click=lambda a=ato: self._confirmar_exclusao(a)) \
                         .props("flat round dense").tooltip("Apagar daqui")

    def _confirmar_exclusao(self, ato: dict) -> None:
        enviado = ato["situacao"] == ato_manual.ENVIADO
        with self.tela.camada, gui.dialog() as dialogo, gui.card().classes("cartao w-[480px] max-w-full q-pa-lg"):
            gui.label("Apagar este ato manual?").classes("text-h6")
            gui.label(ato["titulo"] or "(sem Nossa Identificação)").classes("text-subtitle2 t-forte")
            gui.label("Ele sai só da lista do Edodocs: o documento já enviado continua no e-docs, nada é "
                      "cancelado lá." if enviado else "O rascunho será apagado e não dá para desfazer.") \
                .classes("text-body2 t-suave")
            with gui.row().classes("w-full justify-end gap-2 q-mt-md"):
                gui.button("Não", on_click=dialogo.close).props("flat no-caps")

                def apagar() -> None:
                    dialogo.close()
                    ato_manual.excluir(ato["id"])
                    if self.ato.get("id") == ato["id"]:
                        self.ato = self._ato_novo()
                        self._mostrar_ato()
                    self._desenhar_lista()
                    gui.notify("Apagado", type="info")

                # Já em foco ao abrir: Enter apaga (Esc ou "Não" desiste)
                gui.button("Apagar", icon="delete", color="negative", on_click=apagar) \
                    .props("unelevated no-caps autofocus")
        dialogo.open()
