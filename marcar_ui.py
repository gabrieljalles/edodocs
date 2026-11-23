"""Marcar no Word o que muda de um ato para outro (assistente, passo 2, para quem só tem o
Word): a pessoa vê o texto do ato, seleciona com o mouse um trecho ("ANTONIO EXEMPLO DOS
SANTOS"), clica em "Marcar o que selecionei" e dá um nome ("NOME"). O programa também
sugere trechos que parecem dados (nomes em maiúsculas, matrículas, datas).

Cada marca pode ser:
    texto       o valor vem de uma coluna da planilha com esse nome
    data        a coluna tem a data (dd/mm/aaaa); no ato sai como foi escrita (por extenso ou não)
    opcoes      escolha entre opções (ex.: A PEDIDO: SIM escreve "a pedido, ", NÃO não escreve)
    genero      muda com o sexo da pessoa (coluna SEXO): "nomeado" / "nomeada"
    quantidade  muda com o número de pessoas do ato: "Fica nomeado" / "Ficam nomeados"
    as_vezes    trecho que só aparece às vezes: quando a pessoa responde SIM numa coluna própria
                ou quando outra coluna está preenchida (", a contar de 1º de outubro de 2026")

Um trecho de escolha / às vezes / uma ou várias pode ter outros campos já marcados dentro
dele (", a contar de {{DATA}}"): no texto da regra eles ficam como {{CAMPO}}.

No fim, criar_modelo() troca os trechos por {{NOME}} (mantendo a formatação do Word) e
devolve as ligações do processo e uma linha de exemplo (os valores que estavam no Word),
para a prévia funcionar antes de a planilha ter dados.
"""
import html
import re
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from nicegui import ui as gui

import atos
import modelo_automatico

MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro",
         "outubro", "novembro", "dezembro"]
DATA_EXTENSO = re.compile(r"\b(\d{1,2})º?\s+de\s+(" + "|".join(MESES) + r")\s+de\s+(\d{4})\b", re.IGNORECASE)
DATA_NUMEROS = re.compile(r"\b\d{1,2}/\d{1,2}/\d{4}\b")
MATRICULA = re.compile(r"\b\d{1,3}\.\d{3}-\d\b")
CPF = re.compile(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b")
# Nome em maiúsculas: 3 ou mais palavras (evita "DECRETA:", "DECRETO S/Nº")
NOME_MAIUSCULO = re.compile(r"\b[A-ZÀ-Ý]{2,}(?:\s+(?:DA|DE|DO|DAS|DOS|E|[A-ZÀ-Ý]{2,})){2,}\b")
NOMES_COMUNS = ["NOME", "MATRÍCULA", "CPF", "CARGO", "SECRETARIA", "NÚCLEO", "LOTAÇÃO", "DATA",
                "DATA DE INÍCIO", "DIAS", "PADRÃO", "NÍVEL", "A PEDIDO", "PROCESSO", "Nº DO ATO"]
TIPOS = {"texto": "Texto que muda (nome, matrícula, cargo...)",
         "data": "Data",
         "opcoes": "Escolha entre opções (ex.: a pedido: sim ou não)",
         "genero": "Muda com o sexo da pessoa (ex.: nomeado / nomeada)",
         "quantidade": "Muda com uma pessoa ou várias (ex.: Fica nomeado / Ficam nomeados)",
         "as_vezes": "Só aparece às vezes (ex.: “, a contar de 1º de outubro”)"}
# Tipos cujo texto pode ter outros campos dentro (o resto é um valor só)
COM_CAMPOS_DENTRO = {"opcoes", "quantidade", "as_vezes"}
SEXO = "SEXO"


def genero_sugerido(texto: str) -> tuple[str, str]:
    """(masculino, feminino) a partir de uma das formas: "nomeada" → ("nomeado", "nomeada")."""
    t = texto
    for fem, masc in (("as", "os"), ("a", "o")):
        if t.endswith(fem):
            return t[:-len(fem)] + masc, t
        if t.endswith(masc):
            return t, t[:-len(masc)] + fem
    return t, t

SCRIPT_SELECAO = """
// Guarda o trecho selecionado no texto do Word (parágrafo e posição dentro dele)
if (!window.__edodocsOuvindo) {
window.__edodocsOuvindo = true;
document.addEventListener('mouseup', () => {
  const s = window.getSelection();
  if (!s || !s.rangeCount || s.isCollapsed) return;
  const r = s.getRangeAt(0);
  const no = n => (n.nodeType === 3 ? n.parentElement : n);
  const p = no(r.startContainer).closest('[data-par]');
  const p2 = no(r.endContainer).closest('[data-par]');
  if (!p) return;
  if (p !== p2) { window.__edodocsSelecao = {erro: 'varios'}; return; }
  const antes = document.createRange();
  antes.selectNodeContents(p);
  antes.setEnd(r.startContainer, r.startOffset);
  const ini = antes.toString().length;
  window.__edodocsSelecao = {par: +p.dataset.par, ini: ini, fim: ini + r.toString().length,
                             texto: r.toString()};
});
}
"""


@dataclass
class Marca:
    campo: str
    tipo: str                                   # um dos TIPOS
    texto: str                                  # como estava no Word
    ocorrencias: list[tuple[int, int, int]]     # (parágrafo, início, fim)
    opcoes: list[dict] = field(default_factory=list)  # [{"valor", "texto"}] (tipo opcoes)
    # genero: {"masculino", "feminino"}; quantidade: {"uma", "uma_feminino", "varias"};
    # as_vezes: {"quando": "sim" ou "preenchido", "coluna", "texto"}
    regra: dict = field(default_factory=dict)
    modelo: str = ""   # o trecho com os campos de dentro como {{CAMPO}} (igual a texto se não tem)


def data_do_extenso(texto: str) -> str:
    """'3 de agosto de 2026' / '1º de setembro de 2026' / '03/08/2026' → '03/08/2026'."""
    m = DATA_EXTENSO.search(texto)
    if m:
        return f"{int(m.group(1)):02d}/{MESES.index(m.group(2).lower()) + 1:02d}/{m.group(3)}"
    m = DATA_NUMEROS.search(texto)
    return m.group(0) if m else ""


def tipo_sugerido(texto: str) -> str:
    return "data" if DATA_EXTENSO.fullmatch(texto.strip()) or DATA_NUMEROS.fullmatch(texto.strip()) else "texto"


def nome_sugerido(texto: str) -> str:
    t = texto.strip()
    if MATRICULA.fullmatch(t):
        return "MATRÍCULA"
    if CPF.fullmatch(t):
        return "CPF"
    if tipo_sugerido(t) == "data":
        return "DATA"
    if NOME_MAIUSCULO.fullmatch(t):
        return "NOME"
    return ""


def sugestoes(paragrafos: list[str]) -> list[tuple[int, int, int, str]]:
    """Trechos que parecem dados de uma pessoa: (parágrafo, início, fim, nome sugerido)."""
    achados = []
    for n, texto in enumerate(paragrafos):
        ocupado = []
        for padrao in (CPF, MATRICULA, DATA_EXTENSO, DATA_NUMEROS, NOME_MAIUSCULO):
            for m in padrao.finditer(texto):
                if any(m.start() < f and i < m.end() for i, f in ocupado):
                    continue
                # nome em maiúsculas no começo do parágrafo costuma ser título ("O PREFEITO")
                if padrao is NOME_MAIUSCULO and (m.start() < 3 or texto.strip() == m.group(0)):
                    continue
                ocupado.append((m.start(), m.end()))
                achados.append((n, m.start(), m.end(), nome_sugerido(m.group(0))))
    return sorted(achados)


def _ocorrencias(paragrafos: list[str], texto: str) -> list[tuple[int, int, int]]:
    """Onde o trecho aparece. Se ele começa/termina com letra ou número, só vale a palavra
    inteira ("15" não casa dentro de "2015"); "a pedido, " pode vir colado no que segue."""
    antes = r"(?<![\w])" if re.match(r"\w", texto) else ""
    depois = r"(?![\w])" if re.search(r"\w$", texto) else ""
    padrao = re.compile(antes + re.escape(texto) + depois)
    return [(n, m.start(), m.end()) for n, p in enumerate(paragrafos) for m in padrao.finditer(p)]


def estrutura(word: Path) -> list:
    """Como desenhar o texto do Word: [("p", n)] para cada parágrafo e, para cada tabela,
    ("tabela", [[(colspan, rowspan, [n, ...]) por célula] por linha]); n é a posição do
    parágrafo em textos_do_word (a mesma das marcas)."""
    body = Document(word).element.body
    todos = list(body.iter(qn("w:p")))  # a lista mantém os elementos vivos (o dict usa eles)
    numero = {p: n for n, p in enumerate(todos)}
    saida = []
    for el in body:
        if el.tag == qn("w:tbl"):
            saida.append(("tabela", [[(colspan, rowspan, [numero[p] for p in tc.iter(qn("w:p"))])
                                      for tc, colspan, rowspan in celulas]
                                     for celulas in atos._celulas(el)]))
        else:
            saida += [("p", numero[p]) for p in el.iter(qn("w:p"))]
    return saida


# Trecho que é só um traço ("–" numa célula vazia da tabela) ou uma célula em branco: vira
# campo que, vazio na planilha, deixa como estava (o traço, a célula em branco)
TRACOS = {"-", "–", "—"}


def criar_modelo(word: Path, destino: Path, marcas: list[Marca]) -> tuple[dict, dict]:
    """Grava o modelo (trechos → {{CAMPO}}) e devolve (ligações, linha de exemplo). O campo
    marcado dentro de outro (a data dentro de ", a contar de ...") não vai para o Word: entra
    pelo texto da regra do de fora."""
    externas = [(n, i, f) for m in marcas if m.tipo in COM_CAMPOS_DENTRO for n, i, f in m.ocorrencias]

    def dentro_de_outro(o: tuple[int, int, int]) -> bool:
        return any(o[0] == n and i <= o[1] and o[2] <= f and o != (n, i, f) for n, i, f in externas)

    achados = [modelo_automatico.Achado(m.texto, [m.campo], [o for o in m.ocorrencias if not dentro_de_outro(o)])
               for m in marcas]
    modelo_automatico.aplicar(word, destino, [a for a in achados if a.ocorrencias])
    ligacoes, exemplo = {}, {"_LINHA": 0, "_EXEMPLO": True}
    for m in marcas:
        chave = atos.normalizar(m.campo)
        modelo = m.modelo or m.texto
        if m.regra.get("existente"):  # a regra é a que o processo já tem
            continue
        if m.tipo == "data":
            exemplo[chave] = data_do_extenso(m.texto) or m.texto
            if DATA_EXTENSO.search(m.texto):  # no Word estava por extenso: continua por extenso
                ligacoes[m.campo] = {"tipo": "data_extenso", "coluna": m.campo}
        elif m.tipo == "opcoes":
            opcoes = [o for o in m.opcoes if str(o.get("valor", "")).strip()]
            ligacoes[m.campo] = {"tipo": "opcoes", "coluna": m.campo, "opcoes": opcoes}
            exemplo[chave] = opcoes[0]["valor"] if opcoes else ""
        elif m.tipo == "genero":
            masc, fem = m.regra.get("masculino", ""), m.regra.get("feminino", "")
            ligacoes[m.campo] = {"tipo": "genero", "coluna": SEXO, "masculino": masc, "feminino": fem}
            exemplo.setdefault(SEXO, "F" if m.texto == fem and fem != masc else "M")
        elif m.tipo == "quantidade":
            casos = [{"se": [{"campo": atos.AUTOMATICOS[atos.QUANTIDADE][0], "op": "maior", "valor": "1"}],
                      "texto": m.regra.get("varias", "")}]
            if m.regra.get("uma_feminino", "").strip():
                casos.append({"se": [{"campo": SEXO, "op": "igual", "valor": "F"}], "texto": m.regra["uma_feminino"]})
            ligacoes[m.campo] = {"tipo": "condicao", "casos": casos, "senao": m.regra.get("uma", "")}
        elif m.tipo == "as_vezes":
            if m.regra.get("quando") == "preenchido":
                ligacoes[m.campo] = {"tipo": "condicao", "senao": "", "casos": [
                    {"se": [{"campo": m.regra.get("coluna", ""), "op": "preenchido"}], "texto": modelo}]}
            else:  # coluna própria: SIM escreve o trecho, NÃO não escreve
                ligacoes[m.campo] = {"tipo": "opcoes", "coluna": m.campo, "opcoes": [
                    {"valor": "SIM", "texto": modelo}, {"valor": "NÃO", "texto": ""}]}
                exemplo[chave] = "SIM"
        elif m.texto.strip() in TRACOS or not m.texto.strip():
            ligacoes[m.campo] = {"tipo": "coluna", "coluna": m.campo, "se_vazio": m.texto}
            exemplo[chave] = ""
        else:
            exemplo[chave] = m.texto
    return ligacoes, exemplo


class MarcadorWord:
    """O texto do Word com as marcas; ao_mudar() é chamado quando as marcas mudam."""

    def __init__(self, camada, word: Path, ao_mudar=None, existentes: dict[str, str] | None = None):
        """existentes: campos que o processo já tem ({nome: tipo da ligação}), ao marcar um
        segundo Word (o de várias pessoas). Marcar com um desses nomes usa a mesma regra: não
        precisa preencher as opções de novo."""
        self.camada = camada
        self.word = word
        self.existentes = {atos.normalizar(n): (n, t) for n, t in (existentes or {}).items()}
        self.paragrafos = modelo_automatico.textos_do_word(word)
        self.estrutura = estrutura(word)
        self.marcas: list[Marca] = []
        self.ao_mudar = ao_mudar or (lambda: None)
        gui.run_javascript(SCRIPT_SELECAO)
        with gui.column().classes("w-full gap-2"):
            with gui.row().classes("w-full items-center gap-2"):
                gui.icon("ads_click", color="primary", size="22px")
                gui.label("Selecione com o mouse o que muda de um ato para outro (o nome da pessoa, a "
                          "matrícula, a data...) e clique em Marcar. Dois cliques numa palavra também "
                          "selecionam.").classes("text-body2 t-forte col")
                gui.button("Marcar o que selecionei", icon="border_color", on_click=self._marcar_selecao) \
                    .props("unelevated no-caps")
            self.area_sugestoes = gui.row().classes("w-full items-center gap-1")
            with gui.row().classes("w-full gap-3 no-wrap items-start"):
                self.area_texto = gui.html("", sanitize=False).classes("col folha-texto rounded") \
                    .style("max-height: 52vh; overflow-y: auto; user-select: text; cursor: text")
                # célula vazia da tabela: não dá para selecionar com o mouse, então tem um botão
                # (o clique chega aqui pelo evento com o id desta área). sanitize=False para o
                # onclick do botão ficar: o texto do Word entra sempre escapado (html.escape)
                self.evento_celula = f"edodocs_celula_{self.area_texto.id}"
                gui.on(self.evento_celula, lambda e: self._marcar_celula(int(e.args)))
                with gui.card().classes("cartao q-pa-sm w-72 gap-1").style("overflow: hidden"):
                    gui.label("Marcados").classes("text-subtitle2 t-forte")
                    self.area_marcas = gui.column().classes("w-full gap-1")
        self.desenhar()

    # --- Desenho ---
    def desenhar(self) -> None:
        marcados = {(n, i, f): m for m in self.marcas for n, i, f in m.ocorrencias}

        def paragrafo(n: int, margem: str = "0 0 6px 0", celula: bool = False) -> str:
            texto = self.paragrafos[n]
            trechos = sorted((i, f, m) for (pn, i, f), m in marcados.items() if pn == n)
            if celula and not texto.strip():  # célula em branco: botão para marcar (ou o campo marcado)
                if trechos:
                    campo = html.escape(trechos[0][2].campo)
                    dentro = (f'<mark title="{campo}" style="background:#ffe08a;border-radius:3px;'
                              f'padding:0 2px">[{campo}]</mark>')
                else:
                    dentro = (f"<span onclick=\"emitEvent('{self.evento_celula}', {n})\" "
                              "title=\"Célula em branco: clique para marcar\" "
                              "style=\"cursor:pointer;color:#1976d2;font-size:.85em;border:1px dashed #1976d2;"
                              "border-radius:3px;padding:0 4px;user-select:none\">+ marcar</span>")
                return f'<p data-par="{n}" style="margin:{margem};min-height:1em">{dentro}</p>'
            def trecho(ini: int, fim: int, marcas: list) -> str:
                """O texto de ini a fim com as marcas; a de fora (", a contar de ...") em azul,
                com as de dentro (a data) em amarelo dentro dela."""
                saida, pos = "", ini
                de_fora = [x for x in marcas if not any(o is not x and o[0] <= x[0] and x[1] <= o[1] for o in marcas)]
                for x in sorted(de_fora, key=lambda x: x[0]):
                    i, f, m = x
                    cor = "#cfe3ff" if m.tipo in COM_CAMPOS_DENTRO else "#ffe08a"
                    dentro = [o for o in marcas if o is not x and i <= o[0] and o[1] <= f]
                    saida += html.escape(texto[pos:i])
                    saida += (f'<mark title="{html.escape(m.campo)}" style="background:{cor};border-radius:3px;'
                              f'padding:0 2px">{trecho(i, f, dentro)}</mark>')
                    pos = f
                return saida + html.escape(texto[pos:fim])

            html_par = trecho(0, len(texto), trechos)
            return f'<p data-par="{n}" style="margin:{margem};min-height:1em">{html_par or "&nbsp;"}</p>'

        partes = []
        for tipo, conteudo in self.estrutura:
            if tipo == "p":
                partes.append(paragrafo(conteudo))
                continue
            linhas = []
            for celulas in conteudo:
                tds = "".join(
                    f'<td colspan="{colspan}" rowspan="{rowspan}" style="border:1px solid #888;'
                    f'padding:2px 6px;vertical-align:middle">{"".join(paragrafo(n, "0", True) for n in ns)}</td>'
                    for colspan, rowspan, ns in celulas)
                linhas.append(f"<tr>{tds}</tr>")
            partes.append('<table style="border-collapse:collapse;margin:4px 0 10px 0;width:100%">'
                          f'{"".join(linhas)}</table>')
        self.area_texto.content = "".join(partes)

        self.area_marcas.clear()
        with self.area_marcas:
            if not self.marcas:
                gui.label("Nada marcado ainda.").classes("text-caption t-suave")
            for m in self.marcas:
                with gui.row().classes("w-full items-center gap-1 no-wrap"):
                    with gui.column().classes("gap-0 col").style("min-width:0"):
                        gui.label("{{" + m.campo + "}}").classes("text-body2 text-weight-medium t-forte ellipsis")
                        extra = f" · {len(m.ocorrencias)}x" if len(m.ocorrencias) > 1 else ""
                        tipo = "Mesma regra do processo" if m.regra.get("existente") else TIPOS[m.tipo].split(" (")[0]
                        gui.label(f"{m.texto[:30]}{extra} · {tipo}") \
                            .classes("text-caption t-suave ellipsis")
                    gui.button(icon="close", on_click=lambda m=m: self._tirar(m)) \
                        .props("flat round dense size=sm color=grey-7").tooltip("Desmarcar")

        self.area_sugestoes.clear()
        livres = [(n, i, f, nome) for n, i, f, nome in sugestoes(self.paragrafos)
                  if not any(pn == n and i < pf and pi < f for (pn, pi, pf) in marcados)]
        with self.area_sugestoes:
            if livres:
                gui.label("Parece dado de pessoa:").classes("text-caption t-suave")
                vistos = set()
                for n, i, f, nome in livres[:12]:
                    texto = self.paragrafos[n][i:f]
                    if texto in vistos:
                        continue
                    vistos.add(texto)
                    gui.chip(f"{texto[:32]}" + (f" → {nome}" if nome else ""), icon="add",
                             color=None, on_click=lambda n=n, i=i, f=f: self._perguntar(n, i, f)) \
                        .props("dense clickable").classes("chip-campo").tooltip("Clique para marcar")

    # --- Marcar ---
    async def _marcar_selecao(self) -> None:
        sel = await gui.run_javascript("return window.__edodocsSelecao || null")
        await gui.run_javascript("window.__edodocsSelecao = null; window.getSelection().removeAllRanges()")
        if not sel:
            gui.notify("Selecione antes, com o mouse, o trecho do texto do Word", type="warning")
            return
        if sel.get("erro"):
            gui.notify("Selecione dentro de um parágrafo só", type="warning")
            return
        n, ini, fim = int(sel["par"]), int(sel["ini"]), int(sel["fim"])
        if not self.paragrafos[n][ini:fim].strip(" ,;:."):
            gui.notify("Selecione o trecho que muda (não só espaço ou pontuação)", type="warning")
            return
        self._perguntar(n, ini, fim)

    def _marcar_celula(self, n: int) -> None:
        """Célula vazia da tabela (ex.: CLASS PCD de quem não é PCD): o campo fica no lugar dela."""
        self._perguntar(n, 0, len(self.paragrafos[n]))

    def _perguntar(self, n: int, ini: int, fim: int) -> None:
        """Janela: qual o nome deste campo e de que tipo ele é."""
        selecionado = self.paragrafos[n][ini:fim]
        # Marcas já feitas que ficam inteiras dentro da seleção (a data dentro de ", a contar
        # de 1º de outubro"): o trecho novo leva elas como {{CAMPO}}. Se a seleção corta uma
        # marca no meio (ou fica dentro de uma), não dá.
        dentro = []
        for m in self.marcas:
            for pn, pi, pf in m.ocorrencias:
                if pn != n or not (ini < pf and pi < fim) or (pi == pf == ini == fim):
                    continue
                if ini <= pi and pf <= fim and (pi, pf) != (ini, fim):
                    dentro.append((pi, pf, m))
                else:
                    gui.notify("Parte desse trecho já está marcada: desmarque antes", type="warning")
                    return
        modelo = selecionado
        for pi, pf, m in sorted(dentro, key=lambda x: x[0], reverse=True):
            modelo = modelo[:pi - ini] + "{{" + m.campo + "}}" + modelo[pf - ini:]
        # Texto (nome, matrícula...): sem os espaços e a pontuação das pontas da seleção
        # (" ANTONIO EXEMPLO, " → "ANTONIO EXEMPLO"). Opções e trechos que somem: a seleção inteira,
        # porque a vírgula de "a pedido, " também some quando a resposta é NÃO.
        aparado = (n, ini + len(selecionado) - len(selecionado.lstrip(" ,;:.")),
                   fim - (len(selecionado) - len(selecionado.rstrip(" ,;:."))))
        texto = self.paragrafos[n][aparado[1]:aparado[2]]
        vazia = not selecionado.strip()  # célula em branco: o trecho é o que estava nela
        if vazia:
            aparado, texto = (n, ini, fim), selecionado
        todas = [] if vazia or dentro else _ocorrencias(self.paragrafos, texto)
        nomes = list(dict.fromkeys([m.campo for m in self.marcas] + [n for n, _ in self.existentes.values()]
                                   + NOMES_COMUNS))
        tipos = {t: r for t, r in TIPOS.items() if not dentro or t in COM_CAMPOS_DENTRO}
        masc, fem = genero_sugerido(texto)
        estado = {"tipo": "as_vezes" if dentro else tipo_sugerido(texto),
                  "opcoes": [{"valor": "SIM", "texto": modelo}, {"valor": "NÃO", "texto": ""}],
                  "genero": {"masculino": masc, "feminino": fem},
                  "quantidade": {"uma": modelo, "uma_feminino": "", "varias": modelo},
                  "as_vezes": {"quando": "preenchido" if dentro else "sim",
                               "coluna": dentro[0][2].campo if dentro else None}}
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[680px] max-w-full q-pa-lg gap-3"):
            gui.label("Marcar um campo").classes("text-h6 t-forte")
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.label("Trecho:").classes("text-body2 t-suave")
                gui.label(modelo if not vazia else "(célula em branco)") \
                    .classes("text-body1 text-weight-medium").style("background:#ffe08a;padding:0 4px")
            if dentro:
                gui.label("Os campos que já estavam marcados dentro dele continuam valendo ("
                          + ", ".join(m.campo for _, _, m in dentro) + ").").classes("text-caption t-suave")
            # Nome livre: a pessoa escreve o que quiser, ou clica numa sugestão
            nome = gui.input("Nome do campo (vira a coluna da planilha)", value=nome_sugerido(texto) or "",
                             placeholder="Escreva um nome ou clique numa sugestão abaixo", autocomplete=nomes) \
                .props('outlined dense clearable bg-color=white input-class=text-grey-10 label-color=grey-8') \
                .classes("w-full")
            with gui.row().classes("w-full items-center gap-1"):
                gui.label("Sugestões:").classes("text-caption t-suave q-mr-xs")
                for sugestao in nomes:
                    gui.chip(sugestao, on_click=lambda s=sugestao: nome.set_value(s)) \
                        .props("dense outline clickable color=primary")

            def sugerir_nome() -> None:
                """Tipos novos ganham um nome pronto, se a pessoa ainda não deu um."""
                if (nome.value or "").strip():
                    return
                tipo = estado["tipo"]
                inicio = " ".join(texto.strip(" ,;:.").split()[:3])  # as primeiras palavras, inteiras
                sugestao = {"genero": f"{masc} OU {fem}", "quantidade": inicio, "as_vezes": inicio}.get(tipo, "")
                sugestao = " ".join(sugestao.upper().split()).strip(" ,;:.")
                if sugestao:
                    nome.set_value(sugestao)

            # Campo que o processo já tem (marcado no Word principal): usa a mesma regra
            aviso_existe = gui.row().classes("w-full items-center gap-2 fundo-previa rounded q-pa-sm no-wrap")
            with aviso_existe:
                gui.icon("link", color="positive", size="22px")
                texto_existe = gui.label("").classes("text-body2 t-forte")
            radio_tipos = gui.radio(tipos, value=estado["tipo"],
                                    on_change=lambda e: (estado.update(tipo=e.value), sugerir_nome(), mostrar())) \
                .props("dense")
            area_tipo = gui.column().classes("w-full gap-1")

            def existente() -> tuple[str, str] | None:
                return self.existentes.get(atos.normalizar(nome.value or ""))

            def ver_nome() -> None:
                """Nome de um campo que já existe: some a escolha do tipo, fica o aviso."""
                achado = existente()
                aviso_existe.set_visibility(bool(achado))
                radio_tipos.set_visibility(not achado)
                area_tipo.set_visibility(not achado)
                if achado:
                    texto_existe.text = (f"{achado[0]} já existe no processo: este trecho usa a mesma regra "
                                         "(não precisa preencher de novo).")

            nome.on_value_change(lambda e: ver_nome())
            # Trecho curto ("–", "AMPLA" numa tabela) costuma aparecer em outros lugares que
            # não têm nada a ver: aí só esta vez, a não ser que a pessoa marque
            todas_vezes = gui.checkbox(f"Marcar todas as vezes que aparece no Word ({len(todas)})",
                                       value=len(texto.strip()) > 5) if len(todas) > 1 else None

            def entrada(rotulo: str, dados: dict, chave: str, dica: str = "") -> None:
                gui.input(rotulo, value=dados.get(chave, ""), placeholder=dica,
                          on_change=lambda e: dados.update({chave: e.value})) \
                    .props("outlined dense").classes("w-full")

            def mostrar() -> None:
                area_tipo.clear()
                tipo = estado["tipo"]
                with area_tipo:
                    if tipo == "opcoes":
                        gui.label("Para cada opção, o que vai escrito no ato (vazio = não escreve nada):") \
                            .classes("text-caption t-suave")
                        for o in estado["opcoes"]:
                            with gui.row().classes("w-full items-center gap-2 no-wrap"):
                                gui.input("Opção", value=o["valor"],
                                          on_change=lambda e, o=o: o.update(valor=e.value)) \
                                    .props("outlined dense").classes("w-40")
                                gui.icon("arrow_forward", color="grey-6")
                                gui.input("Escreve no ato", value=o["texto"],
                                          on_change=lambda e, o=o: o.update(texto=e.value)) \
                                    .props("outlined dense").classes("col")
                        gui.button("Mais uma opção", icon="add",
                                   on_click=lambda: (estado["opcoes"].append({"valor": "", "texto": ""}), mostrar())) \
                            .props("flat dense no-caps size=sm")
                    elif tipo == "genero":
                        gui.label("A planilha ganha a coluna SEXO (F ou M).").classes("text-caption t-suave")
                        with gui.row().classes("w-full gap-2 no-wrap"):
                            with gui.column().classes("col gap-0"):
                                entrada("Homem (M)", estado["genero"], "masculino")
                            with gui.column().classes("col gap-0"):
                                entrada("Mulher (F)", estado["genero"], "feminino")
                    elif tipo == "quantidade":
                        gui.label("O que vai escrito quando o ato tem uma pessoa só e quando tem várias:") \
                            .classes("text-caption t-suave")
                        entrada("Com uma pessoa", estado["quantidade"], "uma", "ex.: Fica nomeado")
                        entrada("Com uma mulher (só se muda; pode deixar em branco)", estado["quantidade"],
                                "uma_feminino", "ex.: Fica nomeada")
                        entrada("Com várias pessoas", estado["quantidade"], "varias", "ex.: Ficam nomeados")
                    elif tipo == "as_vezes":
                        quando = estado["as_vezes"]
                        colunas = [m.campo for m in self.marcas]
                        gui.label("Quando este trecho aparece no ato?").classes("text-caption t-suave")
                        gui.radio({"sim": "Quando eu responder SIM na planilha (coluna com o nome deste campo)",
                                   "preenchido": "Quando outra coluna estiver preenchida"},
                                  value=quando["quando"],
                                  on_change=lambda e: (quando.update(quando=e.value), mostrar())).props("dense")
                        if quando["quando"] == "preenchido":
                            if colunas:
                                gui.select(colunas, value=quando["coluna"] if quando["coluna"] in colunas else None,
                                           label="Coluna", on_change=lambda e: quando.update(coluna=e.value)) \
                                    .props("outlined dense").classes("w-full")
                            else:
                                gui.label("Marque antes a coluna que decide (ex.: a data).") \
                                    .classes("text-caption text-orange-8")

            def regra_do_tipo() -> dict | None:
                tipo = estado["tipo"]
                if tipo == "genero":
                    g = estado["genero"]
                    if not g["masculino"].strip() or not g["feminino"].strip():
                        gui.notify("Escreva como fica para homem e para mulher", type="warning")
                        return None
                    return dict(g)
                if tipo == "quantidade":
                    q = estado["quantidade"]
                    if not q["uma"].strip() or not q["varias"].strip():
                        gui.notify("Escreva como fica com uma pessoa e com várias", type="warning")
                        return None
                    return dict(q)
                if tipo == "as_vezes":
                    quando = estado["as_vezes"]
                    if quando["quando"] == "preenchido" and not quando["coluna"]:
                        gui.notify("Escolha a coluna que decide se o trecho aparece", type="warning")
                        return None
                    return dict(quando)
                return {}

            def salvar() -> None:
                campo = " ".join(str(nome.value or "").split()).upper()
                if not campo:
                    gui.notify("Dê um nome ao campo (ex.: NOME)", type="warning")
                    return
                if any(atos.normalizar(campo) == atos.normalizar(m.campo) for m in self.marcas):
                    gui.notify(f"Já existe um campo {campo}: use outro nome", type="warning")
                    return
                achado = existente()
                if achado:  # a regra já está no processo: aqui só marca o lugar
                    campo, regra = achado[0], {"existente": True}
                    estado["tipo"] = "texto"
                    # opções e regras (", a pedido") levam a seleção inteira, com a pontuação
                    inteiro = achado[1] in ("opcoes", "condicao", "genero")
                else:
                    regra = regra_do_tipo()
                    if regra is None:
                        return
                    inteiro = estado["tipo"] in COM_CAMPOS_DENTRO  # a seleção inteira, com a pontuação
                trecho = selecionado if inteiro else texto
                uma = (n, ini, fim) if inteiro else aparado
                if dentro:  # trecho com campos dentro: só este lugar
                    ocorrencias = [uma]
                else:
                    todas_do_trecho = (_ocorrencias(self.paragrafos, trecho) if not vazia else []) or [uma]
                    ocorrencias = todas_do_trecho if (todas_vezes is None or todas_vezes.value) else [uma]
                    ocupadas = {(pn, pi, pf) for m in self.marcas for pn, pi, pf in m.ocorrencias}
                    ocorrencias = [o for o in ocorrencias
                                   if not any(o[0] == pn and o[1] < pf and pi < o[2] for pn, pi, pf in ocupadas)]
                self.marcas.append(Marca(campo, estado["tipo"], trecho, ocorrencias or [uma],
                                         estado["opcoes"] if estado["tipo"] == "opcoes" else [],
                                         regra, modelo if inteiro else trecho))
                janela.close()
                self.desenhar()
                self.ao_mudar()

            sugerir_nome()
            mostrar()
            ver_nome()
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Marcar", icon="check", on_click=salvar).props("unelevated no-caps")
        janela.open()

    def _tirar(self, marca: Marca) -> None:
        """Desmarca. Um campo que estava dentro de outro leva o de fora junto (o texto do de
        fora dependia dele)."""
        tirar = [marca] + [m for m in self.marcas if m is not marca and "{{" + marca.campo + "}}" in m.modelo]
        for m in tirar:
            self.marcas.remove(m)
        if len(tirar) > 1:
            gui.notify("Desmarquei também " + ", ".join(m.campo for m in tirar[1:]) + " (estava em volta)",
                       type="info")
        self.desenhar()
        self.ao_mudar()
