"""Desenhos do assistente que imitam um Word: a pessoa vê, em cada passo, como o ato vai
ficar (uma pessoa ou várias, o que foi marcado, como as linhas da planilha viram atos).

Tudo é HTML montado aqui (o texto da pessoa entra sempre escapado) e usa o CSS abaixo:

    pagina(...)          uma folha de Word com linhas cinzas no lugar do texto e os campos em
                         amarelo; com várias pessoas, a parte que repete fica num quadro azul
    planilha_para_atos   uma planilha pequena e, ao lado, as folhas que ela vira
    ligacoes             cada coluna da planilha → o campo do Word que ela preenche
"""
import html
import random

NOMES = ["MARIA DA SILVA", "JOÃO PEREIRA", "ANA SOUZA", "CARLOS LIMA", "PAULA ROCHA", "JOSÉ SANTOS",
         "LUCIANA ALVES", "PEDRO COSTA", "FERNANDA MELO", "RAFAEL GOMES"]

CSS = """
.sim-pagina { background:#fff; color:#1f2937; border:1px solid #d6dde8; border-radius:4px;
              box-shadow:0 3px 10px rgba(15,35,75,.16); padding:14px 14px 18px; position:relative;
              font-size:9.5px; line-height:1.55; text-align:left; }
.sim-pagina::after { content:""; position:absolute; top:-1px; right:-1px; width:16px; height:16px;
                     background:linear-gradient(225deg, #f3f5f8 50%, #dfe5ee 50%); border-radius:0 4px 0 0; }
.sim-titulo { height:7px; width:42%; background:#9aa6b8; border-radius:3px; margin:2px auto 9px; }
.sim-linha { display:inline-block; height:6px; background:#dbe1ea; border-radius:3px; vertical-align:middle;
             margin:0 3px 0 0; }
.sim-par { margin:0 0 6px; }
.sim-campo { display:inline-block; background:#ffe08a; color:#5b4300; border-radius:3px; padding:0 4px;
             font-weight:600; font-size:8.5px; line-height:1.4; margin:1px 3px 1px 0; white-space:nowrap; }
.sim-pessoa { background:#cfe3ff; color:#123a7a; }
.sim-repete { border:1.5px dashed #2f6fcf; border-radius:5px; padding:4px 5px 2px; margin:2px 0 6px;
              background:#f4f8ff; position:relative; }
.sim-repete-rotulo { position:absolute; top:-8px; right:6px; background:#2f6fcf; color:#fff;
                     font-size:7.5px; font-weight:600; border-radius:8px; padding:0 6px; }
.sim-cartao-tipo { width:250px; cursor:pointer; border-radius:14px; padding:14px; transition:all .15s;
                   border:2px solid transparent; }
.sim-cartao-tipo:hover { transform:translateY(-2px); box-shadow:0 6px 18px rgba(15,35,75,.18); }
.sim-cartao-tipo.escolhido { border-color:var(--q-primary); box-shadow:0 6px 18px rgba(36,89,168,.28); }
.sim-tabela { border-collapse:collapse; font-size:8.5px; background:#fff; color:#1f2937; width:100%; }
.sim-tabela th { background:#1f7a43; color:#fff; font-weight:600; padding:2px 4px; border:1px solid #c9d3c9;
                 white-space:nowrap; }
.sim-tabela td { padding:2px 4px; border:1px solid #d6dde8; white-space:nowrap; }
.sim-tabela tr.par td { background:#eef6ff; }
.sim-seta { font-size:22px; color:#2f6fcf; text-align:center; line-height:1; }
"""


def _e(texto: str) -> str:
    return html.escape(str(texto))


def _linhas(*larguras: int) -> str:
    """Pedaços de "texto" cinza (larguras em %), um atrás do outro como palavras."""
    return "".join(f'<span class="sim-linha" style="width:{w}%"></span>' for w in larguras)


def _campo(nome: str, pessoa: bool = False) -> str:
    return f'<span class="sim-campo{" sim-pessoa" if pessoa else ""}">{_e(nome)}</span>'


def nomes_aleatorios(quantos: int) -> list[str]:
    return random.sample(NOMES, quantos)


def pagina(varias: bool, campos: list[str] | None = None, da_pessoa: set[str] | None = None,
           exemplos: list[str] | None = None, pessoas: int = 3, titulo: bool = True) -> str:
    """Folha de Word simulada. campos: nomes marcados (aparecem em amarelo); sem campos,
    exemplos são os valores mostrados (nomes de pessoas). varias: o Art. 1º vira a lista
    I, II, III dentro do quadro "repete para cada pessoa" (com os campos de da_pessoa)."""
    campos = list(campos or [])
    exemplos = exemplos or nomes_aleatorios(max(pessoas, 1))
    da_pessoa = set(da_pessoa or [])
    corpo = '<div class="sim-titulo"></div>' if titulo else ""
    corpo += f'<div class="sim-par">{_linhas(30, 40, 22, 55, 35, 18)}</div>'
    # sem saber ainda o que repete (antes de escolher o parágrafo), tudo vai no quadro azul
    do_ato = [c for c in campos if c not in da_pessoa] if varias and da_pessoa else []
    if do_ato:
        corpo += '<div class="sim-par">' + _linhas(20) + "".join(_campo(c) for c in do_ato[:3]) + _linhas(30) + "</div>"
    if not varias:
        if campos:
            meio = "".join(_campo(c) + _linhas(8) for c in campos[:6])
            extra = f'<span class="sim-campo">+{len(campos) - 6}</span>' if len(campos) > 6 else ""
        else:
            meio = _campo(exemplos[0], pessoa=True) + _linhas(14, 25)
            extra = ""
        corpo += f'<div class="sim-par">{_linhas(12, 18)}{meio}{extra}{_linhas(20, 30)}</div>'
    else:
        corpo += f'<div class="sim-par">{_linhas(12, 22, 35, 16)}</div>'
        itens = []
        de_cada = [c for c in campos if c in da_pessoa] or campos[:3]
        for i, numero in enumerate(("I", "II", "III")[:pessoas]):
            if de_cada:
                conteudo = "".join(_campo(c, pessoa=True) for c in de_cada[:3])
            else:
                conteudo = _campo(exemplos[i % len(exemplos)], pessoa=True)
            fim = "." if i == pessoas - 1 else ";"
            itens.append(f'<div style="margin:1px 0"><b>{numero} –</b> {conteudo}{_linhas(16)}{fim}</div>')
        corpo += ('<div class="sim-repete"><span class="sim-repete-rotulo">repete para cada pessoa</span>'
                  + "".join(itens) + "</div>")
    corpo += f'<div class="sim-par">{_linhas(45, 20)}</div>'
    corpo += f'<div class="sim-par" style="text-align:center;margin-top:10px">{_linhas(30)}</div>'
    return f'<div class="sim-pagina">{corpo}</div>'


def cartao_tipo(varias: bool) -> str:
    """Conteúdo do cartão do passo "Tipo de ato": a folha com uma pessoa ou com várias."""
    return pagina(varias, pessoas=3)


def planilha_para_atos(varias: bool, colunas: list[str]) -> str:
    """Uma planilha pequena (3 linhas) e, embaixo, as folhas que ela vira: 3 atos (uma pessoa
    por ato) ou 2 atos, o primeiro com as duas pessoas de mesmo número de ato."""
    nomes = nomes_aleatorios(3)
    colunas = [c for c in colunas if c][:4] or ["NOME"]
    if varias and not any("ATO" in c.upper() or "DECRETO" in c.upper() for c in colunas):
        colunas = ["Nº DO ATO"] + colunas[:3]
    numeros = ["1", "1", "2"] if varias else ["", "", ""]
    linhas = []
    for i, nome in enumerate(nomes):
        celulas = []
        for c in colunas:
            if c.upper().startswith("Nº") or "ATO" in c.upper() or "DECRETO" in c.upper():
                valor = numeros[i] or str(i + 1)
            elif "NOME" in c.upper():
                valor = nome
            else:
                valor = "…"
            celulas.append(f"<td>{_e(valor)}</td>")
        cor = ' class="par"' if varias and numeros[i] == "1" else ""
        linhas.append(f"<tr{cor}>{''.join(celulas)}</tr>")
    tabela = (f'<table class="sim-tabela"><tr>{"".join(f"<th>{_e(c)}</th>" for c in colunas)}</tr>'
              f'{"".join(linhas)}</table>')
    if varias:
        folhas = [(pagina(True, exemplos=nomes[:2], pessoas=2, titulo=False), "Ato 1 · 2 pessoas"),
                  (pagina(False, exemplos=nomes[2:], titulo=False), "Ato 2 · 1 pessoa")]
        explica = "Linhas com o mesmo Nº DO ATO viram um ato só."
    else:
        folhas = [(pagina(False, exemplos=[n], titulo=False), f"Ato {i + 1}") for i, n in enumerate(nomes)]
        explica = "Cada linha da planilha vira um ato."
    grade = "".join(f'<div style="flex:1;min-width:0">{f}<div style="text-align:center;font-size:10px;'
                    f'margin-top:3px;color:#5b6472">{_e(r)}</div></div>' for f, r in folhas)
    return (f'{tabela}<div class="sim-seta">⬇</div>'
            f'<div style="display:flex;gap:6px;transform:scale(.98)">{grade}</div>'
            f'<div style="font-size:11px;color:#5b6472;margin-top:6px;text-align:center">{_e(explica)}</div>')


def ligacoes(pares: list[tuple[str, str, bool]]) -> str:
    """Cada coluna → campo do Word: [(campo, de onde vem, ok)]."""
    linhas = []
    for campo, origem, ok in pares[:12]:
        cor = "#16a34a" if ok else "#ca8a04"
        linhas.append(
            f'<div style="display:flex;align-items:center;gap:6px;margin:3px 0;font-size:11px">'
            f'<span style="background:#1f7a43;color:#fff;border-radius:3px;padding:0 5px;white-space:nowrap;'
            f'max-width:45%;overflow:hidden;text-overflow:ellipsis">{_e(origem)}</span>'
            f'<span style="color:{cor}">➜</span>{_campo(campo)}</div>')
    if len(pares) > 12:
        linhas.append(f'<div style="font-size:11px;color:#5b6472">e mais {len(pares) - 12}...</div>')
    return "".join(linhas)
