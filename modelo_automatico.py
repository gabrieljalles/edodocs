"""Transforma um ato pronto (um Word comum, já preenchido para uma pessoa) em modelo.

A pessoa não precisa saber o que é campo nem escrever {{NOME}}: ela envia um ato que já
fez e a planilha. O programa descobre de qual linha da planilha é o ato (a linha cujos
valores mais aparecem no texto), mostra os trechos que encontrou ("MARIA DA SILVA" é a
coluna NOME, "12.345-6" é a MATRÍCULA...) e, depois que a pessoa confere, troca esses
trechos por {{COLUNA}} mantendo a formatação do Word.

    achados, linha = encontrar(word, registros, colunas)
    aplicar(word, destino, [a for a in achados if a.usar])
"""
import re
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

import atos

# Colunas que o próprio programa grava (status/data de envio): não são dados do ato
CONTROLE = {"STATUS", "STATUSEDOCS", "DATAEDOCS", "STATUSDIARIO", "DATADIARIO", "STATUSPUBLICADO", "CRIADOPOR"}
# Opção "número por extenso" de uma coluna: "DIAS (por extenso)" → campo {{DIAS POR EXTENSO}}
EXTENSO = " (por extenso)"
# Palavras que mudam com o sexo da pessoa (feminino, masculino). As mais longas primeiro.
GENERO = [
    ("à servidora", "ao servidor"), ("da servidora", "do servidor"), ("a servidora", "o servidor"),
    ("servidora", "servidor"), ("lotada", "lotado"), ("nomeada", "nomeado"),
    ("aposentada", "aposentado"), ("designada", "designado"), ("exonerada", "exonerado"),
    ("cedida", "cedido"), ("interessada", "interessado"), ("requerente", "requerente"),
]
_GENERO_OPCAO = re.compile(r"^Muda com (.+?): (.+) / (.+)$")
REFERENCIA = re.compile(r"\b(lei|decreto|portaria|resolu[çc][ãa]o|instru[çc][ãa]o)\b[^\n]{0,25}$", re.IGNORECASE)


@dataclass
class Achado:
    """Um trecho do ato que é dado da planilha."""
    valor: str                      # como está escrito no Word
    colunas: list[str]              # colunas da planilha com esse valor (1ª = sugerida)
    ocorrencias: list[tuple[int, int, int]]  # (parágrafo, início, fim) de cada vez que aparece
    trecho: str = ""                # contexto, para a pessoa reconhecer onde está
    coluna: str = ""                # a escolhida
    usar: bool = True

    def __post_init__(self):
        self.coluna = self.coluna or self.colunas[0]


def _paragrafos(doc) -> list:
    return list(doc.element.body.iter(qn("w:p")))


def _texto(p) -> str:
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def textos_do_word(caminho: Path) -> list[str]:
    """Texto de cada parágrafo, exatamente como fica guardado no Word (base das posições)."""
    return [_texto(p) for p in _paragrafos(Document(caminho))]


def _util(valor: str) -> bool:
    """Valor que vale a pena procurar (evita 'F', 'M', '0', 'SIM'...)."""
    v = valor.strip()
    if not v or v in atos.ERROS_EXCEL:
        return False
    if v.replace(".", "").replace(",", "").isdigit():
        return len(v) >= 2 and int(re.sub(r"\D", "", v) or 0) != 0
    return len(v) >= 4


def _padrao(valor: str) -> re.Pattern:
    """O valor como palavra inteira (15 não casa dentro de 2015), sem diferenciar maiúscula
    e aceitando espaços repetidos."""
    partes = [re.escape(p) for p in valor.strip().split()]
    return re.compile(r"(?<![\w])" + r"\s+".join(partes) + r"(?![\w])", re.IGNORECASE)


def _parte_de_data(texto: str, ini: int, fim: int) -> bool:
    """O número em texto[ini:fim] é pedaço de uma data (dd/mm/aaaa)?"""
    return bool(re.search(r"\d{1,2}/\d{1,2}/$", texto[:ini]) or re.match(r"/\d{1,2}/\d{2,4}", texto[fim:])
                or (re.search(r"\d{1,2}/$", texto[:ini]) and re.match(r"/\d{2,4}", texto[fim:])))


def _contexto(texto: str, ini: int, fim: int, largura: int = 45) -> str:
    a, b = max(0, ini - largura), min(len(texto), fim + largura)
    return ("…" if a else "") + texto[a:ini] + "【" + texto[ini:fim] + "】" + texto[fim:b] + ("…" if b < len(texto) else "")


def melhor_linha(paragrafos: list[str], registros: list[dict], colunas: list[str]) -> dict | None:
    """A linha da planilha cujo conteúdo mais aparece no ato."""
    texto = "\n".join(paragrafos).lower()
    nomes = [(atos.normalizar(c), c) for c in colunas if atos.normalizar(c) not in CONTROLE]
    melhor, pontos_melhor = None, 0
    for r in registros:
        pontos = sum(len(v) for chave, _ in nomes
                     if _util(v := str(r.get(chave, ""))) and v.strip().lower() in texto)
        if pontos > pontos_melhor:
            melhor, pontos_melhor = r, pontos
    return melhor


def encontrar(caminho: Path, registros: list[dict], colunas: list[str]) -> tuple[list[Achado], dict | None]:
    """Os dados da planilha encontrados no ato, na ordem em que aparecem, e a linha usada."""
    paragrafos = textos_do_word(caminho)
    linha = melhor_linha(paragrafos, registros, colunas)
    if linha is None:
        return [], None

    # valor (minúsculo) → colunas com esse valor
    por_valor: dict[str, list[str]] = {}
    escrito: dict[str, str] = {}
    texto_todo = "\n".join(paragrafos).lower()
    for coluna in colunas:
        chave = atos.normalizar(coluna)
        valor = str(linha.get(chave, "")).strip()
        curto = valor.isdigit() and 0 < int(valor) < 10 and atos.por_extenso(int(valor)) in texto_todo
        if chave in CONTROLE or not (_util(valor) or curto):
            continue
        if coluna not in por_valor.setdefault(valor.lower(), []):
            por_valor[valor.lower()].append(coluna)
        escrito[valor.lower()] = valor
    # Números da planilha também escritos por extenso no ato (90 → "noventa")
    for coluna in colunas:
        chave = atos.normalizar(coluna)
        valor = str(linha.get(chave, "")).strip()
        if chave in CONTROLE or not valor.isdigit() or not 0 < int(valor) < 1000:
            continue
        extenso = atos.por_extenso(int(valor))
        if len(extenso) >= 3 and extenso.lower() not in por_valor:
            por_valor.setdefault(extenso.lower(), []).append(coluna + EXTENSO)
            escrito[extenso.lower()] = extenso

    # Palavras que mudam com o sexo (se a planilha tiver uma coluna com F/M)
    sexo = _coluna_sexo(colunas, registros)
    if sexo:
        letra = atos.normalizar(linha.get(atos.normalizar(sexo), ""))[:1]
        for fem, masc in GENERO:
            if fem == masc or letra not in ("F", "M"):
                continue
            termo = fem if letra == "F" else masc
            if termo.lower() not in por_valor and re.search(rf"(?<![\w]){re.escape(termo)}(?![\w])",
                                                              texto_todo, re.IGNORECASE):
                por_valor[termo.lower()] = [f"Muda com {sexo}: {fem} / {masc}"]
                escrito[termo.lower()] = termo

    # Ocorrências de cada valor; os mais longos primeiro, para "SMA 101/2026" ganhar de "2026"
    ocupado: dict[int, list[tuple[int, int]]] = {}
    encontrados: list[tuple[str, int, int, int]] = []
    for valor in sorted(por_valor, key=len, reverse=True):
        padrao = _padrao(escrito[valor])
        for n, texto in enumerate(paragrafos):
            for m in padrao.finditer(texto):
                if any(m.start() < f and i < m.end() for i, f in ocupado.get(n, [])):
                    continue  # está dentro de um valor mais longo já encontrado
                if valor.isdigit() and _parte_de_data(texto, m.start(), m.end()):
                    continue  # o "2026" de "06/07/2026" é da data, não o ano do processo
                ocupado.setdefault(n, []).append((m.start(), m.end()))
                encontrados.append((valor, n, m.start(), m.end()))

    achados: list[Achado] = []
    for valor in por_valor:
        ocorrencias = sorted((n, i, f) for v, n, i, f in encontrados if v == valor)
        if not ocorrencias:
            continue
        cols = por_valor[valor]
        n, i, f = ocorrencias[0]
        if len(ocorrencias) > 1 and (len(cols) > 1 or len(escrito[valor]) <= 10):
            # Mesmo valor em colunas diferentes (ex.: 15 em DIAS e em PADRAO) ou palavra curta
            # que pode ser texto fixo noutra parte ("Torna pública" x "servidora pública"):
            # cada vez que aparece pode ser uma coisa, então é conferida separadamente
            for n, i, f in ocorrencias:
                achados.append(Achado(paragrafos[n][i:f], cols, [(n, i, f)], _contexto(paragrafos[n], i, f)))
        else:
            achados.append(Achado(paragrafos[n][i:f], cols, ocorrencias, _contexto(paragrafos[n], i, f)))
    achados.sort(key=lambda a: a.ocorrencias[0])
    # Número logo depois de "Lei", "Decreto", "Portaria"... costuma ser referência fixa
    # ("Lei Complementar nº 340"), não dado da pessoa: vem desmarcado para a pessoa conferir
    for a in achados:
        n, ini, _ = a.ocorrencias[0]
        antes = paragrafos[n][max(0, ini - 40):ini]
        if a.valor.replace(".", "").isdigit() and REFERENCIA.search(antes):
            a.usar = False
    return achados, linha


def _trocar(p, trocas: list[tuple[int, int, str]]) -> None:
    """Troca trechos do parágrafo (início, fim, novo texto). O texto novo fica no pedaço (run)
    onde o trecho começa, com a formatação dele; o resto do trecho é apagado."""
    nos = list(p.iter(qn("w:t")))
    if not nos:  # parágrafo vazio (ex.: célula de tabela em branco): o texto entra num pedaço novo
        r = p.makeelement(qn("w:r"), {})
        nos = [r.makeelement(qn("w:t"), {})]
        r.append(nos[0])
        p.append(r)
    posicoes = [(no, k) for no in nos for k in range(len(no.text or ""))]
    novos = {id(no): list(no.text or "") for no in nos}
    novos[id(nos[-1])].append("")  # posição depois do último caractere (para inserir no fim)
    posicoes.append((nos[-1], len(nos[-1].text or "")))
    for ini, fim, texto in sorted(trocas, reverse=True):
        primeiro, k0 = posicoes[ini]
        if ini == fim:  # só inserir, sem apagar nada
            novos[id(primeiro)][k0] = texto + novos[id(primeiro)][k0]
            continue
        for pos in range(ini, fim):
            no, k = posicoes[pos]
            novos[id(no)][k] = ""
        novos[id(primeiro)][k0] = texto
    for no in nos:
        no.text = "".join(novos[id(no)])
        no.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _coluna_sexo(colunas: list[str], registros: list[dict]) -> str | None:
    """A coluna com o sexo da pessoa (quase só F/M), se houver."""
    for coluna in colunas:
        valores = [atos.normalizar(r.get(atos.normalizar(coluna), ""))[:1] for r in registros[:200]]
        valores = [v for v in valores if v]
        if len(valores) >= 1 and sum(v in ("F", "M") for v in valores) / len(valores) > 0.9 \
                and any(len(str(r.get(atos.normalizar(coluna), "")).strip()) <= 9 for r in registros[:5]):
            return coluna
    return None


def campo(opcao: str) -> str:
    """Nome do campo no modelo para a opção escolhida ("DIAS (por extenso)" → "DIAS POR EXTENSO",
    "Muda com SEXO: lotada / lotado" → "LOTADA OU LOTADO")."""
    if opcao.endswith(EXTENSO):
        return opcao[:-len(EXTENSO)] + " POR EXTENSO"
    if m := _GENERO_OPCAO.match(opcao):
        return f"{m.group(2)} ou {m.group(3)}".upper()
    return opcao


def ligacoes(achados: list[Achado]) -> dict:
    """Ligações que o processo precisa além de "coluna de mesmo nome" (ex.: por extenso)."""
    regras = {}
    for a in achados:
        if not a.usar:
            continue
        if a.coluna.endswith(EXTENSO):
            regras[campo(a.coluna)] = {"tipo": "extenso", "coluna": a.coluna[:-len(EXTENSO)]}
        elif m := _GENERO_OPCAO.match(a.coluna):
            regras[campo(a.coluna)] = {"tipo": "genero", "coluna": m.group(1),
                                       "feminino": m.group(2), "masculino": m.group(3)}
    return regras


def aplicar(caminho: Path, destino: Path, achados: list[Achado]) -> Path:
    """Grava o modelo: cada trecho escolhido vira {{COLUNA}}."""
    doc = Document(caminho)
    paragrafos = _paragrafos(doc)
    por_paragrafo: dict[int, list[tuple[int, int, str]]] = {}
    for achado in achados:
        if not achado.usar:
            continue
        for n, ini, fim in achado.ocorrencias:
            por_paragrafo.setdefault(n, []).append((ini, fim, "{{" + campo(achado.coluna) + "}}"))
    for n, trocas in por_paragrafo.items():
        _trocar(paragrafos[n], trocas)
    destino.parent.mkdir(parents=True, exist_ok=True)
    doc.save(destino)
    return destino


def titulo_do_modelo(caminho: Path, ano: str = "") -> str:
    """A primeira linha com texto do modelo, no formato dos títulos ({{X}} → {X}). Se o ano
    do número do ato aparecer fixo depois de uma barra ("/2026"), vira {ano do ato}."""
    for texto in textos_do_word(caminho):
        if texto.strip():
            titulo = re.sub(r"\{\{\s*([^{}]+?)\s*\}\}", r"{\1}", " ".join(texto.split()))
            if ano:
                titulo = re.sub(rf"(\}}\s*/\s*){ano}\b", r"\1{ano do ato}", titulo)
            return titulo
    return ""
