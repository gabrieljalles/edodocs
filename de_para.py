"""De-para: listas de código → descrição, para as tabelas de consulta que só trazem o código.

Ex.: a tabela de servidores traz LOTACAO = "0312", mas o ato precisa de "Secretaria Municipal
de Governo". A pessoa cria (em Configurações) o de-para "Secretarias" com cada código e a sua
descrição, e no Preencher sozinho do processo liga a coluna a ele: o que vem da tabela já
entra traduzido. Um de-para serve para todos os processos.

Fica em configs/config.json:
    "de_para": {"Secretarias": {"0312": "Secretaria Municipal de Governo", "0401": ""}}
Código sem descrição (ainda por preencher) entra como está na tabela.
Os códigos são comparados sem pontos, traços e zeros à esquerda ("03.12" = "312").

Às vezes só um pedaço do valor é o código (ex.: os 2 primeiros dígitos da lotação dizem a
secretaria). A "parte" diz qual: {"de": "inicio" ou "fim", "n": 2}; sem parte, o valor todo.
Ela fica na ligação do processo (processo.json, "de_para_parte"), não no de-para.
"""
import csv
import io
import json
import re
from pathlib import Path

import atos
import main

CHAVE_CONFIG = "de_para"
COBERTURA_MINIMA = 0.8  # para sugerir sozinho: quantos dos valores da coluna o de-para conhece


def todos() -> dict[str, dict[str, str]]:
    return dict(main.ler_config().get(CHAVE_CONFIG) or {})


def salvar(nome: str, pares: dict[str, str], antigo: str | None = None) -> None:
    """Grava o de-para (antigo: o nome anterior, se a pessoa renomeou)."""
    lista = todos()
    if antigo and antigo != nome:
        lista.pop(antigo, None)
    lista[nome] = {str(k).strip(): str(v or "").strip() for k, v in pares.items() if str(k).strip()}
    main.salvar_config(**{CHAVE_CONFIG: lista})
    if antigo and antigo != nome:
        _renomear_nos_processos(antigo, nome)


def _renomear_nos_processos(antigo: str, nome: str) -> None:
    import processos  # aqui: processos não precisa do de-para para carregar
    for pasta in usados().get(antigo, []):
        proc = processos.carregar(main.PROCESSOS / pasta)
        for item in (proc.config.get("consulta") or {}).get("puxar") or []:
            if item.get("de_para") == antigo:
                item["de_para"] = nome
        proc.salvar()


def remover(nome: str) -> None:
    lista = todos()
    lista.pop(nome, None)
    main.salvar_config(**{CHAVE_CONFIG: lista})


def chave(codigo) -> str:
    """Como o código é comparado: sem pontos, traços, espaços e zeros à esquerda."""
    n = atos.normalizar(codigo)
    return (n.lstrip("0") or "0") if n.isdigit() else n


def _indice(pares: dict[str, str]) -> dict[str, str]:
    return {chave(k): v for k, v in pares.items() if v}


def parte(valor: str, regra: dict | None = None) -> str:
    """O pedaço do valor que é o código: o todo, ou os n primeiros/últimos caracteres (sem
    contar pontos, traços e espaços: "03.12.001" → 2 primeiros → "03")."""
    n = int((regra or {}).get("n") or 0)
    if not valor or n <= 0:
        return valor
    limpo = atos.normalizar(valor) or str(valor)
    return limpo[-n:] if regra.get("de") == "fim" else limpo[:n]


def descrever_parte(regra: dict | None) -> str:
    """"2 primeiros" / "3 últimos" / "" (o valor todo)."""
    n = int((regra or {}).get("n") or 0)
    if n <= 0:
        return ""
    return f"{n} últimos" if regra.get("de") == "fim" else f"{n} primeiros"


def traduzir(nome: str, valor: str, lista: dict | None = None, regra: dict | None = None) -> str:
    """A descrição do código (ou da parte dele); sem de-para, sem o código ou sem descrição:
    o próprio valor."""
    pares = (lista if lista is not None else todos()).get(nome) or {}
    return _indice(pares).get(chave(parte(valor, regra)), valor) if valor else valor


def cobertura(pares: dict[str, str], valores: list[str], regra: dict | None = None) -> float:
    """0 a 1: dos valores (de uma coluna da tabela), quantos o de-para conhece."""
    valores = [v for v in valores if v]
    if not valores:
        return 0.0
    conhecidos = {chave(k) for k in pares}
    return sum(chave(parte(v, regra)) in conhecidos for v in valores) / len(valores)


def parece_codigo(valores: list[str], linhas: int = 0) -> bool:
    """A coluna tem cara de código (0312, 12.04, A15): curtos, com número, sem ser data, e
    repetidos entre as pessoas (matrícula e CPF, um por pessoa, não contam)."""
    valores = [v for v in valores if v]
    if not valores or (linhas >= 10 and len(valores) > linhas / 2):
        return False
    amostra = valores[:200]
    curtos = [v for v in amostra if 2 <= len(v) <= 10 and any(c.isdigit() for c in v) and " " not in v.strip()
              and not re.fullmatch(r"\d{1,2}/\d{1,2}/\d{2,4}", v)]
    return len(curtos) / len(amostra) >= 0.8


def sugerir(valores: list[str], lista: dict | None = None) -> str | None:
    """O de-para que conhece quase todos os valores da coluna (se houver)."""
    lista = lista if lista is not None else todos()
    notas = [(cobertura(pares, valores[:300]), nome) for nome, pares in lista.items()]
    nota, nome = max(notas, default=(0.0, None))
    return nome if nota >= COBERTURA_MINIMA else None


def de_texto(texto: str) -> dict[str, str]:
    """Pares colados do Excel (código e descrição lado a lado; tab, ; ou ,). Linhas com uma
    coluna só viram código sem descrição."""
    texto = (texto or "").strip("\r\n")
    if not texto:
        return {}
    try:
        dialeto = csv.Sniffer().sniff(texto[:3000], delimiters="\t;,")
    except csv.Error:
        dialeto = csv.excel_tab
    pares = {}
    for linha in csv.reader(io.StringIO(texto), dialeto):
        celulas = [c.strip() for c in linha]
        if celulas and celulas[0]:
            pares[celulas[0]] = " ".join(c for c in celulas[1:] if c)
    return pares


def de_tabela(tabela, coluna_codigo: str, coluna_descricao: str | None = None,
              regra: dict | None = None) -> dict[str, str]:
    """Pares de uma tabela já lida (consultas.Tabela): todos os códigos diferentes da coluna
    (ou da parte dele), com a descrição da outra coluna (ou em branco, para preencher)."""
    pares: dict[str, str] = {}
    for linha in tabela.linhas:
        codigo = parte(linha.get(coluna_codigo, ""), regra)
        if codigo and not pares.get(codigo):
            pares[codigo] = linha.get(coluna_descricao, "") if coluna_descricao else ""
    return pares


def exportar_excel(nomes: list[str], destino: Path, lista: dict | None = None) -> Path:
    """Grava os de-para num Excel (uma folha por de-para: CÓDIGO e DESCRIÇÃO), para levar a
    outro computador ou editar no Excel. O nome inteiro fica em E1 (o da folha é cortado)."""
    import openpyxl
    lista = lista if lista is not None else todos()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    usados_nomes: set[str] = set()
    for nome in nomes:
        folha = re.sub(r"[\[\]:*?/\\]", "_", nome)[:31] or "De-para"
        n = 2
        while folha.lower() in usados_nomes:
            folha = f"{folha[:28]} {n}"
            n += 1
        usados_nomes.add(folha.lower())
        ws = wb.create_sheet(folha)
        ws.append(["CÓDIGO", "DESCRIÇÃO", "", "DE-PARA", nome])
        for codigo, descricao in (lista.get(nome) or {}).items():
            ws.append([str(codigo), descricao])
        for celula in ws["A"]:
            celula.number_format = "@"  # código como texto: 0312 continua 0312
        ws.column_dimensions["A"].width = 16
        ws.column_dimensions["B"].width = 60
    if not wb.worksheets:
        wb.create_sheet("De-para").append(["CÓDIGO", "DESCRIÇÃO"])
    wb.save(destino)
    return destino


def ler_excel(arquivo: Path) -> dict[str, dict[str, str]]:
    """{nome: pares} de um Excel de de-para (o que exportar_excel grava; numa planilha feita à
    mão, cada folha com código e descrição nas colunas A e B vira um de-para)."""
    import openpyxl
    wb = openpyxl.load_workbook(arquivo, read_only=True, data_only=True)
    saida: dict[str, dict[str, str]] = {}
    try:
        for ws in wb.worksheets:
            linhas = list(ws.iter_rows(values_only=True))
            if not linhas:
                continue
            cab = [atos.formatar(v) for v in linhas[0]]
            nome = cab[4] if len(cab) > 4 and atos.normalizar(cab[3]) == "DEPARA" and cab[4] else ws.title
            pares = {}
            for linha in linhas[1:]:
                codigo = atos.formatar(linha[0]) if linha else ""
                if codigo:
                    pares[codigo] = atos.formatar(linha[1]) if len(linha) > 1 else ""
            if pares:
                saida[nome] = pares
    finally:
        wb.close()
    return saida


def usados(pastas: list[Path] | None = None) -> dict[str, list[str]]:
    """{de-para: processos que o usam no Preencher sozinho}."""
    saida: dict[str, list[str]] = {}
    for pasta in pastas if pastas is not None else main.listar_processos():
        try:
            config = json.loads((pasta / "processo.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        for item in (config.get("consulta") or {}).get("puxar") or []:
            if item.get("de_para") and pasta.name not in saida.get(item["de_para"], []):
                saida.setdefault(item["de_para"], []).append(pasta.name)
    return saida


def juntar(nome: str, pares: dict[str, str]) -> int:
    """Junta pares que vieram de fora (ex.: processos importados) ao de-para daqui, sem trocar
    descrição que já existe. Devolve quantos códigos entraram ou ganharam descrição."""
    atual = dict(todos().get(nome) or {})
    indice = {chave(k): k for k in atual}
    novos = 0
    for codigo, descricao in pares.items():
        k = indice.get(chave(codigo))
        if k is None:
            atual[codigo] = descricao
            novos += 1
        elif descricao and not atual[k]:
            atual[k] = descricao
            novos += 1
    if novos or nome not in todos():
        salvar(nome, atual)
    return novos
