"""Incluir, editar e excluir linhas da planilha de dados pelo programa, sem abrir o Excel.

Vale para qualquer planilha: as colunas são descobertas pelo cabeçalho e pelas linhas que já
têm dados. Cada coluna pode ser:
  - digitada: a pessoa preenche (texto, número ou data, conforme o que já existe na coluna);
  - calculada: tem fórmula; o programa não mexe, só copia a fórmula para a linha nova
    (como o Excel faz ao arrastar a fórmula para baixo);
  - de controle: status_edocs, data_edocs... (gravadas pelo próprio programa).

A gravação mexe só nas células alteradas, direto no XML do .xlsx (como planilha_status.py):
o openpyxl descartaria os valores calculados das fórmulas e os vínculos com outros arquivos.
Depois de gravar, recalcular_no_excel() pede ao Excel (se estiver instalado) que recalcule e
salve, para os campos calculados (nome, cargo...) aparecerem na hora. Sem o Excel, eles se
atualizam quando a planilha for aberta nele (a gravação marca "recalcular ao abrir").
"""
import os
import re
import tempfile
import time
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import openpyxl
from lxml import etree

import atos
import trava
from planilha_status import N, NS, _caminho_aba, _letra, _numero_coluna, em_uso

# Colunas gravadas pelo próprio programa (não aparecem para editar)
CONTROLE = {"status_edocs", "data_edocs", "status_diario", "data_diario", "status_publicado", "status",
            "criado_por"}
AMOSTRA = 300  # linhas lidas para descobrir tipos e fórmulas
_ORIGEM_DATA = date(1899, 12, 30)  # dia zero das datas do Excel
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_TIPOS = "http://schemas.openxmlformats.org/package/2006/content-types"


class PlanilhaAberta(Exception):
    """O arquivo está aberto no Excel (não dá para gravar). quem = quem abriu ("" se não se sabe)."""

    def __init__(self, quem: str = ""):
        super().__init__(f"planilha aberta no Excel{f' por {quem}' if quem else ''}")
        self.quem = quem


def _um_por_vez(gravar):
    """Planilha na rede, usada por várias pessoas: só um computador grava por vez (os outros
    esperam a vez e gravam em cima do que ele salvou, sem apagar)."""
    def com_trava(planilha, *args, **kwargs):
        try:
            with trava.trava(Path(planilha)):
                return gravar(planilha, *args, **kwargs)
        except trava.Ocupado as e:
            raise PlanilhaAberta(f"{e.quem}, que está salvando nela agora") from None
    com_trava.__doc__ = gravar.__doc__
    com_trava.__name__ = gravar.__name__
    return com_trava


@dataclass
class Coluna:
    indice: int        # 1 = A
    nome: str          # como está no cabeçalho
    tipo: str          # "texto", "numero" ou "data"
    formula: bool      # calculada pelo Excel
    controle: bool     # gravada pelo programa

    @property
    def editavel(self) -> bool:
        return not (self.formula or self.controle)


# --- Leitura ----------------------------------------------------------------------

def _aba(wb, aba: str | None):
    aba = aba or atos.ABA
    return wb[aba] if aba in wb.sheetnames else wb.active


def ler_estrutura(planilha: Path, aba: str | None) -> tuple[list[Coluna], dict[int, dict[int, str]]]:
    """Numa leitura só (o arquivo é grande: cada leitura leva segundos): as colunas com o
    tipo de cada uma e os valores de todas as linhas {linha: {coluna: texto}}."""
    wb_f = openpyxl.load_workbook(planilha, read_only=True)
    wb_v = openpyxl.load_workbook(planilha, read_only=True, data_only=True)
    try:
        ws_f, ws_v = _aba(wb_f, aba), _aba(wb_v, aba)
        formulas = list(ws_f.iter_rows(min_row=2, max_row=AMOSTRA + 1, values_only=True))
        todas = list(ws_v.iter_rows(min_row=1, values_only=True))
    finally:
        wb_f.close()
        wb_v.close()
    cab = todas[0] if todas else ()
    valores = todas[1:AMOSTRA + 1]
    linhas = {n: {i + 1: _texto(v) for i, v in enumerate(l)} for n, l in enumerate(todas[1:], 2)}
    return _colunas(cab, formulas, valores), linhas


def esquema(planilha: Path, aba: str | None) -> list[Coluna]:
    """Colunas da planilha com o tipo de cada uma (pelas linhas que já têm dados)."""
    return ler_estrutura(planilha, aba)[0]


def _colunas(cab, formulas: list, valores: list) -> list[Coluna]:
    colunas = []
    for i, nome in enumerate(cab):
        if nome is None or not str(nome).strip():
            continue
        nome = str(nome).strip()
        eh_formula = any(isinstance(l[i], str) and l[i].startswith("=")
                         for l in formulas if i < len(l))
        tipos = {"data": 0, "numero": 0, "texto": 0}
        for l in valores:
            v = l[i] if i < len(l) else None
            if isinstance(v, (datetime, date)):
                tipos["data"] += 1
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                tipos["numero"] += 1
            elif v not in (None, ""):
                tipos["texto"] += 1
        tipo = max(tipos, key=lambda t: (tipos[t], t == "texto"))
        colunas.append(Coluna(i + 1, nome, tipo, eh_formula, nome.lower() in CONTROLE))
    return colunas


def _texto(v) -> str:
    return atos.formatar(v)


def ler_linha(planilha: Path, aba: str | None, linha: int) -> dict[int, str]:
    """{número da coluna: valor como texto} de uma linha."""
    wb = openpyxl.load_workbook(planilha, read_only=True, data_only=True)
    try:
        valores = next(_aba(wb, aba).iter_rows(min_row=linha, max_row=linha, values_only=True), ())
    finally:
        wb.close()
    return {i + 1: _texto(v) for i, v in enumerate(valores)}


def _vazio(v) -> bool:
    return v is None or (isinstance(v, str) and (not v.strip() or v.strip() in atos.ERROS_EXCEL))


def ultima_linha(planilha: Path, aba: str | None, colunas: list[Coluna]) -> int:
    """Última linha com algum valor digitado (as de baixo podem ter só fórmulas prontas ou
    restos de erro do Excel, como #N/A colado como valor)."""
    digitadas = [c.indice - 1 for c in colunas if c.editavel]
    wb = openpyxl.load_workbook(planilha, read_only=True, data_only=True)
    try:
        ultima = 1
        for n, valores in enumerate(_aba(wb, aba).iter_rows(min_row=2, values_only=True), 2):
            if any(i < len(valores) and not _vazio(valores[i]) for i in digitadas):
                ultima = n
    finally:
        wb.close()
    return ultima


# --- Gravação no XML ---------------------------------------------------------------

def _linha_xml(dados, numero: int, criar: bool = True):
    row = dados.find(f"m:row[@r='{numero}']", N)
    if row is None and criar:
        row = etree.Element(f"{{{NS}}}row", r=str(numero))
        depois = next((r for r in dados if int(r.get("r")) > numero), None)
        if depois is None:
            dados.append(row)
        else:
            depois.addprevious(row)
    return row


def _celula(row, col: int, criar: bool = True):
    ref = f"{_letra(col)}{row.get('r')}"
    celula = row.find(f"m:c[@r='{ref}']", N)
    if celula is None and criar:
        celula = etree.Element(f"{{{NS}}}c", r=ref)
        depois = next((c for c in row if _numero_coluna(c.get("r")) > col), None)
        if depois is None:
            row.append(celula)
        else:
            depois.addprevious(celula)
        spans = row.get("spans")
        if spans:
            ini, fim = (int(x) for x in spans.split(":"))
            row.set("spans", f"{min(ini, col)}:{max(fim, col)}")
    return celula


def _limpar(celula) -> None:
    for filho in list(celula):
        celula.remove(filho)
    celula.attrib.pop("t", None)


def _numero(texto: str) -> float | None:
    texto = texto.strip().replace(" ", "")
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", texto):  # 1.234,56
        texto = texto.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d+,\d+", texto):  # 12,5
        texto = texto.replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None


def _escrever(celula, texto: str, tipo: str) -> None:
    """Grava o valor com o tipo certo (número e data como número do Excel, para fórmulas e
    PROCV continuarem funcionando). O estilo da célula (formato de data etc.) é mantido."""
    _limpar(celula)
    texto = (texto or "").strip()
    if not texto:
        return
    if tipo == "data" and (dia := atos.ler_data(texto)):
        etree.SubElement(celula, f"{{{NS}}}v").text = str((dia - _ORIGEM_DATA).days)
        return
    if tipo == "numero" and (numero := _numero(texto)) is not None:
        etree.SubElement(celula, f"{{{NS}}}v").text = (
            str(int(numero)) if numero.is_integer() else repr(numero))
        return
    celula.set("t", "inlineStr")
    is_ = etree.SubElement(celula, f"{{{NS}}}is")
    etree.SubElement(is_, f"{{{NS}}}t").text = texto


def _fora_de_aspas(formula: str, trocar) -> str:
    """Aplica 'trocar' só nos trechos da fórmula fora de "textos" e 'nomes de aba'."""
    partes = re.split(r"(\"[^\"]*\"|'[^']*')", formula)
    return "".join(p if i % 2 else trocar(p) for i, p in enumerate(partes))


_REF = re.compile(r"(?<![A-Za-z_\d.\]])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])")


def deslocar_formula(formula: str, linhas: int) -> str:
    """Ajusta as referências relativas de linha, como o Excel ao arrastar para baixo:
    A2 → A3, $A2 → $A3; A$2 e $A$2 ficam iguais."""
    def trocar(trecho: str) -> str:
        return _REF.sub(lambda m: f"{m[1]}{m[2]}{m[3]}{m[4] if m[3] else int(m[4]) + linhas}", trecho)
    return _fora_de_aspas(formula, trocar)


def _formula_de(dados, celula) -> tuple[str, int] | None:
    """(texto da fórmula, linha de onde ela vem) de uma célula com fórmula, inclusive as
    "compartilhadas" do Excel (só a primeira célula do grupo guarda o texto)."""
    f = celula.find("m:f", N)
    if f is None:
        return None
    linha = int(re.search(r"\d+", celula.get("r")).group())
    if f.text:
        return f.text, linha
    si = f.get("si")
    if f.get("t") == "shared" and si is not None:
        mestre = dados.find(f".//m:f[@t='shared'][@si='{si}'][@ref]", N)
        if mestre is not None and mestre.text:
            origem = int(re.search(r"\d+", mestre.getparent().get("r")).group())
            return mestre.text, origem
    return None


def _ajustar_dimensao(sheet, dados) -> None:
    """O tamanho da aba gravado no arquivo ("A1:W120") cobre as linhas novas: quem lê só até
    ele (o próprio programa, sem o Excel para corrigir) não deixa de ver a linha incluída
    e grava a próxima por cima dela."""
    dim = sheet.find("m:dimension", N)
    if dim is None:
        return
    linhas = [int(r.get("r")) for r in dados.findall("m:row", N) if (r.get("r") or "").isdigit()]
    colunas = [_numero_coluna(c.get("r")) for c in dados.iter(f"{{{NS}}}c") if c.get("r")]
    if not linhas or not colunas:
        return
    m = re.match(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$", dim.get("ref", ""))
    inicio = f"{m.group(1)}{m.group(2)}" if m else "A1"
    ultima_col = max(colunas + ([_numero_coluna(m.group(3))] if m and m.group(3) else []))
    ultima_linha_ = max(linhas + ([int(m.group(4))] if m and m.group(4) else []))
    dim.set("ref", f"{inicio}:{_letra(ultima_col)}{ultima_linha_}")


@_um_por_vez
def _salvar(planilha: Path, aba: str | None, alterar, recalcular_ao_abrir: bool = True) -> None:
    """Abre o .xlsx, deixa alterar(dados da aba) mexer nas linhas e grava de volta.
    Lança PlanilhaAberta se ela estiver aberta no Excel (aqui ou por outra pessoa na rede).
    recalcular_ao_abrir: marca para o Excel recalcular tudo ao abrir (use False quando o
    programa já vai recalcular pelo Excel logo em seguida: recalcular tudo é lento)."""
    if (quem := em_uso(planilha)) is not None:
        raise PlanilhaAberta(quem)
    for tentativa in range(20):
        try:
            with zipfile.ZipFile(planilha) as z:
                nome_aba = _caminho_aba(z, aba or atos.ABA)
                arquivos = {i.filename: (i, z.read(i.filename)) for i in z.infolist()}
            break
        except PermissionError:
            if tentativa == 19:
                raise PlanilhaAberta("feche a planilha no Excel para salvar") from None
            time.sleep(0.25)

    sheet = etree.fromstring(arquivos[nome_aba][1])
    dados = sheet.find("m:sheetData", N)
    criou_formula = alterar(dados)
    _ajustar_dimensao(sheet, dados)
    novos = {nome_aba: etree.tostring(sheet, xml_declaration=True, encoding="UTF-8", standalone=True)}

    # Excel recalcula tudo ao abrir (os valores calculados gravados aqui ficam velhos)
    if recalcular_ao_abrir:
        novos["xl/workbook.xml"] = _com_recalculo(arquivos["xl/workbook.xml"][1])

    # Fórmulas novas não estão na "cadeia de cálculo" salva: sem ela, o Excel refaz sozinho
    if criou_formula and "xl/calcChain.xml" in arquivos:
        del arquivos["xl/calcChain.xml"]
        rels = etree.fromstring(arquivos["xl/_rels/workbook.xml.rels"][1])
        for r in list(rels):
            if r.get("Target", "").endswith("calcChain.xml"):
                rels.remove(r)
        novos["xl/_rels/workbook.xml.rels"] = etree.tostring(rels, xml_declaration=True,
                                                             encoding="UTF-8", standalone=True)
        tipos = etree.fromstring(arquivos["[Content_Types].xml"][1])
        for t in list(tipos):
            if t.get("PartName", "").endswith("calcChain.xml"):
                tipos.remove(t)
        novos["[Content_Types].xml"] = etree.tostring(tipos, xml_declaration=True,
                                                      encoding="UTF-8", standalone=True)

    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=planilha.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as saida:
            for nome, (info, conteudo) in arquivos.items():
                saida.writestr(info, novos.get(nome, conteudo))
        for tentativa in range(20):  # na rede, outro computador pode estar lendo por um instante
            try:
                os.replace(tmp, planilha)
                break
            except PermissionError:
                if tentativa == 19:
                    raise PlanilhaAberta("feche a planilha no Excel para salvar") from None
                time.sleep(0.25)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def _com_recalculo(workbook_xml: bytes) -> bytes:
    wb = etree.fromstring(workbook_xml)
    calc = wb.find("m:calcPr", N)
    if calc is None:
        calc = etree.SubElement(wb, f"{{{NS}}}calcPr")
    calc.set("fullCalcOnLoad", "1")
    return etree.tostring(wb, xml_declaration=True, encoding="UTF-8", standalone=True)


def marcar_recalculo(planilha: Path) -> None:
    """Marca a planilha para o Excel recalcular tudo ao abrir (quando não deu para recalcular
    agora, ex.: computador sem Excel)."""
    _salvar(planilha, None, lambda dados: False, recalcular_ao_abrir=True)


# --- Operações -----------------------------------------------------------------------

@_um_por_vez
def editar_linha(planilha: Path, aba: str | None, linha: int, colunas: list[Coluna],
                 valores: dict[int, str], recalcular_ao_abrir: bool = True) -> None:
    """Grava os valores {coluna: texto} na linha (só colunas digitadas)."""
    por_indice = {c.indice: c for c in colunas}

    def alterar(dados) -> bool:
        row = _linha_xml(dados, linha)
        for col, texto in valores.items():
            if por_indice[col].editavel:
                _escrever(_celula(row, col), texto, por_indice[col].tipo)
        return False

    _salvar(planilha, aba, alterar, recalcular_ao_abrir)


def incluir_linha(planilha: Path, aba: str | None, colunas: list[Coluna],
                  valores: dict[int, str], recalcular_ao_abrir: bool = True) -> int:
    """Cria a linha logo depois da última com dados: copia da linha de cima as fórmulas e o
    formato, e grava os valores digitados. Devolve o número da linha nova."""
    return incluir_linhas(planilha, aba, colunas, [valores], recalcular_ao_abrir)[0]


@_um_por_vez
def incluir_linhas(planilha: Path, aba: str | None, colunas: list[Coluna],
                   lista: list[dict[int, str]], recalcular_ao_abrir: bool = True) -> list[int]:
    """Várias linhas novas de uma vez (uma leitura e uma gravação da planilha, em vez de uma
    por linha): cada uma copia as fórmulas e o formato da última linha com dados. Devolve os
    números das linhas novas, na ordem da lista."""
    modelo = ultima_linha(planilha, aba, colunas)
    novas = [modelo + 1 + i for i in range(len(lista))]
    por_indice = {c.indice: c for c in colunas}

    def alterar(dados) -> bool:
        criou = False
        de_cima = _linha_xml(dados, modelo, criar=False)
        for nova, valores in zip(novas, lista):
            row = _linha_xml(dados, nova)
            if de_cima is not None:
                for attr in ("ht", "customHeight", "spans"):
                    if de_cima.get(attr) and not row.get(attr):
                        row.set(attr, de_cima.get(attr))
                for celula_cima in de_cima.findall("m:c", N):
                    col = _numero_coluna(celula_cima.get("r"))
                    celula = _celula(row, col)
                    if celula_cima.get("s") and not celula.get("s"):
                        celula.set("s", celula_cima.get("s"))
                    coluna = por_indice.get(col)
                    if coluna and coluna.formula and celula.find("m:f", N) is None:
                        origem = _formula_de(dados, celula_cima)
                        if origem:
                            texto, linha_origem = origem
                            _limpar(celula)
                            if celula_cima.get("t") == "str":
                                celula.set("t", "str")
                            f = etree.SubElement(celula, f"{{{NS}}}f")
                            f.text = deslocar_formula(texto, nova - linha_origem)
                            criou = True
            for col, texto in valores.items():
                if por_indice[col].editavel:
                    _escrever(_celula(row, col), texto, por_indice[col].tipo)
        return criou

    _salvar(planilha, aba, alterar, recalcular_ao_abrir)
    return novas


@_um_por_vez
def limpar_linha(planilha: Path, aba: str | None, linha: int, colunas: list[Coluna]) -> None:
    """Apaga os valores digitados e os de controle da linha (as fórmulas ficam e passam a
    dar vazio). É a exclusão possível sem o Excel."""
    def alterar(dados) -> bool:
        row = _linha_xml(dados, linha, criar=False)
        if row is not None:
            for c in colunas:
                if not c.formula and (celula := _celula(row, c.indice, criar=False)) is not None:
                    _limpar(celula)
        return False

    _salvar(planilha, aba, alterar)


# --- Excel -------------------------------------------------------------------------

def tem_excel() -> bool:
    """O Excel (e o pywin32, para controlá-lo) estão disponíveis neste computador?"""
    try:
        import winreg

        import win32com.client  # noqa: F401
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application"))
        return True
    except (ImportError, OSError):
        return False


def _folha(livro, aba: str | None):
    folhas = [livro.Worksheets(i).Name for i in range(1, livro.Worksheets.Count + 1)]
    nome = aba or atos.ABA
    return livro.Worksheets(nome if nome in folhas else folhas[0])


@_um_por_vez
def _com_excel(planilha: Path, acao) -> bool:
    """Abre a planilha num Excel escondido, executa acao(planilha_do_excel) e salva.
    False se não houver Excel (ou pywin32) neste computador, ou se der erro."""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")  # instância só nossa
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        livro = excel.Workbooks.Open(str(planilha.resolve()), UpdateLinks=0)
        acao(livro)
        livro.Save()
        livro.Close(False)
        return True
    except Exception:
        return False
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def recalcular_no_excel(planilha: Path, aba: str | None, linhas: list[int]) -> bool:
    """Recalcula pelo Excel as fórmulas das linhas mexidas (nome, cargo... aparecem na hora).
    Só as linhas pedidas: recalcular a planilha inteira (milhares de PROCV) é lento."""
    def calcular(livro) -> None:
        folha = _folha(livro, aba)
        for linha in linhas:
            folha.Rows(linha).Dirty()  # o Excel não sabe que mudamos a linha por fora
            folha.Rows(linha).Calculate()

    return _com_excel(planilha, calcular)


@_um_por_vez
def excluir_linhas(planilha: Path, aba: str | None, linhas: list[int], colunas: list[Coluna]) -> str:
    """Várias linhas de uma vez, de baixo para cima (excluir uma não muda o número das de
    cima). Como excluir_linha: "excluida" (com o Excel) ou "limpa"."""
    ordem = sorted(set(linhas), reverse=True)

    def apagar(livro) -> None:
        folha = _folha(livro, aba)
        for linha in ordem:
            folha.Rows(linha).Delete()

    if (quem := em_uso(planilha)) is not None:
        raise PlanilhaAberta(quem)
    if _com_excel(planilha, apagar):
        return "excluida"
    for linha in ordem:
        limpar_linha(planilha, aba, linha, colunas)
    return "limpa"


def excluir_linha(planilha: Path, aba: str | None, linha: int, colunas: list[Coluna]) -> str:
    """Exclui a linha. Com o Excel, apaga a linha de verdade (as de baixo sobem e as fórmulas
    se ajustam). Sem ele, limpa os valores digitados. Devolve "excluida" ou "limpa"."""
    def apagar(livro) -> None:
        _folha(livro, aba).Rows(linha).Delete()

    if (quem := em_uso(planilha)) is not None:
        raise PlanilhaAberta(quem)
    if _com_excel(planilha, apagar):
        return "excluida"
    limpar_linha(planilha, aba, linha, colunas)
    return "limpa"
