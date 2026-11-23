"""Montar com IA: a pessoa junta atos que já fez (Word ou PDF) e a IA monta o processo.

A IA não mexe em arquivo nenhum. Ela lê o guia (assets/ia/guia_processos.md) e o texto dos
exemplos e devolve um plano em JSON. Este módulo faz o resto, aqui no computador:

    exemplos = ler_arquivos(arquivos, pasta)          texto de cada arquivo (.doc vira .docx)
    resultado = montar_com_ia(exemplos, pasta, ...)   pede o plano, monta e confere

  1. o Word: copia o exemplo base e aplica as trocas do plano sem perder a formatação; no
     ato de várias pessoas, põe {{REPETIR}}/{{FIM REPETIR}} na 1ª pessoa e tira as outras;
  2. a planilha (planilha_nova.criar), vazia, e as regras (processo.json);
  3. a conferência: refaz cada exemplo com os dados que a IA leu dele e compara com o
     original. O que não bate volta para a IA corrigir (até RODADAS vezes).

O resultado abre no assistente, na prévia, para a pessoa conferir e salvar (ver
rascunho_do_resultado).
"""
import copy
import difflib
import json
import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

import atos
import ia
import planilha_nova
import processos
from modelo_automatico import _trocar

EXTENSOES = (".docx", ".doc", ".pdf")
MAX_ARQUIVOS = 15
LIMITE_POR_ARQUIVO = 40_000   # caracteres de cada exemplo mandados à IA
RODADAS = 2                   # quantas vezes a IA pode corrigir o plano
MAX_PERGUNTAS = 20            # perguntas da IA para a pessoa responder antes de terminar
# O que o plano pode dizer do processo (modelo, planilha e aba o Edodocs preenche; quem
# assina e as colunas da tabela a pessoa escolhe no assistente)
CHAVES_DO_PLANO = {"coluna_nome", "coluna_numero", "titulo_edocs", "titulo_diario", "titulo_no_ato",
                   "titulo_na_primeira_linha", "juntar", "vai_ao_diario", "precisa_anexo", "obrigatorios",
                   "edocs", "diario", "ligacoes"}
INVALIDOS = re.compile(r'[<>:"/\\|?*]')
_ASPAS = "\"“”„″"
_APOSTROFOS = "'‘’´`"
_TRACOS = "-–—‐‑"


@dataclass
class Exemplo:
    nome: str                # como a IA vê (o nome do arquivo)
    texto: str
    word: Path | None        # o .docx (o .doc convertido); None num PDF: não serve de base


@dataclass
class Comparacao:
    arquivo: str
    igual: bool
    parecido: int            # % de semelhança
    diferencas: list[tuple[str, str]] = field(default_factory=list)  # (no exemplo, no Edodocs)


@dataclass
class Resultado:
    plano: dict
    config: dict             # o processo.json (modelo, outros_modelos e planilha: nomes em pasta)
    pasta: Path              # onde ficaram o(s) Word(s) e a planilha
    planilha: Path
    problemas: list[str]     # o que ainda não ficou certo (a pessoa vê e corrige no assistente)
    comparacoes: list[Comparacao]
    rodadas: int             # quantas correções a IA fez


# --- Leitura dos arquivos ------------------------------------------------------------------
def texto_do_word(word: Path) -> str:
    """O texto do Word como o Edodocs o vê ao montar o ato (tabelas: células com " | ")."""
    _, _, bloco = atos._abrir_modelo(word)
    return "\n".join(atos._texto_bloco(el) for el in atos.blocos(bloco, False)).strip()


def _texto_pdf(pdf: Path) -> str:
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(pdf)
    try:
        return "\n".join(p.get_textpage().get_text_range() for p in doc).strip()
    finally:
        doc.close()


def _doc_para_docx(origem: Path, destino: Path) -> Path:
    """Word antigo (.doc) para .docx, pelo próprio Word do computador (escondido)."""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise RuntimeError(f"{origem.name}: abra no Word e salve como .docx") from None
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(origem.resolve()), ReadOnly=True, AddToRecentFiles=False,
                                  ConfirmConversions=False)
        doc.SaveAs2(str(destino.resolve()), FileFormat=16)  # 16 = .docx
        doc.Close(False)
        return destino
    except Exception:
        raise RuntimeError(f"{origem.name}: não consegui abrir no Word. Abra e salve como .docx") from None
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def ler_arquivos(arquivos: list[Path], pasta: Path) -> list[Exemplo]:
    """O texto de cada arquivo. Os Words são copiados para pasta/exemplos (o original da
    pessoa nunca é mexido). Arquivo que não dá para ler gera RuntimeError com o motivo."""
    destino = pasta / "exemplos"
    destino.mkdir(parents=True, exist_ok=True)
    exemplos, nomes = [], set()
    for caminho in arquivos:
        nome = caminho.name
        if nome.lower() in nomes:  # dois arquivos com o mesmo nome (de pastas diferentes)
            nome = f"{len(exemplos) + 1} - {nome}"
        nomes.add(nome.lower())
        extensao = caminho.suffix.lower()
        if extensao == ".pdf":
            texto, word = _texto_pdf(caminho), None
            if len(texto.strip()) < 30:
                raise RuntimeError(f"{caminho.name}: este PDF é uma imagem (digitalizado), não tem texto")
        elif extensao in (".docx", ".doc"):
            word = destino / (Path(nome).stem + ".docx")
            if extensao == ".doc":
                _doc_para_docx(caminho, word)
            else:
                shutil.copy2(caminho, word)
            try:
                texto = texto_do_word(word)
            except Exception:
                raise RuntimeError(f"{caminho.name}: não consegui ler este Word") from None
            nome = word.name
        else:
            raise RuntimeError(f"{caminho.name}: só Word (.docx, .doc) ou PDF")
        exemplos.append(Exemplo(nome, texto, word))
    if not any(e.word for e in exemplos):
        raise RuntimeError("Junte pelo menos um Word (.docx ou .doc): o modelo sai dele. PDF só ajuda a "
                           "IA a ver as variações.")
    return exemplos


# --- Pedido à IA --------------------------------------------------------------------------
def pedido(exemplos: list[Exemplo], descricao: str = "") -> str:
    partes = ["Monte um processo do Edodocs a partir destes atos já feitos. Responda só com o "
              "plano em JSON (seção 6 do guia)."]
    if descricao.strip():
        partes.append("O que a pessoa explicou sobre o processo:\n" + descricao.strip())
    partes.append(f"São {len(exemplos)} arquivo(s). Só os Words podem ser \"base\" de um modelo.")
    for e in exemplos:
        tipo = "Word, pode ser base" if e.word else "PDF, só texto: não pode ser base"
        texto = e.texto if len(e.texto) <= LIMITE_POR_ARQUIVO else e.texto[:LIMITE_POR_ARQUIVO] + "\n[...cortado]"
        partes.append(f'=== ARQUIVO: "{e.nome}" ({tipo}) ===\n{texto}\n=== FIM DE "{e.nome}" ===')
    return "\n\n".join(partes)


def pedido_de_correcao(original: str, plano: dict, problemas: list[str]) -> str:
    return (original + "\n\n=== O SEU PLANO ANTERIOR ===\n" + json.dumps(plano, ensure_ascii=False)
            + "\n\n=== O QUE O EDODOCS ENCONTROU AO MONTAR E REFAZER OS EXEMPLOS ===\n"
            + "\n".join(f"- {p}" for p in problemas)
            + "\n\nCorrija esses pontos e devolva o plano inteiro de novo (só o JSON).")


# --- Perguntas da IA para a pessoa ----------------------------------------------------------
def perguntas_do_plano(plano: dict) -> list[dict]:
    """As perguntas do plano, limpas: {"pergunta", "por_que", "contexto", "opcoes"}, com 2 a 4
    opções {"texto", "como_fica"}. A 1ª opção é a que o plano já segue (a sugestão da IA)."""
    saida = []
    for p in _lista(plano.get("perguntas")):
        if not isinstance(p, dict):
            continue
        texto = str(p.get("pergunta") or "").strip()
        opcoes, vistos = [], set()
        for o in _lista(p.get("opcoes")):
            o = o if isinstance(o, dict) else {"texto": o}
            rotulo = str(o.get("texto") or "").strip()
            if rotulo and rotulo not in vistos:
                vistos.add(rotulo)
                opcoes.append({"texto": rotulo, "como_fica": str(o.get("como_fica") or "").strip()})
        if texto and len(opcoes) >= 2:
            saida.append({"pergunta": texto, "por_que": str(p.get("por_que") or "").strip(),
                          "contexto": str(p.get("contexto") or "").strip(), "opcoes": opcoes[:4]})
    return saida[:MAX_PERGUNTAS]


def precisa_ajustar(respostas: list[dict]) -> bool:
    """Só volta para a IA se alguma resposta não for a que o plano já segue."""
    return any(not r.get("sugerida") for r in respostas)


def pedido_de_respostas(original: str, plano: dict, respostas: list[dict]) -> str:
    linhas = [f'- Pergunta: {r["pergunta"]}\n  Resposta: {r["resposta"]}' for r in respostas]
    return (original + "\n\n=== O SEU PLANO ===\n" + json.dumps(plano, ensure_ascii=False)
            + "\n\n=== O QUE A PESSOA RESPONDEU ÀS SUAS PERGUNTAS ===\n" + "\n".join(linhas)
            + "\n\nAjuste o plano conforme as respostas (mude só o que elas pedem) e devolva o plano "
              "inteiro de novo (só o JSON), com \"perguntas\": [].")


# --- Conferência do plano (antes de montar) ----------------------------------------------
def _lista(valor) -> list:
    return valor if isinstance(valor, list) else []


def _citadas_nos_titulos(texto: str) -> list[str]:
    especiais = {processos.NUMERO_DO_ATO, processos.ANO_DO_ATO}
    return [m.group(1) for m in processos._TOKEN.finditer(texto or "")
            if atos.normalizar(m.group(1)) not in especiais]


def conferir_plano(plano: dict, exemplos: list[Exemplo]) -> list[str]:
    """O que impede de montar ou deixaria o processo errado, em frases que a IA entende."""
    problemas = []
    nome = str(plano.get("nome") or "").strip()
    if not nome or INVALIDOS.search(nome):
        problemas.append('"nome" vazio ou com caracteres proibidos (< > : " / \\ | ? *)')
    modelos = [m for m in _lista(plano.get("modelos")) if isinstance(m, dict)]
    if not modelos:
        return problemas + ['"modelos" está vazio']
    if sum(bool(m.get("principal")) for m in modelos) != 1:
        problemas.append('exatamente um modelo precisa ter "principal": true')
    words = {e.nome: e for e in exemplos if e.word}
    arquivos = set()
    for m in modelos:
        if m.get("base") not in words:
            problemas.append(f'o modelo "{m.get("arquivo")}" tem base "{m.get("base")}", que não é um dos Words '
                             f'recebidos ({", ".join(words)})')
        arquivo = str(m.get("arquivo") or "")
        if not arquivo.lower().endswith(".docx") or INVALIDOS.search(arquivo) or arquivo.lower() in arquivos:
            problemas.append(f'o "arquivo" do modelo ("{arquivo}") precisa terminar em .docx, ser único e não '
                             'ter caracteres proibidos')
        arquivos.add(arquivo.lower())
        if not m.get("principal") and not _lista(m.get("se")):
            problemas.append(f'o modelo "{arquivo}" não é o principal e está sem condições em "se"')
    colunas = {atos.normalizar(c.get("nome")) for c in _lista(plano.get("colunas")) if isinstance(c, dict)}
    if not colunas:
        problemas.append('"colunas" está vazio')
    proc = plano.get("processo") if isinstance(plano.get("processo"), dict) else {}
    ligacoes = proc.get("ligacoes") if isinstance(proc.get("ligacoes"), dict) else {}
    conhecidos = colunas | {atos.normalizar(c) for c in ligacoes} | set(atos.AUTOMATICOS)

    def citar(onde: str, nomes) -> None:
        faltam = sorted({n for n in nomes if n and atos.normalizar(n) not in conhecidos})
        if faltam:
            problemas.append(f"{onde} cita {', '.join(faltam)}, que não está em \"colunas\" nem em \"ligacoes\"")

    citadas = [proc.get("coluna_nome"), proc.get("coluna_numero"), (proc.get("juntar") or {}).get("coluna")]
    citadas += _lista(proc.get("obrigatorios"))
    citar('"processo"', [c for c in citadas if isinstance(c, str)])
    for chave in ("titulo_edocs", "titulo_diario", "titulo_no_ato"):
        citar(f'"{chave}"', _citadas_nos_titulos(proc.get(chave, "")))
    for campo, regra in ligacoes.items():
        if not isinstance(regra, dict) or regra.get("tipo") not in processos.TIPOS:
            problemas.append(f'a regra de "{campo}" tem tipo desconhecido ({regra.get("tipo") if isinstance(regra, dict) else regra})')
            continue
        if regra.get("coluna"):
            citar(f'a regra de "{campo}"', [regra["coluna"]] + ([regra["ate"]] if regra.get("ate") else []))
        textos = [regra.get("senao", ""), regra.get("valor", "")]
        for caso in _lista(regra.get("casos")) + _lista(regra.get("opcoes")):
            textos.append(caso.get("texto", "") if isinstance(caso, dict) else "")
            citar(f'a condição de "{campo}"', [c.get("campo") for c in _lista(caso.get("se")) if isinstance(c, dict)]
                  if isinstance(caso, dict) else [])
        citar(f'o texto da regra de "{campo}"', [n for t in textos for n in atos.CHAVES.findall(str(t or ""))])
    for m in modelos:
        citar(f'o modelo "{m.get("arquivo")}"',
              [n for t in _lista(m.get("trocas")) if isinstance(t, dict) for n in atos.CHAVES.findall(str(t.get("para", "")))
               if atos.normalizar(n) not in (atos.REPETIR, atos.FIM_REPETIR)])
        citar(f'as condições do modelo "{m.get("arquivo")}"',
              [c.get("campo") for c in _lista(m.get("se")) if isinstance(c, dict)])
    if not _lista(plano.get("exemplos")):
        problemas.append('"exemplos" está vazio: preciso dos dados de cada exemplo para conferir')
    return problemas


# --- Montagem do Word ---------------------------------------------------------------------
def _padrao(trecho: str) -> re.Pattern:
    """O trecho como a IA escreveu, aceitando as pequenas diferenças de digitação: espaços
    (inclusive o rígido do Word), aspas curvas/retas e os vários traços."""
    partes, espaco = [], False
    for ch in trecho.strip():
        if ch.isspace():
            if not espaco:
                partes.append(r"\s+")
            espaco = True
            continue
        espaco = False
        if ch in _ASPAS:
            partes.append(f"[{re.escape(_ASPAS)}]")
        elif ch in _APOSTROFOS:
            partes.append(f"[{re.escape(_APOSTROFOS)}]")
        elif ch in _TRACOS:
            partes.append(f"[{re.escape(_TRACOS)}]")
        else:
            partes.append(re.escape(ch))
    return re.compile("".join(partes))


def _texto(el) -> str:
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def _texto_linha(tr) -> str:
    return " | ".join(" ".join(" ".join(_texto(p).split()) for p in tc.iter(qn("w:p"))).strip()
                      for tc in tr.findall(qn("w:tc")))


def _achar(body, trecho: str):
    """O parágrafo do corpo (ou a linha de tabela, se o trecho tem " | ") com esse texto."""
    padrao = _padrao(trecho)
    if " | " in trecho:
        candidatos = [(tr, _texto_linha(tr)) for tbl in body.iter(qn("w:tbl")) for tr in tbl.findall(qn("w:tr"))]
    else:
        candidatos = [(el, _texto(el)) for el in body if el.tag == qn("w:p")]
    iguais = [el for el, texto in candidatos if padrao.fullmatch(texto.strip())]
    contem = [el for el, texto in candidatos if padrao.search(texto)]
    return (iguais or contem or [None])[0]


def _vazio(el) -> bool:
    return el is not None and el.tag == qn("w:p") and not _texto(el).strip()


def _marcar_paragrafo(primeiro, ultimo) -> None:
    for texto, vizinho, antes in (("{{REPETIR}}", primeiro, True), ("{{FIM REPETIR}}", ultimo, False)):
        p = primeiro.makeelement(qn("w:p"), {})
        ppr = primeiro.find(qn("w:pPr"))
        if ppr is not None:
            p.append(copy.deepcopy(ppr))
        r = p.makeelement(qn("w:r"), {})
        t = r.makeelement(qn("w:t"), {})
        t.text = texto
        r.append(t)
        p.append(r)
        if antes:
            vizinho.addprevious(p)
        else:
            vizinho.addnext(p)


def _aplicar_trocas(body, trocas: list[dict], arquivo: str) -> list[str]:
    problemas = []
    paragrafos = list(body.iter(qn("w:p")))
    for troca in trocas:
        de, para = str(troca.get("de") or ""), str(troca.get("para") or "")
        if not de.strip():
            continue
        if "\n" in de.strip():
            problemas.append(f'{arquivo}: a troca "{de[:80]}" passa de um parágrafo para outro; divida em '
                             'trocas de um parágrafo só')
            continue
        padrao, achou = _padrao(de), False
        for p in paragrafos:
            texto = _texto(p)
            if not padrao.search(texto):
                continue
            campos = [m.span() for m in atos.CHAVES.finditer(texto)]  # não mexe num {{CAMPO}} já posto
            trechos = [(m.start(), m.end(), para) for m in padrao.finditer(texto)
                       if not any(a < m.end() and m.start() < b for a, b in campos)]
            if trechos:
                _trocar(p, trechos)
                achou = True
        if not achou:
            problemas.append(f'{arquivo}: não achei o trecho "{de[:120]}" no exemplo base (copie exatamente '
                             'como está no texto, sem quebrar linha)')
    return problemas


def montar_word(modelo: dict, base: Path, destino: Path) -> list[str]:
    """Grava o Word do modelo a partir do exemplo base. Devolve os problemas (trecho não
    achado...): o que deu para fazer, fica feito."""
    doc = Document(base)
    body = doc.element.body
    arquivo = destino.name
    problemas, repetir = [], modelo.get("repetir") if isinstance(modelo.get("repetir"), dict) else None
    alvo, ultimo = None, None
    if repetir and str(repetir.get("primeiro") or "").strip():
        alvo = _achar(body, repetir["primeiro"])
        if alvo is None:
            problemas.append(f'{arquivo}: não achei o trecho da 1ª pessoa ("{repetir["primeiro"][:120]}")')
        separado = False
        for trecho in _lista(repetir.get("apagar")):
            el = _achar(body, str(trecho)) if str(trecho).strip() else None
            if el is None or el is alvo:
                problemas.append(f'{arquivo}: não achei o trecho de outra pessoa para apagar ("{str(trecho)[:120]}")')
                continue
            antes = el.getprevious()
            if el.tag == qn("w:p") and _vazio(antes) and antes is not alvo:
                antes.getparent().remove(antes)  # a linha em branco que separava as pessoas
                separado = True
            el.getparent().remove(el)
        if alvo is not None and alvo.tag == qn("w:p"):
            depois = alvo.getnext()
            ultimo = depois if separado and _vazio(depois) else alvo
    problemas += _aplicar_trocas(body, _lista(modelo.get("trocas")), arquivo)
    if alvo is not None:
        if alvo.tag == qn("w:tr"):
            paragrafos = list(alvo.iter(qn("w:p")))
            atos._escrever_na_ponta(paragrafos[0], "{{REPETIR}}", True)
            atos._escrever_na_ponta(paragrafos[-1], "{{FIM REPETIR}}", False)
        else:
            _marcar_paragrafo(alvo, ultimo)
    destino.parent.mkdir(parents=True, exist_ok=True)
    doc.save(destino)
    return problemas


# --- Montagem do processo -----------------------------------------------------------------
def _colunas(plano: dict) -> list[dict]:
    saida, vistas = [], set()
    for c in _lista(plano.get("colunas")):
        nome = " ".join(str(c.get("nome") or "").split()) if isinstance(c, dict) else ""
        if not nome or atos.normalizar(nome) in vistas or atos.normalizar(nome) in atos.AUTOMATICOS:
            continue
        vistas.add(atos.normalizar(nome))
        tipo = c.get("tipo") if c.get("tipo") in ("texto", "numero", "data", "lista") else "texto"
        opcoes = [str(o) for o in _lista(c.get("opcoes")) if str(o).strip()]
        saida.append({"nome": nome, "tipo": "lista" if opcoes else ("texto" if tipo == "lista" else tipo),
                      "opcoes": opcoes})
    return saida


def montar(plano: dict, exemplos: list[Exemplo], pasta: Path) -> tuple[dict, Path, list[str]]:
    """Grava em pasta os Words e a planilha do plano. Devolve (config, planilha, problemas)."""
    words = {e.nome: e.word for e in exemplos if e.word}
    for velho in pasta.glob("*.docx"):  # sobra de uma rodada anterior
        velho.unlink()
    problemas, outros, principal = [], [], ""
    modelos = [m for m in _lista(plano.get("modelos")) if isinstance(m, dict)]
    for m in modelos:
        destino = pasta / str(m["arquivo"])
        problemas += montar_word(m, words[m["base"]], destino)
        if m.get("principal"):
            principal = destino.name
        else:
            outros.append({"modelo": destino.name, "se": _lista(m.get("se"))})
    nome = str(plano["nome"]).strip()
    planilha = planilha_nova.criar(pasta / f"{nome} - planilha.xlsx", _colunas(plano))
    proc = {k: v for k, v in (plano.get("processo") or {}).items() if k in CHAVES_DO_PLANO}
    for chave in ("edocs", "diario"):  # o que a IA não disse (setor...) fica como o padrão
        if isinstance(proc.get(chave), dict):
            proc[chave] = dict(processos.PADRAO[chave], **proc[chave])
    juntar = (proc.get("juntar") or {}).get("coluna") if isinstance(proc.get("juntar"), dict) else ""
    config = dict(proc, modelo=principal, outros_modelos=outros, planilha=planilha.name,
                  aba=planilha_nova.ABA, juntar={"coluna": juntar or ""},
                  _tipo="varias" if juntar else "uma")
    return config, planilha, problemas


def pessoas_do_exemplo(plano: dict, config: dict, exemplo: dict) -> list[dict]:
    """As linhas da planilha que a IA leu de um exemplo (chaves como o Edodocs lê a planilha).
    As pessoas do mesmo exemplo ficam no mesmo ato."""
    juntar = atos.normalizar((config.get("juntar") or {}).get("coluna", ""))
    linhas = []
    for i, pessoa in enumerate(p for p in _lista(exemplo.get("pessoas")) if isinstance(p, dict)):
        linha = {atos.normalizar(k): "" if v is None else str(v).strip() for k, v in pessoa.items()}
        linha["_LINHA"] = i + 2
        linhas.append(linha)
    if juntar and linhas:
        for linha in linhas:
            linha[juntar] = linhas[0].get(juntar) or "1"
    return linhas


def _linhas(texto: str) -> list[str]:
    return [" ".join(linha.split()) for linha in texto.splitlines() if linha.strip()]


def comparar(arquivo: str, original: str, montado: str) -> Comparacao:
    a, b = _linhas(original), _linhas(montado)
    parecido = round(100 * difflib.SequenceMatcher(None, "\n".join(a), "\n".join(b)).ratio())
    diferencas = []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if op != "equal":
            diferencas.append(("\n".join(a[i1:i2]), "\n".join(b[j1:j2])))
    return Comparacao(arquivo, not diferencas, parecido, diferencas)


def refazer_exemplos(plano: dict, exemplos: list[Exemplo], config: dict, pasta: Path) -> list[Comparacao]:
    """Monta cada exemplo com os dados que a IA leu dele e compara com o original."""
    proc = processos.Processo(pasta, config)
    por_nome = {e.nome: e for e in exemplos}
    saida = []
    for exemplo in _lista(plano.get("exemplos")):
        if not isinstance(exemplo, dict) or exemplo.get("arquivo") not in por_nome:
            continue
        original = por_nome[exemplo["arquivo"]]
        linhas = pessoas_do_exemplo(plano, config, exemplo)
        if not linhas:
            saida.append(Comparacao(original.nome, False, 0, [("", "o plano não tem as pessoas deste exemplo")]))
            continue
        try:
            texto = proc.textos(proc.agrupar(linhas)[:1])[0]
        except Exception as e:
            saida.append(Comparacao(original.nome, False, 0, [("", f"erro ao montar: {type(e).__name__} {e}")]))
            continue
        saida.append(comparar(original.nome, original.texto, texto))
    return saida


def problemas_das_comparacoes(comparacoes: list[Comparacao], exemplos: list[Exemplo]) -> list[str]:
    problemas = []
    for c in comparacoes:
        if c.igual:
            continue
        for esperado, saiu in c.diferencas[:6]:
            problemas.append(f'refazendo "{c.arquivo}": no exemplo está "{esperado[:300]}", mas saiu "{saiu[:300]}"')
    conferidos = {c.arquivo for c in comparacoes}
    faltam = [e.nome for e in exemplos if e.nome not in conferidos]
    if faltam:
        problemas.append('"exemplos" não tem os dados de: ' + ", ".join(faltam))
    return problemas


def montar_com_ia(exemplos: list[Exemplo], pasta: Path, descricao: str = "", aviso=lambda texto: None,
                  pedir=None) -> Resultado:
    """Pede o plano, monta e confere; o que não bate volta para a IA corrigir. aviso(texto)
    conta à tela em que pé está. pedir(sistema, pedido) troca a IA (testes)."""
    pedir = pedir or ia.pedir_json
    guia = ia.guia_processos()
    original = pedido(exemplos, descricao)
    aviso("A IA está lendo os seus atos e montando o processo...")
    plano = pedir(guia, original)
    return _conferir_e_corrigir(plano, exemplos, pasta, original, guia, pedir, aviso)


def ajustar_com_respostas(resultado: Resultado, exemplos: list[Exemplo], respostas: list[dict], pasta: Path,
                          descricao: str = "", aviso=lambda texto: None, pedir=None) -> Resultado:
    """Manda as respostas da pessoa para a IA ajustar o plano; monta e confere de novo em pasta."""
    pedir = pedir or ia.pedir_json
    guia = ia.guia_processos()
    original = pedido(exemplos, descricao)
    aviso("A IA está ajustando o processo com as suas respostas...")
    plano = pedir(guia, pedido_de_respostas(original, resultado.plano, respostas))
    if isinstance(plano, dict):
        plano["perguntas"] = []  # já respondidas: não pergunta de novo
    return _conferir_e_corrigir(plano, exemplos, pasta, original, guia, pedir, aviso)


def _conferir_e_corrigir(plano, exemplos, pasta, original, guia, pedir, aviso) -> Resultado:
    pasta.mkdir(parents=True, exist_ok=True)
    rodada, problemas, comparacoes, config, planilha = 0, [], [], {}, None
    while True:
        problemas = conferir_plano(plano, exemplos)
        if not problemas:
            aviso("Montando o Word e a planilha e conferindo cada exemplo...")
            config, planilha, problemas = montar(plano, exemplos, pasta)
            comparacoes = refazer_exemplos(plano, exemplos, config, pasta)
            problemas += problemas_das_comparacoes(comparacoes, exemplos)
        if not problemas or rodada >= RODADAS:
            break
        rodada += 1
        aviso(f"A conferência achou {len(problemas)} ponto(s) para acertar. A IA está corrigindo "
              f"({rodada}ª de {RODADAS} correções)...")
        plano = pedir(guia, pedido_de_correcao(original, plano, problemas))
    if planilha is None:  # o plano nunca ficou bom o bastante para montar
        raise ia.ErroIA("a IA não conseguiu montar um plano válido: " + "; ".join(problemas[:4]))
    return Resultado(plano, config, pasta, planilha, problemas, comparacoes, rodada)


def rascunho_do_resultado(resultado: Resultado, nome: str, respostas: list[dict] | None = None) -> dict:
    """O resultado no formato do rascunho do assistente (assistente.ler_rascunho), para abrir
    direto em "Como vai ficar". Os arquivos ficam em resultado.pasta ("_pasta")."""
    config = dict(resultado.config)
    arquivos = {"modelo": config["modelo"], "planilha": resultado.planilha.name}
    for i, outro in enumerate(config.get("outros_modelos") or []):
        arquivos[f"outro{i}"] = outro["modelo"]
    plano = resultado.plano
    exemplos = [e for e in _lista(plano.get("exemplos")) if isinstance(e, dict)]
    juntar = atos.normalizar((config.get("juntar") or {}).get("coluna", ""))
    pessoas, usados = [], set()
    for n, exemplo in enumerate(exemplos, 1):  # todos os exemplos, para ver cada um na prévia
        linhas = pessoas_do_exemplo(plano, config, exemplo)
        if juntar and linhas:  # cada exemplo é um ato (dois "S/N" não podem virar um só)
            valor = linhas[0].get(juntar) or str(n)
            valor = valor if valor not in usados else f"{valor} ({n})"
            usados.add(valor)
            for linha in linhas:
                linha[juntar] = valor
        pessoas += linhas
    for i, linha in enumerate(pessoas):
        linha["_LINHA"] = i + 2
        linha["_EXEMPLO"] = True
    return {"nome": nome, "config": config, "modo_planilha": "criar", "passo": "previa", "marcas": [],
            "planilha": "", "arquivos": arquivos, "_pasta": str(resultado.pasta),
            "exemplo": pessoas[0] if pessoas else None, "pessoas_exemplo": pessoas,
            "ia": {"explicacao": str(plano.get("explicacao") or ""),
                   "avisos": [str(d) for d in _lista(plano.get("avisos")) + _lista(plano.get("duvidas"))],
                   "respostas": [(r["pergunta"], r["resposta"]) for r in respostas or []],
                   "problemas": resultado.problemas,
                   "conferidos": [(c.arquivo, c.igual, c.parecido) for c in resultado.comparacoes]}}
