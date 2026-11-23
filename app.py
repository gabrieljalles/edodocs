"""Edodocs: a janela do programa (NiceGUI). É a única forma de usar o programa.

    python app.py          (ou o Edodocs.exe gerado pelo build.py)

Os fluxos ficam no main.py (lançar no Edocs, lançar no Diário, conferir a publicação).
Eles rodam numa thread separada para a janela não travar; tudo o que eles mostram ou
perguntam passa pelo ui.py, que aqui usa o motor Janela: mensagens vão para a página
Execução e perguntas viram caixas de diálogo.
"""
import asyncio
import json
import os
import sys
import queue
import re
import threading
import time
from concurrent.futures import Future
from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote

# No executável sem console (build.py) não existe saída de texto: o que o programa e o
# servidor escreveriam vai para configs/edodocs.log (útil para entender um erro). Se a
# pasta do programa não aceitar gravação, o log vai para %LOCALAPPDATA%\Edodocs.
LOG = None
if getattr(sys, "frozen", False) and (sys.stdout is None or sys.stderr is None):
    for _pasta in (Path(sys.executable).parent / "configs",
                   Path(os.environ.get("LOCALAPPDATA", Path.home())) / "Edodocs"):
        try:
            _pasta.mkdir(parents=True, exist_ok=True)
            sys.stdout = sys.stderr = open(_pasta / "edodocs.log", "a", encoding="utf-8", buffering=1)
            LOG = _pasta / "edodocs.log"
            break
        except OSError:
            continue
    else:
        sys.stdout = sys.stderr = open(os.devnull, "w")
    print(f"\n--- {datetime.now():%d/%m/%Y %H:%M:%S} Edodocs abrindo ---")
    # Arquivos que vieram num .zip baixado ou de e-mail têm a marca "veio da internet", e o
    # Windows pode se recusar a carregar as DLLs da janela por causa dela: tira a marca
    for _arq in Path(getattr(sys, "_MEIPASS", ".")).rglob("*"):
        if _arq.suffix.lower() in (".dll", ".pyd", ".exe"):
            try:
                os.remove(f"{_arq}:Zone.Identifier")
            except OSError:
                pass

# Edodocs.exe --site edocs|diario: só a janela do site já logado (ver janela_site.py)
if "--site" in sys.argv and __name__ == "__main__":
    import janela_site
    janela_site.rodar(sys.argv[sys.argv.index("--site") + 1])
    os._exit(0)

from nicegui import app as napp
from nicegui import background_tasks
from nicegui import run
from nicegui import ui as gui

from abrir import abrir, mostrar_na_pasta
import atos
import consultas
import consultas_ui
import de_para
import credenciais
import folha
import main
import planilha_edicao as edicao
import planilha_status
import processos
import simulacao_ui
import ui
from assistente import Assistente
from escolha_pessoas import PainelPessoas, montar_escolha_pessoas

TITULO = "Edodocs"
# Interruptores (ligado/desligado): verde quando ligado
gui.switch.default_props("color=positive")
# Ícones do dodô (no executável do PyInstaller, os arquivos embutidos ficam em sys._MEIPASS)
_ASSETS = Path(getattr(sys, "_MEIPASS", Path(__file__).parent)) / "assets"


def _arquivo(*nomes: str) -> Path:
    """O primeiro arquivo que existir em assets/ (os outros são reserva)."""
    return next((_ASSETS / n for n in nomes if (_ASSETS / n).exists()), _ASSETS / nomes[-1])


ICONE = _arquivo("icone_light.png", "icone.png")                        # modo claro
ICONE_ESCURO = _arquivo("icone_dark.png", "icone_light.png", "icone.png")  # modo escuro
# A barra de título e a barra de tarefas do Windows só aceitam .ico
ICONE_JANELA = _arquivo("icone_light.ico", "icone.ico")
# Manual (assets/manual/manual.html) e o áudio dele, servidos para a página Manual
MANUAL = _ASSETS / "manual"
EXTENSOES_AUDIO = (".mp3", ".m4a", ".wav", ".ogg")
napp.add_static_files("/manual", MANUAL)
# Altura do que fica acima e abaixo do conteúdo das páginas: cabeçalho (72), título da
# página (~60) e margens do painel e da página (~84). Usada para a Execução caber sem rolar.
ALTURA_FIXA = 216
SITES = {"edocs": "e-docs", "diario": "Diário"}
PAUSA_REQUISICAO_PADRAO = 0.25
PAUSA_ENTRE_ATOS_PADRAO = 0.5


def _versao_no_canto() -> None:
    """A versão do programa, discreta, no canto inferior direito da tela."""
    import versao
    gui.label(f"versão {versao.atual()}").classes("text-caption")         .style("position: fixed; right: 14px; bottom: 8px; color: #b0b6bf; z-index: 10; pointer-events: none")


def ia_conectada() -> bool:
    """Tem uma IA em uso com a chave salva (ver ia.conectada)."""
    import ia

    return ia.conectada()


# Bolinhas de status: cor → (classe CSS, significado na legenda)
VERDE, AMARELO, VERMELHO, BRANCO = "verde", "amarelo", "vermelho", "branco"
# Páginas do menu lateral: (chave, nome, ícone)
PAGINAS = [("diario", "Diário Oficial", "newspaper"), ("atos", "Atos", "table_view"),
           ("ato_manual", "Ato manual", "edit_note"),
           ("recebidos", "Caixa de Saída", "inbox"),
           ("execucao", "Execução", "play_circle"), ("processos", "Processos", "folder_special"),
           ("config", "Configurações", "settings"), ("manual", "Manual", "menu_book")]
# Páginas em que o seletor de Processo aparece no topo
PAGINAS_COM_PROCESSO = {"atos", "processos", "execucao", "config"}
LEGENDA = {VERDE: "Concluído", AMARELO: "Aguardando / próximo passo",
           VERMELHO: "Erro", BRANCO: "Ainda não feito"}

CSS = """
.barra-atos .q-btn__content { flex-wrap: nowrap; white-space: nowrap; }
.barra-atos .q-btn-toggle .q-btn { padding: 4px 10px; }
.resultado-busca:hover { background: rgba(36, 89, 168, .10); }
/* Campo preenchido pela tabela de consulta (Nova linha): verde que dá para ler de dia e de noite */
.veio-da-tabela .q-field__control { background: #e8f5e9 !important; box-shadow: inset 4px 0 0 #43a047; }
body.body--dark .veio-da-tabela .q-field__control { background: rgba(67, 160, 71, .22) !important; }
body.body--dark .veio-da-tabela .q-field__native, body.body--dark .veio-da-tabela .q-field__input { color: #fff !important; }
body.body--dark .veio-da-tabela .q-field__label { color: #a5d6a7 !important; }
body.body--dark .chip-uma-pessoa { color: #fff !important; border-color: #8fb4ff !important; }
body { background: #f3f5f8; }
.cabecalho { background: linear-gradient(90deg, #ffffff 0%, #ffffff 50%, #2459a8 100%);
             box-shadow: 0 4px 12px rgba(15, 35, 75, .28); padding: 4px 16px; }
.cab-texto { color: #1b3d78; }
/* Barra de andamento logo abaixo do cabeçalho (vistar tudo, desfazer várias linhas...) */
.barra-andamento { position: absolute; left: 0; right: 0; top: 100%; background: #e8f0fc;
                   color: #1b3d78; border-bottom: 1px solid #c9d8f0; box-shadow: 0 3px 8px rgba(15,35,75,.12); }
body.body--dark .barra-andamento { background: #17284a; color: #dbe6fb; border-bottom-color: #24406f; }
.usuario-login { color: #5f6672; }
.menu-lateral { background: #ffffff; }
.fundo-assistente { background: #f3f5f8; }
.botao-tema { border: 1px solid rgba(255,255,255,.75); border-radius: 6px; color: #ffffff;
              width: 40px; height: 40px; }
body.body--dark .icone-claro, body:not(.body--dark) .icone-escuro { display: none !important; }
/* Lua (modo claro): azul forte e escuro; o sol (modo escuro) continua claro */
body:not(.body--dark) .botao-tema .q-icon { color: #0b2350; }
body.body--dark .botao-tema .q-icon { color: #ffffff; }
.t-forte { color: #1f2937; }
.t-suave { color: #5b6472; }
.t-fraco { color: #9aa3af; }
.fundo-cartao { background: #ffffff; }
.fundo-previa { background: #f8f9fb; }
.chip-campo { background: #dbeafe; color: #1e3a8a; }
/* Assistente: rolagem só na vertical (textos longos, como o nome do setor, encolhem com "…") */
.fundo-assistente .q-scrollarea__content { width: 100% !important; max-width: 100%; }
/* Resumo do Diário de hoje: discreto */
.resumo-diario .q-item { min-height: 32px; padding: 2px 8px; }
.resumo-diario .q-item__label { font-size: 13px; }
.resumo-diario .q-item__section--avatar .q-icon { color: #2459a8; }
body.body--dark .resumo-diario .q-item__section--avatar .q-icon { color: #8fb4ee; }
.resumo-diario .q-focus-helper { display: none; }
.resumo-diario .q-item:hover { background: rgba(36, 89, 168, .06); }
.contagem-diario { cursor: default; opacity: .9; }
.contagem-diario:hover { opacity: 1; }
.divisa-diario { width: 1px; height: 14px; background: rgba(0, 0, 0, .14); }
body.body--dark .divisa-diario { background: rgba(255, 255, 255, .18); }
/* Telas dos sites reproduzidas no assistente */
.campo-site .q-field__control { background: rgba(127,127,127,.08); }
.campo-editavel .q-field__control:before { border-color: var(--q-primary) !important; border-width: 1.5px; }
.barra-editor { background: #334155; color: #e2e8f0; border-radius: 6px 6px 0 0; }
.caixa-editor { height: 360px; overflow: auto; border: 1px solid #cbd5e1; border-top: none; background: #fff; }
/* Janelas de ajuda (botões ⓘ do assistente) */
.ajuda-texto { font-size: 15px; line-height: 1.55; }
.ajuda-texto table { border-collapse: collapse; margin: 6px 0 10px; }
.ajuda-texto th, .ajuda-texto td { border: 1px solid #cbd5e1; padding: 5px 9px; text-align: left; }
.ajuda-texto th { background: #eff6ff; }
.ajuda-texto blockquote { margin: 8px 0; padding: 8px 12px; border-left: 4px solid #2459a8;
                          background: #eff6ff; border-radius: 6px; }
.ajuda-texto code { background: #eef0f3; border-radius: 4px; padding: 0 4px; }
.titulo-icone { color: #2459a8; }
.bolinha { display:inline-block; width:14px; height:14px; border-radius:50%;
           border:1px solid rgba(15,23,42,.18); vertical-align:middle; }
.b-verde { background:#22c55e; } .b-amarelo { background:#facc15; }
.b-vermelho { background:#ef4444; } .b-branco { background:#ffffff; border-color:#94a3b8; }
.cartao { border-radius:14px; box-shadow:0 1px 3px rgba(15,23,42,.08); }
.linha-caixa { border-bottom: 1px solid rgba(100,116,139,.18); min-height: 46px; }
.linha-caixa:last-child { border-bottom: none; }
.linha-caixa:hover { background: rgba(59,130,246,.06); }
.cartao-reto { border-radius:6px; }
.tabela-atos thead tr th { position:sticky; top:0; z-index:1; background:#f8f9fb;
                           font-weight:600; }
.tabela-atos td { white-space:nowrap; }
.log-linha { font-size:13.5px; line-height:1.35; }
.menu-item { border-radius:10px; margin:2px 10px; }
.menu-ativo { background:#dceafc; color:#2459a8; font-weight:600; }
/* Status de login no rodape do menu: borda cinza discreta, alinhada ao traco de cima */
.login-item { border:1px solid #d9dee5; border-radius:4px; margin:6px 0 0 0; }
.bolinha-p { width:9px; height:9px; flex-shrink:0; }
.usuario-pequeno { font-size:12px; }
.icone-pdf { color:#1b3d78; }
.sel-processo .q-field__label { color:#ffffff !important; font-weight:600; }
.previa { white-space:pre-wrap; font-family:Arial, sans-serif; font-size:14px;
          line-height:1.5; text-align:justify; }

/* --- Modo escuro: mesmas classes, cores escuras e texto claro (contraste para leitura) --- */
body.body--dark { background: #0b1220; color: #e5e9f0; }
body.body--dark .cabecalho { background: linear-gradient(90deg, #111c31 0%, #111c31 50%, #24509a 100%);
                             box-shadow: 0 4px 12px rgba(0, 0, 0, .55); }
body.body--dark .cab-texto { color: #e6eeff; }
body.body--dark .usuario-login { color: #a9b4c6; }
body.body--dark .menu-lateral { background: #101a2c; }
body.body--dark .fundo-assistente { background: #0b1220 !important; }
body.body--dark .q-card, body.body--dark .q-table__container, body.body--dark .q-menu,
body.body--dark .q-expansion-item { background: #152238; color: #e5e9f0; }
body.body--dark .cartao { box-shadow: 0 1px 3px rgba(0, 0, 0, .45); }
body.body--dark .tabela-atos thead tr th { background: #1c2b45; color: #e5e9f0; }
body.body--dark .menu-ativo { background: #1f3a66; color: #bcd5ff; }
body.body--dark .login-item { border-color: #2c3b55; }
body.body--dark .icone-pdf { color: #e5e9f0; }
body.body--dark .t-forte { color: #e2e8f0; }
body.body--dark .t-suave { color: #aab4c3; }
body.body--dark .t-fraco { color: #7f8a9c; }
body.body--dark .fundo-cartao { background: #152238; }
body.body--dark .fundo-previa { background: #0f1a2d; }
body.body--dark .chip-campo { background: #1f3a66; color: #dbe7ff; }
body.body--dark .ajuda-texto th, body.body--dark .ajuda-texto td { border-color: #33435e; }
body.body--dark .ajuda-texto th, body.body--dark .ajuda-texto blockquote { background: #1c2a44; }
body.body--dark .ajuda-texto code { background: #24324c; }
body.body--dark .q-tab-panels, body.body--dark .q-tab-panel { background: transparent; }
body.body--dark .titulo-icone { color: #8fb4ee; }
/* Azul de textos, ícones e contornos (não dos botões preenchidos) mais claro no escuro:
   o Quasar usa var(--q-primary) com !important, então troca-se a variável nesses elementos */
body.body--dark .q-btn--flat, body.body--dark .q-btn--outline, body.body--dark .q-icon,
body.body--dark .q-spinner, body.body--dark .q-tab, body.body--dark .q-stepper__tab,
body.body--dark .q-badge--outline, body.body--dark .q-toggle, body.body--dark .q-field,
body.body--dark .q-checkbox { --q-primary: #8fb4ee; }
body.body--dark .etapa-icone { filter: brightness(1.9); }
body.body--dark .b-branco { background: #e5e9f0; }
"""


def _sem_marcacao(texto: str) -> str:
    """Tira a marcação do rich ("[bold]...[/]") das mensagens dos fluxos."""
    return re.sub(r"\[/?(?:bold|dim|green|red|yellow|cyan|bold [a-z]+)?\]", "", str(texto))


def _logo(classes: str) -> None:
    """O dodô: um para o modo claro e outro para o escuro (o CSS mostra o do modo atual)."""
    if ICONE_ESCURO == ICONE:
        gui.image(ICONE).classes(classes)
        return
    gui.image(ICONE).classes(f"{classes} icone-claro")
    gui.image(ICONE_ESCURO).classes(f"{classes} icone-escuro")


def _bolinha(cor: str, dica: str = "", pequena: bool = False) -> None:
    with gui.element("span").classes(f"bolinha b-{cor}" + (" bolinha-p" if pequena else "")):
        if dica:
            gui.tooltip(dica)


# --- Motor do ui.py para a janela ---------------------------------------------

class Janela:
    """Recebe as chamadas do ui.py (vindas da thread do fluxo) e as passa para a tela por
    uma fila. Nas perguntas, a thread do fluxo espera a resposta da caixa de diálogo."""

    def __init__(self):
        self.fila: queue.Queue = queue.Queue()
        # Pedido de cancelamento: vale no próximo ponto seguro (início do próximo ato ou
        # próxima pergunta), nunca no meio de um envio, para o status não ficar sem gravar.
        self.cancelar = False

    def _parar_se_cancelado(self) -> None:
        if self.cancelar:
            self.cancelar = False  # uma vez só: o resto (gravar status) segue normal
            raise KeyboardInterrupt

    def _avisar(self, evento: str, **dados) -> None:
        self.fila.put((evento, dados))

    def _perguntar(self, evento: str, **dados):
        self._parar_se_cancelado()
        futuro: Future = Future()
        self.fila.put((evento, dados | {"futuro": futuro}))
        return futuro.result()

    def mensagem(self, tipo: str, texto: str) -> None:
        self._avisar("mensagem", tipo=tipo, texto=_sem_marcacao(texto))

    @contextmanager
    def processando(self, texto: str):
        self._avisar("ocupado", texto=_sem_marcacao(texto))
        try:
            yield
        finally:
            self._avisar("ocupado", texto="")

    def progresso(self, n: int, total: int, descricao: str) -> None:
        self._parar_se_cancelado()  # entre um ato e outro: ponto seguro para parar
        self._avisar("progresso", n=n, total=total, descricao=_sem_marcacao(descricao))

    def escolher(self, pergunta, opcoes, multiplo=False, vazio=None, marcados=None) -> list[int]:
        return self._perguntar("escolher", pergunta=pergunta, opcoes=opcoes,
                               multiplo=multiplo, vazio=vazio, marcados=marcados or [])

    def escolher_pessoas(self, pergunta, pesquisar, recentes, marcados, vazio) -> list[dict]:
        return self._perguntar("escolher_pessoas", pergunta=pergunta, pesquisar=pesquisar,
                               recentes=recentes, marcados=marcados, vazio=vazio)

    def perguntar(self, pergunta: str, senha: bool = False) -> str:
        return self._perguntar("perguntar", pergunta=pergunta, senha=senha)

    def login(self, titulo: str, explicacao: str) -> tuple[str, str]:
        return self._perguntar("login", titulo=titulo, explicacao=explicacao)

    def confirmar(self, pergunta: str, sim: str = "Sim", nao: str = "Não") -> bool:
        return self._perguntar("confirmar", pergunta=pergunta, sim=sim, nao=nao)

    def confirmar_envio(self, ato: str = "", texto: str = "", folha: str = "", word=None,
                        unico: bool = False) -> str:
        return self._perguntar("confirmar_envio", ato=_sem_marcacao(ato), texto=texto, folha=folha, word=word,
                               unico=unico)

    def anexos_dos_atos(self, lista: list[dict]) -> list | str:
        return self._perguntar("anexos_dos_atos", lista=[dict(i, ato=_sem_marcacao(i["ato"])) for i in lista])


# --- Dados da planilha ----------------------------------------------------------

def _planilha_do_processo() -> Path | None:
    """Mesma regra do main.achar_planilha, sem mensagens (para a tabela)."""
    personalizada = main.ler_config().get("planilhas", {}).get(main.processo.name)
    if personalizada:
        return Path(personalizada) if Path(personalizada).is_file() else None
    return main.proc().planilha_padrao()


def _cores(r: dict, chave_edocs: str) -> dict:
    """(cor, texto) de cada etapa: e-docs, Diário e publicação."""
    edocs = r.get(chave_edocs, "")
    diario = r.get("STATUSDIARIO", "")
    pub = r.get("STATUSPUBLICADO", "")

    if edocs == planilha_status.ENVIADO:
        c_edocs = VERDE
    elif planilha_status.eh_erro(edocs):
        c_edocs = VERMELHO
    else:
        c_edocs = BRANCO

    if diario == planilha_status.ENVIADO:
        c_diario = VERDE
    elif planilha_status.eh_erro(diario):
        c_diario = VERMELHO
    elif c_edocs == VERDE:
        c_diario, diario = AMARELO, diario or "Pronto para lançar no Diário"
    else:
        c_diario = BRANCO

    if str(pub).startswith(planilha_status.PUBLICADO):
        c_pub = VERDE
    elif (pub == planilha_status.NAO_ENCONTRADO_FINAL or planilha_status.eh_erro(pub)
          or pub == planilha_status.ERRO_CONFERENCIA):
        c_pub = VERMELHO
    elif str(pub).startswith(planilha_status.NAO_ENCONTRADO):
        c_pub = AMARELO
    elif c_diario == VERDE:
        c_pub, pub = AMARELO, "Aguardando conferência no Diário Oficial"
    else:
        c_pub = BRANCO

    return {"edocs": (c_edocs, edocs or "Ainda não enviado"),
            "diario": (c_diario, diario or "Ainda não enviado"),
            "pub": (c_pub, pub or "Ainda não conferido")}


# A data de publicação não tem coluna própria: vem do status da conferência,
# ex.: "Publicado com sucesso no diário oficial 7450 (01/10/2026)" → 01/10/2026
DATA_PUBLICACAO = "_DATAPUBLICACAO"


def _data_publicacao(status: str) -> str:
    if not str(status).startswith(planilha_status.PUBLICADO):
        return ""
    m = re.search(r"\((\d{2}/\d{2}/\d{4})\)", status)
    return m.group(1) if m else ""


def _data_br(iso: str) -> str:
    """"2026-10-20" → "20/10/2026" ("?" sem data)."""
    try:
        return f"{date.fromisoformat(iso):%d/%m/%Y}"
    except (TypeError, ValueError):
        return "?"


# --- Campos do formulário de linha: escolher em vez de digitar ----------------------------
SIM_NAO = {"SIM", "NAO", "S", "N"}
MAX_ESCOLHAS = 15        # até quantos valores diferentes a coluna vira lista de escolha
MAX_SUGESTOES = 400      # sugestões ao digitar (secretaria, cargo...)


def _opcoes_das_colunas(colunas: list, linhas: dict, proc) -> dict[int, tuple[str, list[str]]]:
    """Para cada coluna digitada, como ela aparece no formulário de linha:
    ("escolha", valores) para colunas com poucos valores (SIM/NÃO, F/M, EFETIVO/COMISSÃO...),
    ("sugestao", valores) para texto que se repete (secretaria, cargo), em que dá para digitar
    e o programa completa. Os valores vêm do que já existe na planilha e das regras do
    processo (ex.: a regra "A PEDIDO é SIM" já põe SIM e NÃO na lista)."""
    das_regras: dict[str, set[str]] = {}
    datas: set[str] = set()  # colunas que o processo usa como data (calendário)

    def anotar(cond: dict) -> None:
        if cond.get("op", "igual") in ("igual", "diferente"):
            for v in str(cond.get("valor", "")).split(";"):
                if v.strip():
                    das_regras.setdefault(atos.normalizar(cond.get("campo", "")), set()).add(v.strip())

    for regra in proc.config.get("ligacoes", {}).values():
        if regra.get("tipo") == "condicao":
            for caso in regra.get("casos") or []:
                for cond in caso.get("se") or []:
                    anotar(cond)
        elif regra.get("tipo") == "genero" and regra.get("coluna"):
            das_regras.setdefault(atos.normalizar(regra["coluna"]), set()).update({"F", "M"})
        elif regra.get("tipo") == "opcoes" and regra.get("coluna"):
            das_regras.setdefault(atos.normalizar(regra["coluna"]), set()).update(
                o["valor"] for o in regra.get("opcoes") or [] if o.get("valor"))
        elif regra.get("tipo") in ("data_extenso", "periodo") and regra.get("coluna"):
            datas.add(atos.normalizar(regra["coluna"]))
            if regra.get("ate"):
                datas.add(atos.normalizar(regra["ate"]))
    for outro in proc.config.get("outros_modelos") or []:
        for cond in outro.get("se") or []:
            anotar(cond)
    for caixa in proc.caixas:  # caixinha da tabela: marcada é SIM
        das_regras.setdefault(atos.normalizar(caixa), set()).add("SIM")

    saida = {}
    for c in colunas:
        if not c.editavel or c.tipo != "texto":
            continue
        vistos: dict[str, str] = {}
        preenchidas = 0
        for valores in linhas.values():
            v = str(valores.get(c.indice, "")).strip()
            if v:
                preenchidas += 1
                vistos.setdefault(atos.normalizar(v), v)
        for v in das_regras.get(atos.normalizar(c.nome), ()):
            vistos.setdefault(atos.normalizar(v), v)
        if atos.normalizar(c.nome) in {atos.normalizar(x) for x in proc.caixas}:
            vistos = {"SIM": "SIM"}  # caixinha: SIM ou vazio
        elif {n for n in vistos if n} & SIM_NAO and set(vistos) <= SIM_NAO | {""}:
            vistos = {"SIM": "SIM", "NAO": "NÃO"}  # sim/não: sempre as duas opções, escritas igual
        valores = sorted(vistos.values(), key=str.lower)
        if atos.normalizar(c.nome) in datas:
            saida[c.indice] = ("data", [])
            continue
        if not valores:
            continue
        if all(atos.ler_data(v) for v in valores):  # data escrita como texto: calendário
            saida[c.indice] = ("data", [])
            continue
        if len(valores) <= MAX_ESCOLHAS and (atos.normalizar(c.nome) in das_regras or preenchidas > len(valores)):
            saida[c.indice] = ("escolha", valores)
        elif preenchidas > len(valores) * 1.25:  # repete bastante: vale sugerir ao digitar
            saida[c.indice] = ("sugestao", valores[:MAX_SUGESTOES])
    return saida


def _campo_da_coluna(c, valor: str, opcoes: tuple[str, list[str]] | None, obrigatorio: bool = False):
    """O campo de uma coluna no formulário de linha: lista de escolha, texto com sugestões,
    data com calendário ou texto/número. Obrigatório: o rótulo leva *."""
    rotulo = c.nome + (" *" if obrigatorio else "")
    if opcoes and opcoes[0] == "escolha":
        lista = list(dict.fromkeys(opcoes[1] + ([valor] if valor else [])))
        campo = gui.select(lista, label=rotulo, value=valor or None, with_input=True,
                           new_value_mode="add-unique", clearable=True) \
            .props("outlined dense stack-label options-dense").classes("w-full") \
            .tooltip("Escolha na lista (ou digite um valor novo)")
        # Valor novo digitado vale mesmo sem apertar Enter (ao sair do campo)
        digitado = {"texto": ""}
        campo.on("input-value", lambda e: digitado.update(texto=str(e.args or "")))
        # Escolheu na lista: o pedaço digitado para filtrar não pode voltar ao sair do campo
        campo.on_value_change(lambda: digitado.update(texto=""))

        def confirmar_digitado() -> None:
            texto = digitado["texto"].strip()
            digitado["texto"] = ""
            if texto and campo.value != texto:
                _por_valor(campo, texto)

        campo.on("blur", confirmar_digitado)
        return campo
    if opcoes and opcoes[0] == "sugestao":
        return gui.input(rotulo, value=valor, autocomplete=opcoes[1]) \
            .props("outlined dense stack-label").classes("w-full") \
            .tooltip("Comece a digitar: o Edodocs completa com o que já existe na planilha (Tab aceita)")
    campo = gui.input(rotulo, value=valor).props("outlined dense stack-label").classes("w-full")
    if c.tipo == "data" or (opcoes and opcoes[0] == "data"):
        # Só os números: as barras entram sozinhas (01102026 → 01/10/2026)
        campo.props('placeholder="dd/mm/aaaa" mask="##/##/####" inputmode=numeric')
        with campo.add_slot("append"):
            with gui.icon("event").classes("cursor-pointer"):
                with gui.menu() as menu:
                    gui.date(mask="DD/MM/YYYY", on_change=lambda: menu.close()).bind_value(campo) \
                        .props("minimal")
    elif c.tipo == "numero":
        campo.props("inputmode=decimal")
    return campo


def _por_valor(campo, valor: str) -> None:
    """Põe o valor no campo (na lista de escolha, acrescenta se ainda não estiver nela)."""
    if isinstance(campo, gui.select):
        if valor and valor not in campo.options:
            campo.set_options(list(campo.options) + [valor])
        campo.value = valor or None
    else:
        campo.value = valor


def carregar_atos() -> dict:
    """Lê a planilha do processo e monta linhas/colunas da tabela (roda fora da tela)."""
    proc = main.proc()
    modelo = proc.modelo()
    if not modelo:
        return {"erro": f"Nenhum Word (.docx) em entradas/processos/{main.processo.name}/"}
    planilha = _planilha_do_processo()
    if not planilha:
        return {"erro": "Planilha do processo não encontrada (edite o processo no menu Processos)"}

    # Colunas na ordem que a pessoa escolheu (padrão: número do ato, datas de cada etapa,
    # campos do Word e caixinhas), só as visíveis
    tabela = proc.colunas_tabela()
    colunas = [(c["nome"], c["rotulo"]) for c in tabela if c["ver"]]
    caixas = [c["nome"] for c in tabela if c["caixa"]]

    registros = proc.ler_planilha(planilha)
    chave_edocs = main.preparado_chave_edocs(registros)
    linhas = []
    resumo = {"total": 0, "edocs": 0, "diario": 0, "publicado": 0, "erro": 0,
              "pend_edocs": 0, "pend_diario": 0, "pend_pub": 0}
    for r in registros:
        cores = _cores(r, chave_edocs)
        if not proc.vai_ao_diario:
            cores["diario"] = cores["pub"] = (BRANCO, "Este processo não é publicado no Diário Oficial")
        valores = proc.valores(r)  # como vai no ato (ligações do processo, ex.: DIA EXT)
        linha = {"id": r["_LINHA"], "linha": r["_LINHA"], "ordem": r["_LINHA"],
                 "criado_por": str(r.get("CRIADOPOR", "") or "").strip()}
        for etapa, (cor, texto) in cores.items():
            linha[f"{etapa}_cor"], linha[f"{etapa}_txt"] = cor, texto
        linha |= {nome: valores.get(nome, "") for nome, _ in colunas}
        linha[DATA_PUBLICACAO] = _data_publicacao(r.get("STATUSPUBLICADO", ""))
        for nome in caixas:  # caixinha marcada = SIM (ou X, OK...) na planilha
            linha[nome] = atos.normalizar(r.get(nome, "")) in processos.MARCADO
        # O que pode ser desfeito (há status gravado na planilha)
        linha["tem_edocs"] = bool(r.get(chave_edocs) or r.get("DATAEDOCS"))
        linha["tem_diario"] = bool(r.get("STATUSDIARIO") or r.get("DATADIARIO") or r.get("STATUSPUBLICADO"))
        # Já saiu no Diário Oficial: é público, os status ficam travados
        linha["publicado"] = main.publicado(r)
        linhas.append(linha)

        resumo["total"] += 1
        resumo["edocs"] += cores["edocs"][0] == VERDE
        resumo["diario"] += cores["diario"][0] == VERDE
        resumo["publicado"] += cores["pub"][0] == VERDE
        resumo["erro"] += any(c == VERMELHO for c, _ in cores.values())
        resumo["pend_edocs"] += cores["edocs"][0] == BRANCO
        resumo["pend_diario"] += cores["diario"][0] == AMARELO
        resumo["pend_pub"] += cores["pub"][0] == AMARELO
    linhas += _grupos_da_tabela(proc, registros, linhas)
    return {"linhas": linhas, "colunas": colunas, "caixas": caixas, "resumo": resumo,
            "planilha": str(planilha), "registros": {r["_LINHA"]: r for r in registros},
            "modelo": modelo}


def _faixa(numeros: list[int], mostrar: int | None = 2) -> str:
    """Linhas de um ato para mostrar: seguidas viram faixa ("2–49, 274–276"); separadas,
    "4, 7". Com muitos pedaços, mostra só os primeiros e "+ N" (a lista toda vai na dica)."""
    numeros = sorted(set(numeros))
    pedacos, inicio = [], numeros[0]
    for anterior, n in zip(numeros, numeros[1:] + [None]):
        if n is None or n != anterior + 1:
            pedacos.append(str(inicio) if inicio == anterior else
                           f"{inicio}, {anterior}" if anterior == inicio + 1 else f"{inicio}–{anterior}")
            inicio = n
    if mostrar and len(pedacos) > mostrar + 1:
        return ", ".join(pedacos[:mostrar]) + f" + {len(pedacos) - mostrar}"
    return ", ".join(pedacos)


def _grupos_da_tabela(proc, registros: list[dict], linhas: list[dict]) -> list[dict]:
    """Uma linha de grupo para cada ato com mais de uma pessoa ("▸ 2 pessoas"), que abre e
    mostra as linhas que vão nesse ato. As pessoas ganham "grupo" e uma "ordem" logo abaixo
    do grupo (a tabela ordena por ela), para ficarem juntas mesmo longe na planilha."""
    if not proc.coluna_juntar:
        return []
    por_linha = {l["id"]: l for l in linhas}
    da_pessoa = proc.colunas_da_pessoa() | {atos.normalizar(proc.config["coluna_nome"])}
    nome = atos.normalizar(proc.config["coluna_nome"])
    grupos = []
    for ato in proc.agrupar(registros):
        numeros = processos.linhas(ato)
        if len(numeros) < 2:
            por_linha[numeros[0]] |= {"sozinho": True, "base": numeros[0]}  # "1 pessoa", sem abrir/fechar
            continue
        chave = f"ato{min(numeros)}"
        topo = max(numeros)
        filhos = [por_linha[n] for n in sorted(numeros, reverse=True)]
        for k, filho in enumerate(filhos, 1):
            filho |= {"grupo": chave, "filho": True, "ordem": topo + 0.9 - k / 1000}
        primeira = por_linha[min(numeros)]
        cabeca = {c: ("" if c in da_pessoa else v) for c, v in primeira.items()}
        cabeca |= {"id": chave, "grupo": chave, "cabeca": True, "qtd": len(numeros), "filhos": numeros,
                   "linha": "", "linhas_txt": _faixa(numeros),
                   "linhas_dica": _faixa(numeros, None), "ordem": topo + 0.9, "base": min(numeros)}
        if nome in cabeca:
            cabeca[nome] = f"{primeira.get(nome, '')} e mais {len(numeros) - 1}"
        grupos.append(cabeca)
    return grupos


def diario_de_hoje(baixar: bool = True) -> dict:
    """O Diário Oficial de hoje, numa pasta só dele (entradas/Diario_hoje):
    apaga os PDFs de outros dias (não acumula) e, se o de hoje ainda não estiver lá, baixa
    todas as edições do dia (inclusive as extras, "7455-A"). Roda ao abrir o app."""
    import diario

    hoje = date.today()
    pasta = main.DIARIO_HOJE
    pasta.mkdir(parents=True, exist_ok=True)
    prefixo = f"Diario_Oficial_{hoje:%Y-%m-%d}"
    for pdf in pasta.glob("*.pdf"):
        if not pdf.name.startswith(prefixo):
            try:
                pdf.unlink()
            except OSError:
                pass  # aberto em outro programa: some na próxima vez
    arquivos = sorted(pasta.glob(f"{prefixo}*.pdf"))
    if arquivos or not baixar:
        return {"arquivos": arquivos, "baixou": False}
    for dia, numero, id_pdf in diario.edicoes_no_site():
        if dia == hoje:
            destino = pasta / f"{prefixo}_ed{numero}.pdf"
            destino.write_bytes(diario._baixar(f"{diario.SITE}/diario-oficial/versao-pdf/{id_pdf}"))
    return {"arquivos": sorted(pasta.glob(f"{prefixo}*.pdf")), "baixou": True}


def conferir_diario_de_hoje(arquivos: list[Path], registros: list[dict] | None = None) -> dict:
    """Procura os atos do processo em uso nos PDFs do Diário de hoje e marca na planilha
    como publicados os que ainda não estavam (como a etapa 3, só que com o Diário de hoje).
    registros: os já lidos pela tela (evita ler a planilha de novo).
    Devolve {"atos": {pdf: [(descrição, "novo" | "ja")]}, "nao_gravou": motivo ou ""}."""
    import diario

    proc = main.proc()
    planilha = _planilha_do_processo()
    if registros is None:
        registros = proc.ler_planilha(planilha) if planilha else []
    padrao = proc.padrao_conferencia()
    status = None
    saida: dict = {"atos": {}, "nao_gravou": ""}
    for pdf in arquivos:
        texto = diario._texto_pdf(pdf)
        rotulo = diario.edicao(texto, pdf)
        numero = diario.info_edicao(texto)[0]
        m = re.search(r"_ed([\w-]+)\.pdf$", pdf.name)  # edição extra ("7455-A")
        if m and numero and m.group(1).startswith(numero):
            rotulo = rotulo.replace(numero, m.group(1), 1)
        achados = {diario.chave(n, a) for n, a in padrao.findall(texto)}
        lista = []
        for r in registros:
            if proc.chave_conferencia(r) not in achados:
                continue
            if status is None:
                status = main.StatusPlanilha(planilha, "publicado", com_data=False)
            if str(r.get(status.chave(), "")).startswith(planilha_status.PUBLICADO):
                lista.append((proc.descricao(r), "ja"))
            else:
                status.pendentes[status.col_status][r["_LINHA"]] = f"{planilha_status.PUBLICADO} {rotulo}"
                lista.append((proc.descricao(r), "novo"))
        saida["atos"][pdf] = lista
    # Grava tudo de uma vez; planilha aberta no Excel não trava a tela (avisa e tenta depois)
    if status and any(status.pendentes.values()):
        try:
            planilha_status.gravar(planilha, status.aba, status.reservadas,
                                   {c: {} for c in status.reservadas} | status.pendentes)
        except PermissionError:
            quem = planilha_status.em_uso(planilha)
            saida["nao_gravou"] = f"a planilha está aberta no Excel{f' por {quem}' if quem else ''}"
        except Exception as e:
            saida["nao_gravou"] = f"{type(e).__name__}"
    return saida


def testar_login(site: str) -> tuple[str, str, str]:
    """(cor, usuário, explicação) do login no site, entrando de verdade."""
    cred = credenciais.carregar(main.CREDENCIAIS, site)
    if not cred:
        return BRANCO, "", "Nenhum usuário cadastrado"
    if site == "edocs":
        import edocs_http as modulo
    else:
        import diario_http as modulo
    try:
        modulo.entrar(*cred)
    except modulo.LoginRecusado:
        return VERMELHO, cred[0], "Usuário ou senha recusados pelo site"
    except Exception as e:
        return VERMELHO, cred[0], f"Sem conexão com o site ({type(e).__name__})"
    return VERDE, cred[0], "Logado"


# --- Tela -----------------------------------------------------------------------

class Tela:
    def __init__(self):
        self.motor = Janela()
        ui.usar(self.motor)
        self._pendentes: list = []  # perguntas abertas: (janela, futuro)
        self.rodando = False
        self.dados: dict = {}
        self.filtro = "todos"
        self.abertos: set[str] = set()  # atos com várias pessoas abertos na tabela
        self.login: dict[str, tuple[str, str, str]] = {s: (AMARELO, "", "Verificando...") for s in SITES}

        existentes = main.listar_processos()
        salvo = main.ler_config().get("processo")
        main.processo = next((p for p in existentes if p.name == salvo), existentes[0] if existentes else None)

        gui.add_css(CSS + folha.CSS + simulacao_ui.CSS)
        self.tema = gui.dark_mode(value=main.ler_config().get("tema") == "escuro")
        gui.colors(primary="#2459a8", secondary="#2f78c4", positive="#16a34a",
                   negative="#dc2626", warning="#ca8a04")
        self.camada = gui.element()  # onde as caixas de diálogo são criadas
        self._cabecalho()
        if main.processo is None:
            # Primeiro uso: oferece criar o primeiro processo pelo assistente
            _versao_no_canto()
            with gui.column().classes("w-full items-center q-pa-xl gap-3"):
                _logo("w-40 h-40")
                gui.label("Bem-vindo! Vamos criar o seu primeiro processo").classes("text-h5")
                gui.label("Você vai escolher o seu Word modelo e a sua planilha, e dizer qual coluna "
                          "preenche cada campo do Word.").classes("text-body1 t-forte")
                gui.button("Criar processo", icon="add", on_click=lambda: Assistente.novo(self)) \
                    .props("size=lg no-caps")
                # Versão nova do programa: traz os processos exportados da versão antiga
                gui.button("Já tenho processos exportados: importar", icon="file_download",
                           on_click=self._importar_processos).props("flat no-caps")
            return
        self._menu_e_paginas()
        # Primeira vez que abre esta instalação: oferece trazer os processos da versão antiga
        if not main.ler_config().get("boas_vindas_vista"):
            gui.timer(0.6, self._boas_vindas, once=True)
        gui.timer(0.1, self._bombear)
        gui.timer(0.2, self.recarregar, once=True)
        gui.timer(0.5, self.verificar_logins, once=True)
        self._ultimo_evento = time.monotonic()
        gui.timer(1.0, self._pulso)
        gui.timer(1.0, self._diario_de_hoje, once=True)  # baixa o Diário do dia, se faltar

    # Cabeçalho: nome, processo e situação dos logins
    def _cabecalho(self) -> None:
        # Degradê da esquerda (branco) para a direita (azul); no lado branco o texto é escuro
        # Degradê, sombra (profundidade) e espaçamento: classe "cabecalho" no CSS (tem
        # versão escura); no lado claro do degradê o texto é escuro ("cab-texto")
        with gui.header().classes("items-center cabecalho"):
            if main.processo is not None:
                gui.button(icon="menu", on_click=lambda: self.menu.toggle()) \
                    .props("flat round dense").classes("cab-texto").tooltip("Menu")
            _logo("w-16 h-16")
            with gui.column().classes("gap-0 cab-texto"):
                gui.label(TITULO).classes("text-h6 text-weight-bold")
                gui.label("Atos no e-docs e no Diário Oficial · Prefeitura de Uberlândia") \
                    .classes("text-caption")
            gui.space()
            if main.processo is not None:
                nomes = [p.name for p in main.listar_processos()]
                self.sel_processo = gui.select(nomes, value=main.processo.name, label="Processo",
                                               on_change=self._trocar_processo) \
                    .props("dark dense outlined options-dark=false").classes("w-56 sel-processo")
            escuro = self.tema.value
            self.botao_tema = gui.button(icon="light_mode" if escuro else "dark_mode",
                                         on_click=self._trocar_tema) \
                .props("flat dense square").classes("botao-tema q-ml-sm")
            self.botao_tema.tooltip("Modo claro" if escuro else "Modo escuro")
            # Barra de andamento (andamento() mostra, andamento(None) esconde)
            with gui.element("div").classes("barra-andamento") as self.barra:
                with gui.row().classes("w-full items-center gap-3 no-wrap q-px-lg q-py-xs"):
                    gui.spinner(size="18px", color="primary")
                    self.barra_texto = gui.label("").classes("text-body2 text-weight-medium ellipsis col")
                    self.barra_conta = gui.label("").classes("text-body2 text-weight-bold")
                self.barra_progresso = gui.linear_progress(value=0, show_value=False, size="4px",
                                                           color="primary").props("instant-feedback")
            self.barra.set_visibility(False)

    def andamento(self, texto: str | None, feitos: int = 0, total: int = 0) -> None:
        """Mostra o que está acontecendo na barra abaixo do cabeçalho ("Vistando: Portaria
        123", 3 de 10). Sem total, a barra corre sem fim; texto None esconde a barra."""
        if not hasattr(self, "barra"):
            return
        if texto is None:
            self.barra.set_visibility(False)
            return
        self.barra_texto.text = texto
        self.barra_conta.text = f"{feitos} de {total}" if total else ""
        if total:
            self.barra_progresso.props(remove="indeterminate")
            self.barra_progresso.set_value(feitos / total)
        else:
            self.barra_progresso.props("indeterminate")
        self.barra.set_visibility(True)

    def _trocar_tema(self) -> None:
        """Alterna entre o modo claro e o escuro (fica salvo para a próxima vez)."""
        escuro = not self.tema.value
        self.tema.set_value(escuro)
        main.salvar_config(tema="escuro" if escuro else "claro")
        self.botao_tema.props(f"icon={'light_mode' if escuro else 'dark_mode'}")
        self.botao_tema.tooltip("Modo claro" if escuro else "Modo escuro")
        if hasattr(self, "quadro_manual"):
            self.quadro_manual.props(f'src="{self._endereco_manual()}"')

    def _desenhar_logins(self) -> None:
        """Situação dos logins, no rodapé do menu lateral (clicar leva às Configurações)."""
        if not hasattr(self, "chips"):
            return
        self.chips.clear()
        with self.chips:
            for site, nome in SITES.items():
                cor, usuario, dica = self.login[site]
                # Uma linha só: "● e-docs (10717…)". A cor da bolinha diz a situação; o
                # usuário completo e o motivo aparecem ao passar o mouse.
                usuario_txt = usuario or "sem usuário"
                with gui.item(on_click=lambda: self.ir("config")).props("dense") \
                        .classes("login-item q-py-none").style("min-height:30px"):
                    # Usuário completo; se não couber, "…" no fim (o parêntese fica)
                    with gui.row().classes("w-full items-center gap-2 no-wrap overflow-hidden"):
                        _bolinha(cor, pequena=True)
                        gui.label(nome).classes("text-body2 text-weight-medium")
                        with gui.row().classes("items-center gap-0 no-wrap col overflow-hidden"):
                            gui.label("(").classes("usuario-login usuario-pequeno")
                            gui.label(usuario_txt).classes("usuario-login usuario-pequeno ellipsis") \
                                .style("min-width:0")
                            gui.label(")").classes("usuario-login usuario-pequeno")
                    gui.tooltip(f"{nome}: {usuario or 'sem usuário'} · {dica}")

    def _menu_e_paginas(self) -> None:
        """Menu lateral (abre pelo botão ☰ e fecha ao escolher ou clicar fora) e as páginas."""
        self.menu = gui.left_drawer(value=False).props("behavior=mobile width=270") \
            .classes("menu-lateral q-pa-none")
        self.itens_menu = {}
        with self.menu, gui.column().classes("w-full h-full no-wrap gap-0"):
            with gui.row().classes("items-center gap-2 q-pa-md"):
                _logo("w-12 h-12")
                gui.label("Menu").classes("text-h6")
            gui.separator()
            with gui.list().classes("w-full q-mt-sm"):
                for chave, nome, icone in PAGINAS:
                    with gui.item(on_click=lambda c=chave: self.ir(c)).classes("menu-item") as item:
                        with gui.item_section().props("avatar"):
                            gui.icon(icone)
                        with gui.item_section():
                            gui.item_label(nome)
                    self.itens_menu[chave] = item
            # Rodapé do menu: situação dos logins, separada das páginas por um traço
            gui.space()
            gui.separator()
            self.chips = gui.list().classes("w-full q-py-sm")
            self._desenhar_logins()

        with gui.row().classes("items-center gap-2 q-px-md q-pt-md"):
            self.icone_pagina = gui.icon("", size="28px").classes("titulo-icone")
            self.titulo_pagina = gui.label("").classes("text-h5 text-weight-medium t-forte")
        with gui.tab_panels(value="diario").classes("w-full bg-transparent") as self.paginas:
            with gui.tab_panel("atos"):
                self._painel_atos()
            with gui.tab_panel("recebidos"):
                from recebidos_ui import PaginaRecebidos
                self.recebidos = PaginaRecebidos(self)
            with gui.tab_panel("ato_manual"):
                from ato_manual_ui import PaginaAtoManual
                self.ato_manual = PaginaAtoManual(self)
            with gui.tab_panel("execucao"):
                self._painel_execucao()
            with gui.tab_panel("diario"):
                self._painel_diario()
            with gui.tab_panel("processos"):
                self._painel_processos()
            with gui.tab_panel("config"):
                self._painel_config()
            with gui.tab_panel("manual").classes("q-pt-sm"):
                self._painel_manual()
        self.ir("diario")  # o app abre no Diário Oficial

    def ir(self, pagina: str) -> None:
        """Mostra a página e marca o item no menu."""
        if not hasattr(self, "paginas"):
            return
        self.paginas.set_value(pagina)
        # O seletor de processo só onde o processo importa (no Diário, Recebidos e Manual, não)
        self.sel_processo.set_visibility(pagina in PAGINAS_COM_PROCESSO)
        for chave, item in self.itens_menu.items():
            item.classes(add="menu-ativo") if chave == pagina else item.classes(remove="menu-ativo")
        _, nome, icone = next(p for p in PAGINAS if p[0] == pagina)
        self.icone_pagina.name = icone
        self.titulo_pagina.text = nome
        self.menu.hide()
        if pagina == "diario" and not getattr(self, "diario_carregado", False):
            background_tasks.create(self._carregar_diario())
        if pagina == "recebidos" and not self.recebidos.carregado:
            background_tasks.create(self.recebidos.carregar())
        if pagina == "ato_manual":
            self.ato_manual.ao_abrir()

    # --- Abrir o site já logado (botão do globo) ---
    def _abrir_site(self, site: str) -> None:
        import janela_site

        nome, url = janela_site.SITES[site][:2]
        if not credenciais.carregar(main.CREDENCIAIS, site):
            gui.notify(f"Cadastre seu usuário do {nome} em Configurações para abrir já logado",
                       type="warning")
            self.ir("config")
            return
        if not tem_webview2():
            # Sem o componente da janela: abre o login no navegador, para entrar à mão
            import webbrowser
            webbrowser.open(url)
            gui.notify(f"Abrimos o {nome} no navegador: entre com o seu usuário e senha", type="info")
            return
        janela_site.abrir(site)
        gui.notify(f"Abrindo o {nome} já logado (numa janela separada)...", type="positive")

    # --- Página Manual ---
    def _endereco_manual(self) -> str:
        return "/manual/manual.html" + ("?tema=escuro" if self.tema.value else "")

    def _painel_manual(self) -> None:
        """O manual dentro do app, com o áudio (assets/manual/audio.mp3) para quem prefere ouvir."""
        # Qualquer arquivo de áudio na pasta do manual (o mais recente, se houver mais de um)
        audios = sorted((a for a in MANUAL.glob("*") if a.suffix.lower() in EXTENSOES_AUDIO),
                        key=lambda a: a.stat().st_mtime, reverse=True)
        audio = audios[0] if audios else None
        with gui.card().classes("cartao w-full q-pa-md"):
            with gui.row().classes("w-full items-center gap-4 no-wrap"):
                gui.icon("headphones", color="primary", size="32px")
                with gui.column().classes("gap-0 col"):
                    gui.label("Ouvir o manual").classes("text-subtitle1 text-weight-medium t-forte")
                    gui.label("Tá cozido de preguiça? Dê o play e ouça o  resumo sobre o "
                              "programa." if audio else "O áudio do manual ainda vai ser adicionado. "
                              "Por enquanto, leia abaixo.").classes("text-caption t-suave")
                if audio:
                    gui.audio(f"/manual/{quote(audio.name)}").classes("w-[420px] max-w-full")
                else:
                    gui.button("Ouvir", icon="play_arrow").props("unelevated no-caps disable")
                gui.button("Abrir em outra janela", icon="open_in_new",
                           on_click=lambda: abrir(MANUAL / "manual.html")).props("flat no-caps")
        self.quadro_manual = gui.element("iframe").props(f'src="{self._endereco_manual()}"') \
            .classes("w-full q-mt-md rounded-borders") \
            .style("height: calc(100vh - 330px); min-height: 400px; border: 0; background: transparent")

    # --- Aba Atos ---
    def _painel_atos(self) -> None:
        self.resumo = gui.grid(columns=5).classes("w-full gap-4")

        # As três etapas, na ordem, ocupando a largura toda
        acoes = [
            ("1", "Lançar no e-docs", "send", "primary", "#2459a8",
             "Cria cada ato no e-docs com vistos e assinantes e envia.",
             "pend_edocs", "ainda não enviado(s)", main.lancar_edocs_http),
            ("2", "Lançar no Diário", "newspaper", "secondary", "#2f78c4",
             "Cadastra no Diário Oficial os atos já enviados no e-docs.",
             "pend_diario", "pronto(s) para o Diário", main.lancar_diario_http),
            ("3", "Conferir publicação", "fact_check", "#1b3d78", "#1b3d78",
             "Procura nos Diários Oficiais publicados os atos enviados.",
             "pend_pub", "aguardando conferência", main.verificar_diario),
        ]
        self.botoes, self.contadores, self.avisos_diario = [], {}, []
        with gui.grid(columns=3).classes("w-full gap-4 q-mt-md"):
            for numero, nome, icone, cor, hexa, descricao, chave, sufixo, funcao in acoes:
                with gui.card().classes("cartao q-pa-lg w-full").style(f"border-top:4px solid {hexa}"):
                    with gui.row().classes("w-full items-center gap-3 no-wrap"):
                        gui.label(numero).classes("text-white text-weight-bold text-center") \
                            .style(f"background:{hexa};width:30px;height:30px;line-height:30px;"
                                   "border-radius:50%")
                        gui.label(nome).classes("text-h6")
                        gui.space()
                        gui.icon(icone, size="30px").style(f"color:{hexa}").classes("etapa-icone")
                    gui.label(descricao).classes("text-body2 t-suave")
                    if chave != "pend_edocs":
                        aviso = gui.label("Este processo não é publicado no Diário Oficial "
                                          "(para mudar: Processos)").classes("text-caption text-orange-8")
                        aviso.set_visibility(False)
                        self.avisos_diario.append(aviso)
                    with gui.row().classes("w-full items-center q-mt-sm"):
                        self.contadores[chave] = (gui.label("").classes("text-subtitle2"), sufixo)
                        gui.space()
                        site = {"pend_edocs": "edocs", "pend_diario": "diario"}.get(chave)
                        if site:
                            gui.button(icon="public", on_click=lambda s=site: self._abrir_site(s)) \
                                .props("flat round dense").classes("t-suave") \
                                .tooltip("Abrir o " + ("e-docs" if site == "edocs" else "Diário Oficial")
                                         + " já logado, para ver alguma coisa no site")
                        self.botoes.append(gui.button(
                            "Começar", icon="play_arrow", color=cor,
                            on_click=lambda n=nome, f=funcao: self.iniciar(n, f)).props("unelevated text-color=white"))

        with gui.card().classes("cartao w-full q-mt-md q-pa-none"):
            # Filtros, botões, busca e os ícones de pasta/recarregar numa linha só; em janela
            # estreita, o grupo da direita desce inteiro (a pasta nunca fica longe da busca)
            with gui.row().classes("w-full items-center q-pa-md gap-2 barra-atos"):
                gui.toggle({"todos": "Todos", "edocs": "Pendentes e-docs", "diario": "Pendentes Diário",
                            "pub": "Aguardando publicação", "erro": "Com erro"},
                           value="todos", on_change=lambda e: self._filtrar(e.value)) \
                    .props("no-caps unelevated toggle-color=primary")
                with gui.row().classes("items-center gap-2 no-wrap").style("margin-left: auto"):
                    gui.button("Colunas", icon="view_column", on_click=self._editar_colunas) \
                        .props("flat dense no-caps").tooltip("Mudar a ordem, mostrar/esconder colunas e criar caixinhas")
                    gui.button("Nova linha", icon="add", on_click=lambda: self._abrir_linha(None)) \
                        .props("outline no-caps").tooltip("Incluir uma linha na planilha")
                    self.busca = gui.input(placeholder="Buscar nome, matrícula, portaria...") \
                        .props("dense outlined clearable").classes("w-56")
                    self.busca.on_value_change(lambda: self._filtrar(self.filtro))
                    with self.busca.add_slot("prepend"):
                        gui.icon("search")
                    gui.button(icon="folder_open", on_click=self._abrir_pasta_planilha).props("flat round dense") \
                        .tooltip("Abrir a pasta onde está a planilha deste processo")
                    gui.button(icon="refresh", on_click=self.recarregar).props("flat round dense") \
                        .tooltip("Ler a planilha de novo")
            with gui.row().classes("q-px-md q-pb-sm gap-4 text-caption t-suave items-center"):
                for cor, texto in LEGENDA.items():
                    with gui.row().classes("items-center gap-1"):
                        _bolinha(cor)
                        gui.label(texto)
                gui.label("· clique numa linha para editar ou ver o ato montado").classes("italic")
            self.area_tabela = gui.column().classes("w-full")

    def _habilitar_etapas(self) -> None:
        """Botões das etapas: o Diário e a conferência só para processo que vai ao Diário."""
        if not getattr(self, "botoes", None):
            return
        diario = main.proc().vai_ao_diario
        for i, botao in enumerate(self.botoes):
            botao.set_enabled(not self.rodando and (i == 0 or diario))
        for aviso in self.avisos_diario:
            aviso.set_visibility(not diario)

    def _desenhar_resumo(self) -> None:
        self.resumo.clear()
        r = self.dados.get("resumo")
        if not r:
            return
        itens = [("Atos na planilha", r["total"], "groups", "primary"),
                 ("Enviados no e-docs", r["edocs"], "send", "primary"),
                 ("Enviados ao Diário", r["diario"], "newspaper", "secondary"),
                 ("Publicados", r["publicado"], "verified", "positive"),
                 ("Com erro", r["erro"], "error", "negative" if r["erro"] else "grey-6")]
        self._habilitar_etapas()
        for chave, (rotulo, sufixo) in self.contadores.items():
            rotulo.text = f"{r[chave]} {sufixo}"
            rotulo.classes(replace="text-subtitle2 " + ("t-suave" if not r[chave] else ""))
        with self.resumo:
            for rotulo, valor, icone, cor in itens:
                with gui.card().classes("cartao q-pa-md w-full"):
                    with gui.row().classes("items-center gap-3 no-wrap"):
                        gui.icon(icone, size="30px", color=cor)
                        with gui.column().classes("gap-0"):
                            gui.label(str(valor)).classes("text-h5 text-weight-bold leading-none")
                            gui.label(rotulo).classes("text-caption t-suave")

    def _desenhar_tabela(self) -> None:
        self.area_tabela.clear()
        with self.area_tabela:
            if "erro" in self.dados:
                with gui.row().classes("q-pa-lg items-center gap-3"):
                    gui.icon("warning", color="warning", size="32px")
                    gui.label(self.dados["erro"]).classes("text-body1")
                return
            colunas = [{"name": "linha", "label": "Linha", "field": "ordem", "sortable": True,
                        "align": "left"},
                       # Quem enviou ao e-docs (primeiro nome da conta); vazio nos atos de antes
                       {"name": "criado_por", "label": "Criado por", "field": "criado_por", "sortable": True,
                        "align": "left"},
                       {"name": "edocs", "label": "e-docs", "field": "edocs_cor", "align": "center"},
                       {"name": "diario", "label": "Diário", "field": "diario_cor", "align": "center"},
                       {"name": "pub", "label": "Publicado", "field": "pub_cor", "align": "center"}]
            colunas += [{"name": nome, "label": rotulo, "field": nome, "sortable": True, "align": "left"}
                        for nome, rotulo in self.dados["colunas"]]
            # Linhas marcadas nas caixinhas da esquerda: ações em lote (desfazer status)
            self.barra_selecao = gui.row().classes("w-full items-center gap-2 q-px-md q-py-sm fundo-previa")
            self.barra_selecao.set_visibility(False)
            # Mais novas primeiro: a linha maior da planilha (o ato mais recente) fica no topo
            self._cabecas_marcadas: set[str] = set()
            self.tabela = gui.table(rows=self._linhas_filtradas(), columns=colunas, row_key="id",
                                    selection="multiple", on_select=lambda e: self._mostrar_selecao(),
                                    pagination={"rowsPerPage": 15, "sortBy": "linha", "descending": True}) \
                .classes("tabela-atos w-full").props("flat dense separator=horizontal "
                                                     "rows-per-page-label='Linhas por página'")
            self.tabela.bind_filter_from(self.busca, "value")
            for etapa in ("edocs", "diario", "pub"):
                self.tabela.add_slot(f"body-cell-{etapa}", f"""
                    <q-td :props="props">
                      <span :class="'bolinha b-' + props.row.{etapa}_cor"></span>
                      <q-tooltip>{{{{ props.row.{etapa}_txt }}}}</q-tooltip>
                    </q-td>""")
            self.tabela.add_slot("body-cell", """
                <q-td :props="props">
                  <span v-if="props.value !== '' && props.value !== null">{{ props.value }}</span>
                  <span v-else class="t-fraco">—</span>
                </q-td>""")
            # Ato com várias pessoas: a linha do grupo abre/fecha as pessoas dele
            self.tabela.add_slot("body-cell-linha", """
                <q-td :props="props">
                  <span v-if="props.row.cabeca" class="row no-wrap items-center">
                    <q-icon :name="props.row.aberto ? 'expand_more' : 'chevron_right'" size="20px" />
                    <span class="q-mr-sm" style="white-space: nowrap">{{ props.row.linhas_txt }}</span>
                    <q-chip dense square color="primary" text-color="white" icon="groups" class="q-ma-none">
                      {{ props.row.qtd }} pessoas
                    </q-chip>
                    <q-tooltip>Um ato só com {{ props.row.qtd }} pessoas (linhas {{ props.row.linhas_dica }}): clique para ver quem vai nele</q-tooltip>
                    <q-btn flat round dense size="sm" icon="person_add" color="primary" class="q-ml-xs"
                           @click.stop="$parent.$emit('mais_pessoa', props.row.base)">
                      <q-tooltip>Adicionar outra pessoa a este ato</q-tooltip>
                    </q-btn>
                  </span>
                  <span v-else-if="props.row.filho" class="t-suave" style="padding-left: 22px">↳ {{ props.row.linha }}</span>
                  <span v-else-if="props.row.sozinho" class="row no-wrap items-center" style="padding-left: 20px">
                    <span class="q-mr-sm">{{ props.row.linha }}</span>
                    <q-chip dense square outline color="primary" icon="person" class="q-ma-none chip-uma-pessoa">1 pessoa</q-chip>
                    <q-btn flat round dense size="sm" icon="person_add" color="primary" class="q-ml-xs"
                           @click.stop="$parent.$emit('mais_pessoa', props.row.base)">
                      <q-tooltip>Adicionar outra pessoa a este ato</q-tooltip>
                    </q-btn>
                  </span>
                  <span v-else>{{ props.row.linha }}</span>
                </q-td>""")
            # Caixinhas criadas pela pessoa: marcar grava SIM na planilha (o clique não abre a linha)
            for nome in self.dados.get("caixas", []):
                self.tabela.add_slot(f"body-cell-{nome}", f"""
                    <q-td :props="props" @click.stop>
                      <q-checkbox v-if="!props.row.cabeca" dense :model-value="props.row.{nome}"
                        @update:model-value="v => $parent.$emit('caixa', {{id: props.row.id, nome: '{nome}', valor: v}})" />
                    </q-td>""")
            self.tabela.on("caixa", self._marcar_caixa)
            self.tabela.on("rowClick", lambda e: self._clicar_linha(e.args[1]))
            self.tabela.on("mais_pessoa", lambda e: background_tasks.create(self._abrir_linha(None, base=int(e.args))))

    async def _marcar_caixa(self, e) -> None:
        """Marca/desmarca a caixinha de uma linha: grava SIM (ou vazio) na coluna da planilha."""
        linha, nome, valor = e.args["id"], e.args["nome"], bool(e.args["valor"])
        if self.rodando:
            gui.notify("Aguarde o fim da execução para mexer na planilha", type="warning")
            return
        proc = main.proc()
        rotulo = next((c for c in proc.caixas if atos.normalizar(c) == nome), nome)
        planilha = Path(self.dados["planilha"])
        aba = proc.aba or atos.ABA  # como o StatusPlanilha (sem a aba, a primeira)

        def gravar() -> None:
            cols = planilha_status.colunas(planilha, aba, [rotulo])
            planilha_status.gravar(planilha, aba, cols, {rotulo: {linha: "SIM" if valor else ""}})

        try:
            await run.io_bound(gravar)
        except PermissionError:
            quem = planilha_status.em_uso(planilha)
            gui.notify(f"A planilha está aberta no Excel{f' por {quem}' if quem else ''}: feche e marque de novo",
                       type="warning")
            valor = not valor  # volta a caixinha como estava
        except Exception as e:
            gui.notify(f"Não consegui gravar na planilha: {type(e).__name__}", type="negative")
            valor = not valor
        for item in self.dados.get("linhas", []):
            if item["id"] == linha:
                item[nome] = valor
        registro = self.dados.get("registros", {}).get(linha)
        if registro is not None:
            registro[nome] = "SIM" if valor else ""
        self.tabela.rows = self._linhas_filtradas()
        self.tabela.update()

    def _editar_colunas(self) -> None:
        """Janela Colunas: ordem, quais aparecem e caixinhas (salva no processo.json)."""
        from colunas_ui import EditorColunas
        proc = main.proc()
        proc.config.setdefault("caixas", [])
        proc.config.setdefault("colunas_tabela", [])
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[560px] max-w-full q-pa-lg gap-3"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("view_column", color="primary", size="28px")
                gui.label("Colunas da tabela").classes("text-h6 t-forte")
            with gui.element("div").classes("w-full").style("max-height: 60vh; overflow-y: auto"):
                EditorColunas(proc.config, proc.colunas_tabela)

            def pronto() -> None:
                proc.salvar()
                janela.close()
                gui.timer(0.01, self.recarregar, once=True)

            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Salvar", icon="save", on_click=pronto).props("unelevated no-caps")
        janela.open()

    # --- Várias linhas marcadas: desfazer o status de todas de uma vez ---
    def _sincronizar_grupos(self) -> None:
        """Marcar o grupo de um ato marca (e mostra) as pessoas dele; desmarcar desmarca. Se
        uma pessoa do grupo for desmarcada, o grupo também sai."""
        marcadas = list(getattr(self.tabela, "selected", []) or [])
        cabecas = {r["grupo"] for r in marcadas if r.get("cabeca")}
        antes = getattr(self, "_cabecas_marcadas", set())
        pessoas = [l for l in self.dados.get("linhas", []) if l.get("filho")]
        ids = {r["id"] for r in marcadas}
        novos, sairam = cabecas - antes, antes - cabecas
        marcadas += [l for l in pessoas if l["grupo"] in novos and l["id"] not in ids]
        marcadas = [r for r in marcadas if not (r.get("filho") and r["grupo"] in sairam)]
        ids = {r["id"] for r in marcadas}
        incompletos = {l["grupo"] for l in pessoas if l["grupo"] in cabecas and l["id"] not in ids}
        marcadas = [r for r in marcadas if not (r.get("cabeca") and r["grupo"] in incompletos)]
        self._cabecas_marcadas = cabecas - incompletos
        if novos:
            self.abertos |= novos  # abre o grupo, para ver as pessoas marcadas
        self.tabela.selected = marcadas
        self._filtrar(self.filtro)

    def _mostrar_selecao(self) -> None:
        self._sincronizar_grupos()
        marcadas = getattr(self.tabela, "selected", []) or []
        self.barra_selecao.clear()
        self.barra_selecao.set_visibility(bool(marcadas))
        if not marcadas:
            return
        with self.barra_selecao:
            gui.icon("checklist", color="primary", size="22px")
            gui.label(f"{len(self._linhas_marcadas())} linha(s) marcada(s)").classes("text-body2 text-weight-medium t-forte")
            gui.space()
            diario = main.proc().vai_ao_diario
            if diario:
                gui.button("Desfazer Diário", icon="undo", color="orange-8",
                           on_click=lambda: self._confirmar_desfazer_varios("diario")) \
                    .props("outline dense no-caps") \
                    .tooltip("Apaga o status do Diário (e a conferência da publicação) das linhas marcadas")
            gui.button("Desfazer e-docs e Diário" if diario else "Desfazer e-docs", icon="restart_alt",
                       color="orange-8", on_click=lambda: self._confirmar_desfazer_varios("tudo")) \
                .props("unelevated dense no-caps") \
                .tooltip("As linhas marcadas voltam ao começo, como se nunca tivessem sido enviadas")
            gui.button("Excluir linhas", icon="delete", color="negative",
                       on_click=self._confirmar_exclusao_varias).props("outline dense no-caps")                 .tooltip("Tira da planilha as linhas marcadas")
            gui.button("Desmarcar", icon="close", on_click=self._limpar_selecao).props("flat dense no-caps")

    def _linhas_marcadas(self) -> set[int]:
        """As linhas da planilha marcadas (o grupo de um ato marca todas as pessoas dele)."""
        saida: set[int] = set()
        for r in getattr(self.tabela, "selected", []) or []:
            saida |= set(r["filhos"]) if r.get("cabeca") else {r["id"]}
        return saida

    def _limpar_selecao(self) -> None:
        self._cabecas_marcadas = set()
        self.tabela.selected = []
        self.tabela.update()
        self._mostrar_selecao()

    def _confirmar_desfazer_varios(self, etapa: str) -> None:
        if self.rodando:
            gui.notify("Aguarde o fim da execução para mexer na planilha", type="warning")
            return
        linhas = sorted(self._linhas_marcadas())
        registros = self.dados.get("registros", {})
        publicadas = [l for l in linhas if main.publicado(registros.get(l) or {})]
        tudo = etapa == "tudo"
        with self.camada, gui.dialog() as dialogo, gui.card().classes("cartao w-[540px] q-pa-lg gap-2"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("undo", color="orange-8", size="26px")
                gui.label(f"Desfazer {'e-docs e Diário' if tudo else 'Diário'} de {len(linhas)} linha(s)?") \
                    .classes("text-h6")
            gui.label("As linhas voltam ao começo: apaga os status e as datas do e-docs e do Diário. "
                      "Os atos podem ser enviados de novo." if tudo else
                      "Apaga o status do Diário, a data de veiculação e a conferência da publicação. Os "
                      "atos voltam a ficar prontos para lançar no Diário.").classes("text-body2")
            with gui.element("div").classes("w-full fundo-previa rounded q-pa-sm") \
                    .style("max-height: 160px; overflow-y: auto"):
                for l in linhas:
                    texto = main.proc().descricao(registros[l]) if l in registros else f"linha {l}"
                    gui.label(f"Linha {l} · {texto}" + (" (publicado: fica como está)" if l in publicadas else "")) \
                        .classes("text-caption " + ("t-fraco" if l in publicadas else "t-forte"))
            gui.label("Só muda a planilha: nada é apagado no site.").classes("text-caption t-suave")
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=dialogo.close).props("flat no-caps")
                botao = gui.button("Desfazer", icon="undo", color="orange-8").props("unelevated no-caps")

        async def desfazer() -> None:
            planilha = Path(self.dados["planilha"])
            botao.props("loading")
            dialogo.close()
            self.andamento(f"Desfazendo o status de {len(linhas)} linha(s) na planilha...")
            resultado: dict = {}

            def gravar() -> None:
                try:
                    resultado.update(main.desfazer_varios(planilha, linhas, etapa))
                except PermissionError:
                    raise edicao.PlanilhaAberta(planilha_status.em_uso(planilha))

            try:
                await self._quando_livre(planilha, gravar)
            except edicao.PlanilhaAberta:
                self.andamento(None)
                return
            except Exception as e:
                self.andamento(None)
                gui.notify(f"Não consegui desfazer: {type(e).__name__} {e}".strip(), type="negative")
                return
            self.andamento(None)
            partes = [f"{len(resultado['desfeitas'])} linha(s) desfeita(s)"]
            if resultado["publicadas"]:
                partes.append(f"{len(resultado['publicadas'])} já publicada(s) ficaram como estavam")
            if resultado["sem_status"]:
                partes.append(f"{len(resultado['sem_status'])} não tinham status")
            gui.notify("; ".join(partes), type="positive" if resultado["desfeitas"] else "info", multi_line=True)
            await self.recarregar()

        botao.on_click(desfazer)
        dialogo.open()

    def _confirmar_desfazer(self, linha: int, etapa: str, janela=None) -> None:
        """Confirma e apaga o status (e a data) de uma etapa da linha."""
        if self.rodando:
            gui.notify("Aguarde o fim da execução para mexer na planilha", type="warning")
            return
        registro = self.dados.get("registros", {}).get(linha) or {}
        ato = main.proc().descricao(registro) if registro else f"linha {linha}"
        if main.publicado(registro):
            gui.notify("Este ato já foi publicado no Diário Oficial: o envio não pode ser desfeito",
                       type="warning")
            return
        linha_tabela = next((l for l in self.dados.get("linhas", []) if l["id"] == linha), {})
        if etapa == "edocs" and linha_tabela.get("tem_diario"):
            gui.notify("Desfaça primeiro o envio no Diário: as etapas se desfazem de trás para frente",
                       type="warning")
            return
        if etapa == "edocs":
            titulo = "Desfazer envio no e-docs?"
            texto = ("Apaga o status do e-docs e a data de envio desta linha na planilha. O ato volta "
                     "a aparecer como não enviado e pode ser enviado de novo.")
        else:
            titulo = "Desfazer envio no Diário?"
            texto = ("Apaga o status do Diário, a data de veiculação e a conferência da publicação "
                     "desta linha. O ato volta a ficar pronto para lançar no Diário.")
        with self.camada, gui.dialog() as dialogo, gui.card().classes("cartao w-[500px] q-pa-lg"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("undo", color="orange-8", size="26px")
                gui.label(titulo).classes("text-h6")
            gui.label(ato).classes("text-subtitle2 t-forte")
            gui.label(texto).classes("text-body2")
            gui.label("Só muda a planilha: nada é apagado no site.").classes("text-caption t-suave")
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=dialogo.close).props("flat no-caps")
                botao = gui.button("Desfazer", icon="undo", color="orange-8").props("unelevated no-caps")

                async def desfazer() -> None:
                    planilha = Path(self.dados["planilha"])
                    botao.props("loading")

                    def gravar() -> None:
                        try:
                            main.desfazer_status(planilha, linha, etapa)
                        except PermissionError:
                            raise edicao.PlanilhaAberta(planilha_status.em_uso(planilha))

                    try:
                        await self._quando_livre(planilha, gravar)
                    except edicao.PlanilhaAberta:
                        botao.props(remove="loading")
                        return
                    except main.AtoPublicado:
                        dialogo.close()
                        gui.notify("Este ato já foi publicado no Diário Oficial: o envio não pode ser "
                                   "desfeito", type="warning")
                        await self.recarregar()
                        return
                    except main.DiarioAntes:
                        dialogo.close()
                        gui.notify("Desfaça primeiro o envio no Diário: as etapas se desfazem de trás "
                                   "para frente", type="warning")
                        await self.recarregar()
                        return
                    except Exception as e:
                        botao.props(remove="loading")
                        gui.notify(f"Não consegui desfazer: {type(e).__name__} {e}".strip(), type="negative")
                        return
                    dialogo.close()
                    if janela is not None:
                        janela.close()
                    gui.notify(f"Status {'do e-docs' if etapa == 'edocs' else 'do Diário'} desfeito: {ato}",
                               type="positive")
                    await self.recarregar()

                botao.on_click(desfazer)
        dialogo.open()

    def _linhas_filtradas(self) -> list[dict]:
        linhas = self.dados.get("linhas", [])
        filtros = {
            "todos": lambda l: True,
            "edocs": lambda l: l["edocs_cor"] == BRANCO,
            "diario": lambda l: l["diario_cor"] == AMARELO,
            "pub": lambda l: l["pub_cor"] == AMARELO,
            "erro": lambda l: VERMELHO in (l["edocs_cor"], l["diario_cor"], l["pub_cor"]),
        }
        passa = filtros[self.filtro]
        busca = bool(getattr(self, "busca", None) and (self.busca.value or "").strip())
        com_alguem = {l["grupo"] for l in linhas if l.get("filho") and passa(l)}
        saida = []
        for l in linhas:
            if l.get("cabeca"):
                if l["grupo"] in com_alguem:
                    saida.append(dict(l, aberto=l["grupo"] in self.abertos))
            # Pessoa de um ato com várias: só com o grupo aberto (ou buscando, para achar)
            elif passa(l) and (not l.get("filho") or busca or l["grupo"] in self.abertos):
                saida.append(l)
        return saida

    def _clicar_linha(self, linha: dict) -> None:
        """Clique na linha do grupo abre/fecha as pessoas do ato; nas outras, abre a linha."""
        if not linha.get("cabeca"):
            background_tasks.create(self._abrir_linha(linha["id"]))
            return
        self.abertos ^= {linha["grupo"]}
        self._filtrar(self.filtro)

    def _filtrar(self, filtro: str) -> None:
        self.filtro = filtro
        if hasattr(self, "tabela") and "linhas" in self.dados:
            self.tabela.rows = self._linhas_filtradas()
            self.tabela.update()

    def _abrir_pasta_planilha(self) -> None:
        """A pasta da planilha do processo em uso, com a planilha já selecionada."""
        planilha = consultas_ui.planilha_de(main.processo)
        if planilha is None:
            gui.notify("Não achei a planilha deste processo: confira em Processos (Alterar, ao lado da planilha)",
                       type="warning")
            return
        mostrar_na_pasta(planilha)

    async def recarregar(self) -> None:
        self.area_tabela.clear()
        with self.area_tabela:
            with gui.row().classes("q-pa-lg items-center gap-3"):
                gui.spinner(size="lg")
                gui.label("Lendo a planilha...").classes("t-suave")
        try:
            dados = await run.io_bound(carregar_atos)
        except Exception as e:
            dados = {"erro": f"Não consegui ler a planilha: {type(e).__name__} {e}"}
        if dados is None:  # o app está fechando
            return
        self.dados = dados
        self._desenhar_resumo()
        self._desenhar_tabela()
        self._ler_estrutura()  # prepara o editor de linhas em segundo plano

    # --- Editar a planilha (nova linha, editar, excluir) ---
    def _ler_estrutura(self) -> asyncio.Task | None:
        """Lê (em segundo plano) as colunas e os valores da planilha para o editor de linhas.
        Começa logo depois de a tabela carregar, para a janela da linha abrir na hora."""
        if "planilha" not in self.dados:
            return None
        planilha = Path(self.dados["planilha"])
        chave = (str(planilha), planilha.stat().st_mtime, main.proc().aba)
        tarefa = getattr(self, "_estrutura", None)
        if tarefa is None or self._estrutura_chave != chave:
            self._estrutura_chave = chave
            self._estrutura = background_tasks.create(
                run.io_bound(edicao.ler_estrutura, planilha, main.proc().aba))
        return self._estrutura

    async def _abrir_linha(self, linha: int | None, base: int | None = None) -> None:
        """Janela de uma linha: Dados (editar) e Ato montado. linha=None inclui uma nova; com
        base (uma linha de um ato já salvo), a nova pessoa entra nesse ato: vem com o que é do
        ato (número, cargo...) e os dados da pessoa em branco."""
        if "planilha" not in self.dados:
            return
        if self.rodando:
            gui.notify("Aguarde o fim da execução para mexer na planilha", type="warning")
            return
        registro = self.dados.get("registros", {}).get(linha) if linha else None
        campos: dict = {}
        estado: dict = {}
        with self.camada, gui.dialog() as dialogo, \
                gui.card().classes("cartao w-[920px] max-w-full q-pa-none gap-0"):
            with gui.row().classes("w-full items-center q-px-lg q-pt-md q-pb-sm no-wrap"):
                gui.icon("add_circle" if linha is None else "edit_note", color="primary", size="26px")
                do_ato = self.dados.get("registros", {}).get(base) if base else None
                gui.label(f"Nova pessoa no ato · {main.proc().descricao(do_ato)}" if do_ato else
                          "Nova pessoa no ato" if base else "Nova linha" if linha is None else
                          f"Linha {linha}" + (f" · {main.proc().descricao(registro)}" if registro else "")) \
                    .classes("text-h6 ellipsis")
                gui.space()
                gui.button(icon="close", on_click=dialogo.close).props("flat round dense")
            with gui.tabs().props("dense no-caps align=left active-color=primary indicator-color=primary") \
                    .classes("q-px-md t-forte") as abas:
                gui.tab("dados", "Dados", icon="edit")
                if registro:
                    gui.tab("ato", "Ato montado", icon="description")
                    gui.tab("status", "Status", icon="flag")
            gui.separator()
            with gui.tab_panels(abas, value="dados").classes("w-full"):
                with gui.tab_panel("dados").classes("q-pa-none"):
                    with gui.scroll_area().classes("h-[430px] w-full"):
                        area_dados = gui.column().classes("w-full")
                        with area_dados, gui.row().classes("q-pa-lg items-center gap-3"):
                            gui.spinner(size="lg")
                            gui.label("Lendo a planilha...").classes("t-suave")
                if registro:
                    with gui.tab_panel("ato").classes("q-pa-md"):
                        area_ato = gui.column().classes("w-full")
                        with area_ato:
                            gui.spinner()
                    with gui.tab_panel("status").classes("q-pa-md"):
                        self._painel_status(dialogo, linha, registro)
            gui.separator()
            # Linha nova: dá para juntar várias numa lista e salvar todas de uma vez (uma
            # gravação e um recálculo no Excel, em vez de um por linha, que é o demorado)
            fila: list[dict] = []
            if linha is None:
                area_fila = gui.row().classes("w-full items-center gap-1 q-px-lg q-pt-sm")
            with gui.row().classes("w-full items-center q-px-lg q-py-sm"):
                if linha is not None:
                    gui.button("Excluir", icon="delete", color="negative",
                               on_click=lambda: self._confirmar_exclusao(
                                   dialogo, linha, estado.get("colunas", []), registro)) \
                        .props("flat no-caps")
                    if main.proc().coluna_juntar:
                        gui.button("Outra pessoa deste ato", icon="person_add",
                                   on_click=lambda: (dialogo.close(), background_tasks.create(
                                       self._abrir_linha(None, base=linha)))) \
                            .props("flat no-caps").tooltip("Inclui mais uma pessoa neste mesmo ato")
                gui.space()
                gui.button("Cancelar", on_click=dialogo.close).props("flat no-caps")
                if linha is None:
                    # Processo com várias pessoas por ato: a próxima pessoa do mesmo ato só
                    # precisa dos dados dela (o cargo, o edital... vêm da linha anterior)
                    da_pessoa = main.proc().colunas_da_pessoa() if main.proc().coluna_juntar else set()
                    mesmo_ato = gui.button("Outra pessoa do mesmo ato", icon="person_add",
                                           on_click=lambda: repetir_ultima(so_do_ato=True)) \
                        .props("flat no-caps") \
                        .tooltip("A próxima pessoa vai no mesmo ato que a última da lista: copia o que é "
                                 "igual para o ato (número, cargo...) e deixa em branco os dados da pessoa. "
                                 "Sem número do ato, as duas ganham o mesmo")
                    mesmo_ato.set_visibility(False)
                    repetir = gui.button("Repetir a última", icon="content_copy",
                                         on_click=lambda: repetir_ultima()).props("flat no-caps") \
                        .tooltip("Preenche o formulário com a última linha da lista (para mudar só o que for diferente)")
                    gui.button("Adicionar à lista", icon="playlist_add", on_click=lambda: adicionar_a_lista()) \
                        .props("outline no-caps").tooltip("Guarda esta linha e limpa o formulário para a próxima")
                salvar = gui.button("Salvar", icon="save").props("unelevated no-caps")
                salvar.disable()
                if linha is None:
                    salvar.on_click(lambda: self._salvar_novas(
                        dialogo, salvar, estado.get("colunas", []),
                        fila + ([valores_do_formulario()] if any(valores_do_formulario().values()) else []), campos,
                        estado.get("numero_da_base")))
                else:
                    salvar.on_click(lambda: self._salvar_linha(dialogo, salvar, linha,
                                                               estado.get("colunas", []), campos))

        def valores_do_formulario() -> dict[int, str]:
            return {i: str(c.value or "").strip() for i, c in campos.items()}

        def desenhar_fila() -> None:
            area_fila.clear()
            repetir.set_visibility(bool(fila))
            mesmo_ato.set_visibility(bool(fila) and bool(da_pessoa))
            # Salvar grava a lista e também a linha do formulário, se estiver preenchida
            salvar.text = f"Salvar tudo ({len(fila)} na lista)" if fila else "Salvar"
            with area_fila:
                if not fila:
                    gui.icon("lightbulb", color="amber-7", size="18px")
                    gui.label("Vai incluir várias? Preencha uma, clique em Adicionar à lista e siga para a "
                              "próxima. No fim, Salvar grava todas de uma vez.").classes("text-caption t-suave")
                    return
                gui.label(f"Na lista ({len(fila)}):").classes("text-caption text-weight-medium t-forte")
                for i, valores in enumerate(fila):
                    resumo = " · ".join(v for v in valores.values() if v)[:40] or f"Linha {i + 1}"
                    gui.chip(resumo, removable=True, color=None, on_value_change=lambda e, i=i: tirar(i)) \
                        .props("dense square").classes("chip-campo")

        def tirar(i: int) -> None:
            del fila[i]
            desenhar_fila()

        def adicionar_a_lista() -> None:
            valores = valores_do_formulario()
            if not any(valores.values()):
                gui.notify("Preencha pelo menos um campo antes de adicionar", type="warning")
                return
            if not self._obrigatorios_ok(estado.get("colunas", []), valores, campos):
                return
            fila.append(valores)
            if estado.get("busca"):
                estado["busca"].limpar_marcas()
            for c in campos.values():
                _por_valor(c, "")
            if campos:
                next(iter(campos.values())).run_method("focus")
            desenhar_fila()

        def mesmo_numero(nomes: dict) -> None:
            """O que diz que duas linhas são do mesmo ato é o mesmo valor na coluna de juntar
            (ex.: Nº DO DECRETO). Se a última da lista está sem, ganha o próximo número livre,
            que depois é copiado para a pessoa nova."""
            juntar = main.proc().coluna_juntar
            indice = next((i for i, n in nomes.items() if n == juntar and i in campos), None)
            if indice is None or str(fila[-1].get(indice, "")).strip():
                return
            usados = [str(v.get(indice, "")) for v in estado.get("linhas", {}).values()]
            usados += [str(f.get(indice, "")) for f in fila]
            numero = max((int(u) for u in (x.strip() for x in usados) if u.isdigit()), default=0) + 1
            fila[-1][indice] = str(numero)
            rotulo = next(c.nome for c in estado["colunas"] if c.indice == indice)
            gui.notify(f"{rotulo} {numero} nas duas pessoas: é isso que as coloca no mesmo ato",
                       type="info")
            desenhar_fila()

        def repetir_ultima(so_do_ato: bool = False) -> None:
            """Preenche com a última linha da lista; so_do_ato deixa em branco as colunas da
            pessoa (nome, matrícula...) e põe o cursor na primeira delas."""
            if not fila:
                return
            nomes = {c.indice: atos.normalizar(c.nome) for c in estado.get("colunas", [])}
            if estado.get("busca"):
                estado["busca"].limpar_marcas()
            if so_do_ato:
                mesmo_numero(nomes)
            primeira = None
            for i, c in campos.items():
                da_pessoa_ = so_do_ato and nomes.get(i) in da_pessoa
                _por_valor(c, "" if da_pessoa_ else fila[-1].get(i, ""))
                if da_pessoa_ and primeira is None:
                    primeira = c
            if primeira is not None:
                primeira.run_method("focus")

        if linha is None:
            desenhar_fila()
        dialogo.open()
        if registro:
            background_tasks.create(self._mostrar_ato(area_ato, registro))

        try:
            colunas, linhas = await self._ler_estrutura()
        except Exception as e:
            area_dados.clear()
            with area_dados:
                gui.label(f"Não consegui ler a planilha: {type(e).__name__}").classes("text-negative q-pa-md")
            return
        estado["colunas"] = colunas
        estado["linhas"] = linhas
        obrigatorias = {atos.normalizar(c) for c in main.proc().obrigatorios}
        valores = linhas.get(linha, {}) if linha else {}
        opcoes = _opcoes_das_colunas(colunas, linhas, main.proc())
        area_dados.clear()
        with area_dados:
            area_busca = gui.column().classes("w-full q-px-md q-pt-sm")
            with gui.grid(columns=2).classes("w-full gap-x-4 gap-y-1 q-px-md q-pt-sm"):
                for c in colunas:
                    if not c.editavel:
                        continue
                    campos[c.indice] = _campo_da_coluna(c, valores.get(c.indice, ""), opcoes.get(c.indice),
                                                        atos.normalizar(c.nome) in obrigatorias)
                    if atos.normalizar(c.nome) in obrigatorias:  # preenchido, sai o vermelho
                        campos[c.indice].on_value_change(
                            lambda e: e.value and e.sender.props(remove="error error-message"))
            calculadas = [c for c in colunas if c.formula]
            if calculadas:
                with gui.expansion(f"Calculados pelo Excel ({len(calculadas)})", icon="functions") \
                        .classes("w-full q-px-sm q-mt-sm t-forte"):
                    if linha is None:
                        gui.label("Preenchidos pela própria planilha depois de salvar.") \
                            .classes("text-caption t-suave q-px-sm")
                    else:
                        with gui.grid(columns=2).classes("w-full gap-x-6 gap-y-1 q-px-sm"):
                            for c in calculadas:
                                with gui.column().classes("gap-0"):
                                    gui.label(c.nome).classes("text-caption t-suave")
                                    gui.label(valores.get(c.indice, "") or "—").classes("text-body2")
        if base and linha is None:
            self._do_ato(base, estado, campos)
        salvar.enable()
        await self._busca_na_tabela(area_busca, estado, colunas, campos)

    def _do_ato(self, base: int, estado: dict, campos: dict) -> None:
        """Nova pessoa num ato já salvo: o formulário vem com o que é do ato (da linha base) e
        os dados da pessoa em branco. Se o ato ainda não tem número na coluna de juntar, as
        duas linhas ganham o próximo número livre (a base é gravada junto ao salvar)."""
        proc = main.proc()
        da_pessoa = proc.colunas_da_pessoa() | {atos.normalizar(proc.config["coluna_nome"])}
        nomes = {c.indice: atos.normalizar(c.nome) for c in estado["colunas"]}
        valores = estado["linhas"].get(base, {})
        primeira = None
        for i, campo in campos.items():
            pessoa = nomes.get(i) in da_pessoa
            _por_valor(campo, "" if pessoa else valores.get(i, ""))
            if pessoa and primeira is None:
                primeira = campo
        juntar = next((i for i, n in nomes.items() if n == proc.coluna_juntar and i in campos), None)
        if juntar is not None and not str(valores.get(juntar, "")).strip():
            usados = [str(v.get(juntar, "")).strip() for v in estado["linhas"].values()]
            numero = str(max((int(u) for u in usados if u.isdigit()), default=0) + 1)
            _por_valor(campos[juntar], numero)
            estado["numero_da_base"] = (base, juntar, numero)
            rotulo = next(c.nome for c in estado["colunas"] if c.indice == juntar)
            with campos[juntar]:  # a janela roda em segundo plano: o aviso precisa de um lugar na tela
                gui.notify(f"{rotulo} {numero} na linha {base} e na pessoa nova: é isso que as coloca no "
                           "mesmo ato", type="info")
        if primeira is not None:
            primeira.run_method("focus")

    async def _busca_na_tabela(self, area, estado: dict, colunas: list, campos: dict) -> None:
        """Processo ligado a uma tabela de consulta: a caixa que acha a pessoa e preenche a linha."""
        consulta = main.proc().config.get("consulta")
        if not consulta:
            return
        await consultas_ui.pedir_atualizacao(self.camada, consulta["tabela"])  # passou do prazo: a tabela nova
        with area, gui.row().classes("w-full items-center gap-2 no-wrap q-pa-sm fundo-cartao rounded") \
                as carregando:  # tabela grande demora: mostra que está vindo
            gui.spinner(size="22px", color="primary")
            gui.label(f"Carregando a tabela {consulta['tabela']} para ajudar a preencher a linha...") \
                .classes("text-body2 t-suave")
        try:
            tabela = await run.io_bound(consultas.ler_cadastrada, consulta["tabela"])
        except Exception as e:
            carregando.delete()
            with area, gui.row().classes("items-center gap-2 no-wrap") as aviso:
                gui.icon("warning", color="orange-8")
                if isinstance(e, (KeyError, FileNotFoundError)):
                    gui.label(f"A tabela {consulta['tabela']} (de onde vêm os dados das pessoas) não está neste "
                              "computador. Escolha o arquivo do Excel dela:").classes("text-caption text-orange-8")
                else:
                    gui.label(f"Não consegui ler a tabela {consulta['tabela']} ({type(e).__name__}): preencha à "
                              "mão ou escolha o arquivo de novo:").classes("text-caption text-orange-8")

                async def achar() -> None:
                    if await consultas_ui.achar_tabela(self.camada, consulta["tabela"]):
                        aviso.delete()
                        await self._busca_na_tabela(area, estado, colunas, campos)

                gui.button("Escolher a tabela...", icon="upload_file", on_click=achar).props("outline dense no-caps")
            return
        carregando.delete()
        por_nome = {atos.normalizar(c.nome): campos[c.indice] for c in colunas if c.indice in campos}
        with area:
            estado["busca"] = consultas_ui.BuscaNaTabela(consulta, tabela, por_nome, _por_valor)

    @staticmethod
    def _explicar_etapa(chave: str, cor: str, texto: str, registro: dict) -> str:
        """O que aconteceu até agora nesta etapa, em frases (aba Status da linha)."""
        data_edocs, data_diario = registro.get("DATAEDOCS", ""), registro.get("DATADIARIO", "")
        if chave == "edocs":
            if cor == VERDE:
                return (f"Enviado com sucesso no e-docs{f' em {data_edocs}' if data_edocs else ''}. "
                        "O documento foi criado com o texto, os vistos, os assinantes e o destinatário, "
                        "e seguiu para visto.")
            if cor == VERMELHO:
                return (f"O envio ao e-docs deu erro: {texto}. Confira no e-docs se o documento chegou "
                        "antes de enviar de novo.")
            return "Ainda não foi enviado ao e-docs. Ele entra na próxima vez que você usar Lançar no e-docs."
        if chave == "diario":
            if cor == VERDE:
                return (f"Cadastrado com sucesso no Diário Oficial"
                        f"{f', para sair na edição de {data_diario}' if data_diario else ''}.")
            if cor == VERMELHO:
                return f"O cadastro no Diário Oficial deu erro: {texto}. Confira no site do Diário."
            if cor == AMARELO:
                return ("Pronto para o Diário Oficial: já foi enviado no e-docs. Ele entra na próxima vez "
                        "que você usar Lançar no Diário.")
            return "Ainda não pode ir para o Diário Oficial: primeiro ele precisa ser enviado no e-docs."
        # publicação
        if cor == VERDE:
            edicao = texto.replace(planilha_status.PUBLICADO, "").strip()
            return (f"Publicado no Diário Oficial{f', na edição {edicao}' if edicao else ''}. O ato já é "
                    "público: não há mais nada a fazer com ele.")
        if cor == VERMELHO:
            if texto == planilha_status.NAO_ENCONTRADO_FINAL:
                return ("Não foi encontrado no Diário Oficial depois de 3 conferências. Procure no site do "
                        "Diário se ele saiu com outro número ou se ficou para outra data.")
            return f"A conferência da publicação deu erro: {texto}."
        if cor == AMARELO:
            feitas = planilha_status.tentativas_feitas(texto)
            if feitas:
                return (f"Ainda não apareceu no Diário Oficial: já foi procurado {feitas} de 3 vezes. A "
                        "próxima conferência olha o Diário do dia útil seguinte.")
            return ("Aguardando a publicação. Depois que o Diário"
                    f"{f' de {data_diario}' if data_diario else ''} sair, use Conferir publicação "
                    "(ou veja o Diário de hoje, que confere sozinho).")
        return "Ainda não foi lançado no Diário Oficial, então não há publicação para conferir."

    def _painel_status(self, janela, linha: int, registro: dict) -> None:
        """Situação de cada etapa da linha, com a opção de desfazer o envio."""
        linha_tabela = next((l for l in self.dados.get("linhas", []) if l["id"] == linha), {})
        etapas = [("e-docs", "edocs", registro.get("DATAEDOCS", "")),
                  ("Diário Oficial", "diario", registro.get("DATADIARIO", "")),
                  ("Publicação", "pub", "")]
        with gui.column().classes("w-full gap-3 min-h-[400px]"):
            for nome, chave, data in etapas:
                cor = linha_tabela.get(f"{chave}_cor", BRANCO)
                with gui.row().classes("w-full items-center gap-3 no-wrap fundo-cartao rounded q-pa-sm"):
                    gui.html(f'<span class="bolinha b-{cor}"></span>')
                    with gui.column().classes("gap-0 col"):
                        gui.label(nome).classes("text-subtitle2 t-forte")
                        texto = linha_tabela.get(f"{chave}_txt") or ""
                        gui.label(self._explicar_etapa(chave, cor, texto, registro)).classes("text-body2 t-forte")
                        # O que está gravado de fato na planilha (sem os textos de "pendente")
                        gravado = {"edocs": registro.get("STATUSEDOCS") or registro.get("STATUS", ""),
                                   "diario": registro.get("STATUSDIARIO", ""),
                                   "pub": registro.get("STATUSPUBLICADO", "")}[chave]
                        if gravado:
                            gui.label(f"Na planilha: {gravado}" + (f" · {data}" if data else "")) \
                                .classes("text-caption t-fraco")
                    if linha_tabela.get("publicado") and chave in ("edocs", "diario"):
                        with gui.row().classes("items-center gap-1 no-wrap"):
                            gui.icon("lock", color="positive", size="18px")
                            gui.label("Travado: já publicado").classes("text-caption t-suave")
                        continue
                    if chave == "edocs" and linha_tabela.get("tem_edocs") and linha_tabela.get("tem_diario"):
                        with gui.row().classes("items-center gap-1 no-wrap"):
                            gui.icon("arrow_downward", color="grey-6", size="18px")
                            gui.label("Para desfazer, desfaça primeiro o Diário").classes("text-caption t-suave")
                    elif chave == "edocs" and linha_tabela.get("tem_edocs"):
                        gui.button("Desfazer envio no e-docs", icon="undo", color="orange-8",
                                   on_click=lambda: self._confirmar_desfazer(linha, "edocs", janela)) \
                            .props("flat no-caps")
                    if chave == "diario" and linha_tabela.get("tem_diario"):
                        gui.button("Desfazer envio no Diário", icon="undo", color="orange-8",
                                   on_click=lambda: self._confirmar_desfazer(linha, "diario", janela)) \
                            .props("flat no-caps")
            if linha_tabela.get("publicado"):
                gui.label("Este ato já saiu no Diário Oficial: é público, então os status do e-docs e do "
                              "Diário não podem mais ser desfeitos.").classes("text-caption t-forte")
            else:
                gui.label("Desfazer só muda a planilha (apaga o status e a data): nada é apagado nos sites. "
                          "As etapas se desfazem de trás para frente: primeiro o Diário, depois o e-docs.") \
                    .classes("text-caption t-suave")

    def _ato_inteiro(self, registro: dict) -> dict:
        """O ato da linha, com as outras pessoas que entram nele (e esta linha como está)."""
        proc = main.proc()
        if proc.coluna_juntar and "_LINHA" in registro:
            todos = [registro if r["_LINHA"] == registro["_LINHA"] else r
                     for r in self.dados.get("registros", {}).values()]
            return proc.ato_da_linha(todos, registro["_LINHA"]) or registro
        return registro

    async def _mostrar_ato(self, area, registro: dict, altura: int = 400) -> None:
        proc = main.proc()
        registro = self._ato_inteiro(registro)
        try:
            texto = (await run.io_bound(proc.textos, [registro]))[0]
            pagina = (await run.io_bound(proc.folhas, [registro]))[0]
        except Exception as e:
            texto, pagina = f"Não consegui montar o ato: {type(e).__name__}", ""
        area.clear()
        with area:
            if pagina:
                with gui.row().classes("w-full justify-end"):
                    botao_word = gui.button("Abrir no Word (imprimir)", icon="print") \
                        .props("flat dense no-caps").tooltip("Abre uma cópia deste ato no Word, para imprimir")
                    botao_word.on_click(lambda: self._abrir_no_word(
                        lambda: main.word_do_ato(registro, proc), botao_word))
            if "[" in texto:
                faltam = sorted(set(re.findall(r"\[[^\]]+\]", texto)))
                gui.label("Campos vazios: " + ", ".join(faltam)).classes("text-negative text-caption")
            if pagina:
                gui.html(f'<div class="folha-area" style="height:{altura}px">{pagina}</div>').classes("w-full")
            else:
                with gui.scroll_area().classes("w-full border rounded q-pa-sm").style(f"height:{altura}px"):
                    gui.label(texto).classes("previa")

    async def _salvar_linha(self, dialogo, botao, linha: int | None, colunas: list, campos: dict) -> None:
        valores = {i: str(c.value or "").strip() for i, c in campos.items()}
        if linha is None and not any(valores.values()):
            gui.notify("Preencha pelo menos um campo", type="warning")
            return
        if not self._obrigatorios_ok(colunas, valores, campos):
            return
        planilha, aba = Path(self.dados["planilha"]), main.proc().aba
        botao.props("loading")

        def gravar() -> tuple[int, bool]:
            # Com o Excel: grava, e ele recalcula só a linha (rápido). Sem: marca para
            # recalcular quando a planilha for aberta no Excel.
            excel = edicao.tem_excel()
            if linha is None:
                numero = edicao.incluir_linha(planilha, aba, colunas, valores, recalcular_ao_abrir=not excel)
            else:
                edicao.editar_linha(planilha, aba, linha, colunas, valores, recalcular_ao_abrir=not excel)
                numero = linha
            recalculou = excel and edicao.recalcular_no_excel(planilha, aba, [numero])
            if excel and not recalculou:
                edicao.marcar_recalculo(planilha)
            return numero, recalculou

        try:
            numero, recalculou = await self._quando_livre(planilha, gravar)
        except edicao.PlanilhaAberta:  # a pessoa cancelou a espera
            botao.props(remove="loading")
            return
        except Exception as e:
            botao.props(remove="loading")
            gui.notify(f"Não consegui salvar: {type(e).__name__} {e}".strip(), type="negative")
            return
        dialogo.close()
        gui.notify(f"Linha {numero} {'incluída' if linha is None else 'salva'}"
                   + ("" if recalculou else " · os campos calculados se atualizam ao abrir no Excel"),
                   type="positive")
        await self.recarregar()

    @staticmethod
    def _obrigatorios_ok(colunas: list, valores: dict, campos: dict | None = None, quem: str = "") -> bool:
        """Os campos obrigatórios do processo estão preenchidos? Se não, avisa quais e deixa
        os campos vazios em vermelho."""
        obrigatorias = {atos.normalizar(c) for c in main.proc().obrigatorios}
        vazias = [c for c in colunas if c.editavel and atos.normalizar(c.nome) in obrigatorias
                  and not str(valores.get(c.indice, "")).strip()]
        for c in colunas:
            if campos and c.indice in campos and atos.normalizar(c.nome) in obrigatorias:
                if c in vazias:
                    campos[c.indice].props('error error-message="Obrigatório"')
                else:
                    campos[c.indice].props(remove="error error-message")
        if vazias:
            onde = f" n{quem}" if quem else ""
            gui.notify(f"Campo(s) obrigatório(s) vazio(s){onde}: " + ", ".join(c.nome for c in vazias),
                       type="warning")
        return not vazias

    async def _salvar_novas(self, dialogo, botao, colunas: list, lista: list[dict], campos: dict | None = None,
                            numero_da_base: tuple | None = None) -> None:
        """Inclui todas as linhas novas de uma vez (uma gravação e um recálculo no Excel).
        numero_da_base (linha, coluna, número): o ato já salvo que ainda não tinha número
        ganha o mesmo das pessoas novas."""
        if not lista:
            gui.notify("Preencha pelo menos um campo", type="warning")
            return
        formulario = {i: str(c.value or "").strip() for i, c in (campos or {}).items()}
        for n, valores in enumerate(lista, 1):
            # a do formulário (a última) fica com os campos vazios em vermelho
            do_formulario = n == len(lista) and valores == formulario
            if not self._obrigatorios_ok(colunas, valores, campos if do_formulario else None,
                                         "" if do_formulario or len(lista) == 1 else f"a {n}ª linha da lista"):
                return
        planilha, aba = Path(self.dados["planilha"]), main.proc().aba
        botao.props("loading")

        def gravar() -> tuple[list[int], bool]:
            excel = edicao.tem_excel()
            if numero_da_base:
                base, coluna, numero = numero_da_base
                edicao.editar_linha(planilha, aba, base, colunas, {coluna: numero}, recalcular_ao_abrir=not excel)
            numeros = edicao.incluir_linhas(planilha, aba, colunas, lista, recalcular_ao_abrir=not excel)
            recalculou = excel and edicao.recalcular_no_excel(planilha, aba, numeros)
            if excel and not recalculou:
                edicao.marcar_recalculo(planilha)
            return numeros, recalculou

        try:
            numeros, recalculou = await self._quando_livre(planilha, gravar)
        except edicao.PlanilhaAberta:  # a pessoa cancelou a espera
            botao.props(remove="loading")
            return
        except Exception as e:
            botao.props(remove="loading")
            gui.notify(f"Não consegui salvar: {type(e).__name__} {e}".strip(), type="negative")
            return
        dialogo.close()
        quais = f"Linha {numeros[0]}" if len(numeros) == 1 else f"{len(numeros)} linhas ({numeros[0]} a {numeros[-1]})"
        gui.notify(f"{quais} incluída(s)"
                   + ("" if recalculou else " · os campos calculados se atualizam ao abrir no Excel"),
                   type="positive")
        await self.recarregar()

    async def _quando_livre(self, planilha: Path, tarefa):
        """Executa tarefa() (em segundo plano) quando a planilha não estiver aberta no Excel.
        Se estiver (aqui ou por outra pessoa na rede), mostra "aguardando" e continua sozinho
        quando ela for fechada. Lança PlanilhaAberta se a pessoa cancelar a espera."""
        while True:
            quem = await run.io_bound(planilha_status.em_uso, planilha)
            if quem is not None and not await self._esperar_fechar(planilha, quem):
                raise edicao.PlanilhaAberta(quem)
            try:
                return await run.io_bound(tarefa)
            except edicao.PlanilhaAberta:
                continue  # abriram no meio do caminho: espera de novo

    async def _esperar_fechar(self, planilha: Path, quem: str) -> bool:
        """Janela "aguardando a planilha ser fechada". True quando fechar; False se cancelar."""
        cancelou = asyncio.Event()
        with self.camada, gui.dialog().props("persistent") as dialogo, \
                gui.card().classes("cartao w-[460px] q-pa-lg"):
            with gui.row().classes("items-center gap-3 no-wrap"):
                gui.spinner(size="md")
                gui.label("Aguardando a planilha ser fechada").classes("text-h6")
            gui.label(f"A planilha está aberta no Excel{f' por {quem}' if quem else ''}. "
                      "Assim que ela for fechada, o Edodocs salva sozinho — sem risco de um "
                      "sobrescrever o outro.").classes("text-body2")
            with gui.row().classes("w-full justify-end"):
                gui.button("Cancelar", on_click=lambda: (cancelou.set(), dialogo.close())) \
                    .props("flat no-caps")
        dialogo.open()
        while not cancelou.is_set():
            await asyncio.sleep(2)
            if await run.io_bound(planilha_status.em_uso, planilha) is None:
                dialogo.close()
                return True
        return False

    def _confirmar_exclusao_varias(self) -> None:
        """Exclui da planilha as linhas marcadas na tabela (o grupo de um ato: todas as dele)."""
        if self.rodando:
            gui.notify("Aguarde o fim da execução para mexer na planilha", type="warning")
            return
        linhas = sorted(self._linhas_marcadas())
        if not linhas:
            return
        registros = self.dados.get("registros", {})
        enviadas = [l for l in linhas if (r := registros.get(l)) and planilha_status.ENVIADO in (
            r.get(main.preparado_chave_edocs([r]), ""), r.get("STATUSDIARIO", ""))]
        nome = atos.normalizar(main.proc().config["coluna_nome"])
        with self.camada, gui.dialog() as dialogo, gui.card().classes("cartao w-[480px] q-pa-lg gap-2"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("delete", color="negative", size="26px")
                gui.label(f"Excluir {len(linhas)} linha(s)?").classes("text-h6")
            with gui.element("div").classes("w-full").style("max-height: 30vh; overflow-y: auto"):
                for l in linhas:
                    gui.label(f"Linha {l} · {(registros.get(l) or {}).get(nome, '')}").classes("text-body2 t-forte")
            gui.label("As linhas saem da planilha. Não dá para desfazer pelo programa.").classes("text-body2")
            if enviadas:
                gui.label(f"Atenção: {len(enviadas)} delas já foram enviadas. Excluir a linha não cancela o envio.")                     .classes("text-body2 text-negative")
            with gui.row().classes("w-full justify-end gap-2 q-mt-sm"):
                gui.button("Cancelar", on_click=dialogo.close).props("flat no-caps")
                excluir = gui.button(f"Excluir {len(linhas)}", icon="delete", color="negative")                     .props("unelevated no-caps")

                async def confirmar() -> None:
                    excluir.props("loading")
                    planilha, aba = Path(self.dados["planilha"]), main.proc().aba
                    try:
                        colunas, _ = await self._ler_estrutura()
                        resultado = await self._quando_livre(
                            planilha, lambda: edicao.excluir_linhas(planilha, aba, linhas, colunas))
                    except edicao.PlanilhaAberta:  # a pessoa cancelou a espera
                        excluir.props(remove="loading")
                        return
                    except Exception as e:
                        excluir.props(remove="loading")
                        gui.notify(f"Não consegui excluir: {type(e).__name__}", type="negative")
                        return
                    dialogo.close()
                    self._limpar_selecao()
                    gui.notify(f"{len(linhas)} linha(s) excluída(s)" if resultado == "excluida" else
                               f"{len(linhas)} linha(s) apagada(s) (sem o Excel, as fórmulas ficam vazias)",
                               type="positive")
                    await self.recarregar()

                excluir.on_click(confirmar)
        dialogo.open()

    def _confirmar_exclusao(self, janela, linha: int, colunas: list, registro: dict | None) -> None:
        enviado = registro and planilha_status.ENVIADO in (
            registro.get(main.preparado_chave_edocs([registro]), ""), registro.get("STATUSDIARIO", ""))
        with self.camada, gui.dialog() as dialogo, gui.card().classes("cartao w-[440px] q-pa-lg"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("delete", color="negative", size="26px")
                gui.label(f"Excluir a linha {linha}?").classes("text-h6")
            gui.label("A linha sai da planilha. Não dá para desfazer pelo programa.").classes("text-body2")
            if enviado:
                gui.label("Atenção: este ato já foi enviado. Excluir a linha não cancela o envio.") \
                    .classes("text-body2 text-negative")
            with gui.row().classes("w-full justify-end gap-2 q-mt-sm"):
                gui.button("Cancelar", on_click=dialogo.close).props("flat no-caps")
                excluir = gui.button("Excluir", icon="delete", color="negative").props("unelevated no-caps")

                async def confirmar() -> None:
                    excluir.props("loading")
                    planilha, aba = Path(self.dados["planilha"]), main.proc().aba
                    try:
                        resultado = await self._quando_livre(
                            planilha, lambda: edicao.excluir_linha(planilha, aba, linha, colunas))
                    except edicao.PlanilhaAberta:  # a pessoa cancelou a espera
                        excluir.props(remove="loading")
                        return
                    except Exception as e:
                        excluir.props(remove="loading")
                        gui.notify(f"Não consegui excluir: {type(e).__name__}", type="negative")
                        return
                    dialogo.close()
                    janela.close()
                    gui.notify(f"Linha {linha} excluída" if resultado == "excluida" else
                               f"Linha {linha} apagada (sem o Excel, as fórmulas da linha ficam vazias)",
                               type="positive")
                    await self.recarregar()

                excluir.on_click(confirmar)
        dialogo.open()

    # --- Página Diário Oficial ---
    def _painel_diario(self) -> None:
        with gui.row().classes("w-full items-center"):
            gui.label("Edições recentes do Diário Oficial de Uberlândia. Baixe para conferir as "
                      "publicações ou abra o PDF.").classes("text-body1 t-forte")
            gui.space()
            gui.button("Atualizar", icon="refresh", on_click=self._carregar_diario).props("flat no-caps")
            gui.button("Abrir pasta", icon="folder_open", on_click=lambda: abrir(main.DIARIO)) \
                .props("outline no-caps")
        # Diário de hoje: baixado sozinho na primeira abertura do dia
        with gui.card().classes("cartao cartao-reto w-full q-pa-md q-mt-sm").style("border-left:4px solid #2459a8"):
            self.area_hoje = gui.column().classes("w-full gap-1")
        gui.label("Edições recentes").classes("text-subtitle1 text-weight-medium t-forte q-mt-md")
        self.area_diario = gui.column().classes("w-full")

    async def _diario_de_hoje(self, baixar: bool = True) -> None:
        """Mostra o Diário de hoje (e baixa, se for a primeira abertura do dia)."""
        if not hasattr(self, "area_hoje"):
            return
        # Troca de processo no meio da conferência: só a última chamada desenha
        self._vez_hoje = getattr(self, "_vez_hoje", 0) + 1
        vez = self._vez_hoje
        hoje = date.today()
        self.area_hoje.clear()
        with self.area_hoje:
            with gui.row().classes("items-center gap-2"):
                gui.icon("today", size="26px").classes("titulo-icone")
                gui.label(f"Diário de hoje · {hoje:%d/%m/%Y}").classes("text-h6 t-forte")
                gui.space()
            with gui.row().classes("items-center gap-2"):
                gui.spinner()
                espera = gui.label("Procurando o Diário de hoje...").classes("t-suave")
        try:
            resultado = await run.io_bound(diario_de_hoje, baixar)
        except Exception as e:
            resultado = {"arquivos": [], "erro": f"{type(e).__name__}"}
        arquivos = resultado["arquivos"]
        nome = main.processo.name
        conferido = {"atos": {}, "nao_gravou": ""}
        if arquivos:
            espera.text = f"Conferindo os atos de {nome} no Diário de hoje..."
            # Usa a planilha que a tela já leu (se for a deste processo), em vez de ler de novo
            planilha = str(_planilha_do_processo())
            for _ in range(60):
                if self.dados and self.dados.get("planilha") == planilha or vez != self._vez_hoje:
                    break
                await asyncio.sleep(0.5)
            if vez != self._vez_hoje:
                return
            registros = (list(self.dados["registros"].values())
                         if self.dados and self.dados.get("planilha") == planilha else None)
            if not main.proc().vai_ao_diario:
                registros = []  # não vai ao Diário: nada a conferir nem a marcar
            try:
                conferido = await run.io_bound(conferir_diario_de_hoje, arquivos, registros)
            except Exception as e:
                conferido["nao_gravou"] = f"não consegui conferir ({type(e).__name__})"
        if vez != self._vez_hoje:
            return
        atos = conferido["atos"]
        novos = sum(sit == "novo" for lista in atos.values() for _, sit in lista)
        if novos and not conferido["nao_gravou"]:
            gui.notify(f"{novos} ato(s) de {nome} marcado(s) como publicado(s) na planilha",
                       type="positive")
            background_tasks.create(self.recarregar())

        self.area_hoje.clear()
        with self.area_hoje:
            with gui.row().classes("w-full items-center gap-2"):
                gui.icon("today", size="26px").classes("titulo-icone")
                gui.label(f"Diário de hoje · {hoje:%d/%m/%Y}").classes("text-h6 t-forte")
                gui.space()
                gui.button("Verificar de novo", icon="refresh",
                           on_click=lambda: self._diario_de_hoje(True)).props("flat dense no-caps")
            if "erro" in resultado:
                with gui.row().classes("items-center gap-2"):
                    gui.icon("cloud_off", color="warning")
                    gui.label(f"Não consegui acessar o site do Diário ({resultado['erro']}). "
                              "Confira a internet e clique em Verificar de novo.").classes("t-forte")
                return
            if not arquivos:
                with gui.row().classes("items-center gap-2"):
                    gui.icon("schedule", color="warning")
                    gui.label("O Diário de hoje ainda não foi publicado no site. O Edodocs tenta de "
                              "novo quando for aberto (ou clique em Verificar de novo).").classes("t-forte")
                return
            for pdf in arquivos:
                numero = pdf.stem.split("_ed")[-1]
                with gui.row().classes("w-full items-center gap-3 no-wrap"):
                    gui.icon("picture_as_pdf", size="26px").classes("icone-pdf")
                    with gui.column().classes("gap-0 col"):
                        gui.label(f"Edição {numero}").classes("text-subtitle1 text-weight-medium t-forte")
                        achados = atos.get(pdf, [])
                        if not achados:
                            gui.label(f"Nenhum ato de {nome} nesta edição").classes("text-body2 t-suave")
                        else:
                            gui.label(f"{len(achados)} ato(s) de {nome} nesta edição:") \
                                .classes("text-body2 t-suave")
                    gui.button("Abrir", icon="open_in_new", on_click=lambda p=pdf: abrir(p)) \
                        .props("unelevated no-caps")
                if achados:
                    with gui.row().classes("w-full gap-1 q-pl-xl"):
                        for descricao, situacao in achados:
                            novo = situacao == "novo"
                            # O processo entre parênteses: dá para ver de qual processo é cada ato
                            gui.chip(f"{descricao} ({nome})", icon="new_releases" if novo else "check_circle",
                                     color=None).classes("chip-campo").props("dense") \
                                .tooltip("Marcado agora como publicado na planilha" if novo and
                                         not conferido["nao_gravou"] else
                                         "Achado, mas não gravado na planilha" if novo else
                                         "Já estava marcado como publicado na planilha")
            if conferido["nao_gravou"]:
                with gui.row().classes("items-center gap-2 q-mt-xs"):
                    gui.icon("warning", color="warning")
                    gui.label(f"Os atos achados não foram marcados na planilha: {conferido['nao_gravou']}. "
                              "Feche a planilha e clique em Verificar de novo.").classes("text-body2 t-forte")
            gui.label("Os atos achados são marcados como publicados na planilha. O PDF fica na pasta "
                      "entradas/Diario_hoje e é trocado pelo do dia seguinte (não acumula). Para outros "
                      "dias, use Conferir publicação em Atos ou baixe as edições abaixo.") \
                .classes("text-caption t-fraco")

    async def _carregar_diario(self) -> None:
        import diario

        self.diario_carregado = True
        self.area_diario.clear()
        with self.area_diario:
            with gui.row().classes("q-pa-lg items-center gap-3"):
                gui.spinner(size="lg")
                gui.label("Buscando as edições no site do Diário...").classes("t-suave")
        try:
            edicoes = await run.io_bound(diario.edicoes_no_site)
        except Exception as e:
            self.diario_carregado = False
            self.area_diario.clear()
            with self.area_diario, gui.row().classes("q-pa-lg items-center gap-3"):
                gui.icon("cloud_off", color="warning", size="32px")
                gui.label(f"Não consegui acessar o site do Diário ({type(e).__name__}). "
                          "Confira a internet e clique em Atualizar.")
            return

        # Atos deste processo publicados em cada edição (pela conferência)
        por_edicao: dict[str, int] = {}
        for r in self.dados.get("registros", {}).values():
            m = re.search(r"oficial\s+(\S+)\s*\(", str(r.get("STATUSPUBLICADO", "")))
            if str(r.get("STATUSPUBLICADO", "")).startswith(planilha_status.PUBLICADO) and m:
                por_edicao[m.group(1)] = por_edicao.get(m.group(1), 0) + 1

        self.area_diario.clear()
        with self.area_diario, gui.card().classes("cartao cartao-reto w-full q-pa-none"):
            with gui.row().classes("w-full q-px-md q-py-sm text-caption t-suave text-weight-medium"):
                gui.label("Data").classes("w-28")
                gui.label("Edição").classes("w-28")
                gui.label(f"Atos de {main.processo.name}").classes("col")
                gui.label("PDF").classes("w-56 text-right")
            for dia, numero, id_pdf in edicoes:
                gui.separator()
                destino = main.DIARIO / f"Diario_Oficial_{dia:%Y-%m-%d}_ed{numero}.pdf"
                with gui.row().classes("w-full items-center q-px-md q-py-xs no-wrap"):
                    gui.label(f"{dia:%d/%m/%Y}").classes("w-28 text-body2")
                    gui.label(numero).classes("w-28 text-body2 text-weight-medium")
                    qtd = por_edicao.get(numero, 0)
                    with gui.row().classes("col items-center gap-1"):
                        if qtd:
                            _bolinha(VERDE)
                            gui.label(f"{qtd} publicado(s)").classes("text-body2")
                        else:
                            gui.label("—").classes("t-fraco")
                    with gui.row().classes("w-56 justify-end gap-1"):
                        self._botoes_pdf(destino, f"{diario.SITE}/diario-oficial/versao-pdf/{id_pdf}")

    def _botoes_pdf(self, destino: Path, url: str) -> None:
        """Abrir e apagar (se já baixado) ou baixar o PDF da edição. Só esta linha se redesenha
        depois de baixar ou apagar (sem buscar a lista no site de novo: é rápido)."""
        area = gui.row().classes("items-center gap-0 no-wrap justify-end")
        self._desenhar_botoes_pdf(area, destino, url)

    def _desenhar_botoes_pdf(self, area, destino: Path, url: str) -> None:
        import diario

        area.clear()
        de_hoje = main.DIARIO_HOJE / destino.name  # o do dia, baixado ao abrir o app
        arquivo = destino if destino.exists() else de_hoje if de_hoje.exists() else None

        def apagar() -> None:
            try:
                arquivo.unlink()
            except OSError:
                gui.notify("Não consegui apagar: o PDF está aberto. Feche-o e tente de novo.",
                           type="warning")
                return
            gui.notify(f"Apagado: {arquivo.name}", type="info")
            self._desenhar_botoes_pdf(area, destino, url)

        async def baixar() -> None:
            botao.props("loading")
            try:
                conteudo = await run.io_bound(diario._baixar, url)
                destino.parent.mkdir(parents=True, exist_ok=True)
                destino.write_bytes(conteudo)
            except Exception as e:
                botao.props(remove="loading")
                gui.notify(f"Não consegui baixar: {type(e).__name__}", type="negative")
                return
            gui.notify(f"Baixado: {destino.name}", type="positive")
            abrir(destino)
            self._desenhar_botoes_pdf(area, destino, url)

        with area:
            if arquivo:
                gui.button("Abrir", icon="picture_as_pdf", on_click=lambda: abrir(arquivo)) \
                    .props("flat dense no-caps")
                gui.button(icon="delete_outline", on_click=apagar) \
                    .props("flat round dense size=sm").classes("t-fraco").tooltip("Apagar o PDF baixado")
                return
            botao = gui.button("Baixar", icon="download", on_click=baixar).props("flat dense no-caps")

    # --- Aba Execução ---
    def _painel_execucao(self) -> None:
        # A página não rola: ocupa a altura que sobra na janela (tirando cabeçalho, título da
        # página e margens) e só o quadro do registro, que cresce durante a execução, rola.
        with gui.column().classes("w-full no-wrap gap-0").style(f"height: calc(100vh - {ALTURA_FIXA}px)"):
            self._quadros_execucao()

    def _quadros_execucao(self) -> None:
        with gui.card().classes("cartao cartao-reto w-full q-px-md q-py-sm"):
            with gui.row().classes("w-full items-center"):
                self.exec_titulo = gui.label("Nada em execução").classes("text-h6")
                gui.space()
                self.exec_ocupado = gui.row().classes("items-center gap-2 text-primary")
                self.exec_cancelar = gui.button("Cancelar execução", icon="stop_circle", color="negative",
                                                on_click=self._cancelar_execucao).props("flat no-caps")
                self.exec_cancelar.set_visibility(False)
            self.exec_ato = gui.label("").classes("text-body2 t-suave")
            self.exec_barra = gui.linear_progress(value=0, show_value=False, size="10px") \
                .props("rounded color=primary").classes("q-my-sm")
            self.exec_barra.set_visibility(False)
            # Para não parecer travado: quanto tempo desde a última novidade
            self.exec_pulso = gui.label("").classes("text-caption t-fraco")
        with gui.card().classes("cartao cartao-reto w-full q-mt-sm q-pa-none col gap-0").style("min-height:0"):
            # Barrinha fina e animada enquanto algo roda (sinal discreto de que está trabalhando)
            self.exec_atividade = gui.linear_progress(show_value=False, size="3px") \
                .props("indeterminate color=primary").classes("w-full")
            self.exec_atividade.set_visibility(False)
            self.log_area = gui.scroll_area().classes("w-full col")
            with self.log_area:
                self.log = gui.column().classes("w-full q-pa-md gap-1")
            # Só depois que a execução termina: voltar para a página Atos
            with gui.row().classes("w-full justify-center q-pa-sm") as self.exec_voltar:
                gui.button("Voltar para Atos", icon="arrow_back", on_click=lambda: self.ir("atos")) \
                    .props("unelevated no-caps")
            self.exec_voltar.set_visibility(False)

    def _ocupado(self, texto: str) -> None:
        """O que está sendo feito agora. Sem nada específico, mas rodando: "Trabalhando..."."""
        self.exec_ocupado.clear()
        if texto:
            with self.exec_ocupado:
                gui.spinner(size="sm")
                gui.label(texto)
        elif self.rodando:
            with self.exec_ocupado:
                gui.spinner("dots", size="sm")
                gui.label("Trabalhando...").classes("t-suave")

    def _pulso(self) -> None:
        """A cada segundo: se faz tempo que não chega novidade, mostra há quanto tempo (para
        a pessoa ver que não travou). Some quando chega uma mensagem ou abre uma pergunta."""
        if not self.rodando or any(not f.done() for _, f in self._pendentes):
            self.exec_pulso.text = ""
            return
        parado = int(time.monotonic() - self._ultimo_evento)
        self.exec_pulso.text = f"Ainda trabalhando · última atualização há {parado} s" if parado >= 4 else ""

    def _registrar(self, tipo: str, texto: str) -> None:
        icones = {"ok": ("check_circle", "positive"), "erro": ("cancel", "negative"),
                  "aviso": ("warning", "warning"), "info": ("info", "primary")}
        with self.log:
            if tipo == "titulo":
                with gui.row().classes("w-full items-center gap-2 q-mt-sm"):
                    gui.separator().classes("col")
                    gui.label(texto).classes("text-caption text-weight-bold t-suave")
                    gui.separator().classes("col")
            else:
                icone, cor = icones[tipo]
                with gui.row().classes("w-full items-start gap-2 no-wrap log-linha"):
                    gui.label(datetime.now().strftime("%H:%M:%S")).classes("t-fraco text-caption q-mt-xs")
                    gui.icon(icone, color=cor, size="20px")
                    gui.label(texto).classes("text-negative" if tipo == "erro" else "")
        self.log_area.scroll_to(percent=1.0)

    # --- Rodar um fluxo ---
    def iniciar(self, nome: str, funcao, subtitulo: str | None = None) -> None:
        if self.rodando:
            return
        self.rodando = True
        for b in self.botoes:
            b.disable()
        self.sel_processo.disable()
        self.log.clear()
        self.exec_titulo.text = nome
        self.exec_ato.text = subtitulo if subtitulo is not None else f"Processo: {main.processo.name}"
        self.exec_barra.value = 0
        self.exec_barra.set_visibility(False)
        self.motor.cancelar = False
        self._ultimo_evento = time.monotonic()
        self.exec_atividade.set_visibility(True)
        self._ocupado("")
        self.exec_cancelar.set_text("Cancelar execução")
        self.exec_cancelar.enable()
        self.exec_cancelar.set_visibility(True)
        self.exec_voltar.set_visibility(False)
        self.ir("execucao")
        threading.Thread(target=self._rodar, args=(funcao,), daemon=True).start()

    def _rodar(self, funcao) -> None:
        try:
            funcao()
        except KeyboardInterrupt:  # cancelado pela pessoa (botão ou "Cancelar" numa pergunta)
            ui.aviso("Execução cancelada")
        except Exception as e:
            ui.erro(f"Erro inesperado: {type(e).__name__} {e}")
        finally:
            self.motor._avisar("fim")

    def _cancelar_execucao(self) -> None:
        """Cancela no próximo ponto seguro. Se houver uma pergunta aberta, fecha e para já."""
        abertas = [(d, f) for d, f in self._pendentes if not f.done()]
        for dialogo, futuro in abertas:
            dialogo.close()
            futuro.set_exception(KeyboardInterrupt())
        self._pendentes.clear()
        if not abertas:
            self.motor.cancelar = True
            self._registrar("aviso", "Cancelando: o ato em andamento termina antes de parar "
                                     "(para não ficar enviado pela metade)")
        self.exec_cancelar.set_text("Cancelando...")
        self.exec_cancelar.disable()

    async def _fim(self) -> None:
        self.rodando = False
        self.motor.cancelar = False
        self._pendentes.clear()
        self.exec_cancelar.set_visibility(False)
        self.exec_atividade.set_visibility(False)
        self.exec_pulso.text = ""
        self._habilitar_etapas()
        self.sel_processo.enable()
        self._ocupado("")
        self.exec_titulo.text += " · concluído"
        self._registrar("titulo", "Fim")
        self.exec_voltar.set_visibility(True)
        if hasattr(self, "ato_manual"):
            self.ato_manual.depois_da_execucao()
        await self.recarregar()
        # O fluxo apaga o usuário se o login for recusado: atualiza os indicadores
        for site in SITES:
            if not credenciais.carregar(main.CREDENCIAIS, site):
                self.login[site] = (BRANCO, "", "Nenhum usuário cadastrado")
        self._desenhar_logins()

    async def _bombear(self) -> None:
        """Atende o que os fluxos mandaram (mensagens e perguntas)."""
        for _ in range(100):
            try:
                evento, d = self.motor.fila.get_nowait()
            except queue.Empty:
                return
            self._ultimo_evento = time.monotonic()
            if evento == "mensagem":
                self._registrar(d["tipo"], d["texto"])
            elif evento == "ocupado":
                self._ocupado(d["texto"])
            elif evento == "progresso":
                self.exec_barra.set_visibility(True)
                self.exec_barra.value = d["n"] / d["total"]
                self.exec_ato.text = f"Ato {d['n']} de {d['total']} · {d['descricao']}"
            elif evento == "fim":
                await self._fim()
            else:
                getattr(self, f"_dialogo_{evento}")(**d)
                self._pendentes.append((self._ultimo_dialogo, d["futuro"]))

    # --- Caixas de diálogo (respondem ao fluxo pelo futuro) ---
    def _dialogo(self, largura: str = "w-[520px]"):
        dialogo = gui.dialog().props("persistent")
        with self.camada, dialogo:
            cartao = gui.card().classes(f"cartao {largura} max-w-full q-pa-lg")
        self._ultimo_dialogo = dialogo
        return dialogo, cartao

    def _cancelar_pergunta(self, dialogo, futuro: Future) -> None:
        """'Cancelar' numa pergunta: a execução para ali (nada é enviado a partir dela)."""
        dialogo.close()
        if not futuro.done():
            futuro.set_exception(KeyboardInterrupt())

    @staticmethod
    def _responder(dialogo, futuro: Future, valor) -> None:
        dialogo.close()
        if not futuro.done():
            futuro.set_result(valor)

    def _dialogo_escolher(self, pergunta, opcoes, multiplo, vazio, marcados, futuro) -> None:
        dialogo, cartao = self._dialogo()
        ordem: list[int] = list(marcados)
        with cartao:
            gui.label(pergunta).classes("text-h6")
            if not multiplo:
                for i, opcao in enumerate(opcoes):
                    gui.button(opcao, on_click=lambda i=i: self._responder(dialogo, futuro, [i])) \
                        .props("outline no-caps align=left").classes("w-full")
                with gui.row().classes("w-full justify-end"):
                    gui.button("Cancelar", on_click=lambda: self._cancelar_pergunta(dialogo, futuro)) \
                        .props("flat no-caps")
                dialogo.open()
                return
            gui.label("Marque na ordem desejada." + (f" Sem marcar ninguém: {vazio}." if vazio else "")
                      + (" Já vem marcada a escolha da última vez." if marcados else "")) \
                .classes("text-caption t-suave")
            caixas = []

            def marcar(i: int, marcado: bool) -> None:
                if marcado and i not in ordem:
                    ordem.append(i)
                elif not marcado and i in ordem:
                    ordem.remove(i)
                for j, caixa in enumerate(caixas):
                    caixa.text = (f"{ordem.index(j) + 1}º · " if j in ordem else "") + opcoes[j]
                confirmar.set_enabled(bool(ordem) or vazio is not None)

            with gui.scroll_area().classes("max-h-[360px] w-full"):
                for i, opcao in enumerate(opcoes):
                    rotulo = (f"{ordem.index(i) + 1}º · " if i in ordem else "") + opcao
                    caixas.append(gui.checkbox(rotulo, value=i in ordem,
                                               on_change=lambda e, i=i: marcar(i, e.value)))
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=lambda: self._cancelar_pergunta(dialogo, futuro)) \
                    .props("flat no-caps")
                confirmar = gui.button("Confirmar", icon="check",
                                       on_click=lambda: self._responder(dialogo, futuro, list(ordem)))
                confirmar.set_enabled(bool(ordem) or vazio is not None)
        dialogo.open()

    def _dialogo_escolher_pessoas(self, pergunta, pesquisar, recentes, marcados, vazio,
                                  futuro) -> None:
        dialogo, cartao = self._dialogo("w-[640px]")
        montar_escolha_pessoas(cartao, pergunta, pesquisar, recentes, marcados, vazio,
                               lambda lista: self._responder(dialogo, futuro, lista),
                               lambda: self._cancelar_pergunta(dialogo, futuro))
        dialogo.open()

    def _dialogo_perguntar(self, pergunta, senha, futuro) -> None:
        dialogo, cartao = self._dialogo()
        with cartao:
            gui.label(pergunta).classes("text-body1")
            campo = gui.input(password=senha, password_toggle_button=senha).props("outlined autofocus") \
                .classes("w-full")
            campo.on("keydown.enter", lambda: self._responder(dialogo, futuro, campo.value.strip()))
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=lambda: self._cancelar_pergunta(dialogo, futuro)) \
                    .props("flat no-caps")
                gui.button("OK", on_click=lambda: self._responder(dialogo, futuro, campo.value.strip()))
        dialogo.open()

    def _dialogo_login(self, titulo, explicacao, futuro) -> None:
        dialogo, cartao = self._dialogo()
        with cartao:
            with gui.row().classes("items-center gap-2"):
                gui.icon("lock", color="primary", size="28px")
                gui.label(titulo).classes("text-h6")
            gui.label(explicacao).classes("text-caption t-suave")
            usuario = gui.input("Usuário").props("outlined autofocus").classes("w-full")
            senha = gui.input("Senha", password=True, password_toggle_button=True) \
                .props("outlined").classes("w-full")

            def salvar() -> None:
                if usuario.value.strip() and senha.value:
                    self._responder(dialogo, futuro, (usuario.value.strip(), senha.value))
                else:
                    gui.notify("Preencha usuário e senha", type="warning")

            senha.on("keydown.enter", salvar)
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=lambda: self._cancelar_pergunta(dialogo, futuro)) \
                    .props("flat no-caps")
                gui.button("Entrar", icon="login", on_click=salvar)
        dialogo.open()

    def _dialogo_confirmar(self, pergunta, sim, nao, futuro) -> None:
        dialogo, cartao = self._dialogo()
        with cartao:
            gui.label(pergunta).classes("text-body1").style("white-space: pre-line")
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button(nao, on_click=lambda: self._responder(dialogo, futuro, False)).props("flat no-caps")
                gui.button(sim, on_click=lambda: self._responder(dialogo, futuro, True))
        dialogo.open()

    async def _abrir_no_word(self, gerar, botao=None) -> None:
        """Grava o ato num Word (gerar() devolve o caminho) e abre, para imprimir."""
        if botao is not None:
            botao.props("loading")
        try:
            caminho = await run.io_bound(gerar)
        except Exception as e:
            gui.notify(f"Não consegui montar o Word: {e if isinstance(e, RuntimeError) else type(e).__name__}",
                       type="negative")
            return
        finally:
            if botao is not None:
                botao.props(remove="loading")
        abrir(caminho)
        gui.notify("Abri o ato no Word. Para imprimir: Ctrl+P. Esta cópia não muda nada no envio.",
                   type="positive", timeout=6000)

    def _dialogo_confirmar_envio(self, ato, texto, futuro, folha="", word=None, unico=False) -> None:
        dialogo, cartao = self._dialogo("w-[860px]")
        with cartao:
            with gui.row().classes("items-center gap-2 w-full"):
                gui.icon("outgoing_mail", color="primary", size="28px")
                gui.label("Está tudo ok para o envio?").classes("text-h6")
                gui.space()
                if word is not None:
                    botao_word = gui.button("Abrir no Word (imprimir)", icon="print") \
                        .props("flat no-caps").tooltip("Abre uma cópia deste ato no Word, para imprimir")
                    botao_word.on_click(lambda: self._abrir_no_word(word, botao_word))
            gui.label(ato).classes("text-subtitle1 text-weight-medium")
            if "[" in texto:
                faltam = sorted(set(re.findall(r"\[[^\]]+\]", texto)))
                with gui.row().classes("items-center gap-1 text-negative"):
                    gui.icon("warning")
                    gui.label("Campos vazios no ato: " + ", ".join(faltam))
            if folha:
                gui.html(f'<div class="folha-area" style="height:440px">{folha}</div>').classes("w-full")
            else:
                with gui.scroll_area().classes("h-[420px] w-full border rounded q-pa-sm fundo-previa"):
                    gui.label(texto).classes("previa")
            # Da esquerda para a direita, do mais seguro ao envio, com o mesmo espaço entre todos
            # (justify-between) para diminuir a chance de clicar no botão errado
            with gui.row().classes("w-full items-center justify-between q-mt-sm"):
                gui.button("Cancelar" if unico else "Cancelar tudo", icon="block", color="negative",
                           on_click=lambda: self._responder(dialogo, futuro, "cancelar")).props("flat")
                if not unico:  # um ato só (Ato manual): não há "este" e "os próximos"
                    gui.button("Pular este ato", icon="skip_next",
                               on_click=lambda: self._responder(dialogo, futuro, "pular")).props("outline")
                    # Conferiu o primeiro e está tudo certo: envia este e os próximos sem perguntar
                    gui.button("Aceitar todos os próximos", icon="done_all", color="positive",
                               on_click=lambda: self._responder(dialogo, futuro, "todos")) \
                        .props("outline").tooltip("Envia este ato e os próximos desta vez sem perguntar de novo")
                gui.button("Enviar", icon="send", color="positive",
                           on_click=lambda: self._responder(dialogo, futuro, "enviar"))
        dialogo.open()

    def _dialogo_anexos_dos_atos(self, lista, futuro) -> None:
        """Processo com anexo: antes de começar, os PDFs de cada ato (num ato com várias
        pessoas, os de cada uma, com o nome da pessoa na frente, e/ou os do ato todo, como
        o ofício, que vão uma vez só com o nome original; todos no mesmo documento). Ato sem
        nenhum arquivo é pulado. Depois disso o envio segue sem parar."""
        dialogo, cartao = self._dialogo("w-[820px]")
        # escolhidos[(ato, dono)]: dono "" = o ato todo
        escolhidos: dict[tuple[int, str], list[Path]] = {
            (n, dono): [] for n, item in enumerate(lista) for dono in ["", *(item.get("pessoas") or [])]}
        areas: dict[tuple[int, str], object] = {}
        marcas: dict[int, object] = {}
        with cartao:
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("attach_file", color="primary", size="28px")
                with gui.column().classes("gap-0"):
                    gui.label("Anexos dos atos").classes("text-h6")
                    gui.label("Escolha agora os PDFs de cada ato: depois o envio segue sem parar para perguntar.") \
                        .classes("text-body2 t-suave")
            with gui.element("div").classes("w-full").style("max-height: 58vh; overflow-y: auto"):
                for n, item in enumerate(lista):
                    pessoas = item.get("pessoas")
                    with gui.card().classes("w-full q-pa-sm q-mb-sm fundo-cartao").props("flat bordered"):
                        with gui.row().classes("w-full items-center gap-2 no-wrap"):
                            marcas[n] = gui.icon("radio_button_unchecked", size="22px")
                            gui.label(f"Ato {n + 1} · {item['ato']}").classes("text-subtitle2 text-weight-medium t-forte col ellipsis")
                            if pessoas:
                                # Começa fechado (só o anexo do ato todo); abrir mostra pessoa por pessoa
                                botao_abrir = gui.button(f"{len(pessoas)} pessoas", icon="expand_more") \
                                    .props("unelevated dense no-caps size=sm color=primary") \
                                    .tooltip("Ver as pessoas deste ato, para escolher um PDF para cada uma")
                                gui.button("Ofício do ato todo...", icon="folder_open",
                                           on_click=lambda n=n: escolher(n, "")).props("flat dense no-caps") \
                                    .tooltip("Um PDF para o ato inteiro (ex.: o ofício): vai uma vez só, sem "
                                             "o nome de uma pessoa, e já basta para o ato ser enviado")
                            else:
                                gui.button("Escolher PDF(s)...", icon="folder_open",
                                           on_click=lambda n=n: escolher(n, "")).props("flat dense no-caps")
                        if pessoas:
                            areas[(n, "")] = gui.column().classes("w-full gap-1")
                            with gui.column().classes("w-full gap-1") as lista_pessoas:
                                for dono in pessoas:
                                    with gui.row().classes("w-full items-center gap-2 no-wrap q-pl-lg"):
                                        gui.icon("person", color="primary", size="18px")
                                        gui.label(dono or "(sem nome)").classes("text-body2 t-forte col")
                                        gui.button("Escolher PDF(s)...", icon="folder_open",
                                                   on_click=lambda n=n, d=dono: escolher(n, d)).props("flat dense no-caps")
                                    areas[(n, dono)] = gui.column().classes("w-full gap-1 q-pl-lg")
                            lista_pessoas.set_visibility(False)

                            def abrir_fechar(_, lista_pessoas=lista_pessoas, botao=botao_abrir) -> None:
                                aberto = not lista_pessoas.visible
                                lista_pessoas.set_visibility(aberto)
                                botao.props(f"icon={'expand_less' if aberto else 'expand_more'}")

                            botao_abrir.on_click(abrir_fechar)
                        else:
                            areas[(n, "")] = gui.column().classes("w-full gap-1")

            def desenhar(chaves=None) -> None:
                for chave in chaves or areas:
                    area = areas[chave]
                    area.clear()
                    do_ato_todo = chave[1] == "" and bool(lista[chave[0]].get("pessoas"))
                    with area:
                        if not escolhidos[chave] and not do_ato_todo:
                            gui.label("Nenhum arquivo escolhido ainda.").classes("text-caption t-suave q-pl-sm")
                        for arquivo in escolhidos[chave]:
                            with gui.row().classes("w-full items-center gap-2 no-wrap fundo-previa rounded q-px-md q-py-xs"):
                                gui.icon("description", color="primary")
                                gui.label(arquivo.name).classes("text-body2 t-forte col ellipsis")
                                if do_ato_todo:
                                    gui.badge("ato todo", color="primary").props("outline")
                                gui.button(icon="close", on_click=lambda c=chave, a=arquivo: (escolhidos[c].remove(a), desenhar([c]))) \
                                    .props("flat round dense size=sm").tooltip("Tirar este arquivo")
                prontos = 0
                for n, item in enumerate(lista):
                    donos = item.get("pessoas") or [""]
                    com = [d for d in donos if escolhidos[(n, d)]]
                    completo = len(com) == len(donos) or bool(escolhidos[(n, "")])  # o do ato todo basta
                    algum = bool(com) or completo
                    prontos += algum
                    marcas[n].props(f'name={"check_circle" if completo else "error" if algum else "radio_button_unchecked"} '
                                    f'color={"positive" if completo else "orange-8" if algum else "grey-6"}')
                total = sum(len(v) for v in escolhidos.values())
                continuar.set_enabled(bool(total))
                continuar.text = f"Começar ({prontos} de {len(lista)} atos com anexo)"
                sem = len(lista) - prontos
                aviso.text = f"{sem} ato(s) sem anexo serão pulados." if sem and total else ""

            async def escolher(n: int, dono: str) -> None:
                recusados = []
                for arquivo in await self._escolher_arquivos_do_computador():
                    if arquivo.suffix.lower() != ".pdf":  # o e-docs só recebe PDF como anexo
                        recusados.append(arquivo.name)
                    elif arquivo not in escolhidos[(n, dono)]:
                        escolhidos[(n, dono)].append(arquivo)
                if recusados:
                    gui.notify("Só aceita PDF. Não entrou: " + ", ".join(recusados), type="warning",
                               multi_line=True)
                desenhar([(n, dono)])

            def resposta() -> list:
                saida = []
                for n, item in enumerate(lista):
                    if item.get("pessoas"):  # os do ato todo primeiro, sem nome de pessoa
                        saida.append([str(a) for a in escolhidos[(n, "")]] +
                                     [[str(a), d] for d in item["pessoas"] for a in escolhidos[(n, d)]])
                    else:
                        saida.append([str(a) for a in escolhidos[(n, "")]])
                return saida

            with gui.row().classes("items-center gap-3"):
                with gui.row().classes("items-center gap-1 no-wrap"):
                    gui.icon("picture_as_pdf", color="negative", size="20px")
                    gui.label("Somente arquivos PDF").classes("text-body2 text-weight-medium t-forte")
                aviso = gui.label("").classes("text-body2 text-orange-9")
            with gui.row().classes("w-full items-center justify-between q-mt-sm"):
                gui.button("Cancelar", icon="block", color="negative",
                           on_click=lambda: self._responder(dialogo, futuro, "cancelar")).props("flat")
                continuar = gui.button("Começar", icon="play_arrow", color="positive",
                                       on_click=lambda: self._responder(dialogo, futuro, resposta()))
            desenhar()
        dialogo.open()

    async def _escolher_arquivos_do_computador(self, rotulo: str = "PDF",
                                               extensoes: tuple[str, ...] = (".pdf",)) -> list[Path]:
        """Janela de abrir arquivo do Windows (vários de uma vez); no navegador, envio do
        arquivo (fica uma cópia temporária)."""
        if napp.native.main_window is not None:
            import webview
            tipo = webview.FileDialog.OPEN if hasattr(webview, "FileDialog") else webview.OPEN_DIALOG
            filtro = ";".join(f"*{e}" for e in extensoes)
            escolhidos = await napp.native.main_window.create_file_dialog(
                tipo, allow_multiple=True, file_types=(f"{rotulo} ({filtro})",))
            return [Path(c) for c in escolhidos or []]
        import tempfile
        futuro = asyncio.get_running_loop().create_future()
        recebidos: list[Path] = []
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao q-pa-lg"):
            gui.label(f"Envie o(s) arquivo(s): {rotulo}").classes("text-h6")

            async def recebido(e) -> None:
                destino = Path(tempfile.mkdtemp()) / e.file.name
                await e.file.save(destino)
                recebidos.append(destino)

            gui.upload(on_upload=recebido, auto_upload=True, multiple=True).props(f"accept={','.join(extensoes)}")
            gui.button("Pronto", on_click=lambda: (janela.close(), futuro.done() or futuro.set_result(recebidos)))
        janela.open()
        return await futuro

    # --- Logins ---
    async def verificar_logins(self, *sites: str) -> None:
        for site in sites or SITES:
            atual = credenciais.carregar(main.CREDENCIAIS, site)
            self.login[site] = (AMARELO, atual[0] if atual else "", "Verificando...")
        self._desenhar_logins()
        for site in sites or SITES:
            self.login[site] = await run.io_bound(testar_login, site)
            self._desenhar_logins()
        if hasattr(self, "cfg_usuarios"):
            self._desenhar_usuarios()

    def _trocar_processo(self, e) -> None:
        if not e.value or e.value == main.processo.name:
            return
        main.processo = next(p for p in main.listar_processos() if p.name == e.value)
        main.salvar_config(processo=e.value)
        self._desenhar_processos()
        gui.timer(0.01, self.recarregar, once=True)
        # O Diário é do processo em uso: confere de novo o de hoje (sem baixar) e a lista
        background_tasks.create(self._diario_de_hoje(False))
        if getattr(self, "diario_carregado", False):
            background_tasks.create(self._carregar_diario())

    def _confirmar_excluir_processo(self, pasta: Path) -> None:
        """Pergunta antes de excluir. A pasta não é apagada de vez: vai para
        saidas/Backup de processos (dá para trazer de volta copiando para entradas/processos)."""
        if self.rodando:
            gui.notify("Aguarde o fim da execução para excluir um processo", type="warning")
            return
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[520px] max-w-full q-pa-lg gap-3"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("delete_forever", color="negative", size="30px")
                gui.label(f"Excluir o processo {pasta.name}?").classes("text-h6 t-forte")
            gui.label("Ele sai da lista de processos. Por segurança, a pasta dele (o Word, as configurações e a "
                      "planilha, se estiver dentro dela) vai para saidas\\Backup de processos: se mudar de "
                      "ideia, dá para trazer de volta.").classes("text-body2")
            if pasta == main.processo:
                gui.label("É o processo em uso: o Edodocs passa a usar outro.").classes("text-body2 text-orange-8")

            def excluir() -> None:
                janela.close()
                self._excluir_processo(pasta)

            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Sim, excluir", icon="delete", color="negative", on_click=excluir) \
                    .props("unelevated no-caps")
        janela.open()

    def _renomear_processo(self, pasta: Path) -> None:
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[480px] max-w-full q-pa-lg gap-3"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("drive_file_rename_outline", color="primary", size="28px")
                gui.label("Mudar o nome do processo").classes("text-h6")
            nome = gui.input("Nome do processo", value=pasta.name,
                             validation=lambda v: main.problema_no_nome(v or "", pasta) or None) \
                .props("outlined autofocus").classes("w-full")
            gui.label("A planilha, os vistos e assinantes e o que já foi enviado continuam ligados ao "
                      "processo.").classes("text-caption t-suave")

            def salvar() -> None:
                if (nome.value or "").strip() == pasta.name:
                    janela.close()
                    return
                try:
                    novo = main.renomear_processo(pasta, nome.value or "")
                except ValueError as e:
                    gui.notify(str(e), type="warning")
                    return
                except OSError as e:
                    gui.notify("Não consegui mudar o nome: algum arquivo do processo está aberto (Word ou "
                               f"Excel?). Feche e tente de novo. ({type(e).__name__})", type="negative",
                               multi_line=True)
                    return
                janela.close()
                gui.notify(f"Processo {pasta.name} agora se chama {novo.name}", type="positive")
                self.sel_processo.set_options([p.name for p in main.listar_processos()],
                                              value=main.processo.name if main.processo else None)
                self._desenhar_processos()
                if main.processo == novo:
                    gui.timer(0.01, self.recarregar, once=True)

            nome.on("keydown.enter", salvar)
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Salvar", icon="check", on_click=salvar).props("unelevated no-caps")
        janela.open()

    def _excluir_processo(self, pasta: Path) -> None:
        import shutil
        backup = main.SAIDAS / "Backup de processos"
        destino = backup / f"{pasta.name} (excluído em {datetime.now():%Y-%m-%d %Hh%M%S})"
        try:
            backup.mkdir(parents=True, exist_ok=True)
            shutil.move(str(pasta), str(destino))
        except OSError as e:
            gui.notify(f"Não consegui excluir: algum arquivo do processo está aberto (Word ou Excel?). "
                       f"Feche e tente de novo. ({type(e).__name__})", type="negative", multi_line=True)
            return
        planilhas = main.ler_config().get("planilhas", {})
        planilhas.pop(pasta.name, None)
        main.salvar_config(planilhas=planilhas)
        gui.notify(f"Processo {pasta.name} excluído (guardado em saidas\\Backup de processos)", type="positive")
        restantes = main.listar_processos()
        if not restantes:  # era o último: volta para a tela de boas-vindas
            main.processo = None
            main.salvar_config(processo="")
            gui.navigate.reload()
            return
        if pasta == main.processo:
            main.processo = restantes[0]
            main.salvar_config(processo=main.processo.name)
            gui.timer(0.01, self.recarregar, once=True)
        self.sel_processo.set_options([p.name for p in restantes], value=main.processo.name)
        self._desenhar_processos()

    # --- Página Processos ---
    def _painel_processos(self) -> None:
        if problema := main.problema_na_compartilhada():
            with gui.row().classes("w-full items-center gap-2 no-wrap fundo-cartao rounded q-pa-sm"):
                gui.icon("wifi_off", color="negative", size="26px")
                gui.label(problema).classes("text-body2 t-forte")
        elif raiz := main.pasta_compartilhada():
            with gui.row().classes("w-full items-center gap-2 no-wrap"):
                gui.icon("folder_shared", color="primary")
                gui.label(f"Processos da pasta compartilhada {raiz}: o que você mudar aqui vale para todos "
                          f"(uma cópia de antes fica em \"{main.EXTRAS}\").").classes("text-caption t-suave")
        with gui.row().classes("w-full items-center"):
            gui.label("Cada processo é um Word modelo ligado a uma planilha. Crie o seu ou edite "
                      "um existente.").classes("text-body1 t-forte")
            gui.space()
            gui.button("Exportar processos", icon="file_upload", on_click=self._exportar_processos) \
                .props("flat no-caps").tooltip("Guarda todos os processos num arquivo na sua pasta Downloads "
                                               "(para levar para uma versão nova ou outro computador)")
            gui.button("Importar processos", icon="file_download", on_click=self._importar_processos) \
                .props("flat no-caps").tooltip("Traz os processos de um arquivo exportado")
            gui.button("Novo processo", icon="add", on_click=self._novo_processo).props("no-caps")
        self.lista_processos = gui.grid(columns=2).classes("w-full gap-4 q-mt-md")
        self._desenhar_processos()

    def _boas_vindas(self) -> None:
        """Janela da primeira abertura (o .exe recém-baixado): importar os processos de uma
        versão anterior ou começar com os que vieram no programa. Aparece uma vez só."""
        main.salvar_config(boas_vindas_vista=True)
        processos_aqui = [p.name for p in main.listar_processos()]
        with self.camada, gui.dialog() as janela,                 gui.card().classes("cartao w-[620px] max-w-full q-pa-lg gap-3").style("position: relative"):
            with gui.row().classes("items-center gap-3 no-wrap"):
                _logo("w-20 h-20")
                with gui.column().classes("gap-0"):
                    gui.label("Bem-vindo ao Edodocs!").classes("text-h5 t-forte")
                    gui.label("Que bom ter você aqui. Antes de começar:").classes("text-body1 t-suave")
            with gui.card().classes("w-full fundo-cartao q-pa-md gap-2"):
                with gui.row().classes("items-center gap-2 no-wrap"):
                    gui.icon("history", color="primary", size="26px")
                    gui.label("Já usava o Edodocs (em outra versão ou outro computador)?")                         .classes("text-subtitle1 text-weight-medium t-forte")
                gui.label("Traga os seus processos: na versão antiga, use Processos → Exportar processos; "
                          "o arquivo fica na pasta Downloads. Aqui é só importar.").classes("text-body2 t-suave")
                async def importar() -> None:
                    janela.close()
                    await self._importar_processos()

                gui.button("Importar meus processos", icon="file_download", on_click=importar)                     .props("unelevated no-caps")
            with gui.row().classes("w-full items-center gap-2 no-wrap"):
                gui.icon("rocket_launch", color="primary", size="24px")
                gui.label("Primeira vez? " + (f"Já vieram prontos: {', '.join(processos_aqui)}."
                                              if processos_aqui else "Crie o seu primeiro processo em Processos."))                     .classes("text-body2 t-forte col")
                gui.button("Começar", icon="arrow_forward", on_click=janela.close).props("flat no-caps")
            import versao
            gui.label(f"versão {versao.atual()}").classes("text-caption")                 .style("position: absolute; right: 16px; bottom: 6px; color: #b0b6bf")
        janela.open()

    # --- Exportar / importar processos (ver processos_backup.py) ---
    async def _exportar_processos(self) -> None:
        import processos_backup
        cadastradas = consultas.tabelas()
        usadas = processos_backup.tabelas_usadas(main.PROCESSOS, cadastradas)
        faltando = processos_backup.tabelas_faltando(main.PROCESSOS, cadastradas)
        # Planilha que fica fora da pasta do processo (na rede, ou escolhida em Configurações):
        # sem perguntar, iria só o caminho dela
        fora = {}
        for pasta in main.listar_processos():
            planilha = consultas_ui.planilha_de(pasta)
            if planilha and pasta.resolve() not in planilha.resolve().parents:
                fora[pasta.name] = planilha
        todos_de_para = de_para.todos()
        escolha = (await self._o_que_levar(usadas, cadastradas, fora, todos_de_para, faltando)
                   if usadas or fora or todos_de_para or faltando else ({}, {}, {}))
        if escolha is None:
            return  # cancelou
        levar, planilhas, levar_de_para = escolha
        try:
            arquivo, nomes = await run.io_bound(processos_backup.exportar, main.PROCESSOS,
                                                main.ler_config().get("planilhas", {}), None, levar,
                                                levar_de_para, planilhas)
        except Exception as e:
            gui.notify(f"Não consegui exportar: {e}", type="negative")
            return
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[560px] max-w-full q-pa-lg gap-3"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("check_circle", color="positive", size="30px")
                gui.label(f"{len(nomes)} processo(s) exportado(s)").classes("text-h6 t-forte")
            gui.label("O arquivo está na sua pasta Downloads:").classes("text-body2 t-suave")
            gui.label(arquivo.name).classes("text-body1 text-weight-medium t-forte break-all")
            gui.label(", ".join(nomes)).classes("text-caption t-suave")
            if levar:
                gui.label(f"Foram junto as tabelas de consulta: {', '.join(levar)}").classes("text-caption t-suave")
            if planilhas:
                gui.label(f"Foram junto as planilhas de: {', '.join(planilhas)}").classes("text-caption t-suave")
            if levar_de_para:
                gui.label(f"Foram junto os de-para: {', '.join(levar_de_para)}").classes("text-caption t-suave")
            gui.label("Na versão nova do Edodocs, clique em Importar processos e escolha este arquivo.") \
                .classes("text-body2 t-forte")
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Abrir a pasta Downloads", icon="folder_open",
                           on_click=lambda: mostrar_na_pasta(arquivo)).props("flat no-caps")
                gui.button("Fechar", on_click=janela.close).props("unelevated no-caps")
        janela.open()

    async def _o_que_levar(self, usadas: dict[str, list[str]], cadastradas: dict,
                           fora: dict[str, Path], de_paras: dict[str, dict],
                           faltando: dict[str, list[str]] | None = None) -> tuple[dict, dict, dict] | None:
        """Antes de exportar: quais tabelas de consulta, planilhas de fora da pasta do processo
        e de-para vão junto (assim uma exportação leva tudo). faltando: tabelas que os processos
        usam mas não estão neste computador (a pessoa escolhe o Excel, ou elas não vão).
        Devolve ({tabela: cadastro}, {processo: planilha}, {de-para: pares}) ou None (cancelou)."""
        faltando = faltando or {}
        futuro = asyncio.get_running_loop().create_future()
        marcadas = set(usadas)
        planilhas_marcadas = set(fora)
        de_para_marcados = set(de_paras)
        usados_de_para = de_para.usados()
        contagens = {}

        def terminar(valor) -> None:
            janela.close()
            if not futuro.done():
                futuro.set_result(valor)

        def exportar() -> None:
            levar = {}
            for nome in marcadas:
                dados = dict(consultas.tabelas().get(nome) or cadastradas[nome])
                quando = consultas.atualizada_em(dados)
                if quando:
                    dados["atualizada_em"] = quando.isoformat()
                levar[nome] = dados
            terminar((levar, {n: fora[n] for n in planilhas_marcadas},
                      {n: de_paras[n] for n in de_paras if n in de_para_marcados}))

        with self.camada, gui.dialog().props("persistent") as janela, \
                gui.card().classes("cartao w-[640px] max-w-full q-pa-lg gap-3"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("luggage", color="primary", size="30px")
                gui.label("O que levar junto?").classes("text-h6 t-forte")
            if usadas:
                gui.label("Tabelas de consulta: com elas, no outro computador o Preencher sozinho já chega "
                          "funcionando.").classes("text-body2 t-suave")
                for nome, procs in usadas.items():
                    with gui.column().classes("w-full gap-0"):
                        with gui.row().classes("items-center gap-2 no-wrap"):
                            gui.checkbox(f"Levar junto a tabela {nome}", value=True,
                                         on_change=lambda e, n=nome: (marcadas.add if e.value else marcadas.discard)(n))
                            contagens[nome] = gui.label("").classes("text-body2 t-suave")
                        gui.label(f"usada por: {', '.join(procs)}").classes("text-caption t-suave q-ml-xl")
            for nome, procs in faltando.items():
                with gui.column().classes("w-full gap-1 fundo-cartao rounded q-pa-sm") as caixa:
                    with gui.row().classes("items-start gap-2 no-wrap"):
                        gui.icon("warning", color="orange-8", size="22px")
                        gui.label(f"A tabela {nome} (usada por: {', '.join(procs)}) não está neste computador, "
                                  "então não iria junto e no outro computador o Preencher sozinho não acharia as "
                                  "pessoas. Escolha o arquivo do Excel dela para levar junto:") \
                            .classes("text-body2 t-forte")

                    async def achar(n=nome, c=caixa) -> None:
                        if await consultas_ui.achar_tabela(self.camada, n):
                            marcadas.add(n)
                            c.clear()
                            with c, gui.row().classes("items-center gap-2 no-wrap"):
                                gui.icon("check_circle", color="positive", size="22px")
                                gui.label(f"A tabela {n} vai junto").classes("text-body2 t-forte")

                    gui.button("Escolher a tabela...", icon="upload_file", on_click=achar) \
                        .props("outline no-caps").classes("self-start")
            if fora:
                gui.label("Planilhas dos processos que ficam fora da pasta do processo (sem marcar, vai só o "
                          "caminho, que pode não existir no outro computador):").classes("text-body2 t-suave q-mt-sm")
                for nome, planilha in fora.items():
                    with gui.column().classes("w-full gap-0"):
                        gui.checkbox(f"Levar a planilha de {nome}", value=True,
                                     on_change=lambda e, n=nome: (planilhas_marcadas.add if e.value
                                                                  else planilhas_marcadas.discard)(n))
                        gui.label(str(planilha)).classes("text-caption t-suave q-ml-xl break-all")
            if de_paras:
                gui.label("De-para (código → descrição): no outro computador eles se juntam aos de lá, sem "
                          "trocar nenhuma descrição.").classes("text-body2 t-suave q-mt-sm")
                for nome, pares in de_paras.items():
                    with gui.column().classes("w-full gap-0"):
                        gui.checkbox(f"Levar junto o de-para {nome} ({len(pares)} códigos)", value=True,
                                     on_change=lambda e, n=nome: (de_para_marcados.add if e.value
                                                                  else de_para_marcados.discard)(n))
                        gui.label(f"usado por: {', '.join(usados_de_para[nome])}" if usados_de_para.get(nome)
                                  else "ainda não usado em nenhum processo").classes("text-caption t-suave q-ml-xl")
            with gui.row().classes("items-start gap-2 no-wrap fundo-cartao rounded q-pa-sm"):
                gui.icon("privacy_tip", color="orange-8", size="22px")
                gui.label("Tabelas e planilhas costumam ter dados pessoais (CPF, matrícula, lotação). O arquivo "
                          "exportado fica na sua pasta Downloads: só envie para quem pode ver esses dados.") \
                    .classes("text-body2 t-forte")
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=lambda: terminar(None)).props("flat no-caps")
                gui.button("Exportar", icon="file_upload", on_click=exportar).props("unelevated no-caps")
        janela.open()
        for nome, rotulo in contagens.items():  # quantas pessoas (pode demorar em tabela grande)
            try:
                tabela = await run.io_bound(consultas.ler_cadastrada, nome)
                rotulo.set_text(f"({len(tabela.linhas)} pessoas)")
            except Exception:
                pass
        return await futuro

    async def _importar_processos(self) -> None:
        """Janela de importar: escolhe o arquivo (o mais novo de Downloads já vem marcado) e
        quais processos trazer."""
        import processos_backup
        arquivos = await run.io_bound(processos_backup.exportacoes)
        estado = {"arquivo": arquivos[0] if arquivos else None, "lido": None, "marcados": set(), "tabelas": set(),
                  "usar_planilha": set()}
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[720px] max-w-full q-pa-lg gap-3"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("file_download", color="primary", size="30px")
                gui.label("Importar processos").classes("text-h6 t-forte")
            gui.label("Escolha o arquivo exportado (ele fica na sua pasta Downloads) e os processos que "
                      "quer trazer.").classes("text-body2 t-suave")
            with gui.row().classes("w-full items-center gap-2 no-wrap"):
                seletor = gui.select({str(a): a.name for a in arquivos}, label="Arquivo em Downloads",
                                     value=str(estado["arquivo"]) if estado["arquivo"] else None,
                                     on_change=lambda e: escolher(Path(e.value)) if e.value else None) \
                    .props("outlined dense").classes("col")
                if napp.native.main_window is not None:
                    gui.button("Outro arquivo...", icon="attach_file",
                               on_click=lambda: outro_arquivo()).props("flat no-caps")
            area = gui.column().classes("w-full gap-1")
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                botao = gui.button("Importar", icon="file_download", on_click=lambda: importar()) \
                    .props("unelevated no-caps")

        def pode_importar() -> bool:
            return bool(isinstance(estado["lido"], dict) and (estado["marcados"] or estado["tabelas"]))

        def desenhar() -> None:
            area.clear()
            lido = estado["lido"]
            botao.set_enabled(pode_importar())
            with area:
                if estado["arquivo"] is None:
                    gui.label("Não achei nenhum arquivo exportado na sua pasta Downloads. Na versão antiga, "
                              "clique em Exportar processos (em Processos).").classes("text-body2 t-forte")
                    return
                if isinstance(lido, str):
                    with gui.row().classes("items-center gap-2"):
                        gui.icon("warning", color="warning")
                        gui.label(f"Não consegui ler: {lido}").classes("text-body2 t-forte")
                    return
                if lido is None:
                    gui.spinner()
                    return
                origem = " · ".join(x for x in (f"exportado em {lido['exportado_em']}" if lido["exportado_em"] else "",
                                                f"no computador {lido['computador']}" if lido["computador"] else "") if x)
                if origem:
                    gui.label(origem.capitalize()).classes("text-caption t-suave")
                for p in lido["processos"]:
                    with gui.row().classes("w-full items-center gap-2 no-wrap"):
                        gui.checkbox(p["nome"], value=p["nome"] in estado["marcados"],
                                     on_change=lambda e, n=p["nome"]: marcar(n, e.value))
                        if p["existe"]:
                            gui.label("já existe aqui: o atual fica guardado em saidas\\Backup de processos") \
                                .classes("text-caption text-orange-8")
                    if p["planilha_junto"]:  # a planilha de fora da pasta veio junto
                        origem_existe = bool(p["planilha_origem"]) and Path(p["planilha_origem"]).is_file()
                        with gui.column().classes("gap-0 q-ml-xl"):
                            gui.checkbox(f"Usar a planilha que veio junto ({p['planilha_junto']})",
                                         value=p["nome"] in estado["usar_planilha"],
                                         on_change=lambda e, n=p["nome"]: (estado["usar_planilha"].add if e.value
                                                                           else estado["usar_planilha"].discard)(n)) \
                                .props("dense")
                            gui.label(("Desmarcada: continua usando " + p["planilha_origem"]
                                       + (" (existe neste computador)" if origem_existe
                                          else " (não existe neste computador)")) if p["planilha_origem"] else
                                      "Desmarcada: a planilha fica para você escolher em Processos") \
                                .classes("text-caption " + ("t-suave" if origem_existe else "text-orange-8"))
                if lido["tabelas"]:
                    gui.label("Tabelas de consulta (preencher sozinho)").classes("text-subtitle2 t-forte q-mt-sm")
                    daqui = consultas.tabelas()
                for t in lido["tabelas"]:
                    data = _data_br(t["atualizada_em"])
                    with gui.column().classes("w-full gap-0"):
                        if t["existe"]:
                            local = consultas.atualizada_em(daqui.get(t["nome"]) or {})
                            gui.checkbox(f"Trocar a tabela {t['nome']} daqui pela do arquivo",
                                         value=t["nome"] in estado["tabelas"],
                                         on_change=lambda e, n=t["nome"]: marcar_tabela(n, e.value))
                            gui.label(f"a daqui: dados de {local:%d/%m/%Y} · a do arquivo: dados de {data}"
                                      if local else f"a do arquivo: dados de {data}") \
                                .classes("text-caption t-suave q-ml-xl")
                        else:
                            gui.checkbox(f"Trazer a tabela {t['nome']}", value=t["nome"] in estado["tabelas"],
                                         on_change=lambda e, n=t["nome"]: marcar_tabela(n, e.value))
                            gui.label(f"usada por: {', '.join(t['processos'])} · dados de {data}") \
                                .classes("text-caption t-suave q-ml-xl")
                if lido["de_para"]:
                    gui.label("De-para (vêm junto e se juntam aos daqui, sem trocar descrições): "
                              + ", ".join(f"{n} ({len(p)} códigos)" for n, p in lido["de_para"].items())) \
                        .classes("text-caption t-suave q-mt-sm")

        def marcar(nome: str, valor: bool) -> None:
            (estado["marcados"].add if valor else estado["marcados"].discard)(nome)
            botao.set_enabled(pode_importar())

        def marcar_tabela(nome: str, valor: bool) -> None:
            (estado["tabelas"].add if valor else estado["tabelas"].discard)(nome)
            botao.set_enabled(pode_importar())

        async def escolher(arquivo: Path) -> None:
            estado.update(arquivo=arquivo, lido=None)
            desenhar()
            try:
                lido = await run.io_bound(processos_backup.ler, arquivo, main.PROCESSOS, consultas.tabelas())
            except Exception as e:
                estado["lido"] = str(e)
            else:
                estado["lido"] = lido
                # Vêm marcados os que ainda não existem aqui (substituir é escolha da pessoa)
                estado["marcados"] = {p["nome"] for p in lido["processos"] if not p["existe"]}
                # A planilha que veio junto: marcada quando a original não existe neste computador
                estado["usar_planilha"] = {p["nome"] for p in lido["processos"] if p["planilha_junto"]
                                           and not (p["planilha_origem"] and Path(p["planilha_origem"]).is_file())}
                # Tabela nova vem marcada; a que já existe, só se a do arquivo tiver dados mais novos
                daqui = consultas.tabelas()
                estado["tabelas"] = {
                    t["nome"] for t in lido["tabelas"] if not t["existe"] or (t["atualizada_em"] or "") >
                    (consultas.atualizada_em(daqui.get(t["nome"]) or {}) or date.min).isoformat()}
            desenhar()

        async def outro_arquivo() -> None:
            import webview
            tipo = webview.FileDialog.OPEN if hasattr(webview, "FileDialog") else webview.OPEN_DIALOG
            escolhidos = await napp.native.main_window.create_file_dialog(
                tipo, directory=str(processos_backup.pasta_downloads()),
                file_types=("Processos exportados do Edodocs (*.zip)",))
            if escolhidos:
                caminho = Path(escolhidos[0])
                seletor.set_options({**seletor.options, str(caminho): caminho.name}, value=str(caminho))

        async def importar() -> None:
            nomes = [p["nome"] for p in estado["lido"]["processos"] if p["nome"] in estado["marcados"]]
            botao.props("loading")
            try:
                importados = await run.io_bound(processos_backup.importar, estado["arquivo"], nomes,
                                                main.PROCESSOS, main.SAIDAS / "Backup de processos",
                                                set(estado["usar_planilha"]))
                trazidas = await run.io_bound(processos_backup.importar_tabelas, estado["arquivo"],
                                              sorted(estado["tabelas"]), consultas.pasta(),
                                              main.SAIDAS / "Backup de processos")
            except Exception as e:
                gui.notify(f"Não consegui importar: {type(e).__name__} {e}", type="negative")
                return
            finally:
                botao.props(remove="loading")
            for nome, dados in trazidas.items():
                consultas.cadastrar(nome, dados)
            for nome, pares in estado["lido"]["de_para"].items():  # sem dados pessoais: vêm sempre
                de_para.juntar(nome, pares)
            # A planilha que cada processo usava (se o caminho existir neste computador)
            planilhas = main.ler_config().get("planilhas", {})
            for nome, caminho in estado["lido"]["planilhas"].items():
                if nome in importados and nome not in planilhas and Path(caminho).is_file():
                    planilhas[nome] = caminho
            for nome in set(importados) & estado["usar_planilha"]:  # usa a que veio, na pasta do processo
                planilhas.pop(nome, None)
            main.salvar_config(planilhas=planilhas)
            janela.close()
            gui.notify(f"{len(importados)} processo(s) importado(s): {', '.join(importados)}"
                       + (f" · tabela(s) de consulta: {', '.join(trazidas)}" if trazidas else ""), type="positive")
            if not hasattr(self, "sel_processo"):  # não havia processo: monta a tela inteira
                gui.navigate.reload()
                return
            self.sel_processo.set_options([p.name for p in main.listar_processos()], value=self.sel_processo.value)
            self._desenhar_processos()

        janela.open()
        if estado["arquivo"]:
            await escolher(estado["arquivo"])
        else:
            desenhar()

    def _novo_processo(self) -> None:
        """Com uma IA configurada, a pessoa escolhe: montar ela mesma (o passo a passo) ou pela
        IA, a partir dos Words de exemplo. Sem IA, vai direto para o passo a passo."""
        if not ia_conectada():
            Assistente.novo(self)
            return

        def manual() -> None:
            janela.close()
            Assistente.novo(self)

        def com_ia() -> None:
            janela.close()
            import processo_ia_ui
            processo_ia_ui.JanelaIA(self)

        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[640px] max-w-full q-pa-lg gap-4"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("add_circle", color="primary", size="30px")
                gui.label("Novo processo: como você quer montar?").classes("text-h6 t-forte")
            with gui.row().classes("w-full gap-3 no-wrap items-stretch"):
                with gui.card().classes("cartao col q-pa-md gap-2 cursor-pointer").on("click", manual):
                    gui.icon("checklist", color="primary", size="34px")
                    gui.label("Montar eu mesmo").classes("text-subtitle1 text-weight-bold t-forte")
                    gui.label("Passo a passo: você escolhe um ato que já fez, marca o que muda e liga à "
                              "planilha.").classes("text-body2 t-suave")
                with gui.card().classes("cartao col q-pa-md gap-2 cursor-pointer").on("click", com_ia):
                    gui.icon("auto_awesome", color="primary", size="34px")
                    gui.label("Montar com IA").classes("text-subtitle1 text-weight-bold t-forte")
                    gui.label("Você junta vários atos já feitos (Word ou PDF) e a IA monta o Word, a planilha "
                              "e as regras. Você confere antes de salvar.").classes("text-body2 t-suave")
            with gui.row().classes("w-full justify-end"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
        janela.open()

    def _desenhar_processos(self) -> None:
        if not hasattr(self, "lista_processos"):
            return
        self.lista_processos.clear()
        with self.lista_processos:
            for pasta in main.listar_processos():
                proc = processos.carregar(pasta)
                modelo = proc.modelo()
                atual = pasta == main.processo
                with gui.card().classes("cartao q-pa-lg w-full").style(
                        "border:2px solid #2459a8" if atual else ""):
                    with gui.row().classes("w-full items-center gap-2"):
                        gui.icon("folder_special", color="primary", size="28px")
                        gui.label(pasta.name).classes("text-h6")
                        gui.button(icon="edit", on_click=lambda p=pasta: self._renomear_processo(p)) \
                            .props("flat round dense size=sm color=grey-7").tooltip("Mudar o nome do processo")
                        if atual:
                            gui.badge("em uso", color="primary").props("outline")
                        gui.space()
                        if not (pasta / processos.ARQUIVO).exists():
                            gui.badge("sem configuração", color="warning")
                        gui.button(icon="close", on_click=lambda p=pasta: self._confirmar_excluir_processo(p)) \
                            .props("flat round dense color=negative").tooltip("Excluir este processo")
                    with gui.row().classes("items-center gap-2 no-wrap"):
                        gui.icon("description", color="grey-7" if modelo else "negative")
                        gui.label(modelo.name if modelo else "Word não encontrado").classes(
                            "text-body2 break-all" + ("" if modelo else " text-negative"))
                    em_uso = consultas_ui.planilha_de(pasta)
                    with gui.row().classes("items-center gap-2 no-wrap"):
                        gui.icon("table_chart", color="grey-7" if em_uso else "negative")
                        gui.label(str(em_uso) if em_uso else "Planilha não encontrada").classes(
                            "text-body2 break-all" + ("" if em_uso else " text-negative"))
                        gui.button("Alterar", icon="edit", on_click=lambda p=pasta: self._editar_planilha(p)) \
                            .props("flat dense no-caps").tooltip("Onde fica a planilha deste processo (no "
                                                                 "computador ou numa pasta da rede)")
                    gui.label(f"Título no e-docs: {proc['titulo_edocs']}").classes("text-caption t-suave")
                    pessoas = proc["pessoas"]
                    with gui.row().classes("items-center gap-2 no-wrap"):
                        gui.icon("how_to_reg", color="grey-7")
                        if pessoas.get("fixas") and pessoas.get("assinantes"):
                            vistos = ", ".join(v["nome"] for v in pessoas.get("vistos", [])) or "ninguém"
                            assinam = ", ".join(a["nome"] for a in pessoas["assinantes"])
                            gui.label(f"Fixos · Vistam: {vistos} · Assinam: {assinam}") \
                                .classes("text-body2")
                        else:
                            gui.label("Vistos e assinantes: perguntados a cada envio").classes("text-body2 t-suave")
                    with gui.row().classes("items-center gap-2 no-wrap"):
                        confirmar = proc.config.get("confirmar_envio",
                                                    main.ler_config().get("confirmar_envio", True))
                        gui.switch("Perguntar antes de enviar cada ato", value=bool(confirmar),
                                   on_change=lambda e, p=pasta: self._salvar_confirmacao(p, e.value)) \
                            .props("dense")
                        gui.label("(desligado: envia um atrás do outro, sem parar)").classes("text-caption t-suave")
                    with gui.row().classes("items-center gap-2 no-wrap"):
                        gui.switch("Precisa de anexo", value=proc.precisa_anexo,
                                   on_change=lambda e, p=pasta: self._salvar_precisa_anexo(p, e.value)) \
                            .props("dense")
                        gui.label("(ligado: pede o PDF de cada ato na hora de enviar)") \
                            .classes("text-caption t-suave")
                    with gui.row().classes("items-center gap-2 no-wrap"):
                        gui.icon("rule", color="primary" if proc.obrigatorios else "grey-7")
                        gui.label("Campos obrigatórios: " + (", ".join(proc.obrigatorios) or "nenhum")) \
                            .classes("text-body2" + ("" if proc.obrigatorios else " t-suave"))
                        gui.button("Escolher", icon="edit", on_click=lambda p=pasta: self._editar_obrigatorios(p)) \
                            .props("flat dense no-caps").tooltip("Colunas da planilha que não podem ficar vazias")
                    consulta = consultas_ui.resumo(proc.config.get("consulta"))
                    with gui.row().classes("items-center gap-2 no-wrap"):
                        gui.icon("auto_fix_high", color="positive" if consulta else "grey-7")
                        gui.label(f"Preenche sozinho com: {consulta}" if consulta else
                                  "Preencher sozinho: desligado (cada linha é escrita à mão)") \
                            .classes("text-body2" + ("" if consulta else " t-suave"))
                    with gui.row().classes("items-center gap-2 no-wrap"):
                        gui.switch("Publicado no Diário Oficial", value=proc.vai_ao_diario,
                                   on_change=lambda e, p=pasta: self._salvar_vai_ao_diario(p, e.value)) \
                            .props("dense")
                        gui.label("(desligado: só vai ao e-docs)").classes("text-caption t-suave")
                    with gui.row().classes("gap-2 q-mt-sm"):
                        gui.button("Editar", icon="edit", on_click=lambda p=pasta: Assistente(self, p)) \
                            .props("no-caps")
                        gui.button("Vistos e assinantes", icon="how_to_reg",
                                   on_click=lambda p=pasta: self._editar_pessoas(p)).props("outline no-caps")
                        # Só com o Word e a planilha prontos: as colunas vêm da planilha
                        pronto = bool(modelo) and consultas_ui.planilha_de(pasta) is not None
                        preencher = gui.button("Preencher sozinho", icon="auto_fix_high",
                                               on_click=lambda p=pasta: consultas_ui.PassoAPassoConsulta(
                                                   self.camada, p, self._desenhar_processos)) \
                            .props("outline no-caps")
                        if not pronto:
                            preencher.disable()
                            preencher.tooltip("Primeiro deixe o Word e a planilha do processo prontos (Editar)")
                        else:
                            preencher.tooltip("Ligar a uma tabela (ex.: todos os servidores): na Nova linha, você "
                                              "procura a pessoa e o resto vem preenchido")
                        gui.button("Abrir pasta", icon="folder_open", on_click=lambda p=pasta: abrir(p)) \
                            .props("outline no-caps")
                        if not atual:
                            gui.button("Usar este", on_click=lambda p=pasta: self.sel_processo.set_value(p.name)) \
                                .props("flat no-caps")

    async def _editar_obrigatorios(self, pasta: Path) -> None:
        """Escolher as colunas da planilha que não podem ficar vazias: a linha não é salva
        sem elas e o ato não é enviado."""
        proc = processos.carregar(pasta)
        planilha = consultas_ui.planilha_de(pasta)
        if planilha is None:
            gui.notify("Não achei a planilha deste processo (Editar)", type="warning")
            return
        try:
            colunas, _ = await run.io_bound(edicao.ler_estrutura, planilha, proc.aba)
        except Exception as e:
            gui.notify(f"Não consegui ler a planilha: {type(e).__name__}", type="negative")
            return
        nomes = [c.nome for c in colunas if c.editavel and str(c.nome or "").strip()]
        atuais = {atos.normalizar(c) for c in proc.obrigatorios}
        marcados = {n: atos.normalizar(n) in atuais for n in nomes}
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[520px] max-w-full q-pa-lg gap-2"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("rule", color="primary", size="28px")
                with gui.column().classes("gap-0"):
                    gui.label("Campos obrigatórios").classes("text-h6 t-forte")
                    gui.label(pasta.name).classes("text-caption t-suave")
            gui.label("Marque as colunas que não podem ficar vazias. Sem elas, a linha não é salva "
                      "(Nova linha e editar) e o ato não é enviado ao e-docs.").classes("text-body2 t-suave")
            with gui.element("div").classes("w-full").style("max-height: 50vh; overflow-y: auto"):
                with gui.grid(columns=2).classes("w-full gap-x-4 gap-y-0"):
                    for nome in nomes:
                        gui.checkbox(nome, value=marcados[nome],
                                     on_change=lambda e, n=nome: marcados.update({n: e.value})).props("dense")

            def salvar() -> None:
                proc.config["obrigatorios"] = [n for n in nomes if marcados[n]]
                proc.salvar()
                janela.close()
                gui.notify(f"{pasta.name}: " + (", ".join(proc.config["obrigatorios"]) or "nenhum campo")
                           + " obrigatório(s)", type="positive")
                self._desenhar_processos()

            with gui.row().classes("w-full justify-end gap-2 q-mt-sm"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Salvar", icon="save", on_click=salvar).props("unelevated no-caps")
        janela.open()

    def _salvar_vai_ao_diario(self, pasta: Path, valor: bool) -> None:
        proc = processos.carregar(pasta)
        proc.config["vai_ao_diario"] = valor
        proc.salvar()
        gui.notify(f"{pasta.name}: " + ("publicado no Diário Oficial" if valor
                                         else "só vai ao e-docs (etapas do Diário desligadas)"),
                   type="positive" if valor else "info")
        if pasta == main.processo:
            background_tasks.create(self.recarregar())

    @staticmethod
    def _salvar_precisa_anexo(pasta: Path, valor: bool) -> None:
        proc = processos.carregar(pasta)
        proc.config["precisa_anexo"] = valor
        proc.salvar()
        gui.notify(f"{pasta.name}: " + ("pede o anexo de cada ato ao enviar" if valor else "sem anexo"),
                   type="positive" if valor else "info")

    @staticmethod
    def _salvar_confirmacao(pasta: Path, valor: bool) -> None:
        proc = processos.carregar(pasta)
        proc.config["confirmar_envio"] = valor
        proc.salvar()
        gui.notify(f"{pasta.name}: mostra cada ato antes de enviar" if valor
                   else f"{pasta.name}: envia um atrás do outro, sem perguntar",
                   type="positive" if valor else "warning")

    def _editar_pessoas(self, pasta: Path) -> None:
        """Vistos e assinantes fixos do processo (ou perguntados a cada envio)."""
        proc = processos.carregar(pasta)
        pessoas = json.loads(json.dumps(proc["pessoas"]))  # cópia: só vale ao salvar
        dialogo = gui.dialog().props("persistent")
        with self.camada, dialogo, gui.card().classes("cartao w-[600px] max-w-full q-pa-lg"):
            with gui.row().classes("items-center gap-2"):
                gui.icon("how_to_reg", color="primary", size="26px")
                gui.label(f"Vistos e assinantes · {pasta.name}").classes("text-h6")
            painel = PainelPessoas(pessoas)

            def salvar() -> None:
                if not painel.valido():
                    gui.notify("Escolha quem assina (ou desligue a opção)", type="warning")
                    return
                proc.config["pessoas"] = pessoas
                proc.salvar()
                dialogo.close()
                gui.notify(f"Vistos e assinantes de {pasta.name} salvos", type="positive")
                self._desenhar_processos()

            with gui.row().classes("w-full justify-end gap-2 q-mt-md"):
                gui.button("Cancelar", on_click=dialogo.close).props("flat no-caps")
                gui.button("Salvar", icon="save", on_click=salvar).props("unelevated no-caps")
        dialogo.open()

    def processo_salvo(self, pasta: Path) -> None:
        """Chamado pelo assistente ao salvar: passa a usar o processo e atualiza tudo."""
        if not hasattr(self, "sel_processo"):  # era o primeiro processo: monta a tela inteira
            main.salvar_config(processo=pasta.name)
            gui.navigate.reload()
            return
        main.processo = pasta
        main.salvar_config(processo=pasta.name)
        self.sel_processo.set_options([p.name for p in main.listar_processos()], value=pasta.name)
        self._desenhar_processos()
        gui.timer(0.01, self.recarregar, once=True)

    # --- Aba Configurações ---
    def _painel_config(self) -> None:
        with gui.grid(columns=2).classes("w-full gap-4"):
            with gui.column().classes("gap-4"):
                with gui.card().classes("cartao q-pa-lg w-full"):
                    with gui.row().classes("items-center gap-2"):
                        gui.icon("manage_accounts", color="primary", size="26px")
                        gui.label("Usuários dos sites").classes("text-h6")
                    gui.label("As senhas ficam criptografadas neste computador (configs/credenciais.json).") \
                        .classes("text-caption t-suave")
                    self.cfg_usuarios = gui.column().classes("w-full gap-4")
                    self._desenhar_usuarios()
                # A chave da IA fica junto dos usuários e senhas
                from recebidos_ui import PainelIA
                PainelIA()

            with gui.column().classes("gap-4"):
                self._cartao_pasta_compartilhada()

                import de_para_ui
                de_para_ui.CartaoDePara(self.camada)

                with gui.card().classes("cartao q-pa-lg w-full"):
                    with gui.row().classes("items-center gap-2"):
                        gui.icon("pan_tool", color="primary", size="26px")
                        gui.label("Envio").classes("text-h6")
                    gui.label("Perguntar ou não antes de enviar cada ato agora é escolhido em cada "
                              "processo: alguns podem ir direto, outros você acompanha um a um.") \
                        .classes("text-body2 t-suave")
                    gui.button("Escolher nos processos", icon="folder_special",
                               on_click=lambda: self.ir("processos")).props("flat no-caps")

                with gui.card().classes("cartao q-pa-lg w-full"):
                    with gui.row().classes("items-center gap-2"):
                        gui.icon("folder_open", color="primary", size="26px")
                        gui.label("Pastas").classes("text-h6")
                    with gui.row().classes("gap-2"):
                        for rotulo, pasta in (("Processos", main.PROCESSOS), ("Diários baixados", main.DIARIO),
                                              ("Saídas", main.SAIDAS)):
                            gui.button(rotulo, icon="folder",
                                       on_click=lambda p=pasta: abrir(p)).props("outline no-caps")

            # Velocidade dos sites: escondida atrás de um link discreto, para ninguém mexer à toa
            with gui.row().classes("col-span-2 justify-end"):
                gui.button("Opções avançadas", on_click=self._area_perigosa) \
                    .props("flat dense no-caps size=sm color=grey-6")

    def _cartao_pasta_compartilhada(self) -> None:
        """Uma pasta na rede com os processos, as tabelas de consulta e os de-para de todos."""
        raiz = main.pasta_compartilhada()
        with gui.card().classes("cartao q-pa-lg w-full"):
            with gui.row().classes("items-center gap-2"):
                gui.icon("folder_shared", color="primary", size="26px")
                gui.label("Pasta compartilhada").classes("text-h6")
            gui.label("Para todos usarem os mesmos processos, escolha só UMA pasta no servidor: a pasta dos "
                      "processos. Não precisa escolher processo por processo: dentro dela, cada processo fica "
                      "numa pasta com o nome dele (Nomeação, Exoneração...), com o Word, as regras e a "
                      "planilha, e o Edodocs acha cada um sozinho.").classes("text-body2 t-suave")
            gui.label("As tabelas de consulta e os de-para vão junto. O que um muda vale para todos. Senhas e IA "
                      "continuam em cada computador.").classes("text-body2 t-suave")
            if raiz:
                problema = main.problema_na_compartilhada()
                with gui.row().classes("items-center gap-2 no-wrap"):
                    gui.icon("error" if problema else "check_circle", color="negative" if problema else "positive")
                    gui.label(problema or f"Usando: {raiz}").classes("text-body2 t-forte break-all")
                with gui.row().classes("gap-2"):
                    gui.button("Abrir a pasta", icon="folder_open", on_click=lambda: abrir(raiz)) \
                        .props("outline no-caps")
                    gui.button("Trocar de pasta", icon="drive_file_move",
                               on_click=self._escolher_pasta_compartilhada).props("flat no-caps")
                    gui.button("Usar só este computador", icon="computer",
                               on_click=self._sair_da_compartilhada).props("flat no-caps color=grey-8")
            else:
                gui.label("Hoje: tudo fica neste computador.").classes("text-body2 t-forte")
                gui.button("Escolher a pasta dos processos no servidor...", icon="folder_shared",
                           on_click=self._escolher_pasta_compartilhada).props("unelevated no-caps")

    async def _escolher_pasta_compartilhada(self) -> None:
        futuro = asyncio.get_running_loop().create_future()

        def fechar(valor) -> None:
            janela.close()
            if not futuro.done():
                futuro.set_result(valor)

        async def procurar() -> None:
            if napp.native.main_window is None:
                return
            import webview
            tipo = webview.FileDialog.FOLDER if hasattr(webview, "FileDialog") else webview.FOLDER_DIALOG
            escolhida = await napp.native.main_window.create_file_dialog(tipo)
            if escolhida:
                campo.value = escolhida[0]

        with self.camada, gui.dialog().props("persistent") as janela, \
                gui.card().classes("cartao w-[620px] max-w-full q-pa-lg gap-3"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("folder_shared", color="primary", size="30px")
                gui.label("Pasta compartilhada").classes("text-h6 t-forte")
            gui.label("Escolha (ou cole) a pasta do servidor onde ficam os processos, por exemplo "
                      "\\\\servidor\\Edodocs\\processos. É uma pasta só, para todos os processos: cada "
                      "processo fica dentro dela, numa pasta com o nome dele. Use o mesmo caminho em todos "
                      "os computadores.").classes("text-body2 t-suave")
            with gui.row().classes("w-full items-center gap-2 no-wrap"):
                campo = gui.input("Pasta", value=str(main.pasta_compartilhada() or "")) \
                    .props("outlined dense clearable").classes("col")
                if napp.native.main_window is not None:
                    gui.button("Procurar...", icon="folder_open", on_click=procurar).props("flat no-caps")
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=lambda: fechar(None)).props("flat no-caps")
                gui.button("Usar esta pasta", icon="check",
                           on_click=lambda: fechar((campo.value or "").strip().strip('"'))) \
                    .props("unelevated no-caps")
        janela.open()
        escolhida = await futuro
        if not escolhida:
            return
        raiz = Path(escolhida)
        if not await run.io_bound(raiz.is_dir):
            gui.notify("Não encontrei essa pasta. Confira o caminho e se você tem acesso a ela.", type="negative")
            return
        vazia = not await run.io_bound(lambda: any(main._arquivos(p, "*.docx")
                                                   for p in main.processos_da(raiz).glob("*") if p.is_dir()))
        levar = False
        if vazia and main.listar_processos() and main.pasta_compartilhada() is None:
            levar = await self._perguntar_levar(raiz)
            if levar is None:
                return
        try:
            levados = await run.io_bound(main.levar_para_compartilhada, raiz) if levar else []
        except OSError as e:
            gui.notify(f"Não consegui gravar na pasta ({e.strerror or e}). Você tem permissão de escrita nela?",
                       type="negative", multi_line=True)
            return
        main.salvar_config(pasta_compartilhada=str(raiz))
        main.aplicar_pasta_compartilhada()
        gui.notify(f"Usando a pasta compartilhada {raiz}" + (f". Levei: {', '.join(levados)}" if levados else ""),
                   type="positive", multi_line=True, timeout=8000)
        gui.timer(1.5, gui.navigate.reload, once=True)

    async def _perguntar_levar(self, raiz: Path) -> bool | None:
        """Pasta compartilhada ainda vazia: levar os processos deste computador? None: cancelou."""
        futuro = asyncio.get_running_loop().create_future()

        def fechar(valor) -> None:
            janela.close()
            if not futuro.done():
                futuro.set_result(valor)

        with self.camada, gui.dialog().props("persistent") as janela, \
                gui.card().classes("cartao w-[560px] max-w-full q-pa-lg gap-3"):
            gui.label("A pasta ainda não tem processos").classes("text-h6 t-forte")
            gui.label(f"Levar para {raiz} os processos deste computador ("
                      f"{', '.join(p.name for p in main.listar_processos())}), junto com as tabelas de consulta "
                      "e os de-para? Cada processo vai para uma pasta com o nome dele, e os outros já "
                      "encontram tudo pronto.").classes("text-body2 t-forte")
            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=lambda: fechar(None)).props("flat no-caps")
                gui.button("Não, começar vazia", on_click=lambda: fechar(False)).props("flat no-caps")
                gui.button("Sim, levar tudo", icon="drive_file_move", on_click=lambda: fechar(True)) \
                    .props("unelevated no-caps")
        janela.open()
        return await futuro

    def _sair_da_compartilhada(self) -> None:
        main.salvar_config(pasta_compartilhada="")
        main.aplicar_pasta_compartilhada()
        gui.notify("Voltando a usar só os processos deste computador", type="info")
        gui.timer(1.5, gui.navigate.reload, once=True)

    def _area_perigosa(self) -> None:
        """Janela com as pausas entre requisições e entre atos (abre pelo link discreto)."""
        cfg = main.ler_config()
        with self.camada, gui.dialog() as dialogo, gui.card().classes("cartao w-[520px] max-w-full q-pa-lg"):
            with gui.row().classes("items-center gap-2"):
                gui.icon("warning", color="negative", size="26px")
                gui.label("Velocidade dos sites").classes("text-h6")
            gui.label("Pausas menores deixam mais rápido, mas o site pode bloquear a automação "
                      "(erro 403 / Access Denied). Em dúvida, use o padrão.").classes("text-body2 t-forte")
            with gui.row().classes("w-full gap-4 no-wrap"):
                req = gui.number(f"Entre requisições (padrão {PAUSA_REQUISICAO_PADRAO:g}s)", min=0, step=0.05,
                                 format="%.2f",
                                 value=float(cfg.get("pausa_requisicao", PAUSA_REQUISICAO_PADRAO))) \
                    .props("outlined suffix=s").classes("col")
                entre = gui.number(f"Entre atos (padrão {PAUSA_ENTRE_ATOS_PADRAO:g}s)", min=0, step=0.25,
                                   format="%.2f",
                                   value=float(cfg.get("pausa_entre_atos", PAUSA_ENTRE_ATOS_PADRAO))) \
                    .props("outlined suffix=s").classes("col")
            with gui.row().classes("w-full items-center"):
                gui.button("Voltar ao padrão", on_click=lambda: (
                    req.set_value(PAUSA_REQUISICAO_PADRAO), entre.set_value(PAUSA_ENTRE_ATOS_PADRAO))) \
                    .props("flat no-caps")
                gui.space()
                gui.button("Cancelar", on_click=dialogo.close).props("flat")
                gui.button("Salvar", on_click=lambda: self._salvar_pausas(req.value, entre.value)
                           and dialogo.close())
        dialogo.open()

    def _desenhar_usuarios(self) -> None:
        self.cfg_usuarios.clear()
        with self.cfg_usuarios:
            for site, nome in SITES.items():
                cor, usuario, dica = self.login[site]
                with gui.column().classes("w-full gap-2"):
                    with gui.row().classes("items-center gap-2"):
                        _bolinha(cor)
                        gui.label(nome).classes("text-subtitle1 text-weight-medium")
                        gui.label(f"· {usuario or 'sem usuário'} · {dica}").classes("t-suave")
                    with gui.row().classes("w-full gap-2 items-center no-wrap"):
                        campo_u = gui.input("Usuário", value=usuario).props("outlined dense").classes("col")
                        campo_s = gui.input("Senha", password=True, password_toggle_button=True) \
                            .props("outlined dense").classes("col")
                        gui.button("Salvar e testar", icon="login",
                                   on_click=lambda s=site, u=campo_u, p=campo_s: self._salvar_usuario(s, u, p)) \
                            .props("no-caps")
                        gui.button(icon="refresh", on_click=lambda s=site: self.verificar_logins(s)) \
                            .props("flat round").tooltip("Testar o login de novo")

    async def _salvar_usuario(self, site: str, campo_u, campo_s) -> None:
        if not campo_u.value.strip() or not campo_s.value:
            gui.notify("Preencha usuário e senha", type="warning")
            return
        credenciais.salvar(main.CREDENCIAIS, campo_u.value.strip(), campo_s.value, site)
        gui.notify(f"Usuário do {SITES[site]} salvo. Testando o login...")
        await self.verificar_logins(site)
        cor, _, dica = self.login[site]
        # verificar_logins redesenha a área e apaga o botão clicado: avisar a partir da área
        with self.cfg_usuarios:
            gui.notify(f"{SITES[site]}: {dica}", type="positive" if cor == VERDE else "negative")

    def _editar_planilha(self, pasta: Path) -> None:
        """Onde fica a planilha do processo: a da pasta dele ou outra (no computador ou na rede).
        A escolhida fica em config.json → "planilhas" (de cada computador)."""
        proc = processos.carregar(pasta)
        personalizada = main.ler_config().get("planilhas", {}).get(pasta.name)
        # A planilha em uso de verdade: a escolhida aqui ou a definida ao criar o processo
        em_uso = Path(personalizada) if personalizada else proc.planilha_padrao()
        origem = ("escolhida aqui" if personalizada
                  else "definida ao criar o processo" if proc["planilha"]
                  else "a planilha da pasta do processo")
        with self.camada, gui.dialog() as janela, gui.card().classes("cartao w-[680px] max-w-full q-pa-lg gap-2"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("table_chart", color="primary", size="28px")
                with gui.column().classes("gap-0"):
                    gui.label("Planilha do processo").classes("text-h6")
                    gui.label(pasta.name).classes("text-caption t-suave")
            with gui.row().classes("items-start gap-2 no-wrap"):
                existe = bool(em_uso and em_uso.is_file())
                gui.icon("check_circle" if existe else "error", color="positive" if existe else "negative",
                         size="20px")
                with gui.column().classes("gap-0"):
                    gui.label(str(em_uso) if em_uso else "Nenhuma planilha encontrada") \
                        .classes("text-body2 break-all t-forte")
                    gui.label(f"Em uso · {origem}" if existe else
                              "Não encontrei este arquivo: escolha a planilha abaixo") \
                        .classes("text-caption " + ("t-suave" if existe else "text-negative"))
            with gui.row().classes("items-start gap-2 no-wrap q-mt-xs"):
                gui.icon("dns", color="primary", size="20px")
                gui.label("A planilha pode ficar no seu computador ou numa pasta da rede/servidor, "
                          "desde que você tenha acesso a ela — assim a equipe trabalha na mesma "
                          "planilha. Se alguém estiver com ela aberta no Excel, o Edodocs espera "
                          "fechar antes de salvar.").classes("text-caption t-forte")
            caminho = gui.input("Caminho da planilha (.xlsx), no computador ou na rede",
                                value=str(em_uso) if em_uso else "",
                                placeholder=r"\\servidor\pasta\planilha.xlsx  ou  S:\pasta\planilha.xlsx") \
                .props("outlined dense clearable stack-label").classes("w-full")
            with gui.row().classes("w-full items-center gap-2"):
                if napp.native.main_window is not None:
                    gui.button("Escolher arquivo...", icon="attach_file",
                               on_click=lambda: self._escolher_planilha(caminho)).props("outline no-caps")
                if personalizada:
                    gui.button("Usar a da pasta", on_click=lambda: self._salvar_planilha(pasta, "", janela)) \
                        .props("flat no-caps")
                gui.space()
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Usar este", icon="check",
                           on_click=lambda: self._salvar_planilha(pasta, caminho.value, janela)).props("no-caps")
        janela.open()

    async def _escolher_planilha(self, campo) -> None:
        import webview
        tipo = webview.FileDialog.OPEN if hasattr(webview, "FileDialog") else webview.OPEN_DIALOG
        arquivos = await napp.native.main_window.create_file_dialog(
            tipo, file_types=("Planilhas do Excel (*.xlsx)",))
        if arquivos:
            campo.value = arquivos[0]

    def _salvar_planilha(self, pasta: Path, caminho: str | None, janela) -> None:
        planilhas = main.ler_config().get("planilhas", {})
        caminho = (caminho or "").strip().strip('"')
        if caminho:
            arquivo = Path(caminho)
            if arquivo.suffix.lower() != ".xlsx":
                gui.notify("O arquivo precisa ser uma planilha .xlsx", type="negative")
                return
            if not arquivo.is_file():
                gui.notify("Não encontrei o arquivo. Se estiver na rede, confira o caminho e se "
                           "você tem acesso à pasta.", type="negative", multi_line=True)
                return
            raiz = main.pasta_compartilhada()
            if raiz and not str(arquivo).startswith("\\\\") and raiz.resolve() not in arquivo.resolve().parents:
                gui.notify("Atenção: essa planilha está neste computador, e os outros não vão enxergar. Para "
                           "todos usarem a mesma, coloque-a na pasta compartilhada ou num caminho da rede.",
                           type="warning", multi_line=True, timeout=10000)
            planilhas[pasta.name] = str(arquivo)
        else:
            planilhas.pop(pasta.name, None)
        main.salvar_config(planilhas=planilhas)
        gui.notify(f"Planilha de {pasta.name} alterada", type="positive")
        janela.close()
        self._desenhar_processos()
        if pasta == main.processo:  # a tabela dos atos é a do processo em uso
            gui.timer(0.01, self.recarregar, once=True)

    @staticmethod
    def _salvar_pausas(req, entre) -> bool:
        if req is None or entre is None or req < 0 or entre < 0:
            gui.notify("Use números positivos (ex.: 0.25)", type="negative")
            return False
        main.salvar_config(pausa_requisicao=float(req), pausa_entre_atos=float(entre))
        if req < PAUSA_REQUISICAO_PADRAO or entre < PAUSA_ENTRE_ATOS_PADRAO:
            gui.notify("Salvo. Valores abaixo do padrão: fique de olho se o site começar a recusar",
                       type="warning")
        else:
            gui.notify("Pausas salvas", type="positive")
        return True


ATALHO_NOVO = False  # o atalho da área de trabalho acabou de ser criado (avisa uma vez)


def construir() -> None:
    global ATALHO_NOVO
    Tela()
    if ATALHO_NOVO:
        ATALHO_NOVO = False
        gui.timer(2.0, lambda: gui.notify(
            "Criei um atalho do Edodocs na sua área de trabalho: da próxima vez, é só abrir por ele.",
            type="positive", timeout=9000, icon="desktop_windows"), once=True)


# --- Abertura do programa -----------------------------------------------------------
# A janela própria usa o Microsoft Edge WebView2. Em computador sem ele (ou se a janela não
# abrir por outro motivo), o Edodocs abre no navegador padrão, em vez de ficar rodando
# escondido. E só roda um Edodocs por vez: um 2º clique traz para a frente o que já está aberto.

INSTANCIA = main.BASE / "configs" / "instancia.json"
FECHAR_SEM_ABAS = 120  # segundos sem nenhuma aba aberta (no navegador) até fechar sozinho
_MUTEX = None
_WEBVIEW2 = r"Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"


def tem_webview2() -> bool:
    """O Microsoft Edge WebView2 (necessário para a janela própria) está instalado?"""
    import winreg

    locais = [(winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\WOW6432Node\\" + _WEBVIEW2),
              (winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\" + _WEBVIEW2),
              (winreg.HKEY_CURRENT_USER, "Software\\" + _WEBVIEW2)]
    for raiz, caminho in locais:
        try:
            with winreg.OpenKey(raiz, caminho) as chave:
                versao = winreg.QueryValueEx(chave, "pv")[0]
            if versao and versao != "0.0.0.0":
                return True
        except OSError:
            continue
    return False


def _janela_aberta() -> int:
    """Identificador da janela própria do Edodocs (0 se ela não estiver na tela)."""
    import ctypes

    return ctypes.windll.user32.FindWindowW(None, TITULO)


def _trazer_para_frente(janela: int) -> None:
    import ctypes

    ctypes.windll.user32.ShowWindow(janela, 9)  # SW_RESTORE (se estiver minimizada)
    ctypes.windll.user32.SetForegroundWindow(janela)


def _mensagem(texto: str) -> None:
    import ctypes

    ctypes.windll.user32.MessageBoxW(None, texto, TITULO, 0x40)  # ícone de informação


def _porta_livre() -> int:
    import socket

    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _ja_aberto() -> bool:
    """Se já há um Edodocs rodando, mostra ele e devolve True (este não abre outro)."""
    import ctypes
    import webbrowser

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    # O "mutex" fica com este processo até ele fechar (por isso fica guardado no módulo)
    global _MUTEX
    _MUTEX = kernel32.CreateMutexW(None, False, "Local\\Edodocs")
    if ctypes.get_last_error() != 183:  # ERROR_ALREADY_EXISTS
        return False
    janela = _janela_aberta()
    if janela:
        _trazer_para_frente(janela)
        return True
    try:  # o outro está no navegador: abre de novo o endereço dele
        webbrowser.open(json.loads(INSTANCIA.read_text(encoding="utf-8"))["endereco"])
    except (OSError, ValueError, KeyError):
        _mensagem("O Edodocs já está aberto. Se não aparecer, feche-o pelo Gerenciador de "
                  "Tarefas e abra de novo.")
    return True


def _abrir_no_navegador(endereco: str) -> None:
    """Abre no navegador padrão e encerra o Edodocs quando ninguém estiver usando (sem isso
    ele ficaria rodando escondido depois de fechar a aba)."""
    import webbrowser

    webbrowser.open(endereco)
    conectados = {"n": 0}

    def entrou(_) -> None:
        conectados["n"] += 1

    async def saiu(_) -> None:
        conectados["n"] -= 1
        await asyncio.sleep(FECHAR_SEM_ABAS)  # recarregar a página reconecta logo
        if conectados["n"] <= 0:
            print("Nenhuma aba aberta: fechando o Edodocs")
            napp.shutdown()

    napp.on_connect(entrou)
    napp.on_disconnect(saiu)


def iniciar() -> None:
    global ATALHO_NOVO
    if _ja_aberto():
        return
    main.preparar_pastas()
    try:  # Edodocs novo: as tabelas de consulta e os de-para que vieram no pacote
        import consultas
        for item in consultas.receber_do_pacote():
            print(f"Veio no pacote: {item}")
    except Exception as e:
        print(f"Não consegui ler o que veio no pacote: {e}")
    # Primeira vez neste computador (ou programa movido de pasta): atalho na área de trabalho
    try:
        import atalho
        criado = atalho.criar_se_preciso(main.ler_config().get("atalho_para"))
        if criado:
            main.salvar_config(atalho_para=criado)
            ATALHO_NOVO = True
            print(f"Atalho criado na área de trabalho para {criado}")
    except Exception as e:
        print(f"Não consegui criar o atalho na área de trabalho: {type(e).__name__} {e}")
    porta = _porta_livre()
    endereco = f"http://127.0.0.1:{porta}/"
    try:
        import webview  # noqa: F401  (janela própria)
        nativo = tem_webview2()
    except ImportError:
        nativo = False
    print(f"Endereço {endereco} · janela própria: {'sim' if nativo else 'não (sem WebView2)'}")
    try:
        INSTANCIA.parent.mkdir(parents=True, exist_ok=True)
        INSTANCIA.write_text(json.dumps({"endereco": endereco, "pid": os.getpid()}), encoding="utf-8")
    except OSError:
        pass

    if nativo:
        async def conferir_janela() -> None:
            # A janela abre em poucos segundos. Se não aparecer, vai para o navegador.
            for _ in range(40):
                await asyncio.sleep(0.5)
                if janela := _janela_aberta():
                    _trazer_para_frente(janela)
                    return
            print("A janela própria não abriu: abrindo no navegador")
            _abrir_no_navegador(endereco)
        napp.on_startup(conferir_janela)
    else:
        napp.on_startup(lambda: _abrir_no_navegador(endereco))

    try:
        # window_size liga a janela própria no NiceGUI: só quando ela for usada
        gui.run(construir, title=TITULO, favicon=ICONE_JANELA, native=nativo,
                window_size=(1400, 880) if nativo else None,
                reload=False, language="pt-BR", show_welcome_message=False, port=porta, show=False)
    except Exception as e:
        import traceback
        traceback.print_exc()
        _mensagem(f"O Edodocs não conseguiu abrir: {type(e).__name__}: {e}\n\n"
                  f"Envie o arquivo {LOG or 'configs/edodocs.log'} para quem cuida do programa.")


if __name__ == "__main__":
    # A janela própria roda num processo separado: no executável, sem isso ele abriria o
    # programa de novo em vez da janela
    import multiprocessing
    multiprocessing.freeze_support()
    iniciar()
    # Fechou a janela (ou a última aba): sai de vez. Sem isso, alguma tarefa em segundo
    # plano pode deixar o processo vivo, escondido no Gerenciador de Tarefas.
    sys.stdout.flush()
    os._exit(0)
