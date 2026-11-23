"""Assistente de processo (janela do app.py), feito para quem nunca montou um modelo:

    0. Antes de começar    o que separar, arquivos de exemplo e um treino guiado (só no novo)
    1. Tipo de ato         o nome do processo e se cada ato é de uma pessoa ou de várias
    2. O Word              um ato que a pessoa já fez. Sem planilha, ela marca no texto o que
                           muda (marcar_ui.py); com a planilha, o programa acha os dados sozinho
                           (modelo_automatico.py). No ato de várias pessoas, ela clica no
                           parágrafo de uma pessoa e o programa prepara a repetição.
    3. A planilha          "Já tenho uma planilha" ou "Criar a planilha para mim"
                           (planilha_nova.py: colunas dos campos, listas SIM/NÃO, datas)
    4. Confira os campos   cada campo com um exemplo de valor; quase tudo já vem ligado.
                           Campos que mudam conforme a planilha (regras "Se... escrever...")
                           e atos com várias pessoas (outros Words, trecho que repete) também
                           ficam aqui (janelas em regras_ui.py)
    5. Como vai ficar      o ato completo de qualquer pessoa da planilha, os títulos (tirados
                           da primeira linha do Word) e Salvar

Tudo fica salvo no processo.json da pasta do processo (ver processos.py). Os botões ⓘ abrem
as explicações com exemplos (ajudas.py).
"""
import asyncio
import html
import json
import re
import shutil
import tempfile
import time
from dataclasses import asdict
from datetime import date
from pathlib import Path

from nicegui import app as napp
from nicegui import background_tasks
from nicegui import run
from nicegui import ui as gui

from abrir import abrir
import atos
import diario_http
import main
import modelo_automatico
import marcar_ui
import planilha_nova
import processos
import simulacao_ui
import regras_ui
from ajudas import AJUDAS
from colunas_ui import EditorColunas
from escolha_pessoas import PainelPessoas

INVALIDOS = re.compile(r'[<>:"/\\|?*]')
# O processo novo que ainda não foi salvo: fechar a janela no meio não perde nada
RASCUNHO = main.ENTRADAS / "rascunho do processo"
NOMES_PASSOS = {"inicio": "Antes de começar", "tipo": "Tipo de ato", "ato": "O Word", "planilha": "A planilha",
                "ligar": "Confira os campos", "previa": "Como vai ficar"}
AJUDA_WORD = [
    "Escreva o ato normalmente no Word, já com a formatação que deve sair.",
    "Onde entra um dado da planilha, escreva o nome da coluna entre chaves duplas: "
    "{{NOME}}, {{MATRÍCULA}}, {{CARGO}}.",
    "Deixe o título do ato na primeira linha.",
    "Campos de mala direta do Word («NOME») também funcionam.",
]


# --- Exemplos para ensinar --------------------------------------------------------

def gerar_exemplos(pasta: Path) -> tuple[Path, Path]:
    """Cria 'Exemplo - modelo.docx' e 'Exemplo - planilha.xlsx' que combinam entre si."""
    import openpyxl
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    pasta.mkdir(parents=True, exist_ok=True)
    doc = Document()
    estilo = doc.styles["Normal"]
    estilo.font.name, estilo.font.size = "Arial", Pt(12)

    def paragrafo(partes, centro=False, recuo=True):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if centro else WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(12)
        if recuo and not centro:
            p.paragraph_format.first_line_indent = Pt(42.5)
        for texto, negrito in partes:
            p.add_run(texto).bold = negrito

    paragrafo([("ATO ADMINISTRATIVO SMA Nº {{Nº DO ATO}}", True)], centro=True)
    paragrafo([("O DIRETOR ADMINISTRATIVO DE PESSOAL", True),
               (", no exercício de suas atribuições legais,", False)])
    paragrafo([("RESOLVE:", True)])
    paragrafo([("a) Conceder {{DIAS}} ({{DIAS POR EXTENSO}}) dia(s) de licença {{AO A SERVIDOR}} "
                "{{NOME}}, matrícula nº {{MATRÍCULA}}, ocupante do cargo de {{CARGO}}, "
                "{{LOTADO}} na {{SECRETARIA}}.", False)])
    paragrafo([("b) Este ato administrativo entra em vigor na data de sua publicação.", False)])
    paragrafo([("Uberlândia, (data da publicação).", False)])
    modelo = pasta / "Exemplo - modelo.docx"
    doc.save(modelo)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "atos"
    ws.append(["Nº DO ATO", "NOME", "MATRÍCULA", "CARGO", "DIAS", "SEXO", "SECRETARIA"])
    ws.append(["101/2026", "MARIA DA SILVA", "12.345-6", "AUXILIAR ADMINISTRATIVO", 15, "F",
               "Secretaria Municipal de Saúde"])
    ws.append(["102/2026", "JOÃO SOUZA", "23.456-7", "MOTORISTA", 30, "M",
               "Secretaria Municipal de Obras"])
    ws.append(["103/2026", "ANA LIMA", "34.567-8", "PROFESSORA", 10, "F",
               "Secretaria Municipal de Educação"])
    for coluna, largura in zip("ABCDEFG", (14, 22, 12, 28, 8, 8, 34)):
        ws.column_dimensions[coluna].width = largura
    planilha = pasta / "Exemplo - planilha.xlsx"
    wb.save(planilha)
    return modelo, planilha


# Como o modelo de exemplo liga os campos que não são colunas da planilha
LIGACOES_EXEMPLO = {
    "DIAS POR EXTENSO": {"tipo": "extenso", "coluna": "DIAS"},
    "AO A SERVIDOR": {"tipo": "genero", "coluna": "SEXO", "feminino": "à servidora",
                      "masculino": "ao servidor"},
    "LOTADO": {"tipo": "genero", "coluna": "SEXO", "feminino": "lotada", "masculino": "lotado"},
}


def gerar_ato_pronto_exemplo(pasta: Path) -> tuple[Path, Path]:
    """(planilha, ato pronto da MARIA): o que a pessoa teria na mão para criar um processo."""
    modelo, planilha = gerar_exemplos(pasta)
    proc = processos.Processo(pasta, {"modelo": str(modelo), "ligacoes": LIGACOES_EXEMPLO})
    registros = atos.ler_planilha(planilha, "atos", "NOME")
    doc, sect, bloco = atos._abrir_modelo(modelo)
    copia, _ = atos._preencher(bloco, registros[0], proc.valores)
    for elemento in copia:
        sect.addprevious(elemento)
    ato = pasta / "Exemplo - ato pronto da Maria.docx"
    doc.save(ato)
    return planilha, ato


# --- Leitura dos arquivos (roda fora da tela) ---------------------------------------

def _ler_planilha(planilha: Path, aba: str, coluna_nome: str) -> dict:
    abas, colunas = atos.cabecalho_planilha(planilha, aba or None)
    colunas = list(dict.fromkeys(colunas))  # cabeçalhos repetidos aparecem uma vez só
    registros = atos.ler_planilha(planilha, aba or None, atos.normalizar(coluna_nome)) if coluna_nome else []
    return {"abas": abas, "colunas": colunas, "registros": registros}


def ler_rascunho() -> dict | None:
    try:
        return json.loads((RASCUNHO / "rascunho.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def apagar_rascunho() -> None:
    shutil.rmtree(RASCUNHO, ignore_errors=True)


def _coluna_igual(colunas: list[str], nome: str) -> str | None:
    """A coluna da planilha com o mesmo nome (sem diferenciar acento, espaço, maiúscula)."""
    alvo = atos.normalizar(nome)
    return next((c for c in colunas if atos.normalizar(c) == alvo), None)


def _passo(numero: int, titulo: str, explicacao: str, ajuda=None) -> None:
    """Cabeçalho de um passo: número grande, o que fazer e por quê, em linguagem simples.
    ajuda: o botão ⓘ com a explicação detalhada (ver Assistente._info)."""
    with gui.row().classes("items-start gap-3 no-wrap q-mb-sm"):
        gui.label(str(numero)).classes("text-white text-weight-bold text-center text-h6") \
            .style("background:#2459a8;width:36px;height:36px;line-height:36px;border-radius:50%;flex-shrink:0")
        with gui.column().classes("gap-0"):
            with gui.row().classes("items-center gap-1 no-wrap"):
                gui.label(f"Passo {numero} de 5 · {titulo}").classes("text-h6 t-forte")
                if ajuda:
                    ajuda()
            gui.label(explicacao).classes("text-body2 t-suave")


def _conferencia(ok: bool | None, texto: str) -> None:
    """Uma linha da conferência: ✓ (ok), ⚠ (atenção) ou ⓘ (só informação, ok=None)."""
    icone, cor = {True: ("check_circle", "positive"), False: ("warning", "orange-8"),
                  None: ("info", "primary")}[ok]
    with gui.row().classes("items-start gap-2 no-wrap"):
        gui.icon(icone, color=cor, size="20px").classes("q-mt-xs")
        gui.label(texto).classes("text-body2 t-forte")


class Assistente:
    """Janela de criar/editar processo. pasta=None cria um processo novo."""

    def __init__(self, tela, pasta: Path | None = None, rascunho: dict | None = None):
        """rascunho: o processo novo guardado (ler_rascunho), para continuar de onde parou."""
        self.tela = tela
        self.pasta = pasta
        proc = processos.carregar(pasta) if pasta else processos.Processo(Path("."))
        self.config = proc.config
        self.nome = pasta.name if pasta else ""
        self.modelo: Path | None = proc.modelo() if pasta else None
        self.planilha: Path | None = proc.planilha_padrao() if pasta else None
        self.campos: list[tuple[str, str]] = []
        self.abas: list[str] = []
        self.colunas: list[str] = []
        self.registros: list[dict] = []
        self.achados: list = []          # dados da planilha achados no ato pronto
        self.ato_pronto: Path | None = None
        self.linha_ato: dict | None = None
        self.linha_previa = 0
        self.treino = False              # treino com o exemplo: nada é salvo
        self.ato_exemplo: Path | None = None
        # "tenho" (escolheu uma planilha) ou "criar" (o programa cria); None = ainda não escolheu
        self.modo_planilha: str | None = "tenho" if self.planilha else None
        self.exemplo: dict | None = None  # os dados que estavam no Word marcado (para a prévia)
        self.word: Path | None = None     # o Word escolhido no passo 2 (antes de virar modelo)
        self.marcador: marcar_ui.MarcadorWord | None = None
        self._temp = Path(tempfile.mkdtemp(prefix="edodocs_modelo_"))
        self._retomado = bool(rascunho)
        self._marcas_rascunho: list[dict] = []
        self._passo_rascunho: str | None = None
        self._ia: dict | None = None             # montado pela IA (processo_ia.py): o que contar
        self._pessoas_exemplo: list[dict] = []   # os exemplos que a IA leu, para a prévia
        self._assinatura_rascunho = ""
        self._salvo = False
        if rascunho:
            self._retomar(rascunho)
        self._montar()
        background_tasks.create(self._carregar_existente())

    # --- Montagem da janela ---
    def _montar(self) -> None:
        with self.tela.camada, gui.dialog().props("maximized persistent") as self.dialogo, \
                gui.card().classes("w-full h-full q-pa-none fundo-assistente"):
            with gui.row().classes("w-full items-center q-px-lg q-py-md text-white").style(
                    "background: linear-gradient(90deg,#1d4a8f,#2f6fcf);"):
                gui.icon("folder_special", size="28px")
                gui.label(f"Editar processo: {self.nome}" if self.pasta else "Novo processo") \
                    .classes("text-h6")
                gui.space()
                gui.button("Como funciona?", icon="help_outline", on_click=lambda: self._ajuda("comecar")) \
                    .props("flat no-caps color=white")
                gui.button(icon="close", on_click=self.dialogo.close).props("flat round color=white")
            if not self.pasta:  # processo novo: guarda o que já foi feito a cada poucos segundos
                self._timer_rascunho = gui.timer(2.0, self._guardar_rascunho)
                self.dialogo.on_value_change(lambda e: e.value or self._ao_fechar())
            with gui.scroll_area().classes("w-full col"), gui.row().classes("w-full no-wrap items-start gap-0"):
                with gui.stepper().props("animated header-nav").classes("col bg-transparent") as self.passos:
                    if not self.pasta:
                        with gui.step("inicio", title="Antes de começar", icon="flag"):
                            self._passo_inicio()
                    with gui.step("tipo", title="Tipo de ato", icon="category"):
                        self._passo_tipo()
                    with gui.step("ato", title="O Word", icon="description"):
                        self._passo_ato()
                    with gui.step("planilha", title="A planilha", icon="table_chart"):
                        self._passo_planilha()
                    with gui.step("ligar", title="Confira os campos", icon="link"):
                        self._passo_ligar()
                    with gui.step("previa", title="Como vai ficar", icon="visibility"):
                        self._passo_previa()
                # Ao lado de cada passo, um Word de mentira mostrando como o ato está ficando
                # Afastado das etapas e alinhado com o topo delas
                with gui.column().classes("gap-2") \
                        .style("width: 310px; flex-shrink: 0; position: sticky; top: 0; "
                               "margin: 16px 24px 0 28px") as self.painel_sim:
                    with gui.card().classes("cartao w-full q-pa-md gap-2"):
                        with gui.row().classes("items-center gap-2 no-wrap"):
                            gui.icon("auto_stories", color="primary", size="22px")
                            gui.label("Como o seu ato está ficando").classes("text-subtitle2 t-forte")
                        self.sim_legenda = gui.label("").classes("text-caption t-suave")
                        self.sim_desenho = gui.html("", sanitize=False).classes("w-full")
        self.passos.on_value_change(lambda e: self._atualizar_word_simulado())
        self._atualizar_word_simulado()
        self.dialogo.open()

    # --- Word simulado ao lado dos passos (simulacao_ui.py) ---------------------------------
    def _campos_da_simulacao(self) -> list[str]:
        """O que já foi marcado: as marcas da tela de marcar ou, depois, os campos do Word."""
        if self.marcador and self.marcador.marcas and not self.campos:
            return [m.campo for m in self.marcador.marcas]
        return [r for _, r in self.campos]

    def _da_pessoa_na_simulacao(self, campos: list[str]) -> set[str]:
        """Os campos que repetem para cada pessoa (para ir no quadro azul)."""
        try:
            nomes = set()
            for modelo in self._processo_rascunho().modelos():
                nomes |= atos.campos_que_repetem(modelo)
        except Exception:
            return set()
        return {c for c in campos if atos.normalizar(c) in nomes}

    def _atualizar_word_simulado(self) -> None:
        if not hasattr(self, "sim_desenho"):
            return
        passo = self.passos.value
        # No começo ainda não há ato; na prévia, o ato de verdade já aparece
        self.painel_sim.set_visibility(passo not in ("inicio", "previa"))
        escolheu = bool(self.config.get("_tipo") or self.pasta)
        varias = self.varias
        campos = self._campos_da_simulacao()
        legenda, desenho = "", ""
        if passo == "tipo":
            if not escolheu:
                legenda = "Escolha um dos dois tipos: aqui aparece como o seu ato vai ser."
                desenho = simulacao_ui.pagina(False)
            elif varias:
                legenda = "Um ato com várias pessoas: a parte em azul se repete para cada uma (I, II, III...)."
                desenho = simulacao_ui.pagina(True)
            else:
                legenda = "Um ato por pessoa: o nome (e os outros dados) mudam de um ato para outro."
                desenho = simulacao_ui.pagina(False)
        elif passo == "ato":
            da_pessoa = self._da_pessoa_na_simulacao(campos) if varias and self.modelo else set()
            if not campos:
                legenda = "Marque no texto o que muda de um ato para outro: aparece aqui em amarelo."
                desenho = simulacao_ui.pagina(varias)
            elif varias and self.config.get("outros_modelos"):
                legenda = "Em amarelo, o que muda. Com uma pessoa usa o 1º Word; com 2 ou mais, o 2º."
                desenho = ('<div style="font-size:11px;font-weight:600;margin:2px 0">Com 1 pessoa</div>'
                           + simulacao_ui.pagina(False, campos)
                           + '<div style="font-size:11px;font-weight:600;margin:10px 0 2px">Com 2 ou mais</div>'
                           + simulacao_ui.pagina(True, campos, da_pessoa))
            else:
                legenda = f"Em amarelo, os {len(campos)} campo(s) que mudam: cada um vira uma coluna da planilha."
                desenho = simulacao_ui.pagina(varias, campos, da_pessoa)
        elif passo == "planilha":
            colunas = self.colunas if self.planilha and self.colunas else                 [c["nome"] for c in self._colunas_novas()]
            legenda = "Cada pessoa é uma linha da planilha."
            desenho = simulacao_ui.planilha_para_atos(varias, colunas)
        elif passo == "ligar":
            proc = self._processo_rascunho()
            pares = []
            for _, rotulo in self.campos:
                regra = proc.ligacao(rotulo)
                origem = regra.get("coluna") or processos.TIPOS.get(regra.get("tipo", ""), "regra")
                pares.append((rotulo, origem, self._ligada(regra)))
            legenda = "De onde vem cada campo do Word (verde: tudo certo)."
            desenho = simulacao_ui.ligacoes(pares) or simulacao_ui.pagina(varias)
        self.sim_legenda.text = legenda
        self.sim_desenho.content = desenho

    # --- Ajuda (botões ⓘ) --------------------------------------------------------------
    def _ajuda(self, chave: str) -> None:
        """Janela com a explicação detalhada e exemplos (textos em ajudas.py)."""
        titulo, texto = AJUDAS[chave]
        with self.tela.camada, gui.dialog() as janela, gui.card().classes("cartao w-[720px] max-w-full q-pa-lg"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("lightbulb", color="amber-7", size="28px")
                gui.label(titulo).classes("text-h6 t-forte")
            with gui.element("div").classes("w-full").style("max-height: 62vh; overflow-y: auto"):
                gui.markdown(texto).classes("ajuda-texto w-full")
            with gui.row().classes("w-full justify-end"):
                gui.button("Entendi", icon="check", on_click=janela.close).props("unelevated no-caps")
        janela.open()

    def _info(self, chave: str) -> None:
        """Botão ⓘ que abre a ajuda."""
        gui.button(icon="info_outline", on_click=lambda: self._ajuda(chave)) \
            .props("flat round dense size=sm color=primary").tooltip("O que é isso? Clique para ver exemplos")

    # --- Passo 0: antes de começar (só no processo novo) ------------------------------------
    def _passo_inicio(self) -> None:
        with gui.column().classes("w-full gap-4 q-pl-md"):
            with gui.row().classes("items-center gap-3 no-wrap"):
                gui.icon("waving_hand", color="#39ff14", size="36px")
                with gui.column().classes("gap-0"):
                    gui.label("Vamos criar o seu processo").classes("text-h5 t-forte")
                    gui.label("Você ensina o Edodocs uma vez, com um ato que já fez, e depois ele monta e "
                              "envia os próximos sozinho. Leva uns 5 minutos.").classes("text-body1 t-suave")
            gui.label("Você vai precisar de:").classes("text-subtitle1 text-weight-medium t-forte")
            with gui.row().classes("w-full gap-4"):
                for icone, titulo, texto, ajuda in (
                        ("description", "Um ato que você já fez (Word)",
                         "Do jeito que foi publicado. É dele que o Edodocs aprende.", "ato"),
                        ("table_chart", "A planilha (Excel): só se já tiver",
                         "Se ainda não tem, o Edodocs cria uma para você, com listas para escolher "
                         "(sim/não...) em vez de digitar.", "planilha")):
                    with gui.card().classes("cartao q-pa-md col").style("min-width: 300px"):
                        with gui.row().classes("items-center gap-2 no-wrap"):
                            gui.icon(icone, color="primary", size="28px")
                            gui.label(titulo).classes("text-subtitle1 text-weight-medium t-forte")
                            gui.space()
                            self._info(ajuda)
                        gui.label(texto).classes("text-body2 t-suave")
            with gui.card().classes("w-full fundo-cartao q-pa-md"):
                with gui.row().classes("items-center gap-3 no-wrap"):
                    gui.icon("school", color="primary", size="28px")
                    with gui.column().classes("gap-0 col"):
                        gui.label("Primeira vez? Treine com um exemplo").classes("text-subtitle1 text-weight-medium t-forte")
                        gui.label("O Edodocs preenche tudo com uma planilha e um ato de exemplo (pessoas "
                                  "fictícias). Você faz os passos de verdade, mas nada é salvo.") \
                            .classes("text-body2 t-suave")
                    gui.button("Treinar com um exemplo", icon="play_circle", on_click=self._treinar) \
                        .props("outline no-caps")
            gui.label("Em qualquer passo, os botões ⓘ explicam o que fazer, com exemplos. Dá para voltar a um "
                      "passo clicando no nome dele, lá em cima.").classes("text-caption t-suave")
        with gui.stepper_navigation():
            gui.button("Começar", icon="arrow_forward", on_click=self.passos.next).props("unelevated no-caps")
            gui.button("Ver arquivos de exemplo", icon="folder_open", on_click=self._exemplos) \
                .props("flat no-caps")

    async def _treinar(self) -> None:
        """Treino: planilha e ato de exemplo já escolhidos; no fim, nada é salvo."""
        self.treino = True
        planilha, self.ato_exemplo = await run.io_bound(gerar_ato_pronto_exemplo, self._temp / "treino")
        self.planilha_exemplo = planilha
        self.nome = "Treino (não será salvo)"
        self.inp_nome.value = self.nome
        self.inp_nome.disable()
        self.aviso_treino.set_visibility(True)
        self.aviso_treino_ato.set_visibility(True)
        self._escolher_tipo("uma")
        self.passos.set_value("tipo")
        self.planilha = planilha
        self.config["aba"] = ""
        self._escolher_modo("tenho")
        await self._ler_planilha()

    # --- Passo 1: planilha -------------------------------------------------------------
    # --- Passo 1: tipo de ato ------------------------------------------------------------
    def _passo_tipo(self) -> None:
        _passo(1, "Que tipo de ato é este?",
               "Cada ato fala de uma pessoa só, ou um mesmo ato pode ter várias pessoas (uma lista com "
               "I, II, III...)? Isso muda como o Edodocs monta o ato.", lambda: self._info("varias"))
        with gui.column().classes("w-full gap-4 q-pl-xl"):
            with gui.row().classes("items-center gap-2 fundo-cartao rounded q-pa-sm") as self.aviso_treino:
                gui.icon("school", color="primary")
                gui.label("Treino: o exemplo é de uma pessoa por ato. Clique em Próximo.") \
                    .classes("text-body2 t-forte")
            self.aviso_treino.set_visibility(False)
            if self.pasta:
                gui.label(f"Processo: {self.nome}").classes("text-subtitle1 text-weight-medium t-forte")
            else:
                with gui.row().classes("items-center gap-1 no-wrap"):
                    self.inp_nome = gui.input("Como se chama este processo?", placeholder="Ex.: Licença Prêmio",
                                              value=self.nome,
                                              on_change=lambda e: setattr(self, "nome", (e.value or "").strip())) \
                        .props('outlined stack-label bg-color=white input-class=text-grey-10 label-color=grey-8 '
                               'hint="É só para você reconhecer na lista. Ex.: Exoneração"') \
                        .classes("w-96")
                    self._info("nome")
            # Dois cartões em pé, cada um com uma folha de Word de mentira: uma pessoa (um nome
            # destacado) ou várias (a lista I, II, III)
            self.cartoes_tipo = {}
            with gui.row().classes("gap-6 items-stretch"):
                for chave, icone, titulo, texto, exemplo in (
                        ("uma", "person", "Uma pessoa por ato",
                         "Cada linha da planilha vira um ato.", "Ex.: Licença Prêmio, Compensação"),
                        ("varias", "groups", "Várias pessoas no mesmo ato",
                         "Várias linhas da planilha viram um ato só, com um parágrafo para cada pessoa.",
                         "Ex.: Exoneração de duas pessoas, Nomeação")):
                    with gui.card().classes("cartao sim-cartao-tipo gap-2 items-center") \
                            .on("click", lambda c=chave: self._escolher_tipo(c)) as cartao:
                        with gui.row().classes("w-full items-center gap-2 no-wrap"):
                            gui.icon(icone, color="primary", size="34px")
                            gui.label(titulo).classes("text-subtitle1 text-weight-medium t-forte col")
                            check = gui.icon("check_circle", color="primary", size="26px")
                        gui.html(simulacao_ui.cartao_tipo(chave == "varias"), sanitize=False) \
                            .classes("w-full")
                        gui.label(texto).classes("text-body2 t-suave text-center")
                        gui.label(exemplo).classes("text-caption t-suave text-center")
                    self.cartoes_tipo[chave] = (cartao, check)
            self._mostrar_tipo()
        with gui.stepper_navigation():
            gui.button("Próximo", icon="arrow_forward", on_click=self._ir_word).props("unelevated")
            if not self.pasta:
                gui.button("Voltar", on_click=self.passos.previous).props("flat")

    @property
    def varias(self) -> bool:
        return bool((self.config.get("juntar") or {}).get("coluna")) or self.config.get("_tipo") == "varias"

    def _escolher_tipo(self, tipo: str) -> None:
        self.config["_tipo"] = tipo
        juntar = self.config.setdefault("juntar", {})
        if tipo == "uma":
            juntar["coluna"] = ""
        elif not juntar.get("coluna"):
            juntar["coluna"] = _coluna_igual(self.colunas, "Nº DO ATO") or next(
                (c for c in self.colunas if re.search(r"DECRETO|ATO|PORTARIA|GRUPO", atos.normalizar(c))),
                "Nº DO ATO")
        self._mostrar_tipo()
        self._desenhar_varias()
        if self.modelo:
            self._desenhar_trecho()

    def _mostrar_tipo(self) -> None:
        if not hasattr(self, "cartoes_tipo"):
            return
        atual = "varias" if self.varias else ("uma" if self.config.get("_tipo") or self.pasta else None)
        for chave, (cartao, check) in self.cartoes_tipo.items():
            if chave == atual:
                cartao.classes(add="escolhido")
            else:
                cartao.classes(remove="escolhido")
            check.set_visibility(chave == atual)
        self._atualizar_word_simulado()

    def _ir_word(self) -> None:
        if not self.pasta:
            if not self.nome:
                gui.notify("Escreva um nome para o processo (ex.: Licença Prêmio)", type="warning")
                return
            if INVALIDOS.search(self.nome) and not self.treino:
                gui.notify("O nome não pode ter os caracteres < > : \" / \\ | ? *", type="warning")
                return
            if (main.PROCESSOS / self.nome).exists() and not self.treino:
                gui.notify("Já existe um processo com esse nome", type="warning")
                return
        if not self.config.get("_tipo") and not self.pasta:
            gui.notify("Escolha o tipo de ato: uma pessoa ou várias", type="warning")
            return
        self.passos.next()

    # --- Passo 3: a planilha -------------------------------------------------------------
    def _passo_planilha(self) -> None:
        _passo(3, "A planilha", "Onde ficam os dados de cada ato. Se você já tem uma, escolha. Se não tem, "
                                "o Edodocs cria para você, com uma coluna para cada campo do Word.",
               lambda: self._info("planilha"))
        with gui.column().classes("w-full gap-3 q-pl-xl"):
            self.cartoes_modo = {}
            with gui.row().classes("w-full gap-4"):
                for chave, icone, titulo, texto in (
                        ("tenho", "upload_file", "Já tenho uma planilha",
                         "Escolha o arquivo do Excel. Ele continua onde está (no computador ou na rede)."),
                        ("criar", "auto_awesome", "Criar a planilha para mim",
                         "Uma coluna para cada campo, com listas para escolher (sim/não...) e datas.")):
                    with gui.card().classes("cartao q-pa-md col cursor-pointer").style("min-width: 280px") \
                            .on("click", lambda c=chave: self._escolher_modo(c)) as cartao:
                        with gui.row().classes("items-center gap-3 no-wrap"):
                            gui.icon(icone, color="primary", size="34px")
                            with gui.column().classes("gap-0"):
                                gui.label(titulo).classes("text-subtitle1 text-weight-medium t-forte")
                                gui.label(texto).classes("text-body2 t-suave")
                    self.cartoes_modo[chave] = cartao
            self.area_criar = gui.column().classes("w-full gap-2")
            self.area_tenho = gui.column().classes("w-full gap-3")
        with self.area_tenho:
            with gui.row().classes("items-center gap-3"):
                self.botao_planilha = gui.button("Escolher a planilha...", icon="upload_file",
                                                 on_click=self._escolher_planilha).props("no-caps unelevated")
                self.rotulo_planilha = gui.label("").classes("text-body2 t-suave break-all")
            # Enquanto lê: barra, contador de segundos e botões travados (para não parecer travado)
            with gui.card().classes("w-full fundo-cartao q-pa-md gap-2") as self.lendo_planilha:
                with gui.row().classes("items-center gap-3 no-wrap"):
                    gui.spinner(size="md")
                    with gui.column().classes("gap-0"):
                        self.texto_lendo = gui.label("Lendo a planilha...").classes("text-subtitle2 t-forte")
                        gui.label("Planilhas grandes ou na rede podem levar alguns segundos. "
                                  "Não precisa clicar de novo.").classes("text-caption t-suave")
                gui.linear_progress(show_value=False).props("indeterminate rounded color=primary") \
                    .classes("w-full")
            self.lendo_planilha.set_visibility(False)
            self._lendo_desde = 0.0
            gui.timer(1.0, self._atualizar_lendo)
            self.area_planilha = gui.column().classes("w-full")
            # Os ajustes só aparecem depois que a planilha foi lida (antes não há o que ajustar)
            with gui.expansion("Ajustes (só se precisar)", icon="tune") \
                .classes("w-full fundo-cartao rounded") as self.ajustes:
                with gui.row().classes("items-center gap-4 q-pa-sm"):
                    self.sel_aba = gui.select([], label="Qual folha da planilha?", on_change=self._trocar_aba) \
                        .props("outlined").classes("w-64")
                    self.sel_nome = gui.select([], label="Qual coluna tem o nome da pessoa?",
                                               on_change=self._trocar_coluna_nome) \
                        .props("outlined use-input input-debounce=0").classes("w-80")
                    self._info("ajustes")
                gui.label("Linhas sem nome são ignoradas.").classes("text-caption t-suave q-px-sm")
            self.ajustes.set_visibility(False)
        self._mostrar_modo()
        with gui.stepper_navigation():
            self.botao_proximo_planilha = gui.button("Próximo", icon="arrow_forward",
                                                     on_click=self._ir_ligar).props("unelevated")
            gui.button("Voltar", on_click=self.passos.previous).props("flat")

    def _escolher_modo(self, modo: str) -> None:
        self.modo_planilha = modo
        self._mostrar_modo()
        if modo == "criar":
            self._desenhar_criar()

    def _mostrar_modo(self) -> None:
        if not hasattr(self, "cartoes_modo"):
            return
        for chave, cartao in self.cartoes_modo.items():
            cartao.style(replace="min-width: 280px; outline: 3px solid var(--q-primary)"
                         if chave == self.modo_planilha else "min-width: 280px")
        self.area_tenho.set_visibility(self.modo_planilha == "tenho")
        self.area_criar.set_visibility(self.modo_planilha == "criar")

    def _desenhar_criar(self) -> None:
        """Opção "Criar a planilha para mim": mostra as colunas que ela vai ter e cria."""
        self._atualizar_word_simulado()
        self.area_criar.clear()
        with self.area_criar:
            if not self.campos:
                _conferencia(False, "Primeiro prepare o Word (passo 2): as colunas da planilha saem dos campos dele.")
                return
            colunas = self._colunas_novas()
            gui.label("A planilha vai ter estas colunas:").classes("text-body2 t-forte")
            with gui.row().classes("gap-1"):
                for c in colunas:
                    extra = (" · " + "/".join(planilha_nova._opcoes(c["opcoes"]))) if c["tipo"] == "lista" else \
                        (" · data" if c["tipo"] == "data" else "")
                    gui.chip(c["nome"] + extra, color=None).props("dense square").classes("chip-campo")
            criada = self.planilha is not None and str(self.planilha).startswith(str(self._temp))
            with gui.row().classes("items-center gap-3"):
                gui.button("Criar a planilha" if not criada else "Criar de novo (com as colunas atuais)",
                           icon="auto_awesome", on_click=self._criar_planilha).props("unelevated no-caps")
                if criada:
                    gui.button("Abrir no Excel", icon="open_in_new", on_click=lambda: abrir(self.planilha)) \
                        .props("outline no-caps")
            if criada:
                _conferencia(True, f"Planilha criada: {self.planilha.name}. Ela vai para a pasta do processo "
                                   "quando você salvar. Depois é só preencher: uma linha por pessoa.")
                if self.exemplo:
                    _conferencia(None, "A prévia (passo 5) usa os dados que estavam no seu Word como exemplo.")

    def _colunas_novas(self) -> list[dict]:
        """Colunas da planilha nova: "Nº DO ATO" primeiro (também junta as pessoas de um ato)."""
        config = dict(self.config, coluna_numero="Nº DO ATO",
                      juntar={"coluna": "Nº DO ATO"} if self.varias else {"coluna": ""})
        return processos.Processo(self._temp, config).colunas_planilha_nova(self.campos)

    async def _criar_planilha(self) -> None:
        colunas = self._colunas_novas()
        destino = self._temp / f"{self.nome or 'Processo'} - planilha.xlsx"
        try:
            await run.io_bound(planilha_nova.criar, destino, colunas)
        except Exception as e:
            gui.notify(f"Não consegui criar a planilha: {type(e).__name__}", type="negative")
            return
        self.planilha = destino
        nomes = [c["nome"] for c in colunas]
        self.config.update(aba=planilha_nova.ABA, coluna_numero="Nº DO ATO",
                           coluna_nome=_coluna_igual(nomes, "NOME") or (nomes[1] if len(nomes) > 1 else nomes[0]))
        if self.varias:
            self.config.setdefault("juntar", {})["coluna"] = "Nº DO ATO"
        self.abas, self.colunas = [planilha_nova.ABA], nomes
        # Ainda sem dados: a prévia usa o que estava no Word (se ele foi marcado aqui)
        self.registros = [dict(self.exemplo, NODOATO="1", _LINHA=2)] if self.exemplo else []
        self.linha_previa = 0
        self._desenhar_criar()
        self._desenhar_ligacoes()
        self._desenhar_previa()

    def _lendo(self, etapa: str | None) -> None:
        """Mostra (etapa = o que está lendo) ou esconde o aviso de leitura da planilha."""
        lendo = etapa is not None
        if lendo:
            self._lendo_desde = time.monotonic()
            self.texto_lendo.text = etapa
            self._etapa_lendo = etapa
        self.lendo_planilha.set_visibility(lendo)
        for elemento in (self.botao_planilha, self.botao_proximo_planilha, self.sel_aba, self.sel_nome):
            elemento.set_enabled(not lendo)

    def _atualizar_lendo(self) -> None:
        if self.lendo_planilha.visible:
            segundos = int(time.monotonic() - self._lendo_desde)
            self.texto_lendo.text = self._etapa_lendo + (f" · {segundos} s" if segundos >= 2 else "")

    async def _escolher_planilha(self) -> None:
        caminho = await self._escolher_arquivo("Planilha do Excel", ".xlsx")
        if caminho:
            self.planilha = caminho
            self.config["aba"] = ""
            await self._ler_planilha()

    async def _ler_planilha(self) -> None:
        if not self.planilha:
            return
        self.rotulo_planilha.text = str(self.planilha)
        self.ajustes.set_visibility(False)
        self.area_planilha.clear()
        self._lendo("Abrindo a planilha...")
        try:
            lido = await run.io_bound(_ler_planilha, self.planilha, self.config["aba"], "")
        except Exception as e:
            self._lendo(None)
            gui.notify(f"Não consegui ler a planilha ({type(e).__name__}). Ela está aberta no Excel?",
                       type="negative")
            return
        self.abas, self.colunas = lido["abas"], lido["colunas"]
        if self.config["aba"] not in self.abas:
            self.config["aba"] = atos.ABA if atos.ABA in self.abas else self.abas[0]
        if not _coluna_igual(self.colunas, self.config["coluna_nome"]):
            self.config["coluna_nome"] = _coluna_igual(self.colunas, "NOME") or (self.colunas or [""])[0]
        if not _coluna_igual(self.colunas, self.config["coluna_numero"]):
            self.config["coluna_numero"] = next(
                (c for c in self.colunas if re.search(r"PORTARIA|N.?\s*DO\s*ATO|NUMERO", atos.normalizar(c))),
                (self.colunas or [""])[0])
        self.sel_aba.set_options(self.abas, value=self.config["aba"])
        self.sel_nome.set_options(self.colunas, value=_coluna_igual(self.colunas, self.config["coluna_nome"]))
        await self._ler_registros()
        self.ajustes.set_visibility(True)

    async def _trocar_aba(self, e) -> None:
        if e.value and e.value != self.config["aba"]:
            self.config["aba"] = e.value
            await self._ler_planilha()

    async def _trocar_coluna_nome(self, e) -> None:
        if e.value and e.value != self.config["coluna_nome"]:
            self.config["coluna_nome"] = e.value
            await self._ler_registros()

    async def _ler_registros(self) -> None:
        self._lendo("Lendo as linhas da planilha...")
        try:
            lido = await run.io_bound(_ler_planilha, self.planilha, self.config["aba"],
                                      self.config["coluna_nome"])
        except Exception as e:
            gui.notify(f"Não consegui ler a planilha ({type(e).__name__}). Ela está aberta no Excel?",
                       type="negative")
            return
        finally:
            self._lendo(None)
        self.registros = lido["registros"]
        self.linha_previa = 0
        self._desenhar_planilha()
        self._desenhar_ligacoes()
        self._desenhar_previa()

    def _desenhar_planilha(self) -> None:
        self._atualizar_word_simulado()
        self.area_planilha.clear()
        with self.area_planilha:
            if not self.registros:
                _conferencia(False, "Não achei linhas com dados nesta planilha.")
                with gui.column().classes("q-pl-lg gap-0"):
                    for dica in ("Os títulos das colunas estão na 1ª linha?",
                                 "Os dados estão em outra folha (aba)? Troque em Ajustes, logo abaixo.",
                                 "A coluna do nome está certa? Linhas sem nome são ignoradas (Ajustes)."):
                        gui.label("• " + dica).classes("text-body2 t-suave")
                return
            # Conferência automática: o que o Edodocs entendeu da planilha
            with gui.card().classes("w-full fundo-cartao q-pa-md gap-1"):
                gui.label("O que o Edodocs entendeu da sua planilha").classes("text-subtitle2 t-forte")
                _conferencia(True, f"{len(self.registros)} ato(s) (pessoas) e {len(self.colunas)} coluna(s).")
                nome = self.config["coluna_nome"]
                _conferencia("NOME" in atos.normalizar(nome),
                             f"Nome das pessoas: coluna {nome}." + ("" if "NOME" in atos.normalizar(nome) else
                             " Não achei uma coluna chamada NOME: se estiver errado, troque em Ajustes."))
                numero = self.config["coluna_numero"]
                tem_numero = bool(re.search(r"PORTARIA|N.?\s*DO\s*ATO|NUMERO", atos.normalizar(numero)))
                _conferencia(tem_numero, f"Número do ato: coluna {numero}." + ("" if tem_numero else
                             " Confira no último passo (Ajustar títulos)."))
                sexo = modelo_automatico._coluna_sexo(self.colunas, self.registros)
                _conferencia(None, f"Coluna de sexo (F/M): {sexo}. Palavras como lotada/lotado mudam sozinhas."
                             if sexo else "Sem coluna de sexo (F/M): palavras como lotada/lotado ficam fixas no "
                             "modelo. Se quiser que mudem, crie uma coluna SEXO com F ou M.")
            with gui.row().classes("items-center gap-2 q-mt-sm"):
                gui.label("As primeiras linhas:").classes("text-body2 t-forte")
            mostrar = self.colunas[:8]
            linhas = [{"id": r["_LINHA"], **{c: r.get(atos.normalizar(c), "") for c in mostrar}}
                      for r in self.registros[:4]]
            gui.table(rows=linhas, columns=[{"name": c, "label": c, "field": c, "align": "left"}
                                            for c in mostrar], row_key="id") \
                .props("flat dense bordered").classes("w-full")

    # --- Passo 2: um ato pronto -----------------------------------------------------------
    def _passo_ato(self) -> None:
        _passo(2, "O Word",
               "Escolha o Word de um ato que você já fez. Depois, marque no texto o que muda de um ato para "
               "outro (o nome, a matrícula...). Se já tiver a planilha, o Edodocs pode achar isso sozinho.",
               lambda: self._info("ato"))
        with gui.column().classes("w-full gap-3 q-pl-xl"):
            with gui.row().classes("items-center gap-2 fundo-cartao rounded q-pa-sm") as self.aviso_treino_ato:
                gui.icon("school", color="primary")
                gui.label("Treino: use o ato de exemplo (é da MARIA DA SILVA, que está na planilha).") \
                    .classes("text-body2 t-forte")
                gui.button("Usar o ato de exemplo", icon="description",
                           on_click=lambda: self._escolher_ato(self.ato_exemplo)).props("unelevated no-caps dense")
            self.aviso_treino_ato.set_visibility(False)
            with gui.row().classes("items-center gap-3"):
                gui.button("Escolher o Word...", icon="upload_file", on_click=lambda: self._escolher_ato()) \
                    .props("no-caps unelevated")
                self.botao_abrir_modelo = gui.button(
                    "Abrir o modelo no Word", icon="edit_document",
                    on_click=lambda: self.modelo and abrir(self.modelo)).props("outline no-caps")
                self.botao_abrir_modelo.set_visibility(bool(self.modelo))
            self.area_ato = gui.column().classes("w-full gap-2")
            # Ato de várias pessoas: o parágrafo que repete (aparece depois que o modelo fica pronto)
            self.area_trecho = gui.column().classes("w-full gap-2")
            with gui.expansion("Prefiro preparar o modelo eu mesmo", icon="school") \
                    .classes("w-full fundo-cartao rounded"):
                with gui.column().classes("q-pa-sm gap-1"):
                    with gui.row().classes("items-center gap-1"):
                        gui.label("Como preparar").classes("text-subtitle2 t-forte")
                        self._info("modelo_manual")
                    for i, texto in enumerate(AJUDA_WORD, 1):
                        with gui.row().classes("no-wrap gap-2 items-start"):
                            gui.label(f"{i}.").classes("text-weight-bold text-primary")
                            gui.label(texto).classes("text-body2")
                    gui.button("Gerar um Word e uma planilha de exemplo", icon="auto_awesome",
                               on_click=self._exemplos).props("outline no-caps").classes("q-mt-sm")
        with gui.stepper_navigation():
            gui.button("Próximo", icon="arrow_forward", on_click=self._ir_planilha).props("unelevated")
            gui.button("Voltar", on_click=self.passos.previous).props("flat")

    def _ir_planilha(self) -> None:
        if not self.modelo or not self.campos:
            gui.notify("Escolha o Word e marque o que muda (ou deixe o Edodocs achar com a planilha)",
                       type="warning")
            return
        if self.modo_planilha == "criar":
            self._desenhar_criar()
        self.passos.next()

    async def _escolher_ato(self, caminho: Path | None = None) -> None:
        caminho = caminho or await self._escolher_arquivo("Documento do Word", ".docx")
        if not caminho:
            return
        self.area_ato.clear()
        with self.area_ato, gui.row().classes("items-center gap-2"):
            gui.spinner()
            gui.label("Lendo o ato e procurando os dados da planilha...").classes("t-suave")
        try:
            campos = await run.io_bound(atos.campos_do_modelo, caminho)
        except Exception as e:
            self._aviso_ato(f"Não consegui abrir o Word ({type(e).__name__}). Ele está aberto em outro programa?")
            return
        if campos:  # já é um modelo com campos: não precisa marcar nada
            self.ato_pronto, self.achados = None, []
            await self._usar_modelo(caminho)
            return
        self.word = caminho
        if not self.registros:  # sem planilha: a pessoa marca no texto o que muda
            self._mostrar_marcador()
            return
        try:
            self.achados, self.linha_ato = await run.io_bound(
                modelo_automatico.encontrar, caminho, self.registros, self.colunas)
        except Exception as e:
            self._aviso_ato(f"Não consegui ler o Word ({type(e).__name__}).")
            return
        self.ato_pronto = caminho
        self._desenhar_achados()

    def _mostrar_marcador(self) -> None:
        """O texto do Word para a pessoa marcar o que muda (marcar_ui.py)."""
        self.area_ato.clear()
        self.area_trecho.clear()
        with self.area_ato:
            with gui.row().classes("w-full items-center gap-2 fundo-cartao rounded q-pa-sm"):
                gui.icon("table_chart", color="primary")
                gui.label("Já tem a planilha com os dados deste ato? O Edodocs acha sozinho o que muda.") \
                    .classes("text-body2 t-forte col")
                gui.button("Usar a minha planilha", icon="search", on_click=self._achar_com_planilha) \
                    .props("outline dense no-caps")
            self.marcador = marcar_ui.MarcadorWord(self.tela.camada, self.word,
                                                   ao_mudar=lambda: self._contar_marcas())
            with gui.row().classes("items-center gap-3 q-mt-sm"):
                gui.button("Criar o modelo", icon="auto_fix_high", color="positive",
                           on_click=self._criar_modelo_marcado).props("unelevated no-caps")
                self.resumo_marcas = gui.label("").classes("text-body2 t-suave")
        self._contar_marcas()

    def _contar_marcas(self) -> None:
        if hasattr(self, "resumo_marcas") and self.marcador:
            n = len(self.marcador.marcas)
            self.resumo_marcas.text = (f"{n} campo(s) marcado(s)" if n else
                                       "Marque pelo menos o nome da pessoa antes de criar o modelo")
        self._atualizar_word_simulado()

    async def _criar_modelo_marcado(self) -> None:
        if not self.marcador or not self.marcador.marcas:
            gui.notify("Marque no texto pelo menos um trecho que muda (ex.: o nome da pessoa)", type="warning")
            return
        destino = self._temp / "modelo.docx"
        try:
            ligacoes, self.exemplo = await run.io_bound(marcar_ui.criar_modelo, self.word, destino,
                                                        self.marcador.marcas)
        except Exception as e:
            gui.notify(f"Não consegui criar o modelo: {type(e).__name__}", type="negative")
            return
        self.config["ligacoes"] = ligacoes
        self.ato_pronto = self.word
        await self._usar_modelo(destino, deduzir_titulos=True)

    async def _achar_com_planilha(self) -> None:
        """Tem a planilha: escolhe agora (vale também para o passo 3) e acha os dados no Word."""
        caminho = await self._escolher_arquivo("Planilha do Excel", ".xlsx")
        if not caminho:
            return
        self.planilha = caminho
        self.config["aba"] = ""
        self._escolher_modo("tenho")
        await self._ler_planilha()
        if self.registros and self.word:
            await self._escolher_ato(self.word)

    def _desenhar_trecho(self) -> None:
        """Ato de várias pessoas: escolher o parágrafo de cada pessoa (o que repete)."""
        if not hasattr(self, "area_trecho"):
            return
        self.area_trecho.clear()
        if not (self.varias and self.modelo):
            return
        resumo = regras_ui.resumo_trecho(self.modelo)
        outros = self.config.get("outros_modelos") or []
        with self.area_trecho, gui.card().classes("cartao w-full q-pa-md gap-2"):
            with gui.row().classes("items-center gap-2"):
                gui.icon("groups", color="primary", size="26px")
                gui.label("Quando o ato tem várias pessoas").classes("text-subtitle1 text-weight-medium t-forte")
                self._info("varias")
            if resumo:
                _conferencia(True, f"Neste Word, isto se repete para cada pessoa: “{resumo}”")
            for outro in outros:
                caminho = self._caminho_outro(outro)
                resumo_outro = regras_ui.resumo_trecho(caminho) if caminho.is_file() else None
                with gui.row().classes("w-full items-center gap-2 no-wrap"):
                    _conferencia(bool(resumo_outro),
                                 f"Com {regras_ui.descrever_condicoes(outro.get('se') or [])}: usa {caminho.name}"
                                 + (f", e repete “{resumo_outro}”" if resumo_outro else
                                    " (falta escolher o parágrafo de uma pessoa)"))
                    gui.space()
                    gui.button("Escolher o parágrafo", icon="touch_app",
                               on_click=lambda c=caminho: self._escolher_paragrafo_pessoa(c)) \
                        .props("flat dense no-caps")
                    gui.button(icon="delete", on_click=lambda o=outro: self._tirar_outro_word(o)) \
                        .props("flat round dense color=negative").tooltip("Não usar mais este Word")
            if resumo or outros:
                return
            gui.label("Como é o seu Word?").classes("text-body2 t-forte")
            with gui.row().classes("w-full gap-3 items-stretch"):
                with gui.card().classes("cartao col q-pa-md gap-1 fundo-previa"):
                    gui.label("Ele já tem a lista das pessoas").classes("text-subtitle2 t-forte")
                    gui.label("Cada pessoa tem um parágrafo (“I – PAULO...;” “II – LUIZ...”) ou uma linha numa "
                              "tabela. O Edodocs repete para cada pessoa do ato e tira as das outras.") \
                        .classes("text-body2 t-suave")
                    gui.button("Escolher o parágrafo ou a linha de uma pessoa", icon="touch_app",
                               on_click=lambda: self._escolher_paragrafo_pessoa()) \
                        .props("outline no-caps").classes("q-mt-sm")
                with gui.card().classes("cartao col q-pa-md gap-1 fundo-previa"):
                    gui.label("Ele é de uma pessoa só").classes("text-subtitle2 t-forte")
                    gui.label("Quando são várias pessoas o texto muda (ex.: “Ficam exonerados..., sendo: I – ...; "
                              "II – ...”). Escolha um Word de um ato com várias pessoas: ele é usado quando o ato "
                              "tiver 2 ou mais pessoas.").classes("text-body2 t-suave")
                    gui.button("Escolher o Word de várias pessoas...", icon="note_add",
                               on_click=self._marcar_outro_word).props("outline no-caps").classes("q-mt-sm")

    def _existentes(self) -> dict[str, str]:
        """Campos que o processo já tem e o tipo da ligação de cada um (para o 2º Word)."""
        proc = self._processo_rascunho()
        return {rotulo: proc.ligacao(rotulo).get("tipo", "coluna") for _, rotulo in self.campos}

    async def _marcar_outro_word(self) -> None:
        """O Word para quando o ato tem várias pessoas: marcar o que muda (os campos que já
        existem usam a mesma regra), depois escolher o parágrafo de uma pessoa."""
        caminho = await self._escolher_arquivo("Documento do Word", ".docx")
        if not caminho:
            return
        destino = self._temp / f"{caminho.stem} (várias pessoas).docx"
        try:
            ja_tem_campos = await run.io_bound(atos.campos_do_modelo, caminho)
        except Exception as e:
            gui.notify(f"Não consegui abrir o Word ({type(e).__name__}). Ele está aberto em outro programa?",
                       type="negative")
            return
        if ja_tem_campos:  # já é um modelo: só usa
            shutil.copy2(caminho, destino)
            await self._usar_outro_word(destino, {})
            return
        with self.tela.camada, gui.dialog().props("maximized") as janela, \
                gui.card().classes("cartao w-full q-pa-lg gap-3"):
            with gui.row().classes("w-full items-center gap-2 no-wrap"):
                gui.icon("groups", color="primary", size="28px")
                with gui.column().classes("gap-0 col"):
                    gui.label(f"Word de várias pessoas: {caminho.name}").classes("text-h6 t-forte")
                    gui.label("Marque o que muda, como no outro Word. Use os mesmos nomes (NOME, MATRÍCULA, "
                              "VÍNCULO...): eles aparecem na lista e usam as regras que você já fez. Marque só "
                              "a primeira pessoa; as outras saem no próximo passo.").classes("text-body2 t-suave")
                gui.button(icon="close", on_click=janela.close).props("flat round dense")
            marcador = marcar_ui.MarcadorWord(self.tela.camada, caminho, existentes=self._existentes())

            async def criar() -> None:
                if not marcador.marcas:
                    gui.notify("Marque pelo menos o nome da primeira pessoa", type="warning")
                    return
                try:
                    ligacoes, _ = await run.io_bound(marcar_ui.criar_modelo, caminho, destino, marcador.marcas)
                except Exception as e:
                    gui.notify(f"Não consegui criar o modelo: {type(e).__name__}", type="negative")
                    return
                janela.close()
                await self._usar_outro_word(destino, ligacoes)

            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
                gui.button("Pronto: escolher o parágrafo de uma pessoa", icon="arrow_forward", color="positive",
                           on_click=criar).props("unelevated no-caps")
        janela.open()

    async def _usar_outro_word(self, destino: Path, ligacoes: dict) -> None:
        """Passa a usar o Word de várias pessoas (quando o ato tem 2 ou mais) e já pede o
        parágrafo de uma pessoa. Ligações novas entram; as que o processo já tem ficam."""
        atuais = {atos.normalizar(n) for n in self.config.setdefault("ligacoes", {})}
        for nome, regra in ligacoes.items():
            if atos.normalizar(nome) not in atuais:
                self.config["ligacoes"][nome] = regra
        self.config.setdefault("outros_modelos", []).append(
            {"modelo": str(destino), "se": [{"campo": regras_ui.ROTULO_QUANTIDADE, "op": "maior", "valor": "1"}]})
        await self._depois_de_mudar_word_async()
        self._desenhar_trecho()
        if not regras_ui.resumo_trecho(destino):
            self._escolher_paragrafo_pessoa(destino)

    def _tirar_outro_word(self, outro: dict) -> None:
        self.config["outros_modelos"].remove(outro)
        self._depois_de_mudar_word()
        self._desenhar_trecho()

    def _escolher_paragrafo_pessoa(self, modelo: Path | None = None) -> None:
        modelo = modelo or self.modelo
        paragrafos = atos.paragrafos_do_modelo(modelo)
        tabelas = atos.linhas_das_tabelas(modelo)
        with self.tela.camada, gui.dialog() as janela, \
                gui.card().classes("cartao w-[900px] max-w-full q-pa-lg gap-3"):
            gui.label("Clique no parágrafo ou na linha de uma das pessoas").classes("text-h6 t-forte")
            gui.label("Escolha o que começa com “I –” ou, numa tabela, a linha da primeira pessoa. O Edodocs "
                      "cuida do resto.").classes("text-body2 t-suave")
            with gui.column().classes("w-full gap-1").style("max-height: 60vh; overflow-y: auto"):
                for i, texto in enumerate(paragrafos):
                    if i in tabelas:
                        with gui.column().classes("w-full gap-0 fundo-cartao rounded q-pa-xs"):
                            gui.label("Tabela: clique na linha de uma pessoa").classes("text-caption t-suave q-px-sm")
                            for n, celulas in enumerate(tabelas[i]):
                                texto_linha = " | ".join(" ".join(c.split()) for c in celulas)
                                gui.button(texto_linha[:220] + ("…" if len(texto_linha) > 220 else ""),
                                           on_click=lambda i=i, n=n: escolher(i, n)) \
                                    .props("flat dense no-caps align=left").classes("w-full text-left t-forte")
                        continue
                    if not texto.strip() or atos.CHAVES.fullmatch(texto.strip() or "x"):
                        continue
                    gui.button(texto[:220] + ("…" if len(texto) > 220 else ""),
                               on_click=lambda i=i: escolher(i)) \
                        .props("flat no-caps align=left").classes("w-full text-left t-forte")

            async def escolher(indice: int, linha: int | None = None) -> None:
                janela.close()
                try:
                    if linha is None:
                        feito = await run.io_bound(atos.preparar_trecho_pessoa, modelo, indice)
                    else:
                        feito = await run.io_bound(atos.preparar_linha_pessoa, modelo, indice, linha)
                except PermissionError:
                    gui.notify("O Word está aberto: feche e tente de novo", type="negative")
                    return
                partes = [f"{'o parágrafo' if linha is None else 'a linha'} vai se repetir para cada pessoa"]
                if feito.get("numerou"):
                    partes.append("a numeração (I, II, III...) é automática")
                if feito["removidos"]:
                    partes.append(f"{feito['removidos']} {'parágrafo(s)' if linha is None else 'linha(s)'} "
                                  "das outras pessoas saíram")
                gui.notify("Pronto: " + "; ".join(partes), type="positive", multi_line=True)
                await self._ler_campos()
                self._desenhar_trecho()
                self._desenhar_previa()

            with gui.row().classes("w-full justify-end"):
                gui.button("Cancelar", on_click=janela.close).props("flat no-caps")
        janela.open()

    def _aviso_ato(self, texto: str) -> None:
        self.area_ato.clear()
        with self.area_ato, gui.row().classes("items-center gap-2"):
            gui.icon("warning", color="warning")
            gui.label(texto).classes("t-forte")

    def _desenhar_achados(self) -> None:
        """Os dados encontrados no ato, para a pessoa conferir (marcar/desmarcar, trocar coluna)."""
        self.area_ato.clear()
        with self.area_ato:
            if not self.achados or self.linha_ato is None:
                self._aviso_ato("Não encontrei dados da planilha neste ato.")
                with self.area_ato, gui.column().classes("q-pl-lg gap-0"):
                    gui.label("O que costuma causar isso:").classes("text-body2 t-forte")
                    for dica in ("O ato é de uma pessoa que não está na planilha: escolha o ato de alguém que está.",
                                 "O nome ou a matrícula estão escritos diferente no Word e na planilha "
                                 "(abreviado, sem acento...): deixe igual e tente de novo.",
                                 "É um modelo em branco, sem os dados de ninguém: use um ato já preenchido, "
                                 "ou escreva os campos como {{NOME}} (veja logo abaixo)."):
                        gui.label("• " + dica).classes("text-body2 t-suave")
                return
            proc = processos.Processo(Path("."), self.config)
            with gui.row().classes("items-center gap-2"):
                gui.icon("person_search", color="primary", size="24px")
                gui.label(f"Este ato é de: {proc.descricao(self.linha_ato)} (linha {self.linha_ato['_LINHA']} "
                          "da planilha)").classes("text-subtitle1 text-weight-medium t-forte")
            with gui.row().classes("items-center gap-1 no-wrap"):
                gui.label(f"Encontrei {len(self.achados)} dado(s) da planilha no texto. Deixe marcado o que "
                          "muda de uma pessoa para outra; desmarque o que é texto fixo do ato.") \
                    .classes("text-body2 t-suave")
                self._info("achados")
            with gui.card().classes("cartao cartao-reto w-full q-pa-none gap-0"):
                for n, achado in enumerate(self.achados):
                    if n:
                        gui.separator()
                    self._linha_achado(achado)
            with gui.row().classes("items-center gap-3 q-mt-sm"):
                gui.button("Criar o modelo", icon="auto_fix_high", color="positive",
                           on_click=self._criar_modelo).props("unelevated no-caps")
                self.resumo_achados = gui.label("").classes("text-body2 t-suave")
            self._atualizar_resumo()

    def _linha_achado(self, achado) -> None:
        with gui.row().classes("w-full items-center gap-3 no-wrap q-px-md q-py-xs"):
            gui.checkbox(value=achado.usar, on_change=lambda e, a=achado: self._marcar(a, e.value))
            with gui.column().classes("gap-0 col"):
                # o trecho do ato com o dado destacado
                trecho = html.escape(achado.trecho).replace("【", "<mark>").replace("】", "</mark>")
                gui.html(trecho).classes("text-body2 t-forte")
            gui.icon("arrow_forward", color="grey-6")
            opcoes = list(dict.fromkeys(achado.colunas + self.colunas))
            gui.select(opcoes, value=achado.coluna, label="É a coluna",
                       on_change=lambda e, a=achado: setattr(a, "coluna", e.value or a.coluna)) \
                .props("dense outlined use-input input-debounce=0").classes("w-60")

    def _marcar(self, achado, valor: bool) -> None:
        achado.usar = valor
        self._atualizar_resumo()

    def _atualizar_resumo(self) -> None:
        if hasattr(self, "resumo_achados"):
            marcados = sum(a.usar for a in self.achados)
            self.resumo_achados.text = f"{marcados} dado(s) marcado(s) vão virar campos do modelo"

    async def _criar_modelo(self) -> None:
        marcados = [a for a in self.achados if a.usar]
        if not marcados:
            gui.notify("Marque pelo menos um dado", type="warning")
            return
        destino = self._temp / "modelo.docx"
        try:
            await run.io_bound(modelo_automatico.aplicar, self.ato_pronto, destino, marcados)
        except Exception as e:
            gui.notify(f"Não consegui criar o modelo: {type(e).__name__}", type="negative")
            return
        # Ligações especiais (ex.: dias por extenso) e títulos tirados da 1ª linha do Word
        self.config["ligacoes"] = modelo_automatico.ligacoes(marcados)
        await self._usar_modelo(destino, deduzir_titulos=True)

    async def _usar_modelo(self, modelo: Path, deduzir_titulos: bool | None = None) -> None:
        """Passa a usar este modelo: lê os campos e (num processo novo) deduz os títulos."""
        self.modelo = modelo
        await self._ler_campos()
        if deduzir_titulos or (deduzir_titulos is None and not self.pasta):
            self._deduzir_titulos()
        self.botao_abrir_modelo.set_visibility(True)
        self.area_ato.clear()
        with self.area_ato:
            with gui.row().classes("items-center gap-2"):
                gui.icon("check_circle", color="positive", size="26px")
                gui.label(f"Modelo pronto, com {len(self.campos)} campo(s).") \
                    .classes("text-subtitle1 text-weight-medium t-forte")
            with gui.row().classes("gap-1"):
                for _, rotulo in self.campos:
                    gui.chip("{{" + rotulo + "}}", text_color="white").props("dense square").classes("chip-campo")
            gui.label("Pronto! Clique em Próximo para escolher a planilha.") \
                .classes("text-body2 t-suave")
        self._desenhar_trecho()
        self._desenhar_previa()

    def _deduzir_titulos(self) -> None:
        """Os títulos do e-docs e do Diário vêm da primeira linha do modelo (o título do ato)."""
        ano = ""
        if self.registros:
            numero = processos.Processo(Path("."), self.config).numero(self.registros[0])
            m = re.search(r"/\s*(\d{4})", numero)
            ano = m.group(1) if m else ""
        titulo = modelo_automatico.titulo_do_modelo(self.modelo, ano)
        if titulo and "{" not in titulo and any(atos.normalizar(r) == "NOME" for _, r in self.campos):
            # Título sem número (ex.: "DECRETO S/Nº"): o nome da pessoa diferencia um ato do outro
            titulo += " - {NOME}"
        if titulo and "{" in titulo:
            self.config["titulo_edocs"] = self.config["titulo_diario"] = titulo
            self.config["titulo_na_primeira_linha"] = True

    async def _ler_campos(self) -> None:
        if not self.modelo:
            return
        try:
            # Campos do Word principal e dos outros Words (atos com várias pessoas); os
            # automáticos ({{ITEM}}...) o Edodocs preenche sozinho
            self.campos = await run.io_bound(self._processo_rascunho().campos)
        except Exception as e:
            self.campos = []
            gui.notify(f"Não consegui ler o Word ({type(e).__name__}). Ele está aberto ou corrompido?",
                       type="negative")
        self._desenhar_ligacoes()

    async def _exemplos(self) -> None:
        """Planilha, ato pronto e modelo de exemplo em saidas/Exemplos, para a pessoa ver como são."""
        pasta = main.SAIDAS / "Exemplos"
        await run.io_bound(gerar_ato_pronto_exemplo, pasta)
        abrir(pasta)
        gui.notify("Exemplos abertos: a planilha, um ato pronto e o modelo que o Edodocs cria", type="positive")

    def _ir_ligar(self) -> None:
        if not self.modelo or not self.campos:
            gui.notify("Volte ao passo 2 e prepare o Word", type="warning")
            return
        criada = self.planilha is not None and str(self.planilha).startswith(str(self._temp))
        if not (self.registros or criada):
            gui.notify("Escolha a sua planilha ou clique em Criar a planilha para mim", type="warning")
            return
        self._desenhar_ligacoes()
        self.passos.next()

    # --- Passo 3: conferir os campos ---------------------------------------------------
    def _passo_ligar(self) -> None:
        _passo(4, "Confira os campos",
               "Cada campo do modelo e de onde vem o valor, com um exemplo. Quase tudo já vem certo: "
               "só mexa nos que estiverem com a bolinha vermelha.", lambda: self._info("campos"))
        with gui.column().classes("w-full gap-2 q-pl-xl"):
            with gui.expansion("Mais de uma pessoa no mesmo ato", icon="groups") \
                    .classes("w-full fundo-cartao rounded") as self.expansao_varias:
                self.area_varias = gui.column().classes("w-full gap-2 q-pa-sm")
            with gui.row().classes("w-full gap-4 items-start no-wrap"):
                self.area_ligacoes = gui.column().classes("col gap-2")
        with gui.stepper_navigation():
            gui.button("Próximo", icon="arrow_forward", on_click=self._ir_previa).props("unelevated")
            gui.button("Voltar", on_click=self.passos.previous).props("flat")

    def _regra(self, rotulo: str) -> dict:
        """Ligação do campo (a salva, ajustada aos nomes das colunas, ou a sugerida)."""
        ligacoes = self.config["ligacoes"]
        chave = next((k for k in ligacoes if atos.normalizar(k) == atos.normalizar(rotulo)), None)
        if chave is None:
            chave = rotulo
            ligacoes[chave] = self._sugerir(rotulo)
        elif chave != rotulo:
            ligacoes[rotulo] = ligacoes.pop(chave)
        regra = ligacoes[rotulo]
        if (regra["tipo"] == "coluna" and regra.get("coluna") and self._sem_dados(regra["coluna"])
                and not regra.get("se_vazio")):
            # Ligado a uma coluna sem nenhum valor (ex.: DIA EXT com #NAME? em todas as linhas):
            # nunca é o que a pessoa quer; usa a sugestão, se houver uma melhor
            sugestao = self._sugerir(rotulo)
            if sugestao["tipo"] != "coluna":
                ligacoes[rotulo] = regra = sugestao
        if regra.get("coluna"):
            regra["coluna"] = _coluna_igual(self.colunas, regra["coluna"]) or regra["coluna"]
        elif regra["tipo"] == "coluna":
            regra["coluna"] = _coluna_igual(self.colunas, rotulo) or ""
        return regra

    def _sem_dados(self, coluna: str) -> bool:
        """A coluna existe, mas está vazia em todas as linhas da planilha. Fórmula com erro
        (#NAME?, #VALUE!...) também conta como vazia: é o que acontece com uma coluna de
        "por extenso" feita com função de outro arquivo do Excel."""
        nome = atos.normalizar(coluna)
        # A linha de exemplo (o que estava no Word, com a planilha nova ainda vazia) não conta:
        # um campo que só aparece num texto de opção ({{NÚCLEO}} na opção COMISSÃO) não está
        # no Word de exemplo, e não é por isso que a coluna está errada
        reais = [r for r in self.registros if not r.get("_EXEMPLO")]
        return bool(reais) and not any(str(r.get(nome, "")).strip() for r in reais)

    def _sugerir(self, rotulo: str) -> dict:
        """Ligação sugerida para um campo novo: a coluna de mesmo nome, se ela tiver dados.
        Campo "por extenso" (DIA EXT, DIAS POR EXTENSO, DATA EXTENSO) sem coluna útil vira
        número ou data por extenso da coluna de base (DIA EXT → DIAS)."""
        coluna = _coluna_igual(self.colunas, rotulo) or ""
        if coluna and not self._sem_dados(coluna):
            return {"tipo": "coluna", "coluna": coluna}
        nome = atos.normalizar(rotulo)
        base = re.sub(r"(POREXTENSO|EXTENSO|EXT)$", "", nome)
        if base and base != nome:
            for candidata in self.colunas:
                normal = atos.normalizar(candidata)
                if not normal.startswith(base) or normal == nome or self._sem_dados(candidata):
                    continue
                valores = [str(r.get(normal, "")).strip() for r in self.registros]
                valores = [v for v in valores if v]
                if all(v.replace(".", "").isdigit() for v in valores):
                    return {"tipo": "extenso", "coluna": candidata}
                if all(atos.ler_data(v) for v in valores):
                    return {"tipo": "data_extenso", "coluna": candidata}
        return {"tipo": "coluna", "coluna": coluna}

    def _ligada(self, regra: dict) -> bool:
        if regra["tipo"] == "fixo":
            return bool(regra.get("texto"))
        if regra["tipo"] == "condicao":
            return bool(regra.get("casos") or regra.get("senao"))
        # A coluna existir basta: vazia em todas as linhas pode ser só um campo opcional ainda
        # não usado (as datas da aposentadoria, INICIAIS...). Isso vira aviso (bolinha
        # amarela, ver _cor), não trava o Próximo nem o Salvar
        return bool(_coluna_igual(self.colunas, regra.get("coluna", "")))

    def _cor(self, regra: dict) -> str:
        """Bolinha do campo: verde (ok), amarela (a coluna existe, mas ainda não tem nenhum
        valor na planilha: confira) ou vermelha (falta dizer de onde vem)."""
        if not self._ligada(regra):
            return "vermelho"
        coluna = _coluna_igual(self.colunas, regra.get("coluna", "")) if regra.get("coluna") else ""
        if coluna and regra["tipo"] not in ("fixo", "condicao") and self._sem_dados(coluna)                 and not regra.get("se_vazio"):
            return "amarelo"
        return "verde"

    def _atos(self) -> list[dict]:
        """Os atos da planilha (com várias pessoas, as linhas do mesmo ato juntas)."""
        return self._processo_rascunho().agrupar(self.registros)

    def _registro_previa(self) -> dict | None:
        """O ato escolhido na prévia."""
        lista = self._atos()
        return lista[min(self.linha_previa, len(lista) - 1)] if lista else None

    def _exemplo(self, rotulo: str) -> str:
        """O valor do campo para o ato da prévia (ex.: "quinze")."""
        registro = self._registro_previa()
        if registro is None or not self._ligada(self._regra(rotulo)):
            return ""
        return str(self._processo_rascunho().valores(registro).get(atos.normalizar(rotulo), ""))

    def _desenhar_ligacoes(self) -> None:
        if not hasattr(self, "area_ligacoes"):
            return
        self._atualizar_word_simulado()
        self._desenhar_varias()
        self.area_ligacoes.clear()
        with self.area_ligacoes:
            for _, rotulo in self.campos:
                self._editor(rotulo)
        # Remove ligações de campos que não estão mais no Word (só depois de ler os campos:
        # antes disso a lista está vazia e apagaria as ligações salvas do processo)
        if not self.campos:
            return
        atuais = {atos.normalizar(r) for _, r in self.campos}
        for chave in [k for k in self.config["ligacoes"] if atos.normalizar(k) not in atuais]:
            del self.config["ligacoes"][chave]

    # --- Várias pessoas no mesmo ato (passo 3) ---
    def _desenhar_varias(self) -> None:
        """Juntar linhas num ato, o trecho do Word que repete e os outros Words."""
        if not hasattr(self, "area_varias"):
            return
        juntar = self.config.setdefault("juntar", {})
        outros = self.config.setdefault("outros_modelos", [])
        ligado = bool(juntar.get("coluna"))
        if ligado:
            self.expansao_varias.open()
        self.area_varias.clear()
        with self.area_varias:
            with gui.row().classes("items-center gap-1"):
                gui.switch("Um mesmo ato pode ter mais de uma pessoa (ex.: exoneração de duas pessoas)",
                           value=ligado, on_change=self._ligar_varias)
                self._info("varias")
            if not ligado:
                gui.label("Desligado: cada linha da planilha é um ato.").classes("text-caption t-suave")
                return

            # 1. Que linhas vão juntas
            with gui.row().classes("w-full items-center gap-3 no-wrap"):
                gui.label("1.").classes("text-weight-bold text-primary")
                gui.select(self.colunas, value=_coluna_igual(self.colunas, juntar.get("coluna", "")),
                           label="Linhas com o mesmo valor nesta coluna viram um ato só",
                           on_change=lambda e: self._mudar_juntar(e.value)) \
                    .props("outlined dense use-input input-debounce=0").classes("w-[460px]")
                lista = self._atos()
                varios = sum(len(atos.itens(a)) > 1 for a in lista)
                gui.label(f"Na planilha: {len(lista)} ato(s), {varios} com mais de uma pessoa.") \
                    .classes("text-body2 t-suave")

            # 2. O trecho que repete em cada Word
            with gui.row().classes("w-full items-start gap-3 no-wrap"):
                gui.label("2.").classes("text-weight-bold text-primary")
                with gui.column().classes("col gap-1"):
                    gui.label("Em cada Word, a parte que se repete para cada pessoa:").classes("text-body2 t-forte")
                    for modelo in self._processo_rascunho().modelos():
                        resumo = regras_ui.resumo_trecho(modelo)
                        with gui.row().classes("w-full items-center gap-2 no-wrap"):
                            gui.icon("check_circle" if resumo else "radio_button_unchecked",
                                     color="positive" if resumo else "grey-6")
                            gui.label(f"{modelo.name}: " + (f"“{resumo}”" if resumo else "nada se repete")) \
                                .classes("text-body2 t-forte ellipsis col")
                            gui.button("Escolher o trecho", icon="repeat",
                                       on_click=lambda m=modelo: regras_ui.escolher_trecho(
                                           self.tela.camada, m, self._depois_de_mudar_word)) \
                                .props("outline dense no-caps")
                    gui.label("No trecho que repete, estes campos são preenchidos sozinhos: "
                              + " · ".join("{{" + r + "}} " + ex for r, ex in atos.AUTOMATICOS.values())) \
                        .classes("text-caption t-suave")

            # 3. Outros Words (ex.: um para uma pessoa, outro para várias)
            with gui.row().classes("w-full items-start gap-3 no-wrap"):
                gui.label("3.").classes("text-weight-bold text-primary")
                with gui.column().classes("col gap-1"):
                    with gui.row().classes("items-center gap-1"):
                        gui.label("O texto muda muito com mais de uma pessoa? Use outro Word para esses casos:") \
                            .classes("text-body2 t-forte")
                        self._info("outros_modelos")
                    for i, outro in enumerate(outros):
                        with gui.card().classes("cartao w-full q-pa-sm"):
                            with gui.row().classes("w-full items-center gap-2 no-wrap"):
                                gui.icon("description", color="primary")
                                with gui.column().classes("gap-0 col"):
                                    gui.label(Path(outro.get("modelo", "")).name).classes("text-weight-medium t-forte")
                                    gui.label("Usado quando " + regras_ui.descrever_condicoes(outro.get("se") or [])) \
                                        .classes("text-caption t-suave")
                                gui.button("Quando usar", icon="rule",
                                           on_click=lambda o=outro: self._condicoes_do_modelo(o)) \
                                    .props("flat dense no-caps")
                                gui.button("Abrir", icon="edit_document",
                                           on_click=lambda o=outro: abrir(self._caminho_outro(o))) \
                                    .props("flat dense no-caps")
                                gui.button(icon="delete", on_click=lambda i=i: self._remover_outro(i)) \
                                    .props("flat round dense color=negative").tooltip("Não usar mais este Word")
                    with gui.row().classes("items-center gap-2"):
                        gui.button("Adicionar outro Word...", icon="note_add", on_click=self._adicionar_outro) \
                            .props("outline dense no-caps")
                        gui.label(f"Os outros casos usam {self.modelo.name if self.modelo else 'o Word principal'}.") \
                            .classes("text-caption t-suave")

    def _ligar_varias(self, e) -> None:
        juntar = self.config.setdefault("juntar", {})
        if e.value:
            juntar["coluna"] = juntar.get("coluna") or next(
                (c for c in self.colunas if re.search(r"DECRETO|ATO|PORTARIA|GRUPO", atos.normalizar(c))),
                (self.colunas or [""])[0])
        else:
            juntar["coluna"] = ""
        self.config["_tipo"] = "varias" if e.value else "uma"  # o passo 1 mostra o mesmo
        self._mostrar_tipo()
        self._desenhar_trecho()
        self._desenhar_varias()
        self._desenhar_previa()

    def _mudar_juntar(self, coluna) -> None:
        self.config.setdefault("juntar", {})["coluna"] = coluna or ""
        self.linha_previa = 0
        self._desenhar_varias()
        self._desenhar_previa()

    def _caminho_outro(self, outro: dict) -> Path:
        caminho = Path(outro.get("modelo", ""))
        return caminho if caminho.is_absolute() else (self.pasta or Path(".")) / caminho

    async def _adicionar_outro(self) -> None:
        caminho = await self._escolher_arquivo("Documento do Word", ".docx")
        if not caminho:
            return
        # Trabalha numa cópia (vai para a pasta do processo ao salvar): marcar o trecho que
        # repete não mexe no arquivo original da pessoa
        destino = self._temp / caminho.name
        shutil.copy2(caminho, destino)
        outro = {"modelo": str(destino), "se": [{"campo": regras_ui.ROTULO_QUANTIDADE, "op": "maior", "valor": "1"}]}
        self.config.setdefault("outros_modelos", []).append(outro)
        await self._depois_de_mudar_word_async()
        self._condicoes_do_modelo(outro)

    def _remover_outro(self, i: int) -> None:
        del self.config["outros_modelos"][i]
        self._depois_de_mudar_word()

    def _condicoes_do_modelo(self, outro: dict) -> None:
        outro.setdefault("se", [])
        regras_ui.editar_condicoes(
            self.tela.camada, f"Quando usar {Path(outro['modelo']).name}?",
            "O Edodocs usa este Word quando todas as condições valem; nos outros casos, o Word principal. "
            "Ex.: QUANTIDADE DE PESSOAS é maior que 1.",
            outro["se"], self.colunas, self.registros, self._depois_de_mudar_word)

    def _depois_de_mudar_word(self) -> None:
        background_tasks.create(self._depois_de_mudar_word_async())

    async def _depois_de_mudar_word_async(self) -> None:
        """Um Word mudou (trecho marcado, Word novo): relê os campos e refaz a prévia."""
        await self._ler_campos()
        self._desenhar_previa()

    def _editor(self, rotulo: str) -> None:
        regra = self._regra(rotulo)
        with gui.card().classes("cartao w-full q-pa-sm"):
            with gui.row().classes("w-full items-center gap-3 no-wrap"):
                bolinha = gui.element("span").classes(f"bolinha b-{self._cor(regra)}")
                with gui.column().classes("gap-0 w-56"):
                    gui.label("{{" + rotulo + "}}").classes("text-weight-medium ellipsis t-forte")
                    exemplo = gui.label("").classes("text-caption t-suave")
                gui.select(processos.TIPOS, value=regra["tipo"],
                           on_change=lambda e: self._trocar_tipo(rotulo, e.value)) \
                    .props("dense outlined").classes("w-48")
                detalhes = gui.row().classes("col items-center gap-2 no-wrap")

            def mostrar_exemplo() -> None:
                valor = self._exemplo(rotulo)
                if valor:
                    exemplo.text = f"Ex.: {valor}"
                elif regra["tipo"] == "condicao" and self._ligada(regra):
                    exemplo.text = "Ex.: (nada, neste ato)"  # vazio de propósito pela regra
                elif regra.get("coluna") and _coluna_igual(self.colunas, regra["coluna"])                         and self._sem_dados(regra["coluna"]):
                    exemplo.text = (f"Coluna {regra['coluna']} ainda sem nenhum valor na planilha. Normal se o "
                                    "campo é opcional; se deveria ter, confira (fórmula com erro, como #NAME?)")
                elif regra.get("coluna") and _coluna_igual(self.colunas, regra["coluna"]):
                    exemplo.text = f"Vem da coluna {regra['coluna']} (o exemplo aparece quando ela tiver dados)"
                else:
                    exemplo.text = "Sem exemplo (confira a ligação)"

            def mudar(chave: str, valor) -> None:
                regra[chave] = valor or ""
                bolinha.classes(replace=f"bolinha b-{self._cor(regra)}")
                mostrar_exemplo()

            mostrar_exemplo()
            with detalhes:
                if regra["tipo"] not in ("fixo", "condicao"):
                    gui.select(self.colunas, value=_coluna_igual(self.colunas, regra.get("coluna", "")),
                               label="Coluna da data de início" if regra["tipo"] == "periodo" else "Coluna",
                               on_change=lambda e: mudar("coluna", e.value)) \
                        .props("dense outlined use-input input-debounce=0").classes("col")
                if regra["tipo"] == "periodo":
                    gui.select(self.colunas, value=_coluna_igual(self.colunas, regra.get("ate", "")),
                               label="Coluna da data de fim", on_change=lambda e: mudar("ate", e.value)) \
                        .props("dense outlined use-input input-debounce=0").classes("col")
                if regra["tipo"] == "coluna":
                    gui.input("Se vazio, usar (opcional)", value=regra.get("se_vazio", ""),
                              on_change=lambda e: mudar("se_vazio", e.value)).props("dense outlined").classes("col")
                    # MAIÚSCULAS / Primeiras Maiúsculas (ex.: "CARGO: PROFESSOR" na tabela do concurso)
                    gui.select({"": "Letras como estão", "maiusculas": "MAIÚSCULAS",
                                "titulo": "Primeiras Maiúsculas", "minusculas": "minúsculas"},
                               value=regra.get("formato", ""), label="Letras",
                               on_change=lambda e: mudar("formato", e.value)) \
                        .props("dense outlined options-dense").classes("w-44")
                elif regra["tipo"] == "fixo":
                    gui.input("Texto", value=regra.get("texto", ""),
                              on_change=lambda e: mudar("texto", e.value)).props("dense outlined").classes("col")
                elif regra["tipo"] == "genero":
                    gui.input("Se feminino (F)", value=regra.get("feminino", ""),
                              on_change=lambda e: mudar("feminino", e.value)).props("dense outlined").classes("col")
                    gui.input("Se masculino (M)", value=regra.get("masculino", ""),
                              on_change=lambda e: mudar("masculino", e.value)).props("dense outlined").classes("col")
                elif regra["tipo"] == "condicao":
                    casos = regra.get("casos") or []
                    resumo = f"{len(casos)} regra(s)" + (" + texto para os outros casos" if regra.get("senao") else "")
                    gui.label(resumo if casos or regra.get("senao") else "Nenhuma regra ainda") \
                        .classes("text-body2 t-suave col")

                    def fechou() -> None:
                        # Relê os campos: um {{CAMPO}} novo no texto das regras entra na lista
                        self._depois_de_mudar_word()

                    gui.button("Editar regras", icon="alt_route", on_click=lambda: regras_ui.editar_texto_condicional(
                        self.tela.camada, rotulo, regra, self.colunas, self.registros,
                        lambda: (self._descricao_previa(), self._exemplo(rotulo)), fechou)) \
                        .props("unelevated dense no-caps")

    def _trocar_tipo(self, rotulo: str, tipo: str) -> None:
        anterior = self.config["ligacoes"][rotulo]
        if tipo == "condicao":
            self.config["ligacoes"][rotulo] = {"tipo": tipo, "casos": [], "senao": ""}
        else:
            self.config["ligacoes"][rotulo] = {"tipo": tipo, "coluna": anterior.get("coluna", "")}
        self._desenhar_ligacoes()

    def _descricao_previa(self) -> str:
        registro = self._registro_previa()
        return self._processo_rascunho().descricao(registro) if registro else ""

    def _ir_previa(self) -> None:
        sem_ligacao = [r for _, r in self.campos if not self._ligada(self._regra(r))]
        if sem_ligacao:
            gui.notify("Escolha de onde vêm os campos com bolinha vermelha: " + ", ".join(sem_ligacao),
                       type="warning")
            return
        self._desenhar_previa()
        self.passos.next()

    # --- Passo 4: como vai ficar -----------------------------------------------------------
    def _passo_previa(self) -> None:
        _passo(5, "Veja como o ato vai ficar",
               "Escolha uma pessoa da planilha e confira o ato completo. Se estiver tudo certo, é só salvar.",
               lambda: self._info("titulos"))
        if self._ia is not None:
            self._cartao_da_ia()
        with gui.row().classes("w-full gap-4 items-start no-wrap q-pl-xl"):
            with gui.card().classes("cartao q-pa-md col").style("min-width: 0"):
                with gui.row().classes("w-full items-center"):
                    gui.label("Como vai ficar").classes("text-subtitle1 text-weight-medium t-forte")
                    gui.space()
                    self.sel_previa = gui.select({}, label="Pessoa", on_change=self._trocar_previa) \
                        .props("dense outlined").classes("w-80")
                with gui.tabs(on_change=self._aba_previa) \
                        .props("dense no-caps align=left active-color=primary indicator-color=primary") \
                        .classes("t-forte") as abas:
                    gui.tab("word", "O Word", icon="description")
                    gui.tab("edocs", "No e-docs", icon="send")
                    self.aba_diario = gui.tab("diario", "No Diário Oficial", icon="newspaper")
                with gui.tab_panels(abas, value="word").classes("w-full bg-transparent"):
                    with gui.tab_panel("word").classes("q-pa-none q-pt-sm"):
                        self.previa = gui.html('<div class="folha-area" style="height:560px"></div>') \
                            .classes("w-full")
                    with gui.tab_panel("edocs").classes("q-pa-none q-pt-sm"):
                        self._simulacao_edocs()
                    with gui.tab_panel("diario").classes("q-pa-none q-pt-sm"):
                        self._simulacao_diario()
            with gui.column().classes("w-[440px] gap-3"):
                with gui.card().classes("cartao q-pa-md w-full"):
                    gui.label("Tudo pronto?").classes("text-subtitle1 text-weight-medium t-forte")
                    self.area_checklist = gui.column().classes("w-full gap-1")
                with gui.card().classes("cartao q-pa-md w-full"):
                    with gui.row().classes("items-center gap-1"):
                        gui.label("Títulos").classes("text-subtitle1 text-weight-medium t-forte")
                        self._info("titulos")
                    self.area_titulos = gui.column().classes("w-full gap-2")
                with gui.card().classes("cartao q-pa-md w-full"):
                    with gui.row().classes("items-center gap-2"):
                        gui.icon("newspaper", color="primary")
                        gui.label("Diário Oficial").classes("text-subtitle1 text-weight-medium t-forte")
                        self._info("diario")
                    self.config.setdefault("vai_ao_diario", True)
                    gui.switch("Este processo será publicado no Diário Oficial",
                               value=bool(self.config["vai_ao_diario"]), on_change=self._mudar_vai_ao_diario)
                    self.explica_diario = gui.label("").classes("text-caption t-suave")
                with gui.card().classes("cartao q-pa-md w-full"):
                    with gui.row().classes("items-center gap-2"):
                        gui.icon("pan_tool", color="primary")
                        gui.label("Envio").classes("text-subtitle1 text-weight-medium t-forte")
                        self._info("envio")
                    self.config.setdefault("confirmar_envio",
                                           main.ler_config().get("confirmar_envio", True))
                    gui.switch("Perguntar antes de enviar cada ato", value=bool(self.config["confirmar_envio"]),
                               on_change=lambda e: self.config.update(confirmar_envio=e.value))
                    gui.label("Ligado: você vê cada ato e decide enviar, pular ou cancelar. Desligado: os atos "
                              "vão um atrás do outro, sem parar.").classes("text-caption t-suave")
                    gui.switch("Precisa de anexo", value=bool(self.config.get("precisa_anexo")),
                               on_change=lambda e: self.config.update(precisa_anexo=e.value)).classes("q-mt-sm")
                    gui.label("Ligado: a cada ato, o Edodocs pede o(s) PDF(s) do computador para anexar no "
                              "e-docs (cada ato tem o seu).").classes("text-caption t-suave")
                with gui.card().classes("cartao q-pa-md w-full"):
                    with gui.row().classes("items-center gap-2"):
                        gui.icon("how_to_reg", color="primary")
                        gui.label("Vistos e assinantes").classes("text-subtitle1 text-weight-medium t-forte")
                        self._info("pessoas")
                    self.painel_pessoas = PainelPessoas(self.config["pessoas"],
                                                        ao_mudar=self._mostrar_pessoas_simuladas)
                with gui.card().classes("cartao q-pa-md w-full"):
                    with gui.row().classes("items-center gap-2"):
                        gui.icon("view_column", color="primary")
                        gui.label("Colunas da tabela dos atos").classes("text-subtitle1 text-weight-medium t-forte")
                        self._info("colunas")
                    self.config.setdefault("caixas", [])
                    self.config.setdefault("colunas_tabela", [])
                    self.editor_colunas = EditorColunas(
                        self.config, lambda: self._processo_rascunho().colunas_tabela(self.campos))
                with gui.expansion("Ajustar títulos", icon="title").classes("w-full fundo-cartao rounded"):
                    with gui.column().classes("w-full gap-2 q-pa-sm"):
                        self.sel_numero = gui.select([], label="Coluna do número do ato",
                                                     on_change=lambda e: self._mudar_titulo("coluna_numero", e.value)) \
                            .props("outlined dense use-input input-debounce=0").classes("w-full")
                        self.inp_edocs = gui.input("Título no e-docs", value=self.config["titulo_edocs"],
                                                   on_change=lambda e: self._mudar_titulo("titulo_edocs", e.value)) \
                            .props("outlined dense").classes("w-full")
                        self.inp_diario = gui.input("Título no Diário Oficial", value=self.config["titulo_diario"],
                                                    on_change=lambda e: self._mudar_titulo("titulo_diario", e.value)) \
                            .props("outlined dense").classes("w-full")
                        self.inp_no_ato = gui.input("Título na 1ª linha do ato (opcional)",
                                                    value=self.config.get("titulo_no_ato", ""),
                                                    placeholder="DECRETO S/Nº - EXONERA {NOME}",
                                                    on_change=lambda e: self._mudar_titulo("titulo_no_ato", e.value)) \
                            .props("outlined dense").classes("w-full") \
                            .tooltip("Troca a 1ª linha do Word no ato do e-docs. Sem limite de tamanho; "
                                     "com várias pessoas termina em E OUTROS. Em branco: fica a linha do Word.")
                        gui.label("{COLUNA} vira o valor da coluna. Especiais: {número do ato} e {ano do ato}.") \
                            .classes("text-caption t-suave")
        self._mudar_vai_ao_diario()  # explicação e aba do Diário conforme o processo
        with gui.stepper_navigation():
            gui.button("Está certo, salvar processo", icon="save", color="positive",
                       on_click=self._salvar).props("unelevated no-caps")
            gui.button("Voltar", on_click=self.passos.previous).props("flat")

    # --- Montado pela IA: o que ela fez e o que falta a pessoa fazer -----------------------
    FALTA_FAZER = [
        ("person_search", "Troque a pessoa em “Pessoa” e leia o ato de cada exemplo. Errado? Volte ao "
                          "passo 4 (Confira os campos) e ajuste a regra."),
        ("how_to_reg", "Escolha os vistos e assinantes (ou deixe para escolher a cada envio)."),
        ("title", "Confira os títulos do e-docs e do Diário e se o processo vai ao Diário."),
        ("save", "Clique em “Está certo, salvar processo”. Depois é só preencher a planilha."),
    ]

    def _lista_falta_fazer(self) -> None:
        for n, (icone, texto) in enumerate(self.FALTA_FAZER, 1):
            with gui.row().classes("items-start gap-2 no-wrap"):
                gui.icon(icone, color="primary", size="20px").classes("q-mt-xs")
                gui.label(f"{n}. {texto}").classes("text-body2 t-forte")

    def _cartao_da_ia(self) -> None:
        dados = self._ia
        with gui.card().classes("cartao q-pa-md gap-2 q-mb-md").style(
                "border-left: 5px solid var(--q-primary); margin-left: 48px"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("auto_awesome", color="primary", size="26px")
                gui.label("A IA montou este processo. Agora é com você: confira e finalize.") \
                    .classes("text-subtitle1 text-weight-bold t-forte")
            if dados.get("explicacao"):
                gui.label(dados["explicacao"]).classes("text-body2 t-forte")
            with gui.row().classes("w-full gap-6 items-start"):
                with gui.column().classes("gap-1 col").style("min-width: 260px"):
                    gui.label("O que o Edodocs conferiu").classes("text-caption text-weight-bold t-suave")
                    for arquivo, igual, parecido in dados.get("conferidos") or []:
                        _conferencia(bool(igual), f"{arquivo}: " + ("refeito igual ao original" if igual
                                                                    else f"{parecido}% igual, confira"))
                    for pergunta, resposta in dados.get("respostas") or []:
                        _conferencia(True, f"{pergunta} Você respondeu: {resposta}")
                    for aviso in (dados.get("avisos") or []) + (dados.get("duvidas") or []):
                        _conferencia(None, f"Dica para a planilha: {aviso}")
                    if dados.get("problemas"):
                        _conferencia(False, f"{len(dados['problemas'])} ponto(s) a IA não conseguiu acertar. "
                                            "Veja no ato e ajuste no passo 4.")
                with gui.column().classes("gap-1 col").style("min-width: 260px"):
                    gui.label("Falta você fazer").classes("text-caption text-weight-bold t-suave")
                    self._lista_falta_fazer()

    def _contar_o_que_a_ia_fez(self) -> None:
        """Ao abrir o que a IA montou: o resumo e o pedido para a pessoa finalizar."""
        dados = self._ia
        conferidos = dados.get("conferidos") or []
        iguais = sum(1 for _, igual, _ in conferidos if igual)
        with self.tela.camada, gui.dialog() as janela, \
                gui.card().classes("cartao w-[600px] max-w-full q-pa-lg gap-3"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("auto_awesome", color="primary", size="32px")
                gui.label(f"Pronto! A IA montou o processo “{self.nome}”").classes("text-h6 t-forte")
            if dados.get("explicacao"):
                gui.label(dados["explicacao"]).classes("text-body1")
            if conferidos:
                _conferencia(iguais == len(conferidos),
                             f"O Edodocs refez {len(conferidos)} exemplo(s) com o que a IA montou: {iguais} "
                             "saíram iguais ao original." + ("" if iguais == len(conferidos)
                                                             else " Confira os outros na prévia."))
            gui.label("O Word, a planilha e as regras estão prontos, mas nada foi salvo ainda. Falta você:") \
                .classes("text-body2 t-forte")
            with gui.column().classes("gap-1 q-pl-sm"):
                self._lista_falta_fazer()
            gui.label("Os passos 1 a 4 estão lá em cima, se quiser ver ou mudar algo (o nome do processo fica "
                      "no passo 1).").classes("text-caption t-suave")
            with gui.row().classes("w-full justify-end"):
                gui.button("Conferir e finalizar", icon="fact_check", on_click=janela.close) \
                    .props("unelevated no-caps")
        janela.open()

    def _mudar_vai_ao_diario(self, e=None) -> None:
        """Vai ao Diário: ele recebe o ato sem a 1ª linha (o título vai no campo próprio).
        Não vai: a tela do Diário e o título do Diário somem do passo 4."""
        if e is not None:
            self.config["vai_ao_diario"] = e.value
        vai = bool(self.config.get("vai_ao_diario", True))
        if vai:
            self.config["titulo_na_primeira_linha"] = True
        self.explica_diario.text = (
            "Sim: depois do e-docs, o ato é cadastrado no Diário (sem a 1ª linha do Word, porque o "
            "título vai num campo próprio) e a publicação é conferida." if vai else
            "Não: o processo só vai ao e-docs. Lançar no Diário e Conferir publicação ficam "
            "desligados. Dá para mudar depois em Processos.")
        self.aba_diario.set_visibility(vai)
        self._desenhar_titulos()

    # --- Telas dos sites reproduzidas (passo 4) ---
    @staticmethod
    def _titulo_da_tela(icone: str, texto: str) -> None:
        with gui.row().classes("items-center gap-2 no-wrap q-px-sm q-py-xs rounded-borders") \
                .style("background: var(--q-primary); color: white; width: fit-content"):
            gui.icon(icone, size="18px")
            gui.label(texto).classes("text-caption text-weight-bold")

    @staticmethod
    def _campo_fixo(rotulo: str, valor: str = "", classes: str = "col"):
        """Campo só para ver (como aparece no site); devolve o input para atualizar o valor."""
        return gui.input(rotulo, value=valor).props("outlined dense stack-label readonly") \
            .classes(classes + " campo-site")

    def _simulacao_edocs(self) -> None:
        edocs = self.config["edocs"]
        with gui.card().classes("w-full fundo-cartao q-pa-md gap-2"):
            with gui.row().classes("w-full items-center"):
                self._titulo_da_tela("send", "e-docs · Instrumentos Jurídicos/Normativos")
                gui.space()
                gui.label("Os campos com borda azul você pode mudar aqui.").classes("text-caption t-suave")
                self._info("sites")
            with gui.row().classes("w-full gap-3 no-wrap items-start"):
                # Listas com as opções que ESTE usuário tem no e-docs (buscadas no site)
                posicao = int(edocs.get("setor_posicao") or 1)
                self.sel_setor = gui.select({posicao: edocs.get("setor") or f"{posicao}º setor da sua lista"},
                                            value=posicao, label="Setor", on_change=self._trocar_setor) \
                    .props("outlined dense stack-label options-dense").classes("col campo-editavel") \
                    .style("min-width: 0")
                self.sel_tipo = gui.select([edocs["tipo"]], value=edocs["tipo"], label="Tipo",
                                           on_change=lambda e: e.value and edocs.update(tipo=e.value)) \
                    .props("outlined dense stack-label options-dense").classes("col campo-editavel") \
                    .style("min-width: 0")
                self._campo_fixo("Data", date.today().strftime("%d/%m/%Y"), "w-36")
            with gui.row().classes("w-full items-center gap-2 no-wrap -mt-1"):
                self.status_opcoes = gui.label("As listas de Setor e Tipo vêm do seu e-docs.") \
                    .classes("text-caption t-suave col")
                self.botao_opcoes = gui.button("Buscar no e-docs", icon="sync",
                                               on_click=lambda: self._carregar_opcoes_edocs()) \
                    .props("flat dense no-caps size=sm")
            with gui.row().classes("w-full gap-3 no-wrap"):
                self.sim_identificacao = self._campo_fixo("Nossa Identificação (o título no e-docs)")
                self._campo_fixo("Identificação Sistema", "o e-docs cria ao gravar")
            with gui.row().classes("w-full gap-3 no-wrap items-start"):
                with gui.column().classes("col gap-0").style("min-width: 0"):
                    gui.label("Texto do documento").classes("text-caption text-weight-bold text-primary")
                    with gui.row().classes("w-full items-center gap-3 q-px-sm q-py-xs barra-editor"):
                        for icone in ("undo", "redo", "format_bold", "format_italic", "format_underlined",
                                      "format_align_left", "format_align_center", "format_align_right",
                                      "format_align_justify", "format_indent_decrease", "format_indent_increase"):
                            gui.icon(icone, size="18px")
                    self.sim_texto_edocs = gui.html("").classes("w-full caixa-editor")
                with gui.column().classes("w-48 gap-3 q-pt-md"):
                    self.sim_pessoas = {}
                    for chave, icone, rotulo in (("vistos", "fact_check", "Visto(s)"),
                                                 ("assinantes", "draw", "Assinante(s)"),
                                                 ("anexos", "attach_file", "Doc. Anexo(s)")):
                        with gui.column().classes("gap-1 w-full"):
                            self._titulo_da_tela(icone, rotulo)
                            self.sim_pessoas[chave] = gui.label("").classes("text-caption t-forte")
                    # O grupo de destinatários: a lista vem da conta da pessoa no e-docs
                    with gui.column().classes("gap-1 w-full"):
                        self._titulo_da_tela("groups", "Dest. Publicador(es)")
                        grupo = edocs.get("grupo_destinatario", "")
                        self.sel_grupo = gui.select({grupo: grupo or "o 1º grupo da sua lista"}, value=grupo,
                                                    label="Grupo", on_change=self._trocar_grupo) \
                            .props("outlined dense stack-label options-dense").classes("w-full campo-editavel")
        self._mostrar_pessoas_simuladas()

    def _aba_previa(self, e) -> None:
        """Na primeira vez que a aba do e-docs abre, busca as opções do usuário no site."""
        if e.value == "edocs" and not getattr(self, "_opcoes_buscadas", False):
            self._opcoes_buscadas = True
            background_tasks.create(self._carregar_opcoes_edocs())

    async def _sessao_edocs(self):
        """Sessão no e-docs com o usuário salvo em Configurações (entra uma vez só)."""
        import credenciais
        import edocs_http

        if getattr(self, "_sessao", None) is None:
            dados = credenciais.carregar(main.CREDENCIAIS, "edocs")
            if not dados:
                raise RuntimeError("cadastre seu usuário do e-docs em Configurações para ver as suas opções")
            self._sessao = await run.io_bound(edocs_http.entrar, *dados)
        return self._sessao

    async def _carregar_opcoes_edocs(self) -> None:
        """Setores e tipos que este usuário tem no e-docs, nas listas da tela reproduzida."""
        import edocs_http

        edocs = self.config["edocs"]
        posicao = int(edocs.get("setor_posicao") or 1)
        self.status_opcoes.text = "Buscando as suas opções no e-docs..."
        self.botao_opcoes.props("loading")
        try:
            for tentativa in range(2):
                try:
                    sessao = await self._sessao_edocs()
                    opcoes = await run.io_bound(edocs_http.opcoes_do_documento, sessao, posicao)
                    try:  # os grupos de destinatários da conta (Dest. Publicador)
                        opcoes["grupos"] = await run.io_bound(edocs_http.grupos_destinatarios, sessao)
                    except edocs_http.ErroEdocs:
                        opcoes["grupos"] = []
                    break
                except edocs_http.SessaoExpirada:
                    self._sessao = None  # a sessão caiu: entra de novo
            else:
                raise RuntimeError("a sessão do e-docs caiu; tente de novo")
        except Exception as e:
            self.status_opcoes.text = f"Não deu para buscar as opções no e-docs: {e}"
            return
        finally:
            self.botao_opcoes.props(remove="loading")
        setores, tipos = opcoes["setores"], opcoes["tipos"]
        grupos = opcoes.get("grupos") or []
        if grupos:
            atual = edocs.get("grupo_destinatario", "")
            lista = {g: g for g in grupos}
            if atual and atual not in lista:  # o grupo salvo não aparece para esta conta
                lista[atual] = f"{atual} (não está na sua lista)"
            self.sel_grupo.set_options(lista, value=atual if atual else grupos[0])
            if not atual:
                edocs["grupo_destinatario"] = grupos[0]
        if setores:
            posicao = min(posicao, len(setores))
            self._ajustando_setor = True  # trocar as opções não é a pessoa escolhendo outro setor
            self.sel_setor.set_options({i: f"{i}º · {t}" for i, t in enumerate(setores, 1)}, value=posicao)
            self._ajustando_setor = False
            edocs.update(setor_posicao=posicao, setor=setores[posicao - 1])
        aviso = ""
        if edocs["tipo"] not in tipos:
            aviso = f" O tipo {edocs['tipo']} não aparece para você neste setor: escolha outro."
            tipos = tipos + [edocs["tipo"]]
        self.sel_tipo.set_options(tipos, value=edocs["tipo"])
        self.status_opcoes.text = (f"Opções do seu e-docs: {len(setores)} setor(es), {len(opcoes['tipos'])} "
                                   f"tipo(s) e {len(grupos)} grupo(s) de destinatários." + aviso
                                   + ("" if grupos else " Os grupos aparecem quando você tiver algum documento "
                                                        "na Caixa de Saída."))

    def _trocar_grupo(self, e) -> None:
        if e.value is not None:
            self.config["edocs"]["grupo_destinatario"] = e.value

    async def _trocar_setor(self, e) -> None:
        """Outro setor: os tipos podem ser outros, então busca de novo."""
        if getattr(self, "_ajustando_setor", False) or not e.value:
            return
        edocs = self.config["edocs"]
        if int(e.value) == int(edocs.get("setor_posicao") or 1):
            return
        edocs["setor_posicao"] = int(e.value)
        await self._carregar_opcoes_edocs()

    def _simulacao_diario(self) -> None:
        diario = self.config["diario"]
        with gui.card().classes("w-full fundo-cartao q-pa-md gap-2"):
            with gui.row().classes("w-full items-center"):
                self._titulo_da_tela("newspaper", "Diário Oficial · Cadastrar publicação")
                gui.space()
                gui.label("Os campos com borda azul você pode mudar aqui.").classes("text-caption t-suave")
                self._info("sites")
            with gui.row().classes("w-full gap-3 no-wrap"):
                self.sim_titulo_diario = self._campo_fixo("Título")
                self._campo_fixo("Data de veiculação", atos.proximo_dia_util(date.today()).strftime("%d/%m/%Y"),
                                 "w-44").tooltip("O dia útil seguinte ao envio no e-docs")
            with gui.row().classes("w-full gap-3 no-wrap"):
                # Listas iguais às da tela "Incluir" do Diário. Um valor antigo que não está na
                # lista (escrito à mão) continua aparecendo, para não sumir sem a pessoa ver
                for chave, rotulo in (("secretaria", "Orgão/Secretaria"), ("categoria", "Categoria"),
                                      ("coluna", "Colunas")):
                    opcoes = list(diario_http.OPCOES[chave])
                    atual = next((o for o in opcoes if o.upper() == str(diario.get(chave) or "").upper()),
                                 diario.get(chave) or None)
                    if atual and atual not in opcoes:
                        opcoes.append(atual)
                    diario[chave] = atual or ""
                    gui.select(opcoes, value=atual, label=rotulo,
                               on_change=lambda e, c=chave: diario.update({c: e.value or ""})) \
                        .props('outlined dense stack-label options-dense placeholder="Selecione"') \
                        .classes("col campo-editavel")
            gui.label("Descrição (o texto, sem o título: ele vai no campo Título)") \
                .classes("text-caption text-weight-bold text-primary")
            self.sim_texto_diario = gui.html("").classes("w-full caixa-editor")
            gui.label("Anexo: nenhum").classes("text-caption t-suave")

    def _mostrar_pessoas_simuladas(self) -> None:
        """Quem vista, quem assina e o destinatário, como entram na tela do e-docs."""
        if not hasattr(self, "sim_pessoas"):
            return
        pessoas = self.config["pessoas"]
        fixas = pessoas.get("fixas") and pessoas.get("assinantes")
        if fixas:
            vistos = [f"{i}º {p['nome']}" for i, p in enumerate(pessoas.get("vistos", []), 1)] or ["ninguém"]
        else:
            vistos = ["quem você escolher na hora do envio"]
        self.sim_pessoas["vistos"].text = "\n".join(vistos)
        self.sim_pessoas["assinantes"].text = (
            "\n".join(f"{i}º {p['nome']}" for i, p in enumerate(pessoas["assinantes"], 1)) if fixas
            else "quem você escolher na hora do envio")
        self.sim_pessoas["anexos"].text = "nenhum"
        for rotulo in self.sim_pessoas.values():
            rotulo.style("white-space: pre-line")

    async def _atualizar_simulacao(self, registro: dict) -> None:
        if not hasattr(self, "sim_identificacao"):
            return
        proc = self._processo_rascunho()
        for campo, funcao in ((self.sim_identificacao, proc.titulo_edocs),
                              (self.sim_titulo_diario, proc.titulo_diario)):
            try:
                campo.value = funcao(registro)
            except RuntimeError as e:
                campo.value = f"⚠ {e}"
        try:
            texto = (await run.io_bound(proc.folhas, [registro], False))[0]
            sem_titulo = (await run.io_bound(proc.folhas, [registro], False, True))[0]
        except Exception as e:
            texto = sem_titulo = f"<p>Não consegui montar o texto: {html.escape(type(e).__name__)}</p>"
        self.sim_texto_edocs.content = texto
        self.sim_texto_diario.content = sem_titulo

    def _processo_rascunho(self) -> processos.Processo:
        """Processo com o que já foi escolhido (caminhos absolutos), para as prévias."""
        config = dict(self.config, modelo=str(self.modelo or ""), planilha=str(self.planilha or ""))
        return processos.Processo(self.pasta or main.PROCESSOS / (self.nome or "_novo"), config)

    def _desenhar_previa(self) -> None:
        if not hasattr(self, "sel_previa"):
            return
        if hasattr(self, "editor_colunas"):
            self.editor_colunas.desenhar()  # os campos do Word podem ter mudado
        proc = self._processo_rascunho()
        lista = proc.agrupar(self.registros)
        self.linha_previa = min(self.linha_previa, max(len(lista) - 1, 0))
        self.sel_previa.set_options({i: proc.descricao(r) for i, r in enumerate(lista)},
                                    value=self.linha_previa if lista else None)
        if hasattr(self, "inp_edocs"):
            self.inp_edocs.value = self.config["titulo_edocs"]
            self.inp_diario.value = self.config["titulo_diario"]
            self.inp_no_ato.value = self.config.get("titulo_no_ato", "")
        self._desenhar_titulos()
        self._previa_depois()

    def _trocar_previa(self, e) -> None:
        if e.value is not None and e.value != self.linha_previa:
            self.linha_previa = e.value
            self._previa_depois()
            self._desenhar_titulos()

    def _previa_depois(self, atraso: float = 0.0) -> None:
        """Atualiza a prévia em segundo plano (os elementos que pediram podem ser redesenhados)."""
        async def tarefa() -> None:
            await asyncio.sleep(atraso)
            await self._atualizar_previa()
        background_tasks.create(tarefa())

    async def _atualizar_previa(self) -> None:
        if not (self.modelo and self.registros and hasattr(self, "previa")):
            return
        registro = self._registro_previa()
        try:
            pagina = (await run.io_bound(self._processo_rascunho().folhas, [registro]))[0]
        except Exception as e:
            pagina = f"<p>Não consegui montar a prévia: {html.escape(f'{type(e).__name__} {e}')}</p>"
        self.previa.content = f'<div class="folha-area" style="height:560px">{pagina}</div>'
        await self._atualizar_simulacao(registro)

    def _mudar_titulo(self, chave: str, valor) -> None:
        if valor is not None:
            self.config[chave] = valor
            self._desenhar_titulos()
            self._previa_depois(0.3)  # as telas dos sites mostram o título novo

    def _desenhar_checklist(self) -> None:
        if not hasattr(self, "area_checklist"):
            return
        self.area_checklist.clear()
        sem_ligacao = [r for _, r in self.campos if not self._ligada(self._regra(r))]
        titulo_ok = True
        if self.registros:
            registro = self._registro_previa()
            try:
                self._processo_rascunho().titulo_edocs(registro)
            except RuntimeError:
                titulo_ok = False
        with self.area_checklist:
            criada = self.planilha is not None and str(self.planilha).startswith(str(self._temp))
            qtd_atos = len(self._atos()) if self.registros else 0
            if criada:
                _conferencia(True, "Planilha criada pelo Edodocs (vazia: preencha depois de salvar)")
            else:
                _conferencia(bool(self.registros), f"Planilha com {qtd_atos} ato(s)" if self.registros
                             else "Falta a planilha (passo 3)")
            _conferencia(bool(self.campos), f"Modelo com {len(self.campos)} campo(s)" if self.campos
                         else "Falta o modelo (passo 2)")
            _conferencia(not sem_ligacao, "Todos os campos sabem de onde vêm" if not sem_ligacao
                         else "Campos sem ligação (passo 3): " + ", ".join(sem_ligacao))
            _conferencia(titulo_ok, "Título montado certo" if titulo_ok
                         else "O título está com problema: veja abaixo em Títulos")
            _conferencia(None, "Confira o ato ao lado trocando a pessoa: se ele estiver certo, salve.")

    def _desenhar_titulos(self) -> None:
        self._desenhar_checklist()
        if not hasattr(self, "area_titulos"):
            return
        self.sel_numero.set_options(self.colunas, value=_coluna_igual(self.colunas, self.config["coluna_numero"]))
        self.area_titulos.clear()
        with self.area_titulos:
            if not self.registros:
                gui.label("Escolha a planilha para ver os títulos").classes("t-suave")
                return
            proc = self._processo_rascunho()
            registro = self._registro_previa()
            titulos = [("No e-docs", proc.titulo_edocs)]
            if proc.vai_ao_diario:
                titulos.append(("No Diário Oficial", proc.titulo_diario))
            for rotulo, funcao in titulos:
                try:
                    titulo, cor = funcao(registro), "t-forte"
                except RuntimeError as e:
                    titulo, cor = str(e), "text-negative"
                with gui.column().classes("gap-0"):
                    gui.label(rotulo).classes("text-caption t-suave")
                    gui.label(titulo).classes(f"text-body1 text-weight-medium {cor}")

    # --- Salvar ---
    def _salvar(self) -> None:
        sem_ligacao = [r for _, r in self.campos if not self._ligada(self._regra(r))]
        if sem_ligacao:
            gui.notify("Escolha de onde vêm os campos com bolinha vermelha: " + ", ".join(sem_ligacao),
                       type="warning")
            self.passos.set_value("ligar")
            return
        if self.treino:
            self._fim_do_treino()
            return
        if not self.painel_pessoas.valido():
            gui.notify("Escolha quem assina (ou desligue \"Sempre os mesmos vistos e assinantes\")",
                       type="warning")
            return
        pasta = self.pasta or main.PROCESSOS / self.nome
        pasta.mkdir(parents=True, exist_ok=True)
        # O Word fica dentro da pasta do processo (cópia); a planilha fica onde está, a não ser
        # que tenha vindo pelo navegador (cópia temporária)
        if self.modelo.parent != pasta:
            shutil.copy2(self.modelo, pasta / self.modelo.name)
        outros = []
        for outro in self.config.get("outros_modelos") or []:
            caminho = self._caminho_outro(outro)
            if caminho.is_file() and caminho.parent != pasta:
                shutil.copy2(caminho, pasta / caminho.name)
            outros.append(dict(outro, modelo=caminho.name))
        temporaria = str(self.planilha).startswith(tempfile.gettempdir())
        if temporaria:
            shutil.copy2(self.planilha, pasta / self.planilha.name)
        config = dict(self.config, modelo=self.modelo.name, outros_modelos=outros,
                      planilha=self.planilha.name if temporaria or self.planilha.parent == pasta
                      else str(self.planilha))
        config.pop("_tipo", None)  # só do assistente: no processo, quem diz é o "juntar"
        processos.Processo(pasta, config).salvar()
        if not self.pasta:  # o processo novo foi salvo: o rascunho não serve mais
            self._salvo = True
            apagar_rascunho()
        # A planilha escolhida aqui passa a valer (sem a planilha "pessoal" antiga)
        planilhas = main.ler_config().get("planilhas", {})
        if planilhas.pop(pasta.name, None):
            main.salvar_config(planilhas=planilhas)
        self.dialogo.close()
        gui.notify(f"Processo {pasta.name} salvo", type="positive")
        self.tela.processo_salvo(pasta)

    def _arquivos_proprios(self) -> bool:
        """No treino, a pessoa trocou o exemplo pela planilha E pelo Word dela?"""
        exemplos = self._temp / "treino"
        planilha_propria = bool(self.planilha) and self.planilha != getattr(self, "planilha_exemplo", None)
        if self.ato_pronto is not None:  # o modelo veio de um ato pronto
            word_proprio = self.ato_pronto != self.ato_exemplo
        else:                            # escolheu um modelo com campos
            word_proprio = bool(self.modelo) and exemplos not in self.modelo.parents
        return planilha_propria and word_proprio

    def _fim_do_treino(self) -> None:
        proprios = self._arquivos_proprios()
        with self.tela.camada, gui.dialog() as janela, gui.card().classes("cartao w-[560px] q-pa-lg"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("celebration", color="amber-7", size="32px")
                gui.label("Treino concluído!").classes("text-h6 t-forte")
            if not proprios:
                gui.label("É assim que se cria um processo. Como era um treino, nada foi salvo.").classes("text-body1")
                gui.label("Agora faça o mesmo com a sua planilha e um ato seu: clique em Novo processo.") \
                    .classes("text-body2 t-suave")
                with gui.row().classes("w-full justify-end"):
                    gui.button("Fechar", icon="check",
                               on_click=lambda: (janela.close(), self.dialogo.close())).props("unelevated no-caps")
            else:
                # Treinou com a planilha e o Word dela: dá para aproveitar tudo e salvar de verdade
                gui.label("Você treinou com a sua planilha e o seu Word. Quer manter o que escolheu e salvar "
                          "como um processo de verdade? Assim não precisa começar de novo.").classes("text-body1")
                nome = gui.input("Nome do processo", placeholder="Ex.: Licença Prêmio") \
                    .props("outlined stack-label autofocus").classes("w-full")

                def salvar_de_verdade() -> None:
                    escolhido = (nome.value or "").strip()
                    if not escolhido:
                        gui.notify("Escreva um nome para o processo (ex.: Licença Prêmio)", type="warning")
                        return
                    if INVALIDOS.search(escolhido):
                        gui.notify("O nome não pode ter os caracteres < > : \" / \\ | ? *", type="warning")
                        return
                    if (main.PROCESSOS / escolhido).exists():
                        gui.notify("Já existe um processo com esse nome", type="warning")
                        return
                    self.treino = False
                    self.nome = escolhido
                    janela.close()
                    self._salvar()

                nome.on("keydown.enter", salvar_de_verdade)
                with gui.row().classes("w-full justify-end gap-2"):
                    gui.button("Não, só era treino", on_click=lambda: (janela.close(), self.dialogo.close())) \
                        .props("flat no-caps")
                    gui.button("Manter e salvar o processo", icon="save", color="positive",
                               on_click=salvar_de_verdade).props("unelevated no-caps")
        janela.open()

    # --- Abertura (edição) e escolha de arquivos ---
    async def _carregar_existente(self) -> None:
        if self.pasta:
            self.config.setdefault("_tipo", "varias" if (self.config.get("juntar") or {}).get("coluna") else "uma")
            self._mostrar_tipo()
        if self._retomado:
            self._mostrar_tipo()
            self._mostrar_modo()
        if self.planilha:
            await self._ler_planilha()
        if self.modelo:
            await self._usar_modelo(self.modelo, deduzir_titulos=False)
        elif self._retomado and self.word:  # parou no meio da marcação: as marcas voltam
            self._mostrar_marcador()
            self.marcador.marcas = [marcar_ui.Marca(**dict(m, ocorrencias=[tuple(o) for o in m["ocorrencias"]]))
                                    for m in self._marcas_rascunho]
            self.marcador.desenhar()
            self._contar_marcas()
        else:
            self.area_ato.clear()
        if self._retomado:
            await self._voltar_ao_passo()

    # --- Rascunho: o processo novo fechado antes de salvar continua de onde parou --------
    @classmethod
    def novo(cls, tela) -> None:
        """Novo processo. Se um ficou pela metade, pergunta se continua ou começa do zero."""
        dados = ler_rascunho()
        if not dados:
            cls(tela)
            return
        with tela.camada, gui.dialog() as janela, gui.card().classes("cartao w-[540px] max-w-full q-pa-lg gap-3"):
            with gui.row().classes("items-center gap-2 no-wrap"):
                gui.icon("restore", color="primary", size="30px")
                gui.label("Continuar de onde parou?").classes("text-h6 t-forte")
            nome = dados.get("nome") or "sem nome"
            gui.label(f"Você começou a criar o processo “{nome}” e fechou antes de salvar.").classes("text-body1")
            gui.label(f"Parou em: {NOMES_PASSOS.get(dados.get('passo'), 'Antes de começar')}") \
                .classes("text-body2 t-suave")

            def do_zero() -> None:
                janela.close()
                apagar_rascunho()
                cls(tela)

            def continuar() -> None:
                janela.close()
                cls(tela, rascunho=dados)

            with gui.row().classes("w-full justify-end gap-2"):
                gui.button("Começar do zero", icon="restart_alt", on_click=do_zero).props("flat no-caps")
                gui.button("Continuar", icon="play_arrow", on_click=continuar).props("unelevated no-caps")
        janela.open()

    def _retomar(self, dados: dict) -> None:
        """Volta ao processo guardado. As cópias dos Words (e da planilha criada aqui) vão para
        a pasta temporária: o rascunho continua sendo só cópia."""
        pasta = self._temp / "rascunho"
        caminhos = {}
        # Montado pela IA: os arquivos estão todos numa pasta só (processo_ia.rascunho_do_resultado)
        fora = Path(dados["_pasta"]) if dados.get("_pasta") else None
        for chave, nome in (dados.get("arquivos") or {}).items():
            origem = fora / nome if fora else RASCUNHO / chave / nome
            if origem.is_file():
                (pasta / chave).mkdir(parents=True, exist_ok=True)
                caminhos[chave] = Path(shutil.copy2(origem, pasta / chave / nome))
        self.nome = dados.get("nome") or ""
        self.config.update(dados.get("config") or {})
        self.config["outros_modelos"] = [dict(o, modelo=str(caminhos[f"outro{i}"]))
                                         for i, o in enumerate(self.config.get("outros_modelos") or [])
                                         if f"outro{i}" in caminhos]
        self.modelo = caminhos.get("modelo")
        self.word = caminhos.get("word")
        planilha = caminhos.get("planilha") or (Path(dados["planilha"]) if dados.get("planilha") else None)
        self.planilha = planilha if planilha and planilha.is_file() else None
        self.modo_planilha = dados.get("modo_planilha") or ("tenho" if self.planilha else None)
        self.exemplo = dados.get("exemplo")
        self._marcas_rascunho = dados.get("marcas") or []
        self._passo_rascunho = dados.get("passo")
        self._ia = dados.get("ia")
        self._pessoas_exemplo = dados.get("pessoas_exemplo") or []

    async def _voltar_ao_passo(self) -> None:
        criada = self.planilha is not None and self._temp in self.planilha.parents
        if self.modo_planilha == "criar":
            self._desenhar_criar()
        if criada and not self.registros and self._pessoas_exemplo:  # a IA leu os exemplos
            self.registros = [dict(p) for p in self._pessoas_exemplo]
            self._desenhar_ligacoes()
            self._desenhar_previa()
        elif criada and not self.registros and self.exemplo:  # planilha criada aqui, ainda vazia
            self.registros = [dict(self.exemplo, NODOATO="1", _LINHA=2)]
            self._desenhar_ligacoes()
            self._desenhar_previa()
        passo = self._passo_rascunho if self._passo_rascunho in NOMES_PASSOS else "tipo"
        if passo in ("ligar", "previa") and not (self.modelo and self.campos):
            passo = "ato"
        if passo == "previa":
            self._desenhar_ligacoes()
            self.passos.set_value("ligar")
            self._ir_previa()
        else:
            if passo == "ligar":
                self._desenhar_ligacoes()
            self.passos.set_value(passo)
        with self.dialogo:  # vem de uma tarefa em segundo plano: precisa do lugar na tela
            if self._ia is not None:
                self._contar_o_que_a_ia_fez()
            else:
                gui.notify("Continuando de onde você parou", type="info")

    def _guardar_rascunho(self) -> None:
        """Guarda o processo novo como está, com cópias dos Words e da planilha criada aqui.
        Só grava quando algo mudou."""
        if self.pasta or self.treino or self._salvo:
            return
        if not (self.nome or self.config.get("_tipo") or self.word or self.modelo):
            return  # ainda não fez nada
        marcas = [asdict(m) for m in self.marcador.marcas] if self.marcador and not self.modelo else []
        planilha_aqui = self.planilha is not None and self._temp in self.planilha.parents
        arquivos = {"modelo": self.modelo, "word": self.word, "planilha": self.planilha if planilha_aqui else None}
        for i, outro in enumerate(self.config.get("outros_modelos") or []):
            arquivos[f"outro{i}"] = self._caminho_outro(outro)
        arquivos = {chave: caminho for chave, caminho in arquivos.items() if caminho and caminho.is_file()}
        estado = {"nome": self.nome, "config": self.config, "modo_planilha": self.modo_planilha,
                  "exemplo": self.exemplo, "passo": self.passos.value, "marcas": marcas,
                  "planilha": str(self.planilha or ""),
                  "arquivos": {chave: caminho.name for chave, caminho in arquivos.items()}}
        try:
            texto = json.dumps(estado, ensure_ascii=False, indent=2, default=str)
            assinatura = texto + repr([(str(c), c.stat().st_mtime_ns) for c in arquivos.values()])
            if assinatura == self._assinatura_rascunho:
                return
            RASCUNHO.mkdir(parents=True, exist_ok=True)
            for pasta in RASCUNHO.iterdir():  # arquivos que saíram do processo
                if pasta.is_dir() and pasta.name not in arquivos:
                    shutil.rmtree(pasta, ignore_errors=True)
            for chave, caminho in arquivos.items():
                destino = RASCUNHO / chave
                destino.mkdir(exist_ok=True)
                for velho in destino.iterdir():
                    if velho.name != caminho.name:
                        velho.unlink()
                shutil.copy2(caminho, destino / caminho.name)
            (RASCUNHO / "rascunho.json").write_text(texto, encoding="utf-8")
            self._assinatura_rascunho = assinatura
        except (OSError, TypeError, ValueError):
            pass  # Word aberto, sem permissão...: tenta de novo daqui a pouco

    def _ao_fechar(self) -> None:
        self._guardar_rascunho()
        self._timer_rascunho.cancel()
        if not (self._salvo or self.treino) and ler_rascunho():
            gui.notify("Guardei o que você fez. Para continuar, clique em Novo processo.", type="info",
                       timeout=6000)

    async def _escolher_arquivo(self, rotulo: str, extensao: str) -> Path | None:
        """Janela de abrir arquivo do Windows; no navegador, envio do arquivo (fica uma cópia)."""
        if napp.native.main_window is not None:
            import webview
            tipo = webview.FileDialog.OPEN if hasattr(webview, "FileDialog") else webview.OPEN_DIALOG
            escolhidos = await napp.native.main_window.create_file_dialog(
                tipo, file_types=(f"{rotulo} (*{extensao})",))
            return Path(escolhidos[0]) if escolhidos else None

        futuro = asyncio.get_running_loop().create_future()
        with self.tela.camada, gui.dialog() as janela, gui.card().classes("cartao q-pa-lg"):
            gui.label(f"Envie o arquivo ({extensao})").classes("text-h6")

            async def recebido(e) -> None:
                destino = Path(tempfile.mkdtemp()) / e.file.name
                await e.file.save(destino)
                janela.close()
                if not futuro.done():
                    futuro.set_result(destino)

            gui.upload(on_upload=recebido, auto_upload=True, max_files=1).props(f"accept={extensao}")
            gui.button("Cancelar", on_click=lambda: (janela.close(), futuro.done() or futuro.set_result(None))) \
                .props("flat")
        janela.open()
        return await futuro
