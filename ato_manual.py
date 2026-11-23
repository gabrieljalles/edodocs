"""Ato manual: um ato que não tem processo nem planilha (cada um é diferente), escrito à mão
no próprio Edodocs e enviado ao e-docs, sem precisar abrir o site.

Cada ato fica guardado em entradas/atos_manuais/<id>.json (rascunho, enviado ou com erro),
para a pessoa ver depois na aba "Meus atos manuais". O envio segue o caminho da Execução:
login, vistos e assinantes (perguntados no fim, como na execução padrão), os anexos escolhidos
na tela do ato, conferência do ato e envio.
"""
import html as html_lib
import json
import re
import time
import uuid
from datetime import date, datetime
from pathlib import Path

import lxml.html

import main
import ui

PASTA = main.ENTRADAS / "atos_manuais"
# Chave das escolhas de vistos/assinantes (como o nome do processo na execução padrão)
CHAVE_PESSOAS = "Ato manual"
RASCUNHO, ENVIADO, ERRO = "rascunho", "enviado", "erro"
MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto",
         "Setembro", "Outubro", "Novembro", "Dezembro"]
# Blocos que o editor do e-docs entende como parágrafo/estrutura (o resto fica dentro de <p>)
BLOCOS = {"p", "div", "ul", "ol", "table", "hr", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "pre"}
# As duas telas do e-docs (cada conta tem uma, a outra ou as duas): a pessoa escolhe antes
INSTRUMENTO, GENERICO = "instrumento", "generico"
MODELOS = {INSTRUMENTO: "Instrumentos Jurídicos/Normativos", GENERICO: "Documentos Genéricos"}
# O texto que a tela de Documentos Genéricos já traz (capturas: caixa_entrada.har)
TEXTO_GENERICO = ("<p>&nbsp;</p><p>&nbsp;</p><p><strong>Assunto: </strong></p><p>&nbsp;</p><p>&nbsp;</p>"
                  "<p>&nbsp;</p><p>&nbsp;</p><p>Atenciosamente, </p><p>&nbsp;</p><p>&nbsp;</p>")


def texto_inicial(dia: date | None = None, modelo: str = INSTRUMENTO) -> str:
    """O texto de um ato novo: vazio (a pessoa escreve tudo). O que o e-docs traz (a data por
    extenso nos Instrumentos, TEXTO_GENERICO nos Genéricos) não é posto: atrapalhava."""
    return ""


def novo(modelo: str = "") -> dict:
    """Ato em branco. modelo vazio = a pessoa ainda não escolheu a tela (Instrumento ou Genérico)."""
    hoje = date.today()
    return {"id": "", "criado": "", "alterado": "", "modelo": modelo, "setor_posicao": 0, "setor": "",
            "tipo": "", "data": hoje.strftime("%d/%m/%Y"), "titulo": "",
            "texto": texto_inicial(hoje, modelo), "anexos": [], "grupo": "", "iniciais": "",
            "incluir_visto": False, "situacao": RASCUNHO, "doc": "", "enviado_em": "", "criado_por": "",
            "erro": "", "vistos": [], "assinantes": []}


def nome_do_modelo(ato: dict) -> str:
    return MODELOS.get(ato.get("modelo") or INSTRUMENTO, "")


def pede_vistos(ato: dict) -> bool:
    """Instrumentos sempre têm vistos; Documentos Genéricos só com "Incluir Visto?" = Sim."""
    return ato.get("modelo") != GENERICO or bool(ato.get("incluir_visto"))


# --- Guardar e listar -------------------------------------------------------------

def _arquivo(ato_id: str) -> Path:
    return PASTA / f"{ato_id}.json"


def salvar(ato: dict) -> dict:
    """Grava o ato (cria o id na primeira vez). Devolve o próprio ato."""
    PASTA.mkdir(parents=True, exist_ok=True)
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    if not ato.get("id"):
        ato["id"] = datetime.now().strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:6]
        ato["criado"] = agora
    ato["alterado"] = agora
    temporario = _arquivo(ato["id"]).with_suffix(".tmp")
    temporario.write_text(json.dumps(ato, ensure_ascii=False, indent=2), encoding="utf-8")
    temporario.replace(_arquivo(ato["id"]))
    return ato


def ler(ato_id: str) -> dict | None:
    try:
        dados = json.loads(_arquivo(ato_id).read_text(encoding="utf-8"))
        dados.setdefault("modelo", INSTRUMENTO)  # os de antes da escolha eram todos Instrumentos
        return novo() | dados
    except (OSError, ValueError):
        return None


def listar() -> list[dict]:
    """Todos os atos manuais, o mais recente primeiro."""
    atos = []
    for arquivo in PASTA.glob("*.json") if PASTA.is_dir() else []:
        if ato := ler(arquivo.stem):
            atos.append(ato)
    return sorted(atos, key=lambda a: a["id"], reverse=True)


def excluir(ato_id: str) -> None:
    """Apaga só o registro no Edodocs (um ato já enviado continua no e-docs)."""
    _arquivo(ato_id).unlink(missing_ok=True)


def copia(ato: dict) -> dict:
    """Um ato novo com o mesmo conteúdo (para aproveitar um já enviado)."""
    base = novo(ato.get("modelo") or INSTRUMENTO)
    for chave in ("setor_posicao", "setor", "tipo", "titulo", "texto", "grupo", "iniciais", "incluir_visto"):
        base[chave] = ato.get(chave, base[chave])
    return base


# --- Texto -----------------------------------------------------------------------

def texto_puro(html: str) -> str:
    """O texto sem as marcações (para conferir se tem algo escrito e para a lista)."""
    html = re.sub(r"(?i)<br\s*/?>|</?(div|p|li|h\d)\b[^>]*>", "\n", html or "")
    texto = html_lib.unescape(re.sub(r"<[^>]+>", "", html))
    return "\n".join(linha.strip() for linha in texto.splitlines() if linha.strip())


# Tags sem fechamento no HTML: no e-docs vão fechadas (<br />), como o editor do site grava
VAZIAS = re.compile(r"<(br|col|hr|img)\b([^>]*?)\s*/?>", re.I)


def _como_o_site(html: str) -> str:
    """O HTML no jeito do editor do e-docs (TinyMCE): tags vazias fechadas (<br />, <col />),
    tabela com border-collapse e todo caractere fora do ASCII como entidade (ç → &ccedil;).
    Com acento cru, o gerador de PDF do e-docs diz que o arquivo está corrompido (atos._entidades)."""
    from html.entities import codepoint2name

    html = VAZIAS.sub(lambda m: f"<{m.group(1).lower()}{m.group(2).rstrip()} />", html)

    def tabela(m: re.Match) -> str:
        tag = m.group(0)
        if "border-collapse" in tag:
            return tag
        if re.search(r'style="', tag):
            return re.sub(r'style="', 'style="border-collapse: collapse; ', tag, count=1)
        return tag[:-1].rstrip() + ' style="border-collapse: collapse;">'

    html = re.sub(r"<table\b[^>]*>", tabela, html, flags=re.I)
    return "".join(c if ord(c) < 128 else
                   f"&{codepoint2name[ord(c)]};" if ord(c) in codepoint2name else f"&#{ord(c)};"
                   for c in html)


def html_para_edocs(html: str) -> str:
    """O HTML do editor (texto solto e <div> por linha) em parágrafos <p>, como o editor
    do e-docs grava. Alinhamento, negrito, listas etc. ficam como estão. Sai no jeito do
    site (_como_o_site), para o PDF do e-docs não acusar arquivo corrompido."""
    return _como_o_site(_em_paragrafos(html))


def _em_paragrafos(html: str) -> str:
    try:
        partes = lxml.html.fragments_fromstring(html or "")
    except Exception:
        return html or ""
    saida, solto = [], []

    def fechar() -> None:
        if "".join(solto).strip():
            saida.append(f"<p>{''.join(solto)}</p>")
        solto.clear()

    for parte in partes:
        if isinstance(parte, str):
            solto.append(html_lib.escape(parte))
            continue
        if parte.tag in BLOCOS:
            fechar()
            if parte.tag == "div":
                parte.tag = "p"
            saida.append(lxml.html.tostring(parte, encoding="unicode", with_tail=False))
            if parte.tail:
                solto.append(html_lib.escape(parte.tail))
        else:
            solto.append(lxml.html.tostring(parte, encoding="unicode"))
    fechar()
    return "".join(saida)


def folha(ato: dict) -> str:
    """O ato para a conferência antes de enviar, como fica no editor. Tabela com border="1"
    mostra as bordas finas em todas as células, como no PDF do e-docs."""
    return ('<style>.folha-texto table[border="1"] td, .folha-texto table[border="1"] th '
            '{ border: 1px solid #000; }</style>'
            f'<div class="folha-texto" style="font-family:Arial;font-size:12pt;padding:24px 32px;min-height:440px">'
            f'{html_para_edocs(ato.get("texto", ""))}</div>')


def problemas(ato: dict) -> list[str]:
    """O que falta para enviar (vazio = pode enviar)."""
    faltam = []
    if not str(ato.get("titulo", "")).strip():
        faltam.append("a Nossa Identificação (o título do ato no e-docs)")
    if not ato.get("tipo"):
        faltam.append("o Tipo")
    if not re.fullmatch(r"\d{2}/\d{2}/\d{4}", str(ato.get("data", ""))):
        faltam.append("a Data (dia/mês/ano)")
    if not texto_puro(ato.get("texto", "")):
        faltam.append("o Texto do documento")
    return faltam


# --- Envio (roda na Execução, fora da tela) --------------------------------------------

def _pessoas(usuario: str, pesquisar, pedir_vistos: bool = True) -> tuple[list[dict], list[dict]]:
    """Vistos e assinantes, perguntados como na execução padrão. Quem envia, se vista, vista em 1º.
    pedir_vistos=False: só assinantes (Documentos Genéricos sem visto)."""
    from escolha_pessoas import sem_repetir

    vistos = sem_repetir(main.escolher_vistos(usuario, pesquisar, CHAVE_PESSOAS)) if pedir_vistos else []
    assinantes = sem_repetir(main.escolher_assinantes(usuario, pesquisar, CHAVE_PESSOAS))
    meu_cpf = main._cpf(usuario)
    vistos = [p for p in vistos if p["cpf"] == meu_cpf] + [p for p in vistos if p["cpf"] != meu_cpf]
    main.mostrar_pessoas(vistos, assinantes)
    return vistos, assinantes


def anexos_sumidos(ato: dict) -> list[str]:
    """Anexos escolhidos na tela que não estão mais no computador (movidos ou apagados)."""
    return [Path(a).name for a in ato.get("anexos", []) if not Path(a).is_file()]


def _marcar(ato: dict, **campos) -> None:
    ato.update(campos)
    salvar(ato)


def _pode_enviar(ato: dict | None, ato_id: str) -> str:
    """Por que o ato não pode ir (vazio = pode). Só pergunta no caso de um envio que caiu no meio."""
    if not ato:
        return f"ato manual {ato_id} não encontrado (foi apagado?)"
    nome = ato["titulo"] or "(sem Nossa Identificação)"
    if ato["situacao"] == ENVIADO:
        return f"{nome}: já foi enviado (documento {ato['doc']}). Para mandar de novo, use “Fazer uma cópia”"
    if faltam := problemas(ato):
        return f"{nome}: falta preencher " + "; ".join(faltam)
    # Os anexos são só os escolhidos na tela do ato (Doc. Anexo(s)): nada é perguntado no fim
    if sumidos := anexos_sumidos(ato):
        return (f"{nome}: anexo(s) não encontrado(s) no computador: " + ", ".join(sumidos)
                + ". Tire ou escolha de novo em Doc. Anexo(s)")
    if ato["situacao"] == ERRO and "envio sem confirmação" in ato.get("erro", "") and not ui.confirmar(
            f"{nome}: da última vez o envio do documento {ato['doc']} caiu no meio e pode ter ido. Confira no "
            "e-docs antes. Mandar de novo pode duplicar o ato. Criar outro documento e enviar mesmo assim?",
            sim="Sim, já conferi: não foi", nao="Não, deixar este de fora"):
        return f"{nome}: deixado de fora (confira no e-docs se o documento {ato['doc']} foi)"
    return ""


def enviar(ato_id: str) -> None:
    """Um ato manual só (botão Enviar ao e-docs da tela do ato)."""
    enviar_varios([ato_id])


def enviar_varios(ids: list[str]) -> None:
    """Envia os atos manuais escolhidos. Roda na Execução (as perguntas viram caixas). Entra
    uma vez e pergunta quem vista e quem assina uma vez só: valem para todos os atos."""
    import credenciais
    import edocs_http

    lista = []
    for ato_id in ids:
        ato = ler(ato_id)
        if motivo := _pode_enviar(ato, ato_id):
            ui.erro(f"Fora do envio: {motivo}")
        else:
            lista.append(ato)
    if not lista:
        return

    cfg = main.ler_config()
    edocs_http.PAUSA_REQUISICAO = float(cfg.get("pausa_requisicao", edocs_http.PAUSA_REQUISICAO))
    pausa_entre_atos = float(cfg.get("pausa_entre_atos", main.PAUSA_ENTRE_ATOS))

    usuario, senha = main.pedir_credenciais()
    ui.info(f"Usuário: {usuario}")
    with ui.processando("Entrando no e-docs"):
        try:
            sessao = edocs_http.entrar(usuario, senha)
        except edocs_http.LoginRecusado:
            credenciais.remover(main.CREDENCIAIS, "edocs")
            ui.erro("Login recusado. Usuário do e-docs apagado, tente de novo")
            return
        except Exception as e:
            ui.erro(f"Falha ao entrar no e-docs: {type(e).__name__} {e}".strip())
            return
    ui.ok("Logado")

    # A última etapa antes de enviar: quem vista e quem assina (os mesmos para todos os atos)
    criado_por = main.nome_da_conta(sessao, usuario)  # quem envia (primeiro nome da conta)
    # Vistos só se algum ato tem (os Documentos Genéricos sem "Incluir Visto?" não têm)
    vistos, assinantes = _pessoas(usuario, lambda termo: edocs_http.pesquisar_servidores(sessao, termo),
                                  pedir_vistos=any(pede_vistos(a) for a in lista))
    cpfs_vistos, cpfs_assinantes = [p["cpf"] for p in vistos], [p["cpf"] for p in assinantes]
    if not main.resolver_em_montagem(sessao):
        return

    confirmar = True
    unico = len(lista) == 1
    enviados = pulados = 0
    docs_criados: set[str] = set()  # o e-docs devolver o mesmo número = o anterior ficou em montagem
    try:
        for n, ato in enumerate(lista, 1):
            ui.progresso(n, len(lista), ato["titulo"])
            if n > 1 and pausa_entre_atos:
                time.sleep(pausa_entre_atos)
            vistos_do_ato = vistos if pede_vistos(ato) else []
            _marcar(ato, vistos=main.guardar_pessoas(vistos_do_ato), assinantes=main.guardar_pessoas(assinantes))
            modelo = ato.get("modelo") or INSTRUMENTO
            # Setor, tipo e grupo são de cada ato
            edocs_http.TIPO_ATO = ato["tipo"]
            edocs_http.SETOR_POSICAO = int(ato.get("setor_posicao") or 1)
            edocs_http.GRUPO_DESTINATARIO = ato.get("grupo", "")
            doc = None
            try:
                with ui.processando(f"Gravando {ato['titulo']} no e-docs"):
                    doc = edocs_http.preparar_documento(
                        sessao, ato["titulo"], html_para_edocs(ato["texto"]),
                        cpfs_vistos if vistos_do_ato else [], cpfs_assinantes,
                        avisar=ui.ok, data=ato["data"], modelo=modelo,
                        extras={"iniciais": ato.get("iniciais", ""), "incluir_visto": pede_vistos(ato)})
                for arquivo in [Path(a) for a in ato.get("anexos", [])]:
                    with ui.processando(f"Anexando {arquivo.name}"):
                        edocs_http.anexar_arquivo(sessao, doc, arquivo, avisar=ui.ok)
            except edocs_http.SessaoExpirada:
                ui.erro("A sessão caiu: faça login de novo")
                return
            except edocs_http.ErroEdocs as e:
                if doc:  # o anexo não entrou: o rascunho em montagem travaria os próximos atos
                    e = f"{e}. {edocs_http.descartar_rascunho(sessao, doc)}"
                pulados += 1
                _marcar(ato, situacao=ERRO, erro=str(e)[:400])
                ui.erro(f"Pulado {ato['titulo']}: {e}")
                continue
            if doc in docs_criados:
                _marcar(ato, situacao=ERRO, erro=f"documento {doc} reutilizado: confira no e-docs")
                ui.erro(f"Parado em {ato['titulo']}: o documento {doc} foi reutilizado "
                        "(o ato anterior deve ter ficado em montagem)")
                break
            docs_criados.add(doc)

            if confirmar:
                escolha = ui.confirmar_envio(ato["titulo"], texto_puro(ato["texto"]), folha(ato), unico=unico)
                if escolha in ("cancelar", "pular"):
                    with ui.processando(f"Cancelando o rascunho {doc}"):
                        resultado = edocs_http.descartar_rascunho(sessao, doc)
                    ui.aviso(f"Não enviado: {ato['titulo']}. {resultado}. O ato continua salvo aqui como rascunho")
                    if escolha == "cancelar":
                        break
                    pulados += 1
                    continue
                if escolha == "todos":
                    confirmar = False
                    ui.info("Os próximos atos serão enviados sem perguntar")

            try:
                with ui.processando(f"Enviando {ato['titulo']}"):
                    edocs_http.enviar_documento(sessao, doc, avisar=ui.ok, modelo=modelo)
            except Exception as e:
                # Caiu no meio do envio: não dá para saber se foi. Nunca reenvia sozinho; para aqui.
                _marcar(ato, situacao=ERRO, doc=doc,
                        erro=f"envio sem confirmação ({type(e).__name__} {e}): confira no e-docs"[:400])
                ui.erro(f"Envio de {ato['titulo']} sem confirmação: {type(e).__name__} {e}".strip())
                ui.aviso(f"Confira no e-docs se o documento {doc} foi enviado antes de tentar de novo")
                break
            enviados += 1
            _marcar(ato, situacao=ENVIADO, doc=doc, erro="", enviado_em=time.strftime("%d/%m/%Y %H:%M"),
                    criado_por=criado_por)
            ui.ok(f"Enviado: {ato['titulo']} (documento {doc})")
    finally:
        if len(lista) > 1:
            ui.info(f"Enviados {enviados} · Pulados {pulados} · Total {len(lista)}")
