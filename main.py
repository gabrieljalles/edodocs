"""Os fluxos do Edodocs (lançar no e-docs, lançar no Diário, conferir a publicação),
as configurações e o status na planilha. Quem mostra e pergunta é a janela do app.py."""
import json
import re
import shutil
import sys
import tempfile
import time
from datetime import date, datetime
from pathlib import Path

import atos
import planilha_status
import processos
import trava
import ui


def base_dir() -> Path:
    # No executável do PyInstaller, as pastas ficam ao lado do .exe
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


BASE = base_dir()
ENTRADAS = BASE / "entradas"
SAIDAS = BASE / "saidas"
# Cada pasta aqui dentro é um processo (Licença Prêmio, Compensação...): o Word da mala
# direta e a planilha dele. Para criar um processo novo basta criar a pasta com os dois.
PROCESSOS = ENTRADAS / "processos"
DIARIO = ENTRADAS / "Diario_oficial"
# Só o Diário do dia (aberto pelo app): os de outros dias são apagados ao abrir
DIARIO_HOJE = ENTRADAS / "Diario_hoje"
CREDENCIAIS = BASE / "configs" / "credenciais.json"
# Pausa padrão entre um ato e o seguinte (a entre requisições fica em edocs_http/diario_http)
PAUSA_ENTRE_ATOS = 0.5
CONFIG = BASE / "configs" / "config.json"

# Pasta compartilhada (na rede, escolhida em Configurações): é a pasta dos processos, cada um
# numa pasta com o nome dele. As tabelas de consulta, os de-para e as versões anteriores ficam
# na pasta EXTRAS dentro dela. Senhas, IA, tema... continuam em configs/ de cada computador.
PROCESSOS_DESTE_COMPUTADOR = PROCESSOS
COMPARTILHADAS = ("tabelas_consulta", "de_para", "planilhas")  # chaves do config que vão para a rede
ARQUIVO_COMPARTILHADO = "edodocs - configurações.json"
EXTRAS = "_configurações do Edodocs"
VERSOES = "versões anteriores"  # cópia de cada processo antes de ser mudado
GUARDAR_VERSAO_A_CADA = 600     # segundos: uma cópia por sessão de edição, não uma por clique
VERSOES_GUARDADAS = 30

# Pasta do processo em uso (escolhido ao abrir o programa ou em "Trocar processo")
processo: Path | None = None


def preparar_pastas() -> None:
    for p in (PROCESSOS, DIARIO, DIARIO_HOJE, SAIDAS):
        p.mkdir(parents=True, exist_ok=True)


def _ler_json(arquivo: Path) -> dict:
    try:
        return json.loads(arquivo.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def pasta_compartilhada() -> Path | None:
    """A pasta da rede escolhida em Configurações (None: tudo neste computador)."""
    pasta = _ler_json(CONFIG).get("pasta_compartilhada")
    return Path(pasta) if pasta else None


def processos_da(raiz: Path) -> Path:
    """Onde ficam os processos na pasta escolhida: nela mesma ou, se a pessoa escolheu a pasta
    de cima, na pasta "processos" de dentro."""
    return raiz / "processos" if (raiz / "processos").is_dir() else raiz


def extras_da(raiz: Path) -> Path:
    """Onde ficam as tabelas de consulta, os de-para e as versões anteriores (sempre junto dos
    processos: quem escolheu a pasta de cima vê o mesmo que quem escolheu a dos processos)."""
    return processos_da(raiz) / EXTRAS


def extras_compartilhados() -> Path | None:
    raiz = pasta_compartilhada()
    return extras_da(raiz) if raiz else None


def aplicar_pasta_compartilhada() -> None:
    """Processos da pasta compartilhada (ou deste computador, sem ela)."""
    global PROCESSOS
    raiz = pasta_compartilhada()
    PROCESSOS = processos_da(raiz) if raiz else PROCESSOS_DESTE_COMPUTADOR


def problema_na_compartilhada() -> str:
    """"" se a pasta compartilhada está acessível (ou não é usada); senão, o que dizer."""
    raiz = pasta_compartilhada()
    if raiz and not raiz.is_dir():
        return (f"Não consegui abrir a pasta compartilhada {raiz}. Confira se o computador está na rede "
                "e se você tem acesso a ela.")
    return ""


_cache_compartilhado: dict = {}


def _ler_compartilhado(raiz: Path) -> dict:
    """As configurações da pasta compartilhada (relidas só quando o arquivo muda: a rede é lenta)."""
    arquivo = extras_da(raiz) / ARQUIVO_COMPARTILHADO
    try:
        marca = (str(arquivo), arquivo.stat().st_mtime_ns, arquivo.stat().st_size)
    except OSError:
        return {}
    if _cache_compartilhado.get("marca") != marca:
        _cache_compartilhado.update(marca=marca, dados=_ler_json(arquivo))
    return json.loads(json.dumps(_cache_compartilhado["dados"]))  # cópia: quem lê pode mexer


def ler_config() -> dict:
    cfg = _ler_json(CONFIG)
    if cfg.get("pasta_compartilhada"):
        comum = _ler_compartilhado(Path(cfg["pasta_compartilhada"]))
        for chave in COMPARTILHADAS:
            cfg.pop(chave, None)
            if chave in comum:
                cfg[chave] = comum[chave]
    return cfg


def salvar_config(**valores) -> None:
    local = _ler_json(CONFIG)
    raiz = valores.get("pasta_compartilhada", local.get("pasta_compartilhada"))
    comum = {c: valores.pop(c) for c in list(valores) if c in COMPARTILHADAS} if raiz else {}
    if comum:
        arquivo = extras_da(Path(raiz)) / ARQUIVO_COMPARTILHADO
        arquivo.parent.mkdir(parents=True, exist_ok=True)
        with trava.trava(arquivo):  # relê dentro da trava: não apaga o que outra pessoa salvou
            arquivo.write_text(json.dumps(_ler_json(arquivo) | comum, ensure_ascii=False, indent=2),
                               encoding="utf-8")
    if valores or not comum:
        CONFIG.parent.mkdir(parents=True, exist_ok=True)
        CONFIG.write_text(json.dumps(local | valores, ensure_ascii=False, indent=2), encoding="utf-8")


def guardar_versao(pasta: Path) -> None:
    """Antes de mudar um processo da pasta compartilhada: guarda uma cópia do processo.json e
    dos Word em "versões anteriores/<processo>/<data hora>" (dá para voltar se alguém errar).
    Uma cópia a cada GUARDAR_VERSAO_A_CADA segundos por processo; ficam as VERSOES_GUARDADAS
    mais novas."""
    raiz = pasta_compartilhada()
    try:
        if not raiz or not pasta.is_dir() or pasta.resolve().parent != processos_da(raiz).resolve():
            return
        base = extras_da(raiz) / VERSOES / pasta.name
        antigas = sorted(d for d in base.iterdir() if d.is_dir()) if base.is_dir() else []
        if antigas and time.time() - antigas[-1].stat().st_mtime < GUARDAR_VERSAO_A_CADA:
            return
        destino = base / datetime.now().strftime("%Y-%m-%d %Hh%Mm%Ss")
        destino.mkdir(parents=True, exist_ok=True)
        for arquivo in [pasta / processos.ARQUIVO, *_arquivos(pasta, "*.docx")]:
            if arquivo.is_file():
                shutil.copy2(arquivo, destino / arquivo.name)
        for velha in (antigas + [destino])[:-VERSOES_GUARDADAS]:
            shutil.rmtree(velha, ignore_errors=True)
    except OSError:
        pass  # sem a cópia, a mudança segue (não trava quem está editando)


def levar_para_compartilhada(raiz: Path) -> list[str]:
    """Primeira vez numa pasta compartilhada vazia: copia para lá os processos, as tabelas de
    consulta e os de-para deste computador. Não sobrescreve nada que já exista lá."""
    import consultas
    levados = []
    destino = processos_da(raiz)
    extras = extras_da(raiz)
    extras.mkdir(parents=True, exist_ok=True)
    for pasta in sorted(PROCESSOS_DESTE_COMPUTADOR.iterdir()) if PROCESSOS_DESTE_COMPUTADOR.is_dir() else []:
        if pasta.is_dir() and not (destino / pasta.name).exists() and _arquivos(pasta, "*.docx"):
            shutil.copytree(pasta, destino / pasta.name,
                            ignore=shutil.ignore_patterns("~$*", "~edodocs*"))
            levados.append(pasta.name)
    local = _ler_json(CONFIG)
    comum = _ler_json(extras / ARQUIVO_COMPARTILHADO)
    tabelas = comum.get("tabelas_consulta") or {}
    pasta_tabelas = extras / consultas.NOME_PASTA
    for nome, dados in (local.get("tabelas_consulta") or {}).items():
        origem = Path(dados.get("arquivo") or "")
        if nome in tabelas or not origem.is_file():
            continue
        pasta_tabelas.mkdir(parents=True, exist_ok=True)
        if not (pasta_tabelas / origem.name).exists():
            shutil.copy2(origem, pasta_tabelas / origem.name)
        tabelas[nome] = dict(dados, arquivo=str(pasta_tabelas / origem.name))
        levados.append(f"tabela {nome}")
    de_para = comum.get("de_para") or {}
    for nome, pares in (local.get("de_para") or {}).items():
        if nome not in de_para:
            de_para[nome] = pares
            levados.append(f"de-para {nome}")
    # Planilha escolhida fora da pasta do processo: só vale se os outros também enxergam
    planilhas = comum.get("planilhas") or {}
    for nome, caminho in (local.get("planilhas") or {}).items():
        if nome not in planilhas and str(caminho).startswith("\\\\"):
            planilhas[nome] = caminho
    with trava.trava(extras / ARQUIVO_COMPARTILHADO):
        (extras / ARQUIVO_COMPARTILHADO).write_text(json.dumps(
            _ler_json(extras / ARQUIVO_COMPARTILHADO) | {"tabelas_consulta": tabelas, "de_para": de_para,
                                                        "planilhas": planilhas},
            ensure_ascii=False, indent=2), encoding="utf-8")
    return levados


def _arquivos(pasta: Path, padrao: str) -> list[Path]:
    """Arquivos da pasta, sem os temporários do Office ("~$")."""
    return sorted(p for p in pasta.glob(padrao) if not p.name.startswith("~$"))


def listar_processos() -> list[Path]:
    """Pastas de entradas/processos que têm um modelo Word."""
    if not PROCESSOS.is_dir():
        return []
    return sorted((p for p in PROCESSOS.iterdir()
                   if p.is_dir() and p.name not in (EXTRAS, VERSOES) and _arquivos(p, "*.docx")),
                  key=lambda p: p.name.lower())


NOME_INVALIDO = re.compile(r'[<>:"/\\|?*]')


def problema_no_nome(nome: str, atual: Path | None = None) -> str:
    """Por que o nome não serve para um processo ("" se serve). O nome é a pasta."""
    nome = nome.strip()
    if not nome:
        return "Escreva o nome do processo"
    if NOME_INVALIDO.search(nome) or nome.endswith("."):
        return 'O nome não pode ter < > : " / \\ | ? * nem terminar com ponto'
    if any(p.name.lower() == nome.lower() and p != atual for p in PROCESSOS.iterdir() if p.is_dir()):
        return f"Já existe um processo chamado {nome}"
    return ""


def renomear_processo(pasta: Path, novo: str) -> Path:
    """Muda o nome do processo (a pasta) e o que o config.json guarda pelo nome: o processo em
    uso, a planilha escolhida e os vistos/assinantes lembrados. Caminhos absolutos para dentro
    da pasta no processo.json são corrigidos. Lança ValueError (nome) ou OSError (arquivo
    aberto)."""
    global processo
    novo = novo.strip()
    erro = problema_no_nome(novo, pasta)
    if erro:
        raise ValueError(erro)
    destino = pasta.with_name(novo)
    if destino == pasta:
        return pasta
    pasta.rename(destino)
    arquivo = destino / processos.ARQUIVO
    if arquivo.is_file():
        texto = arquivo.read_text(encoding="utf-8")
        antigo = json.dumps(str(pasta), ensure_ascii=False)[1:-1]  # como fica escrito no JSON
        if antigo in texto:
            arquivo.write_text(texto.replace(antigo, json.dumps(str(destino), ensure_ascii=False)[1:-1]),
                               encoding="utf-8")
    cfg = ler_config()
    mudar = {}
    for chave in ("planilhas", "ultimas_escolhas"):
        por_nome = dict(cfg.get(chave) or {})
        if pasta.name in por_nome:
            por_nome[novo] = por_nome.pop(pasta.name)
            mudar[chave] = por_nome
    if cfg.get("processo") == pasta.name:
        mudar["processo"] = novo
    if mudar:
        salvar_config(**mudar)
    if processo == pasta:
        processo = destino
    return destino


def escolher_processo() -> bool:
    """Define o processo em uso. Com um só, entra direto; com vários, pergunta."""
    global processo
    processos = listar_processos()
    if not processos:
        ui.erro("Nenhum processo em entradas/processos/. "
                "Crie uma pasta com o Word (.docx) e a planilha (.xlsx)")
        return False
    if len(processos) == 1:
        processo = processos[0]
    else:
        processo = processos[ui.escolher("Qual processo?", [p.name for p in processos])[0]]
    ui.titulo(processo.name)
    return True


def perguntar_antes_de_enviar() -> bool:
    """Se o processo em uso mostra cada ato antes de enviar (processo.json "confirmar_envio").
    Processos de antes desta opção seguem o que estava em Configurações."""
    valor = proc().config.get("confirmar_envio")
    return ler_config().get("confirmar_envio", True) if valor is None else bool(valor)


def proc() -> processos.Processo:
    """Configuração do processo em uso (processo.json da pasta dele)."""
    return processos.carregar(processo)


def achar_modelo() -> Path | None:
    modelo = proc().modelo()
    if not modelo:
        ui.erro(f"Nenhum Word (.docx) em entradas/processos/{processo.name}/")
    return modelo


def achar_planilha() -> Path | None:
    """Planilha configurada para o processo (4 → 2) ou a .xlsx da pasta dele."""
    personalizada = ler_config().get("planilhas", {}).get(processo.name)
    if personalizada:
        caminho = Path(personalizada)
        if caminho.is_file():
            return caminho
        ui.erro(f"Planilha configurada não encontrada: {caminho}")
        return None

    planilha = proc().planilha_padrao()
    if not planilha:
        ui.erro(f"Planilha do processo {processo.name} não encontrada "
                "(confira em Configurações ou edite o processo)")
    return planilha


def verificar_diario() -> None:
    """Confere se os atos enviados ao Diário (status_diario = Enviado com sucesso) saíram
    publicados. Cada tentativa olha o Diário de uma data diferente: a data de veiculação
    (data_diario), depois o próximo dia útil, depois o seguinte, até 3 tentativas."""
    import diario

    if not proc().vai_ao_diario:
        ui.aviso(f"O processo {processo.name} não é publicado no Diário Oficial (para mudar: Processos)")
        return
    planilha = achar_planilha()
    if not planilha:
        return

    with ui.processando("Lendo planilha"):
        registros = proc().ler_planilha(planilha)
    status = StatusPlanilha(planilha, "publicado", com_data=False)
    chave_pub = status.chave()

    # Só confere o que já foi enviado ao Diário e ainda não tem resultado final
    enviados = [r for r in registros if r.get("STATUSDIARIO") == planilha_status.ENVIADO]
    candidatos = [r for r in enviados
                  if not planilha_status.conferencia_finalizada(r.get(chave_pub))]
    ja = len(enviados) - len(candidatos)
    if ja:
        ui.info(f"Já conferidos (publicados ou 3 tentativas): {ja}")
    if not candidatos:
        ui.ok("Nenhum ato para conferir" if enviados else "Nenhum ato enviado ao Diário ainda")
        return

    def datas_das_tentativas(r) -> list:
        """Data do Diário de cada tentativa: veiculação, +1 dia útil, +2 dias úteis."""
        veic = atos.ler_data(r.get("DATADIARIO", ""))
        if not veic:
            return []
        datas = [veic]
        while len(datas) < 3:
            datas.append(atos.proximo_dia_util(datas[-1]))
        return datas

    # Baixa do site os Diários das datas que serão conferidas agora (só datas já passadas)
    necessarias = set()
    for r in candidatos:
        datas = datas_das_tentativas(r)
        feitas = planilha_status.tentativas_feitas(r.get(chave_pub))
        necessarias |= {d for d in datas[feitas:] if d <= date.today()}
    if necessarias:
        try:
            with ui.processando("Baixando Diários Oficiais do site"):
                fora_do_site = diario.baixar_diarios(necessarias, DIARIO, avisar=ui.ok)
            for d in sorted(fora_do_site):
                ui.aviso(f"Diário de {d:%d/%m/%Y} não está na lista do site "
                         "(se existir, baixe e coloque em entradas/Diario_oficial/)")
        except Exception as e:
            ui.aviso(f"Não consegui baixar do site ({type(e).__name__}): uso os PDFs da pasta")

    try:
        with ui.processando("Lendo Diários Oficiais"):
            por_data = diario.indexar_pdfs(DIARIO, avisar=ui.aviso,
                                           padrao=proc().padrao_conferencia())
    except Exception as e:
        ui.erro(f"Não consegui ler os PDFs do Diário: {type(e).__name__}")
        return
    for dia in sorted(por_data):
        for rotulo, achados in por_data[dia]:
            ui.info(f"Diário {rotulo}: {len(achados)} ato(s) administrativo(s)")

    publicados = 0
    divergencias = []
    faltando: dict = {}  # data do Diário que falta → quantos atos esperam por ela
    aguardando_data: dict = {}  # data futura → quantos atos
    for r in candidatos:
        numero = proc().chave_conferencia(r)
        datas = datas_das_tentativas(r)
        if not numero or not datas:
            status.marcar(r["_LINHA"], planilha_status.ERRO_CONFERENCIA)
            motivo = "sem número do ato no título do Diário" if not numero else "sem data_diario"
            divergencias.append(f"Linha {r['_LINHA']}: {motivo}")
            continue

        atual = r.get(chave_pub)
        feitas = planilha_status.tentativas_feitas(atual)
        resultado = None
        # Corrige contagem impossível (tentativa registrada para um Diário que ainda não saiu,
        # vinda da versão antiga da conferência)
        if feitas and datas[feitas - 1] > date.today():
            while feitas and datas[feitas - 1] > date.today():
                feitas -= 1
            atual = f"{planilha_status.NAO_ENCONTRADO} {feitas}" if feitas else ""
            resultado = atual or None
            if not feitas:
                status.marcar(r["_LINHA"], "")  # volta para "não conferido"
        while feitas < 3:
            dia = datas[feitas]
            if dia > date.today():
                aguardando_data[dia] = aguardando_data.get(dia, 0) + 1
                break  # o Diário dessa data ainda não saiu
            if dia not in por_data:
                faltando[dia] = faltando.get(dia, 0) + 1
                break  # sem o Diário dessa data: a tentativa não é gasta
            rotulo = next((rot for rot, achados in por_data[dia] if numero in achados), None)
            if rotulo:
                resultado = f"{planilha_status.PUBLICADO} {rotulo}"
                break
            atual = planilha_status.proxima_tentativa(atual)
            feitas += 1
            resultado = atual

        if resultado is None:
            continue
        status.marcar(r["_LINHA"], resultado)
        if resultado.startswith(planilha_status.PUBLICADO):
            publicados += 1
        else:
            divergencias.append(f"Ato Nº {numero} Não foi encontrado ({resultado})")
    status.concluir()

    arquivo_div = f"divergencias - {processo.name}.xlsx"
    try:
        diario.salvar_divergencias(divergencias, SAIDAS / arquivo_div)
    except PermissionError:
        ui.aviso(f"{arquivo_div} aberto no Excel: não foi atualizado")

    for dia in sorted(aguardando_data):
        ui.info(f"Diário de {dia:%d/%m/%Y} ainda não saiu ({aguardando_data[dia]} ato(s) aguardando)")
    for dia in sorted(faltando):
        ui.aviso(f"Falta o Diário de {dia:%d/%m/%Y} ({faltando[dia]} ato(s) aguardando)")
    ui.ok(f"Publicados: {publicados} de {len(candidatos)} atos conferidos")
    if divergencias:
        ui.erro(f"Divergências em {len(divergencias)} atos → saidas\\{arquivo_div}")


SITES = {"edocs": "Edocs", "diario": "Diário"}


def cadastrar_usuario(site: str = "edocs") -> tuple[str, str]:
    import credenciais

    usuario, senha = ui.login(f"Usuário do {SITES[site]}",
                              "A senha será criptografada e salva em configs/credenciais.json")
    credenciais.salvar(CREDENCIAIS, usuario, senha, site)
    return usuario, senha


def pedir_credenciais(site: str = "edocs") -> tuple[str, str]:
    import credenciais

    return credenciais.carregar(CREDENCIAIS, site) or cadastrar_usuario(site)


RECENTES = 5  # quantas pessoas escolhidas por último aparecem para um clique


def _cpf(texto: str) -> str:
    return "".join(c for c in str(texto) if c.isdigit()).zfill(11)


def escolher_pessoas(pergunta: str, pesquisar, usuario: str, vazio: str | None,
                     lembrar: str, sem_mim: bool = True, chave: str = "") -> list[dict]:
    """Escolhe uma ou mais pessoas (na ordem desejada), pesquisando por CPF, nome ou
    matrícula. pesquisar(termo) → [pessoa]. sem_mim: você mesmo não aparece (assinantes);
    nos vistos aparece, para quem vista os próprios atos se incluir.
    Ficam salvas em configs/config.json: a escolha de cada processo ('ultimas_escolhas',
    vem marcada na próxima vez) e as últimas pessoas usadas ('pessoas_recentes').
    vazio: o que significa não escolher ninguém; None torna a resposta obrigatória.
    chave: onde lembrar a escolha (o nome do processo; "Ato manual" nos atos sem processo)."""
    chave = chave or processo.name
    meu_cpf = _cpf(usuario) if sem_mim else None
    cfg = ler_config()
    recentes = [p for p in cfg.get("pessoas_recentes", []) if isinstance(p, dict)]
    escolhas = cfg.get("ultimas_escolhas", {})
    por_cpf = {p["cpf"]: p for p in recentes}
    anteriores = []
    for p in escolhas.get(chave, {}).get(lembrar, []):
        # Versões antigas guardavam só o CPF
        p = p if isinstance(p, dict) else por_cpf.get(_cpf(p), {"cpf": _cpf(p), "nome": f"CPF {p}"})
        if p["cpf"] != meu_cpf:
            anteriores.append(p)

    def pesquisar_sem_mim(termo: str) -> list[dict]:
        return [p for p in pesquisar(termo) if p["cpf"] != meu_cpf]

    escolhidos = ui.escolher_pessoas(pergunta, pesquisar_sem_mim,
                                     [p for p in recentes if p["cpf"] != meu_cpf][:RECENTES],
                                     anteriores, vazio)
    escolhas.setdefault(chave, {})[lembrar] = guardar_pessoas(escolhidos)
    salvar_config(ultimas_escolhas=escolhas)
    lembrar_recentes(escolhidos)
    return escolhidos


def guardar_pessoas(pessoas: list[dict]) -> list[dict]:
    """Só o que serve para mostrar e pesquisar de novo (sem ids internos do site)."""
    return [{k: p.get(k, "") for k in ("cpf", "nome", "matricula", "cargo", "lotacao")}
            for p in pessoas]


def lembrar_recentes(pessoas: list[dict]) -> None:
    """Põe as pessoas no começo dos 'últimos escolhidos' (configs/config.json)."""
    guardar = guardar_pessoas(pessoas)
    recentes = [p for p in ler_config().get("pessoas_recentes", []) if isinstance(p, dict)]
    novos = guardar + [p for p in recentes if p["cpf"] not in {g["cpf"] for g in guardar}]
    salvar_config(pessoas_recentes=novos[:RECENTES * 2])


def recentes_sem(usuario: str = "") -> list[dict]:
    """Os últimos escolhidos (sem o próprio usuário), para um clique."""
    meu_cpf = _cpf(usuario) if usuario else ""
    return [p for p in ler_config().get("pessoas_recentes", [])
            if isinstance(p, dict) and p.get("cpf") != meu_cpf][:RECENTES]


def nome_da_conta(sessao, usuario: str) -> str:
    """O primeiro nome da conta do e-docs (ex.: "Gabriel"), pesquisado pelo CPF no próprio
    e-docs. Fica guardado em configs/config.json ('nomes_contas'), para não pesquisar toda
    vez. "" se não achar (o envio segue do mesmo jeito)."""
    import edocs_http

    cpf = _cpf(usuario)
    guardados = ler_config().get("nomes_contas", {})
    if guardados.get(cpf):
        return guardados[cpf]
    try:
        nome = edocs_http.buscar_servidor(sessao, cpf).get("nome", "")
    except Exception:
        return ""
    primeiro = nome.split()[0].capitalize() if nome.split() else ""
    if primeiro:
        salvar_config(nomes_contas=guardados | {cpf: primeiro})
    return primeiro


def pessoas_do_envio(usuario: str, pesquisar) -> tuple[list[str], list[str]]:
    """CPFs de quem vista e de quem assina, na ordem. Quem envia só vista se estiver na
    lista (nem todo mundo vista os próprios atos); se estiver, vista sempre em 1º (a ordem
    dos vistos é a ordem em que o ato passa). Se o processo tem vistos e assinantes fixos,
    usa esses sem perguntar; senão, pergunta (pesquisando no e-docs)."""
    fixas = proc()["pessoas"]
    if fixas.get("fixas") and fixas.get("assinantes"):
        vistos = fixas.get("vistos", [])
        assinantes = fixas["assinantes"]
        ui.info("Vistos e assinantes fixos deste processo (para mudar: menu Processos)")
    else:
        vistos = escolher_vistos(usuario, pesquisar)
        assinantes = escolher_assinantes(usuario, pesquisar)
    # A mesma pessoa duas vezes (mesmo CPF ou nome) dá "Seq. Assinatura em duplicidade" no e-docs
    from escolha_pessoas import sem_repetir
    vistos, assinantes = sem_repetir(vistos), sem_repetir(assinantes)
    meu_cpf = _cpf(usuario)
    vistos = [p for p in vistos if p["cpf"] == meu_cpf] + [p for p in vistos if p["cpf"] != meu_cpf]
    mostrar_pessoas(vistos, assinantes)
    return [p["cpf"] for p in vistos], [p["cpf"] for p in assinantes]


def mostrar_pessoas(vistos: list[dict], assinantes: list[dict]) -> None:
    """Sempre mostra quem vista e quem assina, na ordem, fixos ou escolhidos."""
    ui.ok("Vistam: " + (" · ".join(f"{i}º {p['nome']}" for i, p in enumerate(vistos, 1)) or "ninguém"))
    ui.ok("Assinam: " + (" · ".join(f"{i}º {p['nome']}" for i, p in enumerate(assinantes, 1))
                         or "ninguém"))


def escolher_vistos(usuario: str, pesquisar, chave: str = "") -> list[dict]:
    """Pergunta quem vai vistar. Você aparece na pesquisa: se vista os próprios atos, se
    inclui (e fica marcado nas próximas vezes)."""
    return escolher_pessoas("Quem vai vistar? (para vistar você mesmo, inclua o seu nome)", pesquisar,
                            usuario, "ninguém", "vistos", sem_mim=False, chave=chave)


def escolher_assinantes(usuario: str, pesquisar, chave: str = "") -> list[dict]:
    """Pergunta quem vai assinar (sem o próprio usuário)."""
    return escolher_pessoas("Quem vai assinar?", pesquisar, usuario, None, "assinantes", chave=chave)


# Colunas de controle na planilha (a ordem define a ordem das colunas)
COLUNAS_PLANILHA = ["status_edocs", "data_edocs", "status_diario", "data_diario",
                    "status_publicado", "criado_por"]
# criado_por: o primeiro nome da conta do e-docs que enviou o ato (vai no fim, depois das que
# já existem, para nenhuma coluna mudar de lugar nas planilhas antigas)
# O que "desfazer" apaga de cada etapa. Desfazer o Diário apaga também a conferência da
# publicação (ela depende do envio ao Diário).
DESFAZER = {"edocs": ["status_edocs", "data_edocs", "criado_por"],
            "diario": ["status_diario", "data_diario", "status_publicado"]}


class AtoPublicado(Exception):
    """O ato já saiu no Diário Oficial: é público, os status não podem mais mudar."""


class DiarioAntes(Exception):
    """O Diário ainda está preenchido: desfaça o Diário antes do e-docs (cascata)."""


def publicado(registro: dict) -> bool:
    """O ato já foi encontrado publicado no Diário Oficial (conferência da publicação)."""
    return str(registro.get("STATUSPUBLICADO", "")).startswith(planilha_status.PUBLICADO)


def desfazer_status(planilha: Path, linha: int, etapa: str) -> None:
    """Apaga o status (e a data, se houver) de uma etapa ('edocs' ou 'diario') na linha.
    Lança PermissionError se a planilha estiver aberta no Excel, AtoPublicado se o ato já
    saiu no Diário Oficial e DiarioAntes se for o e-docs com o Diário ainda preenchido (as
    etapas se desfazem de trás para frente). Confere direto no arquivo, na hora."""
    status = StatusPlanilha(planilha, etapa)
    atual = planilha_status.linha_da_planilha(planilha, status.aba, linha)
    if atual.get(status.reservadas["status_publicado"], "").startswith(planilha_status.PUBLICADO):
        raise AtoPublicado("o ato já foi publicado no Diário Oficial")
    if etapa == "edocs" and any(atual.get(status.reservadas[c], "").strip() for c in DESFAZER["diario"]):
        raise DiarioAntes("desfaça primeiro o envio no Diário")
    valores = {c: {} for c in status.reservadas}
    for coluna in DESFAZER[etapa]:
        valores[coluna][linha] = ""
    planilha_status.gravar(planilha, status.aba, status.reservadas, valores)


def desfazer_varios(planilha: Path, linhas: list[int], etapa: str) -> dict[str, list[int]]:
    """Desfaz o status de várias linhas de uma vez (uma leitura e uma gravação na planilha,
    em vez de uma por linha). etapa: 'diario' (só o Diário) ou 'tudo' (Diário e e-docs: a
    linha volta ao começo). Linhas já publicadas no Diário Oficial ficam como estão.
    Devolve {"desfeitas": [...], "publicadas": [...], "sem_status": [...]}.
    Lança PermissionError se a planilha estiver aberta no Excel."""
    status = StatusPlanilha(planilha, "edocs")
    etapas = ["diario"] if etapa == "diario" else ["diario", "edocs"]
    colunas = [c for e in etapas for c in DESFAZER[e]]
    atuais = planilha_status.linhas_da_planilha(planilha, status.aba, linhas)
    resultado: dict[str, list[int]] = {"desfeitas": [], "publicadas": [], "sem_status": []}
    valores = {c: {} for c in status.reservadas}
    for linha in sorted(set(linhas)):
        atual = atuais.get(linha, {})
        if atual.get(status.reservadas["status_publicado"], "").startswith(planilha_status.PUBLICADO):
            resultado["publicadas"].append(linha)
            continue
        if not any(atual.get(status.reservadas[c], "").strip() for c in colunas):
            resultado["sem_status"].append(linha)
            continue
        for coluna in colunas:
            valores[coluna][linha] = ""
        resultado["desfeitas"].append(linha)
    if resultado["desfeitas"]:
        planilha_status.gravar(planilha, status.aba, status.reservadas, valores)
    return resultado


class StatusPlanilha:
    """Grava, na planilha de dados, as colunas status_<site> e data_<site> de cada ato.
    Se o arquivo estiver aberto no Excel, guarda e tenta de novo na próxima gravação."""

    def __init__(self, planilha: Path, site: str, com_data: bool = True):
        self.planilha = planilha
        self.aba = proc().aba or atos.ABA
        self.col_status = f"status_{site}"
        self.col_data = f"data_{site}" if com_data else None
        # No e-docs, grava também quem enviou (primeiro nome da conta), se souber
        self.col_criado = "criado_por" if site == "edocs" else None
        self.criado_por = ""
        # "status" era o nome antigo da coluna do e-docs (vale para qualquer site que grave)
        apelidos = {"status_edocs": ["status"]}
        # Reserva as colunas de todos os sites numa ordem fixa, para um site nunca pegar a
        # coluna livre que outro ainda vai usar
        proprias = [c for c in (self.col_status, self.col_data, self.col_criado) if c]
        todos = COLUNAS_PLANILHA + [c for c in proprias if c not in COLUNAS_PLANILHA]
        self.reservadas = planilha_status.colunas(planilha, self.aba, todos, apelidos)
        self.cols = {c: self.reservadas[c] for c in proprias}
        self.pendentes: dict[str, dict[int, str]] = {c: {} for c in proprias}
        self.avisou = False
        self._chave = atos.normalizar(self.col_status)
        col_edocs = self.reservadas["status_edocs"]
        if self._cabecalho(col_edocs).lower() == "status":
            if site == "edocs":
                self._chave = "STATUS"
            try:  # renomeia o cabeçalho antigo, mantendo os valores
                planilha_status.gravar(planilha, self.aba, self.reservadas,
                                       {c: {} for c in self.reservadas})
            except (PermissionError, OSError):
                pass  # tenta de novo na próxima gravação (o cabeçalho é sempre reescrito)

    def _cabecalho(self, col: int) -> str:
        return planilha_status.cabecalho(self.planilha, self.aba).get(col, "").strip()

    def chave(self) -> str:
        """Nome da coluna de status como aparece nos registros lidos (normalizado)."""
        return self._chave

    def marcar(self, linha: int | list[int], texto: str, data: str | None = None) -> None:
        """Grava o status na linha (ou em todas as linhas de um ato com várias pessoas).
        Quando enviado, grava também a data: a informada (ex.: data de veiculação no Diário)
        ou, se nenhuma for passada, a data e hora do envio."""
        for n in (linha if isinstance(linha, list) else [linha]):
            self.pendentes[self.col_status][n] = texto
            if texto == planilha_status.ENVIADO and self.col_data:  # data só quando enviado
                self.pendentes[self.col_data][n] = data or datetime.now().strftime("%d/%m/%Y %H:%M")
            if texto == planilha_status.ENVIADO and self.col_criado and self.criado_por:
                self.pendentes[self.col_criado][n] = self.criado_por
        self._gravar()

    def _gravar(self) -> bool:
        if not any(self.pendentes.values()):
            return True
        # Inclui todas as colunas reservadas para gravar os cabeçalhos de todos os sites
        valores = {c: {} for c in self.reservadas} | self.pendentes
        try:
            planilha_status.gravar(self.planilha, self.aba, self.reservadas, valores)
        except PermissionError:
            if not self.avisou:
                quem = planilha_status.em_uso(self.planilha)
                ui.aviso(f"Planilha aberta no Excel{f' por {quem}' if quem else ''}: "
                         "o status será gravado quando ela for fechada")
                self.avisou = True
            return False
        except Exception as e:
            ui.erro(f"Não consegui gravar o status na planilha: {type(e).__name__}")
            return False
        for linhas in self.pendentes.values():
            linhas.clear()
        return True

    def concluir(self) -> None:
        """No fim, grava o que ficou pendente. Se a planilha estiver aberta no Excel (aqui ou
        por outra pessoa na rede), espera ela ser fechada e grava sozinho; de minuto em minuto
        pergunta se continua esperando."""
        while not self._gravar():
            quem = planilha_status.em_uso(self.planilha)
            if quem is None:  # não estava aberta: falhou por outro motivo
                if ui.confirmar("Não consegui gravar o status na planilha. Tentar de novo?",
                                sim="Tentar de novo", nao="Pular"):
                    continue
            else:
                por = f" por {quem}" if quem else ""
                with ui.processando(f"Aguardando a planilha ser fechada no Excel{por} "
                                    "para gravar o status..."):
                    limite = time.time() + 60
                    while time.time() < limite and planilha_status.em_uso(self.planilha) is not None:
                        time.sleep(2)
                if planilha_status.em_uso(self.planilha) is None:
                    continue
                if ui.confirmar(f"A planilha continua aberta no Excel{por}. Continuar esperando?",
                                sim="Esperar mais", nao="Pular"):
                    continue
            qtd = len(self.pendentes[self.col_status])
            ui.erro(f"Status de {qtd} ato(s) não gravado na planilha")
            return


def preparar_atos(site: str, depois_de: str | None = None, sem_titulo: bool = False):
    """Parte comum aos sites: lê a planilha, separa os atos pendentes daquele site (coluna
    status_<site> vazia) e monta texto e HTML de cada um. Com 'depois_de', só entram os
    atos que já estão "Enviado com sucesso" naquele outro site. sem_titulo tira a linha
    do título do texto (o Diário recebe o título num campo próprio).
    Retorna (registros, textos, htmls, status) ou None se não houver o que lançar."""
    modelo = achar_modelo()
    if not modelo:
        return None
    planilha = achar_planilha()
    if not planilha:
        return None

    with ui.processando("Lendo planilha"):
        # Cada ato pode juntar várias linhas (várias pessoas no mesmo ato)
        registros = proc().agrupar(proc().ler_planilha(planilha))
    # O site anterior é lido primeiro: abrir o StatusPlanilha pode renomear o cabeçalho
    # antigo ("status"), e a chave precisa bater com a planilha já lida
    chave_antes = StatusPlanilha(planilha, depois_de).chave() if depois_de else None
    status = StatusPlanilha(planilha, site)
    chave = status.chave()

    def todas(ato: dict, teste) -> bool:
        return all(teste(pessoa) for pessoa in atos.itens(ato))

    if depois_de:
        enviado_antes = lambda r: r.get(chave_antes) == planilha_status.ENVIADO  # noqa: E731
        aguardando = sum(1 for r in registros if not todas(r, enviado_antes))
        registros = [r for r in registros if todas(r, enviado_antes)]
        if aguardando:
            ui.info(f"Ainda não enviados no e-docs (ficam de fora): {aguardando}")

    # Só entram os atos com status vazio (já enviados ou com erro ficam de fora)
    enviados_antes = sum(1 for r in registros if todas(r, lambda p: p.get(chave) == planilha_status.ENVIADO))
    com_erro = sum(1 for r in registros if any(planilha_status.eh_erro(p.get(chave)) for p in atos.itens(r)))
    pendentes = [r for r in registros if todas(r, lambda p: not p.get(chave))]
    # Ato com pessoas já enviadas e outras não (ex.: alguém entrou no grupo depois)
    misturados = [r for r in registros if r not in pendentes and any(not p.get(chave) for p in atos.itens(r))]
    for r in misturados:
        ui.aviso(f"Ficou de fora {proc().descricao(r)}: parte das pessoas do ato já tem status "
                 f"(linhas {', '.join(map(str, processos.linhas(r)))})")
    registros = pendentes
    if enviados_antes or com_erro:
        ui.info(f"Já enviados: {enviados_antes} · Com erro: {com_erro} · Pendentes: {len(registros)}")
    if not registros:
        ui.ok("Nenhum ato pendente na planilha")
        return None

    # Campos do Word sem coluna na planilha sairiam como [CAMPO] no ato
    faltam = proc().campos_sem_ligacao(atos.cabecalho_planilha(planilha, proc().aba)[1])
    if faltam:
        ui.aviso("Campos do Word sem coluna na planilha (edite o processo para ligar): "
                 + ", ".join(faltam))

    with ui.processando("Montando textos dos atos"):
        textos = proc().textos(registros, sem_titulo=sem_titulo)
        htmls = proc().htmls(registros, sem_titulo=sem_titulo)
    ui.ok(f"{len(textos)} ato(s) prontos")
    return registros, textos, htmls, status


def resolver_em_montagem(sessao) -> bool:
    """Antes do 1º envio: documento em montagem no Edocs faz o site gravar o 1º ato por cima
    dele (vistos/assinantes duplicados). Mostra quais são e só cancela com autorização.
    False = não seguir com o envio."""
    import edocs_http

    with ui.processando("Procurando documentos em montagem no e-docs"):
        try:
            docs = edocs_http.documentos_em_montagem(sessao)
        except Exception as e:
            ui.aviso(f"Não deu para conferir documentos em montagem: {type(e).__name__} {e}".strip())
            return True
    if not docs:
        return True
    lista = "\n".join(f"• {d.get('titulo') or '(sem título)'} · {d.get('tipo', '')} · "
                      f"{d.get('data', '')} · {d.get('identificacao') or d['doc']}" for d in docs)
    ui.aviso(f"{len(docs)} documento(s) em montagem no e-docs:\n{lista}")
    if not ui.confirmar(
            f"Você tem {len(docs)} documento(s) em montagem no e-docs (criado e não enviado):"
            f"\n\n{lista}\n\nEnquanto houver um documento em montagem, o e-docs não cria documento "
            "novo: ele abre o em montagem mais recente no lugar, e o Edodocs não consegue enviar. "
            "Cancelar esse(s) documento(s) em montagem? Só eles serão cancelados; nenhum documento já "
            "enviado é mexido. Se algum for um rascunho que você quer manter, escolha Não e termine "
            "ou envie ele no e-docs antes.",
            sim="Cancelar os em montagem", nao="Não, parar o envio"):
        ui.aviso("Envio não iniciado. Envie ou cancele o documento em montagem no e-docs e "
                 "comece de novo.")
        return False
    for d in docs:
        try:
            with ui.processando(f"Cancelando o documento em montagem {d['doc']}"):
                edocs_http.cancelar_em_montagem(sessao, d["doc"], avisar=ui.ok)
        except Exception as e:
            ui.erro(f"Não deu para cancelar o documento {d['doc']}: {e}. Envio não iniciado; "
                    "cancele no e-docs e comece de novo.")
            return False
    return True


def lancar_edocs_http() -> None:
    """Fluxo por HTTP: cria o documento, vistos, assinantes, destinatário e envia, ato por
    ato. Pergunta antes de enviar (se configurado) e grava o status na planilha."""
    import edocs_http

    preparado = preparar_atos("edocs")
    if not preparado:
        return
    registros, textos, htmls, status = preparado

    # Processo com anexo: os anexos de todos os atos são escolhidos agora, de uma vez, para
    # o envio depois seguir sem parar no meio
    escolhidos: list[list] = [[] for _ in registros]
    if proc().precisa_anexo:
        coluna_nome = atos.normalizar(proc().config["coluna_nome"])
        lista = []
        for dados in registros:
            pessoas = [p.get(coluna_nome, "") for p in atos.itens(dados)]
            lista.append({"ato": proc().descricao(dados), "pessoas": pessoas if len(pessoas) > 1 else None})
        resposta = ui.anexos_dos_atos(lista)
        if resposta == "cancelar":
            ui.aviso("Cancelado")
            return
        escolhidos = resposta

    # Pausas configuradas na "Área perigosa" (4 → 4). O padrão são os valores do módulo.
    cfg = ler_config()
    edocs_http.PAUSA_REQUISICAO = float(cfg.get("pausa_requisicao", edocs_http.PAUSA_REQUISICAO))
    # Tipo do documento e setor de cada processo
    edocs_http.TIPO_ATO = proc()["edocs"]["tipo"]
    edocs_http.SETOR_POSICAO = int(proc()["edocs"]["setor_posicao"])
    edocs_http.GRUPO_DESTINATARIO = proc()["edocs"].get("grupo_destinatario", "")
    pausa_entre_atos = float(cfg.get("pausa_entre_atos", PAUSA_ENTRE_ATOS))

    usuario, senha = pedir_credenciais()
    ui.info(f"Usuário: [bold]{usuario}[/]")

    with ui.processando("Entrando no Edocs"):
        try:
            sessao = edocs_http.entrar(usuario, senha)
        except edocs_http.LoginRecusado:
            import credenciais
            credenciais.remover(CREDENCIAIS, "edocs")
            ui.erro("Login recusado. Usuário do Edocs apagado, tente de novo")
            return
        except Exception as e:
            ui.erro(f"Falha ao entrar no Edocs: {type(e).__name__} {e}".strip())
            return
    ui.ok("Logado")
    # Quem está enviando, para a coluna criado_por da planilha
    status.criado_por = nome_da_conta(sessao, usuario)
    if status.criado_por:
        ui.info(f"Os atos vão como criados por {status.criado_por}")

    # Vistos e assinantes pesquisados no próprio Edocs, na ordem escolhida
    def pesquisar(termo: str) -> list[dict]:
        return edocs_http.pesquisar_servidores(sessao, termo)

    vistos, assinantes = pessoas_do_envio(usuario, pesquisar)
    # Só aqui, antes do 1º ato, documentos em montagem podem ser cancelados (com autorização)
    if not resolver_em_montagem(sessao):
        return
    if not perguntar_antes_de_enviar():
        ui.aviso("Este processo envia direto, um atrás do outro, sem perguntar "
                 "(para mudar: Processos → Envio)")

    confirmar = perguntar_antes_de_enviar()
    enviados = pulados = 0
    docs_criados: set[str] = set()  # para detectar se o Edocs reutilizou um documento
    try:
        for n, (dados, texto, corpo, escolha) in enumerate(zip(registros, textos, htmls, escolhidos), 1):
            ato = proc().descricao(dados)
            ui.progresso(n, len(registros), ato)
            # Pausa entre um ato e o seguinte (não antes do primeiro), para o Edocs e o
            # Akamai não verem a sequência inteira como rajada
            if n > 1 and pausa_entre_atos:
                time.sleep(pausa_entre_atos)
            try:
                titulo_ato = proc().titulo_edocs(dados)
                if faltam := proc().faltando(dados):
                    raise RuntimeError("campo(s) obrigatório(s) vazio(s) na planilha: " + "; ".join(faltam))
            except RuntimeError as e:
                pulados += 1
                status.marcar(processos.linhas(dados), f"{planilha_status.ERRO}: {e}"[:250])
                ui.erro(f"Pulado {ato}: {e}")
                continue
            # Processo com anexo: os arquivos escolhidos no começo (sem nenhum, o ato é pulado
            # antes de criar o documento, para não deixar rascunho no e-docs)
            anexos = [anexo_da_pessoa(c) for c in escolha]
            if proc().precisa_anexo and not anexos:
                pulados += 1
                ui.aviso(f"Pulado {ato} (sem anexo)")
                continue
            doc = None
            try:
                with ui.processando(f"Preparando {ato}"):
                    doc = edocs_http.preparar_documento(
                        sessao, titulo_ato, corpo, vistos, assinantes, avisar=ui.ok)
                for arquivo in anexos:
                    with ui.processando(f"Anexando {arquivo.name}"):
                        edocs_http.anexar_arquivo(sessao, doc, arquivo, avisar=ui.ok)
            except edocs_http.SessaoExpirada:
                ui.erro("A sessão caiu: faça login de novo")
                return
            except edocs_http.ErroEdocs as e:
                pulados += 1
                if doc:  # o anexo não entrou: o rascunho em montagem travaria os próximos atos
                    e = f"{e}. {edocs_http.descartar_rascunho(sessao, doc)}"
                status.marcar(processos.linhas(dados), f"{planilha_status.ERRO}: {e}"[:250])
                ui.erro(f"Pulado {ato}: {e}")
                continue

            # Se o Edocs devolveu o mesmo número de outro ato, o documento anterior não foi
            # enviado de verdade (ficou em montagem) e este ato está sendo gravado por cima.
            if doc in docs_criados:
                status.marcar(processos.linhas(dados), f"{planilha_status.ERRO}: documento {doc} reutilizado")
                ui.erro(f"Parado em {ato}: o documento {doc} foi reutilizado "
                        "(o ato anterior deve ter ficado em montagem)")
                ui.aviso("Confira no Edocs os atos já gravados antes de rodar de novo")
                break
            docs_criados.add(doc)

            if confirmar:
                escolha = ui.confirmar_envio(ato, texto, folha_do_ato(dados), lambda d=dados: word_do_ato(d))
                if escolha == "cancelar":
                    with ui.processando(f"Cancelando o rascunho {doc}"):
                        resultado = edocs_http.descartar_rascunho(sessao, doc)
                    ui.aviso(f"Cancelado. {resultado}")
                    break
                if escolha == "todos":
                    # Conferiu e confia: os próximos atos desta vez vão sem perguntar
                    confirmar = False
                    ui.info("Os próximos atos serão enviados sem perguntar")
                if escolha == "pular":
                    pulados += 1
                    # Em montagem, o rascunho faria o e-docs recusar todos os atos seguintes
                    with ui.processando(f"Cancelando o rascunho {doc}"):
                        resultado = edocs_http.descartar_rascunho(sessao, doc)
                    ui.aviso(f"Pulado {ato}. {resultado}")
                    continue

            ui.aviso(f"Esse ato será enviado: [bold]{ato}[/]")
            # O envio nunca é repetido sozinho, para não enviar o mesmo ato duas vezes
            try:
                with ui.processando("Enviando"):
                    edocs_http.enviar_documento(sessao, doc, avisar=ui.ok)
            except edocs_http.ErroEdocs as e:
                status.marcar(processos.linhas(dados), f"{planilha_status.ERRO}: envio: {e}"[:250])
                ui.erro(f"Falha no envio de {ato}: {e}")
                ui.aviso("Confira no Edocs se o ato foi enviado antes de rodar de novo")
                break
            except Exception as e:
                # Caiu no meio do envio (conexão, sessão...): não dá para saber se foi.
                # Marca como erro para o ato nunca ser reenviado sozinho (evita duplicidade).
                status.marcar(processos.linhas(dados), f"{planilha_status.ERRO}: envio sem confirmação "
                                               f"({type(e).__name__}): confira no Edocs"[:250])
                ui.erro(f"Envio de {ato} sem confirmação: {type(e).__name__} {e}".strip())
                ui.aviso("Confira no Edocs se o ato foi enviado antes de rodar de novo")
                break
            enviados += 1
            ui.ok(f"Enviado {ato}")
            status.marcar(processos.linhas(dados), planilha_status.ENVIADO)
    except KeyboardInterrupt:
        ui.aviso("Cancelado")
    finally:
        status.concluir()
        if enviados or pulados:
            ui.info(f"Enviados {enviados} · Pulados {pulados} · Total {len(registros)}")


def anexo_da_pessoa(escolha) -> Path:
    """O arquivo a anexar. Escolhido para uma pessoa de um ato com várias ([caminho,
    nome]), vai com o nome dela na frente ("DIEGO NOBRE - requerimento.pdf", uma cópia):
    no e-docs dá para ver de quem é cada anexo, e arquivos de mesmo nome não se confundem."""
    if isinstance(escolha, str):
        return Path(escolha)
    caminho, pessoa = Path(escolha[0]), " ".join(str(escolha[1] or "").split())
    if not pessoa or atos.normalizar(pessoa) in atos.normalizar(caminho.stem):
        return caminho
    import shutil
    import tempfile
    nome = re.sub(r'[\\/:*?"<>|]', "", f"{pessoa} - {caminho.name}")
    destino = Path(tempfile.mkdtemp(prefix="edodocs_anexo_")) / nome
    shutil.copy2(caminho, destino)
    return destino


PASTA_IMPRESSAO = Path(tempfile.gettempdir()) / "Edodocs - atos para imprimir"


def word_do_ato(dados: dict, processo: processos.Processo | None = None) -> Path:
    """Grava o ato preenchido num Word (numa pasta temporária) e devolve o caminho. Se um
    Word com o mesmo nome estiver aberto, grava com outro nome."""
    proc_ = processo or proc()
    nome = re.sub(r'[<>:"/\\|?*]+', "-", proc_.descricao(dados)).strip(" .-") or "ato"
    nome = f"{proc_.nome} - {nome}"[:120]
    PASTA_IMPRESSAO.mkdir(parents=True, exist_ok=True)
    for n in range(1, 50):
        destino = PASTA_IMPRESSAO / (f"{nome}.docx" if n == 1 else f"{nome} ({n}).docx")
        try:
            if destino.exists():
                destino.unlink()  # cópia antiga: se estiver aberta no Word, não apaga
            return proc_.salvar_word(dados, destino)
        except PermissionError:
            continue
    raise RuntimeError("feche os Words abertos desta pasta e tente de novo")


def folha_do_ato(dados: dict) -> str:
    """O ato como o papel do Word, para a pergunta antes de enviar ("" se não der)."""
    try:
        return proc().folhas([dados])[0]
    except Exception:
        return ""


def preparado_chave_edocs(registros: list[dict]) -> str:
    """Nome (normalizado) da coluna de status do e-docs nos registros lidos."""
    return "STATUSEDOCS" if any("STATUSEDOCS" in r for r in registros) else "STATUS"


def data_publicacao(dados: dict) -> date:
    """1 dia útil depois da data de envio ao e-docs (data_edocs). Se a planilha não tiver
    a data (atos enviados antes da coluna existir), pergunta no terminal."""
    enviado = atos.ler_data(dados.get("DATAEDOCS", ""))
    while enviado is None:
        resp = ui.perguntar(f"Sem data_edocs para {proc().numero(dados) or '?'}. "
                            "Data de envio ao e-docs (dd/mm/aaaa, vazio = hoje)")
        enviado = date.today() if not resp else atos.ler_data(resp)
        if enviado is None:
            ui.aviso("Data inválida")
    return atos.proximo_dia_util(enviado)


def lancar_diario_http() -> None:
    """Fluxo por HTTP: cadastra cada ato no Diário (título num campo próprio, corpo sem a
    linha do título) com a data de veiculação (1 dia útil após o envio no e-docs). Pergunta
    antes de enviar (se configurado) e grava o status na planilha."""
    import diario_http

    if not proc().vai_ao_diario:
        ui.aviso(f"O processo {processo.name} não é publicado no Diário Oficial (para mudar: Processos)")
        return
    # O Diário recebe o título num campo próprio: o corpo vai sem a linha do título
    preparado = preparar_atos("diario", depois_de="edocs", sem_titulo=True)
    if not preparado:
        return
    registros, textos, htmls, status = preparado

    cfg = ler_config()
    diario_http.PAUSA_REQUISICAO = float(cfg.get("pausa_requisicao", diario_http.PAUSA_REQUISICAO))
    # Secretaria, categoria e coluna de cada processo (o valor da gravação só vale para o padrão)
    opcoes = proc()["diario"]
    diario_http.usar_opcoes(opcoes["secretaria"], opcoes["categoria"], opcoes["coluna"])
    pausa_entre_atos = float(cfg.get("pausa_entre_atos", PAUSA_ENTRE_ATOS))

    usuario, senha = pedir_credenciais("diario")
    ui.info(f"Usuário do Diário: [bold]{usuario}[/]")
    if not perguntar_antes_de_enviar():
        ui.aviso("Este processo envia direto, um atrás do outro, sem perguntar "
                 "(para mudar: Processos → Envio)")

    with ui.processando("Entrando no Diário"):
        try:
            sessao = diario_http.entrar(usuario, senha)
        except diario_http.LoginRecusado:
            import credenciais
            credenciais.remover(CREDENCIAIS, "diario")
            ui.erro("Login recusado. Usuário do Diário apagado, tente de novo")
            return
        except Exception as e:
            ui.erro(f"Falha ao entrar no Diário: {type(e).__name__} {e}".strip())
            return
    ui.ok("Logado")

    confirmar = perguntar_antes_de_enviar()
    enviados = pulados = 0
    try:
        for n, (dados, texto, corpo) in enumerate(zip(registros, textos, htmls), 1):
            ato = proc().descricao(dados)
            ui.progresso(n, len(registros), ato)
            if n > 1 and pausa_entre_atos:
                time.sleep(pausa_entre_atos)

            try:
                titulo_ato = proc().titulo_diario(dados)
            except RuntimeError as e:
                pulados += 1
                status.marcar(processos.linhas(dados), f"{planilha_status.ERRO}: {e}"[:250])
                ui.erro(f"Pulado {ato}: {e}")
                continue
            data_pub = data_publicacao(dados)

            if confirmar:
                escolha = ui.confirmar_envio(ato, texto, folha_do_ato(dados), lambda d=dados: word_do_ato(d))
                if escolha == "cancelar":
                    ui.aviso("Cancelado")
                    break
                if escolha == "todos":
                    # Conferiu e confia: os próximos atos desta vez vão sem perguntar
                    confirmar = False
                    ui.info("Os próximos atos serão enviados sem perguntar")
                if escolha == "pular":
                    pulados += 1
                    ui.aviso(f"Pulado {ato}")
                    continue

            ui.aviso(f"Esse ato será enviado: [bold]{ato}[/] · veiculação {data_pub:%d/%m/%Y}")
            # O envio nunca é repetido sozinho, para não publicar o mesmo ato duas vezes
            try:
                with ui.processando("Enviando"):
                    diario_http.enviar_ato(sessao, titulo_ato, corpo, data_pub, avisar=ui.ok)
            except diario_http.ErroDiario as e:
                status.marcar(processos.linhas(dados), f"{planilha_status.ERRO}: {e}"[:250])
                ui.erro(f"Falha no envio de {ato}: {e}")
                ui.aviso("Confira no Diário se o ato foi publicado antes de rodar de novo")
                break
            except Exception as e:
                # Sessão caiu ou conexão falhou no meio do envio: não dá para saber se foi.
                # Marca como erro para o ato nunca ser reenviado sozinho (evita duplicidade).
                status.marcar(processos.linhas(dados), f"{planilha_status.ERRO}: envio sem confirmação "
                                               f"({type(e).__name__}): confira no Diário"[:250])
                ui.erro(f"Envio de {ato} sem confirmação: {type(e).__name__} {e}".strip())
                ui.aviso("Confira no Diário se o ato foi publicado antes de rodar de novo")
                break
            enviados += 1
            ui.ok(f"Publicado {ato} · veiculação {data_pub:%d/%m/%Y}")
            # data_diario = data de veiculação enviada
            status.marcar(processos.linhas(dados), planilha_status.ENVIADO, data_pub.strftime("%d/%m/%Y"))
    except KeyboardInterrupt:
        ui.aviso("Cancelado")
    finally:
        status.concluir()
        if enviados or pulados:
            ui.info(f"Enviados {enviados} · Pulados {pulados} · Total {len(registros)}")


aplicar_pasta_compartilhada()
