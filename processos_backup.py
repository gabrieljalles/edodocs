"""Exportar e importar processos: para levar todos os processos para uma versão nova do
Edodocs (ou para outro computador) sem perder nada.

Exportar gera um .zip na pasta Downloads da pessoa ("Edodocs - processos 2026-10-14 15h30.zip")
com a pasta de cada processo (Words, processo.json e a planilha, se estiver na pasta) e um
manifesto com a lista e o caminho da planilha que cada processo usava neste computador.
Importar lê esse .zip (o mais novo de Downloads, ou outro escolhido) e copia os processos
escolhidos para entradas/processos. Um processo que já existe com o mesmo nome não é apagado:
a pasta atual vai antes para saidas/Backup de processos.

As tabelas de consulta que os processos usam (ex.: Servidores, ver consultas.py) podem ir
junto, em tabelas/ no .zip, com o prazo de atualização. Ao importar, o Excel vai para a
pasta das tabelas e a tabela é cadastrada com o mesmo nome (o processo já chega ligado).

A planilha do processo que fica fora da pasta dele (na rede, ou escolhida em Configurações)
também pode ir junto, em planilhas/<processo>/ no .zip. Ao importar, a pessoa escolhe se usa
essa cópia (ela vai para a pasta do processo) ou a do caminho original.
"""
import json
import os
import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path, PurePosixPath

PREFIXO = "Edodocs - processos"
MANIFESTO = "edodocs_processos.json"
INVALIDOS = re.compile(r'[<>:"/\\|?*]')


def pasta_downloads() -> Path:
    """A pasta Downloads do usuário do Windows, mesmo se ela foi movida (para outro disco ou
    para o OneDrive); sem o Windows responder, a Downloads dentro da pasta do usuário."""
    try:
        import ctypes
        from ctypes import wintypes
        import uuid

        guid = uuid.UUID("{374DE290-123F-4565-9164-39C4925E467B}")  # FOLDERID_Downloads

        class GUID(ctypes.Structure):
            _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD),
                        ("Data3", wintypes.WORD), ("Data4", ctypes.c_ubyte * 8)]

        g = GUID(guid.fields[0], guid.fields[1], guid.fields[2],
                 (ctypes.c_ubyte * 8)(*guid.bytes[8:]))
        caminho = ctypes.c_wchar_p()
        if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(g), 0, None, ctypes.byref(caminho)) == 0:
            pasta = Path(caminho.value)
            ctypes.windll.ole32.CoTaskMemFree(caminho)
            if pasta.is_dir():
                return pasta
    except Exception:
        pass
    pasta = Path.home() / "Downloads"
    pasta.mkdir(parents=True, exist_ok=True)
    return pasta


def _processos(pasta_processos: Path) -> list[Path]:
    """Pastas de processo (as que têm um Word), como em main.listar_processos."""
    if not pasta_processos.is_dir():
        return []
    return sorted((p for p in pasta_processos.iterdir()
                   if p.is_dir() and any(not a.name.startswith("~$") for a in p.glob("*.docx"))),
                  key=lambda p: p.name.lower())


def _tabelas_dos_processos(pasta_processos: Path) -> dict[str, list[str]]:
    """{tabela de consulta: processos que a usam}, cadastrada ou não."""
    todas: dict[str, list[str]] = {}
    for pasta in _processos(pasta_processos):
        try:
            config = json.loads((pasta / "processo.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        nome = (config.get("consulta") or {}).get("tabela")
        if nome:
            todas.setdefault(nome, []).append(pasta.name)
    return todas


def _tem_arquivo(nome: str, cadastradas: dict[str, dict]) -> bool:
    return Path((cadastradas.get(nome) or {}).get("arquivo") or "").is_file()


def tabelas_usadas(pasta_processos: Path, cadastradas: dict[str, dict]) -> dict[str, list[str]]:
    """{tabela de consulta: processos que a usam}, só das cadastradas com o arquivo existindo."""
    return {nome: procs for nome, procs in _tabelas_dos_processos(pasta_processos).items()
            if _tem_arquivo(nome, cadastradas)}


def tabelas_faltando(pasta_processos: Path, cadastradas: dict[str, dict]) -> dict[str, list[str]]:
    """{tabela de consulta: processos que a usam} das que os processos usam mas não estão
    neste computador (não cadastradas aqui, ou o arquivo sumiu): sem aviso, não iriam junto."""
    return {nome: procs for nome, procs in _tabelas_dos_processos(pasta_processos).items()
            if not _tem_arquivo(nome, cadastradas)}


def _arquivo_tabela(nome: str, sufixo: str) -> str:
    return INVALIDOS.sub("_", nome).strip(" .") + sufixo.lower()


def exportar(pasta_processos: Path, planilhas: dict, destino: Path | None = None,
             tabelas: dict[str, dict] | None = None, de_para: dict | None = None,
             planilhas_junto: dict[str, Path] | None = None) -> tuple[Path, list[str]]:
    """Gera o .zip com todos os processos. planilhas: o caminho da planilha que cada processo
    usa neste computador (config "planilhas"). tabelas: as tabelas de consulta que vão junto
    ({nome: cadastro, com "arquivo"}); de_para: os de-para (código → descrição) que vão junto;
    planilhas_junto: {processo: planilha de fora da pasta dele} que vão junto.
    Devolve (arquivo, nomes dos processos)."""
    pastas = _processos(pasta_processos)
    if not pastas:
        raise RuntimeError("não há processos para exportar")
    destino = destino or pasta_downloads()
    arquivo = destino / f"{PREFIXO} {datetime.now():%Y-%m-%d %Hh%M}.zip"
    n = 2
    while arquivo.exists():  # duas exportações no mesmo minuto
        arquivo = destino / f"{PREFIXO} {datetime.now():%Y-%m-%d %Hh%M} ({n}).zip"
        n += 1
    nomes = [p.name for p in pastas]
    manifesto = {
        "versao": 1,
        "exportado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "computador": os.environ.get("COMPUTERNAME", ""),
        "usuario": os.environ.get("USERNAME", ""),
        "processos": nomes,
        "planilhas": {nome: planilhas[nome] for nome in nomes if planilhas.get(nome)},
        "tabelas": {},
        "de_para": de_para or {},
        "planilhas_junto": {nome: Path(c).name for nome, c in (planilhas_junto or {}).items() if nome in nomes},
        "planilhas_origem": {nome: str(c) for nome, c in (planilhas_junto or {}).items() if nome in nomes},
    }
    usadas = tabelas_usadas(pasta_processos, tabelas or {})
    for nome, dados in (tabelas or {}).items():
        manifesto["tabelas"][nome] = {
            **{k: v for k, v in dados.items() if k != "arquivo"},
            "arquivo": f"tabelas/{_arquivo_tabela(nome, Path(dados['arquivo']).suffix)}",
            "processos": usadas.get(nome, []),
        }
    temporario = arquivo.with_suffix(".zip.parcial")
    with zipfile.ZipFile(temporario, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(MANIFESTO, json.dumps(manifesto, ensure_ascii=False, indent=2))
        for pasta in pastas:
            for item in sorted(pasta.rglob("*")):
                if item.is_file() and not item.name.startswith("~$"):  # sem temporários do Office
                    z.write(item, f"processos/{pasta.name}/{item.relative_to(pasta).as_posix()}")
        for nome, dados in (tabelas or {}).items():
            z.write(dados["arquivo"], manifesto["tabelas"][nome]["arquivo"])
        for nome, nome_planilha in manifesto["planilhas_junto"].items():
            z.write(planilhas_junto[nome], f"planilhas/{nome}/{nome_planilha}")
    temporario.replace(arquivo)  # só aparece em Downloads quando está completo
    return arquivo, nomes


def exportacoes(pasta: Path | None = None) -> list[Path]:
    """Os .zip exportados que estão em Downloads, do mais novo para o mais velho."""
    pasta = pasta or pasta_downloads()
    return sorted(pasta.glob(f"{PREFIXO}*.zip"), key=lambda p: p.stat().st_mtime, reverse=True)


def _membros(z: zipfile.ZipFile) -> dict[str, list[zipfile.ZipInfo]]:
    """{processo: arquivos dele no .zip}, só com caminhos seguros (processos/<nome>/...)."""
    saida: dict[str, list[zipfile.ZipInfo]] = {}
    for info in z.infolist():
        partes = PurePosixPath(info.filename.replace("\\", "/")).parts
        if (info.is_dir() or len(partes) < 3 or partes[0] != "processos"
                or any(p in ("", ".", "..") for p in partes) or INVALIDOS.search(partes[1])):
            continue
        saida.setdefault(partes[1], []).append(info)
    return saida


def ler(arquivo: Path, pasta_processos: Path, cadastradas: dict | None = None) -> dict:
    """O que tem no .zip: {"exportado_em", "computador", "processos": [{"nome", "arquivos",
    "existe"}], "planilhas", "tabelas": [{"nome", "processos", "atualizada_em", "existe"}]}.
    cadastradas: as tabelas de consulta deste computador. Lança RuntimeError se não for uma
    exportação do Edodocs."""
    try:
        with zipfile.ZipFile(arquivo) as z:
            membros = _membros(z)
            no_zip = set(z.namelist())
            try:
                manifesto = json.loads(z.read(MANIFESTO).decode("utf-8"))
            except (KeyError, ValueError):
                manifesto = {}
    except zipfile.BadZipFile:
        raise RuntimeError("o arquivo não é um .zip válido (pode ter sido baixado pela metade)")
    if not membros:
        raise RuntimeError("este arquivo não tem processos do Edodocs")
    junto = manifesto.get("planilhas_junto") or {}
    processos = [{"nome": nome, "arquivos": len(infos), "existe": (pasta_processos / nome).exists(),
                  "planilha_junto": junto[nome] if _planilha_junto(no_zip, nome, junto.get(nome)) else "",
                  "planilha_origem": str((manifesto.get("planilhas_origem") or {}).get(nome, ""))}
                 for nome, infos in sorted(membros.items(), key=lambda m: m[0].lower())]
    tabelas = [{"nome": nome, "processos": dados.get("processos") or [],
                "atualizada_em": dados.get("atualizada_em", ""), "existe": nome in (cadastradas or {})}
               for nome, dados in (manifesto.get("tabelas") or {}).items()
               if dados.get("arquivo") in no_zip and _tabela_segura(dados["arquivo"])]
    return {"exportado_em": manifesto.get("exportado_em", ""), "computador": manifesto.get("computador", ""),
            "processos": processos, "planilhas": manifesto.get("planilhas") or {}, "tabelas": tabelas,
            "de_para": {str(n): {str(k): str(v) for k, v in p.items()}
                        for n, p in (manifesto.get("de_para") or {}).items() if isinstance(p, dict)}}


def _planilha_junto(no_zip: set, processo: str, arquivo) -> bool:
    return (bool(arquivo) and isinstance(arquivo, str) and not INVALIDOS.search(arquivo)
            and arquivo not in (".", "..") and f"planilhas/{processo}/{arquivo}" in no_zip)


def _tabela_segura(caminho: str) -> bool:
    partes = PurePosixPath(caminho).parts
    return (len(partes) == 2 and partes[0] == "tabelas" and not INVALIDOS.search(partes[1])
            and partes[1] not in ("", ".", ".."))


def importar_tabelas(arquivo: Path, nomes: list[str], pasta_tabelas: Path, pasta_backup: Path) -> dict[str, dict]:
    """Tira do .zip as tabelas de consulta escolhidas para pasta_tabelas. Um Excel com o mesmo
    nome que já estava lá vai antes para pasta_backup. Devolve {nome: cadastro} ("arquivo" já
    é o caminho novo) para cadastrar no config."""
    saida = {}
    with zipfile.ZipFile(arquivo) as z:
        try:
            manifesto = json.loads(z.read(MANIFESTO).decode("utf-8"))
        except (KeyError, ValueError):
            return {}
        for nome in nomes:
            dados = (manifesto.get("tabelas") or {}).get(nome)
            if not dados or not _tabela_segura(dados.get("arquivo", "")):
                continue
            try:
                info = z.getinfo(dados["arquivo"])
            except KeyError:
                continue
            pasta_tabelas.mkdir(parents=True, exist_ok=True)
            destino = pasta_tabelas / PurePosixPath(dados["arquivo"]).name
            parcial = destino.with_name(destino.name + ".parcial")
            with z.open(info) as origem, open(parcial, "wb") as saida_arq:
                shutil.copyfileobj(origem, saida_arq)
            if destino.exists():
                pasta_backup.mkdir(parents=True, exist_ok=True)
                shutil.move(str(destino), str(pasta_backup / f"{destino.stem} (antes da importação "
                                                             f"{datetime.now():%Y-%m-%d %Hh%M%S}){destino.suffix}"))
            parcial.replace(destino)
            # A data do arquivo é a de quando os dados foram atualizados (não a da importação)
            quando = datetime(*info.date_time).timestamp()
            os.utime(destino, (quando, quando))
            saida[nome] = {k: v for k, v in dados.items() if k not in ("arquivo", "processos")} | {"arquivo": str(destino)}
    return saida


def importar(arquivo: Path, nomes: list[str], pasta_processos: Path, pasta_backup: Path,
             usar_planilha: set[str] | frozenset = frozenset()) -> list[str]:
    """Copia os processos escolhidos do .zip para pasta_processos. Se já existe um processo com
    o mesmo nome, a pasta atual vai antes para pasta_backup (nada é apagado). usar_planilha:
    os processos que passam a usar a planilha que veio junto (ela vai para a pasta do
    processo). Devolve os nomes importados."""
    importados = []
    with zipfile.ZipFile(arquivo) as z:
        membros = _membros(z)
        no_zip = set(z.namelist())
        try:
            junto = json.loads(z.read(MANIFESTO).decode("utf-8")).get("planilhas_junto") or {}
        except (KeyError, ValueError):
            junto = {}
        for nome in nomes:
            infos = membros.get(nome)
            if not infos:
                continue
            destino = pasta_processos / nome
            # Extrai numa pasta ao lado e só depois troca: se falhar no meio, o atual fica como estava
            novo = pasta_processos / f".importando {nome}"
            if novo.exists():
                shutil.rmtree(novo)
            novo.mkdir(parents=True)
            raiz = novo.resolve()
            try:
                for info in infos:
                    partes = PurePosixPath(info.filename.replace("\\", "/")).parts[2:]
                    alvo = novo.joinpath(*partes)
                    if raiz not in alvo.resolve().parents:
                        continue  # caminho estranho: fica de fora
                    alvo.parent.mkdir(parents=True, exist_ok=True)
                    with z.open(info) as origem, open(alvo, "wb") as saida:
                        shutil.copyfileobj(origem, saida)
                planilha = junto.get(nome)
                if nome in usar_planilha and _planilha_junto(no_zip, nome, planilha):
                    with z.open(f"planilhas/{nome}/{planilha}") as origem, open(novo / planilha, "wb") as saida:
                        shutil.copyfileobj(origem, saida)
                    _apontar_planilha(novo, planilha)
            except Exception:
                shutil.rmtree(novo, ignore_errors=True)  # não deixa processo pela metade
                raise
            if destino.exists():
                pasta_backup.mkdir(parents=True, exist_ok=True)
                shutil.move(str(destino), str(pasta_backup / f"{nome} (antes da importação {datetime.now():%Y-%m-%d %Hh%M%S})"))
            novo.rename(destino)
            importados.append(nome)
    return importados


def _apontar_planilha(pasta: Path, planilha: str) -> None:
    """O processo.json passa a usar a planilha que está na pasta dele."""
    arquivo = pasta / "processo.json"
    try:
        config = json.loads(arquivo.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        config = {}
    config["planilha"] = planilha
    arquivo.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
