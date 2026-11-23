"""Gera o Edodocs.exe (a janela do app.py) e um único pacote: pacotes/Edodocs.zip, com o
programa e todos os processos de entradas/processos.

    python build.py                  gera o .exe e o Edodocs.zip
    python build.py --sem-planilhas  sem as planilhas (cada pessoa escolhe a sua em
                                     Configurações ou ao editar o processo)
    python build.py --so-pacote      não gera o .exe de novo (usa o de dist/)

Na outra máquina: descompactar e abrir Edodocs/Edodocs.exe (não precisa instalar Python).
"""
import argparse
import ctypes
import json
import os
import subprocess
import sys
import zipfile
from pathlib import Path

BASE = Path(__file__).parent
NOME = "Edodocs"
DIST = BASE / "dist" / NOME          # pasta do programa (--onedir: abre bem mais rápido)
EXE = DIST / f"{NOME}.exe"
PROCESSOS = BASE / "entradas" / "processos"
PACOTES = BASE / "pacotes"

# Módulos do projeto (os importados dentro de funções também, por garantia)
MODULOS = ["main", "ui", "atos", "processos", "planilha_status", "planilha_edicao", "credenciais",
           "assistente", "versao", "regras_ui", "processos_backup", "colunas_ui", "marcar_ui", "planilha_nova", "simulacao_ui", "escolha_pessoas", "modelo_automatico", "edocs_http", "diario_http",
           "diario", "ajudas", "janela_site", "abrir", "folha", "atalho", "ia", "recebidos_ui", "consultas", "consultas_ui", "de_para", "de_para_ui", "trava",
           "ato_manual", "ato_manual_ui", "processo_ia", "processo_ia_ui"]

PYINSTALLER = [
    "--noconfirm", "--clean", "--onedir", "--windowed", "--name", NOME,
    "--icon", str(BASE / "assets" / "icone_light.ico"),
    "--add-data", f"{BASE / 'assets'}{os.pathsep}assets",
    "--specpath", str(BASE / "build"),
    "--collect-all", "nicegui",          # páginas, scripts e estilos da janela
    "--collect-all", "webview",          # janela nativa (pywebview + Edge WebView2)
    "--collect-all", "curl_cffi",        # acesso ao e-docs e ao Diário (com a libcurl)
    "--collect-all", "pypdfium2", "--collect-all", "pypdfium2_raw",  # leitura dos PDFs
    "--collect-data", "docx",            # modelos internos do python-docx
    "--collect-submodules", "anthropic",  # IA dos resumos (página Recebidos)
    "--hidden-import", "clr",            # pythonnet, usado pela janela nativa
    "--hidden-import", "win32com.shell.shell", "--hidden-import", "pythoncom",  # atalho na área de trabalho
    *[arg for m in MODULOS for arg in ("--hidden-import", m)],
    "app.py",
]

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def etapa(texto: str) -> None:
    print(f"\n== {texto} ==", flush=True)


def falhar(texto: str) -> None:
    print(f"\nERRO: {texto}")
    sys.exit(1)


def rodar(*comando) -> None:
    if subprocess.run(comando, cwd=BASE).returncode != 0:
        falhar("veja as mensagens acima.")


def todos_os_processos() -> list[Path]:
    """Pastas de entradas/processos que têm um Word modelo."""
    existentes = sorted((p for p in PROCESSOS.iterdir() if p.is_dir() and _arquivos(p, "*.docx")),
                        key=lambda p: p.name.lower()) if PROCESSOS.is_dir() else []
    if not existentes:
        falhar(f"nenhum processo em {PROCESSOS.relative_to(BASE)} (pasta com o Word e a planilha)")
    return existentes


def _arquivos(pasta: Path, *padroes: str) -> list[Path]:
    """Arquivos da pasta, sem os temporários do Office ("~$")."""
    return sorted(p for padrao in padroes for p in pasta.glob(padrao)
                  if not p.name.startswith("~$"))


def gerar_executavel() -> None:
    etapa("1/3 Instalando dependências")
    rodar(sys.executable, "-m", "pip", "install", "--upgrade", "-r", "requirements.txt")

    manual = BASE / "assets" / "manual"
    if not any(a.suffix.lower() in (".mp3", ".m4a", ".wav", ".ogg") for a in manual.glob("*")):
        print("AVISO: o manual vai sem áudio (coloque um .mp3 ou .m4a em assets/manual/)")

    etapa("2/3 Gerando o Edodocs.exe")
    # Versão que aparece no canto da tela de boas-vindas (data e código do último commit)
    import versao
    texto = versao.do_git(BASE) or "sem versão"
    (BASE / "assets" / "versao.txt").write_text(texto, encoding="utf-8")
    print(f"Versão: {texto}", flush=True)
    rodar(sys.executable, "-m", "PyInstaller", *PYINSTALLER)
    if not EXE.exists():
        falhar(f"{EXE.relative_to(BASE)} não foi gerado.")


def _na_rede(caminho: Path) -> bool:
    """Caminho de rede (\\\\servidor\\... ou unidade mapeada): outras máquinas também acessam."""
    texto = str(caminho)
    if texto.startswith("\\\\"):
        return True
    unidade = os.path.splitdrive(texto)[0]
    return bool(unidade) and ctypes.windll.kernel32.GetDriveTypeW(unidade + "\\") == 4  # DRIVE_REMOTE


def _config(pasta: Path) -> dict:
    try:
        return json.loads((pasta / "processo.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def arquivos_do_processo(pasta: Path, com_planilhas: bool) -> tuple[dict[str, Path], dict, list[str]]:
    """(arquivos {nome no pacote: origem}, processo.json do pacote, avisos).
    Planilha guardada fora da pasta, neste computador (ex.: Downloads), vai junto para dentro
    da pasta do processo; planilha na rede continua apontando para a rede."""
    config = _config(pasta)
    avisos = []
    arquivos = {a.name: a for a in _arquivos(pasta, "*.docx")}
    planilhas = {a.name: a for a in _arquivos(pasta, "*.xlsx")}
    externa = config.get("planilha", "")
    if externa and Path(externa).is_absolute():
        caminho = Path(externa)
        if _na_rede(caminho):
            avisos.append(f"planilha na rede ({caminho}): quem receber precisa ter acesso a ela")
        elif caminho.is_file():
            planilhas = {caminho.name: caminho}
            config["planilha"] = caminho.name  # no pacote, fica dentro da pasta do processo
        else:
            avisos.append(f"planilha não encontrada ({caminho}): escolha de novo no outro computador")
            config["planilha"] = ""
    if com_planilhas:
        arquivos |= planilhas
    else:
        if not externa or not _na_rede(Path(externa)):
            config["planilha"] = ""
        avisos.append("vai sem planilha: a pessoa escolhe a dela em Configurações")
    if not planilhas and com_planilhas and not (externa and _na_rede(Path(externa))):
        avisos.append("vai sem planilha (nenhuma .xlsx na pasta)")
    if not (pasta / "processo.json").exists():
        avisos.append("sem processo.json: ao abrir, configure o processo em Processos → Editar")
    return arquivos, config, avisos


def levar_tabelas_e_de_para(z: zipfile.ZipFile, raiz: Path, processos: list[Path]) -> None:
    """As tabelas de consulta que os processos usam (o Excel e o cadastro) e os de-para, do
    configs/ deste computador: no outro, o Preencher sozinho já abre funcionando. Vão em
    entradas/, que o programa lê na primeira abertura (consultas.receber_do_pacote)."""
    config_aqui = BASE / "configs" / "config.json"
    try:
        config = json.loads(config_aqui.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return
    cadastradas = config.get("tabelas_consulta") or {}
    levar = {"tabelas": {}, "de_para": config.get("de_para") or {}}
    for pasta in processos:
        nome = (_config(pasta).get("consulta") or {}).get("tabela")
        dados = cadastradas.get(nome) if nome else None
        if not nome or nome in levar["tabelas"]:
            continue
        if not dados or not Path(dados.get("arquivo") or "").is_file():
            print(f"  AVISO {pasta.name}: a tabela {nome} não está cadastrada aqui (ou o arquivo sumiu): "
                  "não vai junto")
            continue
        arquivo = Path(dados["arquivo"])
        z.write(arquivo, raiz / "entradas" / "tabelas de consulta" / arquivo.name)
        levar["tabelas"][nome] = dict(dados, arquivo=arquivo.name)
        print(f"  Tabela {nome} vai junto ({arquivo.name})")
    for nome in levar["de_para"]:
        print(f"  De-para {nome} vai junto")
    if levar["tabelas"] or levar["de_para"]:
        z.writestr(str(raiz / "entradas" / "veio_no_pacote.json"), json.dumps(levar, ensure_ascii=False, indent=2))


def montar_pacote(processos: list[Path], com_planilhas: bool, com_tabelas: bool = True) -> Path:
    """Zip com a pasta Edodocs/ (programa + entradas). configs/ não vai: a senha salva só
    funciona no computador e usuário do Windows onde foi cadastrada (no outro computador
    o programa pede usuário e senha de novo)."""
    PACOTES.mkdir(exist_ok=True)
    destino = PACOTES / f"{NOME}.zip"
    destino.unlink(missing_ok=True)
    # Pacotes de versões anteriores do build (um por processo): não valem mais
    for antigo in PACOTES.glob(f"{NOME} - *.zip"):
        antigo.unlink()
    raiz = Path(NOME)

    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
        # Só o programa: se o .exe de dist/ foi aberto para teste, ele criou configs/ (com
        # senha), entradas/ e saidas/ ao lado, que não podem ir para outras pessoas
        z.write(EXE, raiz / EXE.name)
        for arq in (DIST / "_internal").rglob("*"):
            if arq.is_file():
                z.write(arq, raiz / arq.relative_to(DIST))
        for pasta in processos:
            arquivos, config, avisos = arquivos_do_processo(pasta, com_planilhas)
            for nome_arq, origem in arquivos.items():
                z.write(origem, raiz / "entradas" / "processos" / pasta.name / nome_arq)
            if config:
                z.writestr(str(raiz / "entradas" / "processos" / pasta.name / "processo.json"),
                           json.dumps(config, ensure_ascii=False, indent=2))
            for aviso in avisos:
                print(f"  AVISO {pasta.name}: {aviso}")
        if com_tabelas:
            levar_tabelas_e_de_para(z, raiz, processos)
        # Pastas vazias que o programa usa
        for vazia in ("entradas/Diario_oficial/", "entradas/Diario_hoje/", "saidas/", "configs/"):
            z.writestr(f"{NOME}/{vazia}", "")
    return destino


def main() -> None:
    args = argparse.ArgumentParser(description=__doc__,
                                   formatter_class=argparse.RawDescriptionHelpFormatter)
    args.add_argument("--sem-planilhas", action="store_true",
                      help="não leva as planilhas (cada pessoa escolhe a sua)")
    args.add_argument("--sem-tabelas", action="store_true",
                      help="não leva as tabelas de consulta nem os de-para")
    args.add_argument("--so-pacote", action="store_true",
                      help="não gera o .exe de novo; só monta o pacote com o de dist/")
    opcoes = args.parse_args()

    processos = todos_os_processos()
    if opcoes.so_pacote:
        if not EXE.exists():
            falhar(f"{EXE.relative_to(BASE)} não existe: rode sem --so-pacote")
    else:
        gerar_executavel()

    etapa("3/3 Montando o Edodocs.zip")
    print("Processos: " + ", ".join(p.name for p in processos), flush=True)
    pacote = montar_pacote(processos, not opcoes.sem_planilhas, not opcoes.sem_tabelas)
    print(f"\nPronto! {pacote.relative_to(BASE)} ({pacote.stat().st_size / 1e6:.0f} MB)")
    print("Na outra máquina: descompacte e abra Edodocs\\Edodocs.exe")
    print(f"Para testar aqui antes de enviar: {EXE.relative_to(BASE)}")


if __name__ == "__main__":
    main()
