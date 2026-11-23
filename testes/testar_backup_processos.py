"""Testes offline de exportar/importar processos (processos_backup.py).

  1. exportar junta todos os processos num .zip (sem temporários do Office) com o manifesto;
  2. ler mostra os processos do .zip e quais já existem aqui;
  3. importar copia os escolhidos; o que já existia vai para o backup (nada é apagado);
  4. .zip com caminho malicioso (../) não escreve fora da pasta; .zip quebrado dá erro claro;
  5. a pasta Downloads do Windows é encontrada;
  6. as tabelas de consulta vão junto (quando escolhidas) e voltam cadastráveis, com a data dos dados.

Rode com:  python testes/testar_backup_processos.py
"""
import json
import sys
import tempfile
import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import processos_backup as backup  # noqa: E402

TEMP = Path(tempfile.mkdtemp())
ORIGEM = TEMP / "antigo" / "processos"
DESTINO = TEMP / "novo" / "processos"
BACKUP = TEMP / "novo" / "backup"
DOWNLOADS = TEMP / "Downloads"


def criar(pasta: Path, arquivos: dict[str, str]) -> None:
    for nome, conteudo in arquivos.items():
        (pasta / nome).parent.mkdir(parents=True, exist_ok=True)
        (pasta / nome).write_text(conteudo, encoding="utf-8")


def teste_1_exportar() -> Path:
    criar(ORIGEM / "Exoneração", {"uma.docx": "w1", "várias.docx": "w2", "processo.json": "{}",
                                  "~$uma.docx": "temporário", "dados.xlsx": "planilha"})
    criar(ORIGEM / "Licença Prêmio", {"modelo.docx": "w", "processo.json": "{}"})
    criar(ORIGEM / "Sem Word", {"processo.json": "{}"})  # não é processo
    DOWNLOADS.mkdir(parents=True)
    arquivo, nomes = backup.exportar(ORIGEM, {"Licença Prêmio": r"S:\rede\licenca.xlsx", "Outro": "x"},
                                     DOWNLOADS)
    assert nomes == ["Exoneração", "Licença Prêmio"], nomes
    assert arquivo.parent == DOWNLOADS and arquivo.name.startswith(backup.PREFIXO)
    with zipfile.ZipFile(arquivo) as z:
        nomes_zip = set(z.namelist())
        manifesto = json.loads(z.read(backup.MANIFESTO))
    assert "processos/Exoneração/várias.docx" in nomes_zip and "processos/Exoneração/dados.xlsx" in nomes_zip
    assert not any("~$" in n for n in nomes_zip)
    assert manifesto["planilhas"] == {"Licença Prêmio": r"S:\rede\licenca.xlsx"}
    assert backup.exportacoes(DOWNLOADS) == [arquivo]
    assert not list(DOWNLOADS.glob("*.parcial"))
    print("  [1] exportar junta tudo num .zip em Downloads .......... OK")
    return arquivo


def teste_2_ler(arquivo: Path) -> None:
    criar(DESTINO / "Licença Prêmio", {"modelo.docx": "versão nova", "processo.json": "{}"})
    lido = backup.ler(arquivo, DESTINO)
    assert [(p["nome"], p["existe"]) for p in lido["processos"]] == [("Exoneração", False), ("Licença Prêmio", True)]
    assert lido["processos"][0]["arquivos"] == 4
    print("  [2] ler mostra os processos e quais já existem ......... OK")


def teste_3_importar(arquivo: Path) -> None:
    importados = backup.importar(arquivo, ["Exoneração", "Licença Prêmio"], DESTINO, BACKUP)
    assert importados == ["Exoneração", "Licença Prêmio"]
    assert (DESTINO / "Exoneração" / "várias.docx").read_text(encoding="utf-8") == "w2"
    assert (DESTINO / "Licença Prêmio" / "modelo.docx").read_text(encoding="utf-8") == "w"
    guardados = list(BACKUP.iterdir())
    assert len(guardados) == 1 and (guardados[0] / "modelo.docx").read_text(encoding="utf-8") == "versão nova"
    assert not list(DESTINO.glob(".importando*"))
    print("  [3] importar copia e guarda o que já existia ........... OK")


def teste_4_zip_estranho() -> None:
    ruim = TEMP / "ruim.zip"
    with zipfile.ZipFile(ruim, "w") as z:
        z.writestr("processos/X/../../../fora.txt", "ataque")
        z.writestr("processos/../fora2.txt", "ataque")
        z.writestr("processos/Bom/modelo.docx", "ok")
    assert [p["nome"] for p in backup.ler(ruim, DESTINO)["processos"]] == ["Bom"]
    backup.importar(ruim, ["Bom", "X"], DESTINO, BACKUP)
    assert not list(TEMP.rglob("fora*.txt"))
    quebrado = TEMP / "quebrado.zip"
    quebrado.write_bytes(b"isto nao e zip")
    try:
        backup.ler(quebrado, DESTINO)
        raise AssertionError("devia dar erro")
    except RuntimeError as e:
        assert "zip" in str(e)
    print("  [4] .zip malicioso ou quebrado não estraga nada ........ OK")


def teste_5_downloads() -> None:
    pasta = backup.pasta_downloads()
    assert pasta.is_dir(), pasta
    print(f"  [5] pasta Downloads encontrada ({pasta}) ... OK")


def teste_6_tabelas_de_consulta() -> None:
    import os
    import time
    origem = TEMP / "t6" / "processos"
    criar(origem / "Exoneração", {"uma.docx": "w", "processo.json": json.dumps(
        {"consulta": {"tabela": "Servidores", "chaves": [], "puxar": []}})})
    criar(origem / "Férias", {"modelo.docx": "w", "processo.json": "{}"})
    tabela = TEMP / "t6" / "rede" / "servidores atual.xlsx"
    criar(tabela.parent, {tabela.name: "dados das pessoas"})
    dez_dias = time.time() - 10 * 86400
    os.utime(tabela, (dez_dias, dez_dias))
    cadastradas = {"Servidores": {"arquivo": str(tabela), "aba": "Folha", "atualizar_dias": 7,
                                  "atualizada_em": "2026-10-10"},
                   "Sem arquivo": {"arquivo": str(TEMP / "nao existe.xlsx")}}
    assert backup.tabelas_usadas(origem, cadastradas) == {"Servidores": ["Exoneração"]}

    # Sem as tabelas: o .zip não leva
    sem, _ = backup.exportar(origem, {}, TEMP / "t6")
    assert backup.ler(sem, TEMP / "t6" / "destino")["tabelas"] == []
    time.sleep(1.1)  # outro nome de .zip

    com, _ = backup.exportar(origem, {}, TEMP / "t6", {"Servidores": cadastradas["Servidores"]})
    lido = backup.ler(com, TEMP / "t6" / "destino", {})
    assert lido["tabelas"] == [{"nome": "Servidores", "processos": ["Exoneração"], "atualizada_em": "2026-10-10",
                                "existe": False}], lido["tabelas"]
    assert backup.ler(com, TEMP / "t6" / "destino", {"Servidores": {}})["tabelas"][0]["existe"]

    pasta = TEMP / "t6" / "tabelas de consulta"
    criar(pasta, {"Servidores.xlsx": "a antiga daqui"})
    trazidas = backup.importar_tabelas(com, ["Servidores", "Outra"], pasta, TEMP / "t6" / "backup")
    destino = pasta / "Servidores.xlsx"
    assert trazidas == {"Servidores": {"aba": "Folha", "atualizar_dias": 7, "atualizada_em": "2026-10-10",
                                       "arquivo": str(destino)}}, trazidas
    assert destino.read_text(encoding="utf-8") == "dados das pessoas"
    assert abs(destino.stat().st_mtime - dez_dias) < 3, "a data do arquivo é a dos dados, não a da importação"
    guardadas = list((TEMP / "t6" / "backup").iterdir())
    assert len(guardadas) == 1 and guardadas[0].read_text(encoding="utf-8") == "a antiga daqui"
    assert not list(pasta.glob("*.parcial"))

    # Manifesto mexido apontando para fora: fica de fora
    ruim = TEMP / "t6" / "ruim.zip"
    with zipfile.ZipFile(ruim, "w") as z:
        z.writestr(backup.MANIFESTO, json.dumps({"tabelas": {"X": {"arquivo": "../fora.xlsx"}}}))
        z.writestr("../fora.xlsx", "ataque")
        z.writestr("processos/Bom/modelo.docx", "ok")
    assert backup.ler(ruim, DESTINO)["tabelas"] == []
    assert backup.importar_tabelas(ruim, ["X"], pasta, TEMP / "t6" / "backup") == {}
    print("  [6] tabelas de consulta vão e voltam com o processo .... OK")


def teste_7_planilha_de_fora() -> None:
    origem = TEMP / "t7" / "processos"
    criar(origem / "Nomeação", {"modelo.docx": "w", "processo.json": json.dumps(
        {"planilha": r"S:\\rede\\nomeacao.xlsx", "modelo": "modelo.docx"})})
    rede = TEMP / "t7" / "rede" / "nomeação 2026.xlsx"
    criar(rede.parent, {rede.name: "linhas da planilha"})
    arquivo, _ = backup.exportar(origem, {}, TEMP / "t7", planilhas_junto={"Nomeação": rede, "Outro": rede})
    assert isinstance(arquivo, Path) and arquivo.suffix == ".zip", arquivo
    lido = backup.ler(arquivo, TEMP / "t7" / "destino")
    assert lido["processos"][0]["planilha_junto"] == "nomeação 2026.xlsx", lido["processos"]

    # Sem escolher: vem só o processo, como antes
    backup.importar(arquivo, ["Nomeação"], TEMP / "t7" / "a", TEMP / "t7" / "bk")
    assert not (TEMP / "t7" / "a" / "Nomeação" / "nomeação 2026.xlsx").exists()
    # Usando a que veio: ela vai para a pasta e o processo aponta para ela
    backup.importar(arquivo, ["Nomeação"], TEMP / "t7" / "b", TEMP / "t7" / "bk", {"Nomeação"})
    pasta = TEMP / "t7" / "b" / "Nomeação"
    assert (pasta / "nomeação 2026.xlsx").read_text(encoding="utf-8") == "linhas da planilha"
    config = json.loads((pasta / "processo.json").read_text(encoding="utf-8"))
    assert config["planilha"] == "nomeação 2026.xlsx" and config["modelo"] == "modelo.docx", config
    print("  [7] planilha de fora da pasta vai junto (se escolher) .. OK")


if __name__ == "__main__":
    arquivo = teste_1_exportar()
    teste_2_ler(arquivo)
    teste_3_importar(arquivo)
    teste_4_zip_estranho()
    teste_5_downloads()
    teste_6_tabelas_de_consulta()
    teste_7_planilha_de_fora()
    print("\n7 testes passaram.")
