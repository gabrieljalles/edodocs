"""Pasta compartilhada: processos, tabelas e de-para na rede; cópia antes de mudar; um
computador grava a planilha por vez.

    python testes/testar_pasta_compartilhada.py
"""
import json
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import openpyxl  # noqa: E402

import main  # noqa: E402
import planilha_edicao as edicao  # noqa: E402
import planilha_status  # noqa: E402
import processos  # noqa: E402
import trava  # noqa: E402

TMP = Path(tempfile.mkdtemp(prefix="edodocs_compartilhada_"))
LOCAL = TMP / "este computador"
REDE = TMP / "servidor" / "processos"  # a pessoa escolhe a pasta dos processos
REDE.mkdir(parents=True)
EXTRAS = REDE / "_configurações do Edodocs"
main.CONFIG = LOCAL / "configs" / "config.json"
main.PROCESSOS_DESTE_COMPUTADOR = LOCAL / "entradas" / "processos"
main.ENTRADAS = LOCAL / "entradas"
main.aplicar_pasta_compartilhada()

passou = 0


def ok(texto: str) -> None:
    global passou
    passou += 1
    print(f"  [{passou}] {texto} " + "." * max(2, 55 - len(texto)) + " OK")


def processo_local(nome: str) -> Path:
    pasta = main.PROCESSOS_DESTE_COMPUTADOR / nome
    pasta.mkdir(parents=True)
    import docx
    docx.Document().save(pasta / f"{nome}.docx")
    wb = openpyxl.Workbook()
    wb.active.title = "dados"
    wb.active.append(["NOME", "CARGO"])
    wb.save(pasta / f"{nome} - planilha.xlsx")
    processos.Processo(pasta, {"modelo": f"{nome}.docx", "planilha": f"{nome} - planilha.xlsx"}).salvar()
    return pasta


try:
    # 1. Sem pasta compartilhada: tudo como antes
    processo_local("Exoneração")
    tabela = TMP / "Downloads" / "pessoas.xlsx"
    tabela.parent.mkdir()
    openpyxl.Workbook().save(tabela)
    main.salvar_config(tabelas_consulta={"Pessoas": {"arquivo": str(tabela), "aba": ""}},
                       de_para={"Secretarias": {"1": "Secretaria A"}}, tema="escuro")
    assert main.PROCESSOS == main.PROCESSOS_DESTE_COMPUTADOR
    assert [p.name for p in main.listar_processos()] == ["Exoneração"]
    ok("sem pasta compartilhada, tudo neste computador")

    # 2. Levar para a pasta compartilhada
    levados = main.levar_para_compartilhada(REDE)
    main.salvar_config(pasta_compartilhada=str(REDE))
    main.aplicar_pasta_compartilhada()
    assert main.PROCESSOS == REDE and (REDE / "Exoneração" / "processo.json").is_file()
    assert [p.name for p in main.listar_processos()] == ["Exoneração"]
    assert "tabela Pessoas" in levados and "de-para Secretarias" in levados
    cfg = main.ler_config()
    assert Path(cfg["tabelas_consulta"]["Pessoas"]["arquivo"]).parent == EXTRAS / "tabelas de consulta"
    assert cfg["de_para"] == {"Secretarias": {"1": "Secretaria A"}} and cfg["tema"] == "escuro"
    ok("processos, tabela e de-para levados para a rede")

    # 3. O que é de todos vai para a rede; o que é do computador fica nele
    main.salvar_config(de_para={"Secretarias": {"1": "Secretaria A", "2": "Secretaria B"}}, tema="claro")
    comum = json.loads((EXTRAS / main.ARQUIVO_COMPARTILHADO).read_text(encoding="utf-8"))
    local = json.loads(main.CONFIG.read_text(encoding="utf-8"))
    assert comum["de_para"]["Secretarias"]["2"] == "Secretaria B" and "tema" not in comum
    assert local["tema"] == "claro"
    ok("de-para na rede, tema neste computador")

    # 4. Outro computador (outro config local) enxerga o mesmo
    outro = TMP / "outro computador" / "config.json"
    outro.parent.mkdir(parents=True)
    outro.write_text(json.dumps({"pasta_compartilhada": str(REDE), "tema": "escuro"}), encoding="utf-8")
    meu = main.CONFIG
    main.CONFIG = outro
    assert main.ler_config()["de_para"]["Secretarias"]["2"] == "Secretaria B"
    assert main.ler_config()["tema"] == "escuro"
    main.CONFIG = meu
    ok("outro computador vê os mesmos de-para")

    # 5. Tabela nova escolhida em Downloads é copiada para a rede
    import consultas
    nova = TMP / "Downloads" / "servidores.xlsx"
    openpyxl.Workbook().save(nova)
    consultas.salvar_tabela("Servidores", nova)
    assert Path(consultas.tabelas()["Servidores"]["arquivo"]) == EXTRAS / "tabelas de consulta" / "servidores.xlsx"
    ok("tabela nova vai para a pasta compartilhada")

    # 6. Cópia antes de mudar o processo (uma por sessão de edição)
    pasta = main.PROCESSOS / "Exoneração"
    proc = processos.carregar(pasta)
    proc.config["titulo_edocs"] = "NOVO"
    proc.salvar()
    proc.salvar()
    versoes = list((EXTRAS / main.VERSOES / "Exoneração").iterdir())
    assert len(versoes) == 1
    assert "NOVO" not in (versoes[0] / processos.ARQUIVO).read_text(encoding="utf-8")
    assert (versoes[0] / "Exoneração.docx").is_file()
    assert [p.name for p in main.listar_processos()] == ["Exoneração"]  # a pasta de configurações não é processo
    ok("cópia de antes em 'versões anteriores'")

    # 7. Trava: um grava por vez; trava esquecida é ignorada
    alvo = TMP / "arquivo.txt"
    import threading
    resultado = []

    def outro_tenta() -> None:
        try:
            with trava.trava(alvo, espera=0.5):
                resultado.append("entrou")
        except trava.Ocupado:
            resultado.append("esperou")

    with trava.trava(alvo):
        assert trava.quem_esta(alvo)
        with trava.trava(alvo):  # quem já tem pode pedir de novo
            pass
        t = threading.Thread(target=outro_tenta)
        t.start()
        t.join()
    assert resultado == ["esperou"], resultado
    assert trava.quem_esta(alvo) is None
    velha = trava.arquivo_da_trava(alvo)
    velha.write_text("programa que fechou")
    import os
    os.utime(velha, (time.time() - 3600, time.time() - 3600))
    with trava.trava(alvo, espera=0.5):
        pass
    ok("trava: espera a vez e ignora trava esquecida")

    # 8. Dois computadores gravando a mesma planilha ao mesmo tempo: nenhuma linha se perde
    planilha = pasta / "Exoneração - planilha.xlsx"
    codigo = (f"import sys; sys.path.insert(0, {str(RAIZ)!r}); from pathlib import Path; "
              "import planilha_edicao as ed; p = Path(sys.argv[1]); c = ed.esquema(p, 'dados'); "
              "[ed.incluir_linhas(p, 'dados', c, [{c[0].indice: f'PESSOA {sys.argv[2]}-{i}', c[1].indice: 'X'}], "
              "recalcular_ao_abrir=False) for i in range(6)]")
    outros = [subprocess.Popen([sys.executable, "-c", codigo, str(planilha), str(n)]) for n in range(3)]
    for p in outros:
        p.wait(120)
    nomes = [r[0] for r in openpyxl.load_workbook(planilha).active.iter_rows(min_row=2, values_only=True) if r[0]]
    assert len(nomes) == 18, (len(nomes), nomes)
    assert not list(pasta.glob("~edodocs*"))
    ok("3 computadores, 18 linhas gravadas, nenhuma perdida")

    # 9. Status também espera a vez
    cols = planilha_status.colunas(planilha, "dados", ["status_edocs"])
    planilha_status.gravar(planilha, "dados", cols, {"status_edocs": {2: "Enviado"}})
    assert openpyxl.load_workbook(planilha).active.cell(2, cols["status_edocs"]).value == "Enviado"
    ok("status do envio grava com a trava")

    # 10. Quem escolheu a pasta de cima (com "processos" dentro) também acha tudo
    main.salvar_config(pasta_compartilhada=str(REDE.parent))
    main.aplicar_pasta_compartilhada()
    assert main.PROCESSOS == REDE and main.extras_compartilhados() == EXTRAS
    assert main.ler_config()["de_para"]["Secretarias"]["2"] == "Secretaria B"
    main.salvar_config(pasta_compartilhada=str(REDE))
    main.aplicar_pasta_compartilhada()
    ok("pasta de cima, com 'processos' dentro, também serve")

    # 11. Voltar a usar só este computador
    main.salvar_config(pasta_compartilhada="")
    main.aplicar_pasta_compartilhada()
    assert main.PROCESSOS == main.PROCESSOS_DESTE_COMPUTADOR
    assert main.ler_config()["de_para"] == {"Secretarias": {"1": "Secretaria A"}}
    ok("voltar para este computador")

    # 12. Pasta da rede fora do ar: aviso, sem quebrar
    main.salvar_config(pasta_compartilhada=str(TMP / "não existe"))
    main.aplicar_pasta_compartilhada()
    assert main.problema_na_compartilhada() and main.listar_processos() == []
    ok("pasta da rede fora do ar: avisa")
finally:
    shutil.rmtree(TMP, ignore_errors=True)

print(f"\n{passou} testes passaram.")
