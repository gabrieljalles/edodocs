"""Testes offline do "desfazer status" (planilha de exemplo; a real não é tocada).

  1. desfazer o e-docs (depois do Diário) apaga status_edocs e data_edocs só daquela linha;
  2. desfazer o Diário apaga status_diario, data_diario e a conferência da publicação;
  3. as células apagadas somem (sem texto vazio nem a cor do status) e as outras linhas
     e colunas continuam iguais;
  4. ato já publicado no Diário Oficial não pode ser desfeito (nada muda na planilha);
  5. cascata: com o Diário preenchido, o e-docs não pode ser desfeito (primeiro o Diário);
  6. várias linhas de uma vez: só o Diário, ou e-docs e Diário; a publicada fica como está.

Rode com:  python testes/testar_desfazer.py
"""
import json
import sys
import tempfile
import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import assistente  # noqa: E402
import main  # noqa: E402
import planilha_status  # noqa: E402

TEMP = Path(tempfile.mkdtemp())


def planilha_com_status() -> Path:
    """Exemplo com as 3 pessoas enviadas ao e-docs e ao Diário. A da linha 4 já foi achada
    publicada no Diário Oficial; as outras ainda não ("Não encontrado 1")."""
    _, planilha = assistente.gerar_exemplos(TEMP / f"p{len(list(TEMP.iterdir()))}")
    pasta = planilha.parent / "Processo"
    pasta.mkdir()
    (pasta / "processo.json").write_text(json.dumps({"aba": "atos"}), encoding="utf-8")
    main.processo = pasta
    for site in ("edocs", "diario"):
        status = main.StatusPlanilha(planilha, site)
        for linha in (2, 3, 4):
            status.marcar(linha, planilha_status.ENVIADO, "05/10/2026")
    publicado = main.StatusPlanilha(planilha, "publicado", com_data=False)
    for linha in (2, 3):
        publicado.marcar(linha, f"{planilha_status.NAO_ENCONTRADO} 1")
    publicado.marcar(4, f"{planilha_status.PUBLICADO} 7456 (09/10/2026)")
    return planilha


def ler(planilha: Path) -> dict[int, dict]:
    regs = main.proc().ler_planilha(planilha)
    return {r["_LINHA"]: {k: r.get(k, "") for k in ("NOME", "STATUSEDOCS", "DATAEDOCS", "STATUSDIARIO",
                                                       "DATADIARIO", "STATUSPUBLICADO")} for r in regs}


def teste_1_desfaz_edocs() -> None:
    planilha = planilha_com_status()
    antes = ler(planilha)
    main.desfazer_status(planilha, 3, "diario")
    main.desfazer_status(planilha, 3, "edocs")
    depois = ler(planilha)
    assert depois[3]["STATUSEDOCS"] == "" and depois[3]["DATAEDOCS"] == "", depois[3]
    assert depois[3]["STATUSDIARIO"] == "" and depois[3]["NOME"] == antes[3]["NOME"]
    assert depois[2] == antes[2] and depois[4] == antes[4]
    print("  [1] Diário e depois e-docs: apaga só daquela linha ..... OK")


def teste_2_desfaz_diario() -> None:
    planilha = planilha_com_status()
    antes = ler(planilha)
    main.desfazer_status(planilha, 2, "diario")
    depois = ler(planilha)
    assert depois[2]["STATUSDIARIO"] == depois[2]["DATADIARIO"] == depois[2]["STATUSPUBLICADO"] == ""
    assert depois[2]["STATUSEDOCS"] == planilha_status.ENVIADO and depois[2]["NOME"] == antes[2]["NOME"]
    assert depois[3] == antes[3]
    print("  [2] desfazer Diário: apaga Diário e publicação ......... OK")


def teste_3_celula_some() -> None:
    planilha = planilha_com_status()
    main.desfazer_status(planilha, 3, "diario")
    main.desfazer_status(planilha, 3, "edocs")
    status = main.StatusPlanilha(planilha, "edocs")
    col = planilha_status._letra(status.reservadas["status_edocs"])
    with zipfile.ZipFile(planilha) as z:
        xml = next(z.read(n).decode() for n in z.namelist() if n.startswith("xl/worksheets/sheet"))
    assert f'r="{col}3"' not in xml and f'r="{col}2"' in xml
    print("  [3] a célula apagada some (sem texto vazio nem cor) ..... OK")


def teste_4_publicado_travado() -> None:
    planilha = planilha_com_status()
    antes = ler(planilha)
    for etapa in ("edocs", "diario"):
        try:
            main.desfazer_status(planilha, 4, etapa)
            raise AssertionError(f"desfez o {etapa} de um ato já publicado")
        except main.AtoPublicado:
            pass
    assert ler(planilha) == antes
    assert main.publicado(antes[4]) and not main.publicado(antes[2])
    print("  [4] ato já publicado: nada pode ser desfeito ............ OK")


def teste_5_cascata() -> None:
    planilha = planilha_com_status()
    antes = ler(planilha)
    try:
        main.desfazer_status(planilha, 2, "edocs")
        raise AssertionError("desfez o e-docs com o Diário ainda preenchido")
    except main.DiarioAntes:
        pass
    assert ler(planilha) == antes  # nada mudou
    main.desfazer_status(planilha, 2, "diario")
    main.desfazer_status(planilha, 2, "edocs")  # agora pode
    assert ler(planilha)[2]["STATUSEDOCS"] == ""
    print("  [5] cascata: e-docs só depois do Diário ................ OK")


def teste_6_varias_linhas() -> None:
    planilha = planilha_com_status()
    antes = ler(planilha)
    resultado = main.desfazer_varios(planilha, [2, 3, 4], "diario")
    assert resultado == {"desfeitas": [2, 3], "publicadas": [4], "sem_status": []}, resultado
    depois = ler(planilha)
    for linha in (2, 3):
        assert depois[linha]["STATUSDIARIO"] == depois[linha]["STATUSPUBLICADO"] == "", depois[linha]
        assert depois[linha]["STATUSEDOCS"] == antes[linha]["STATUSEDOCS"]  # o e-docs fica
    assert depois[4] == antes[4]  # publicada: nada muda
    resultado = main.desfazer_varios(planilha, [2, 3], "tudo")
    assert resultado["desfeitas"] == [2, 3]
    depois = ler(planilha)
    assert all(depois[l]["STATUSEDOCS"] == depois[l]["DATAEDOCS"] == "" for l in (2, 3))
    assert main.desfazer_varios(planilha, [2], "tudo") == {"desfeitas": [], "publicadas": [], "sem_status": [2]}
    print("  [6] várias linhas de uma vez ........................... OK")


def main_() -> None:
    print("Testes do desfazer status (offline)\n")
    guardado = main.processo
    try:
        for teste in (teste_1_desfaz_edocs, teste_2_desfaz_diario, teste_3_celula_some,
                      teste_4_publicado_travado, teste_5_cascata, teste_6_varias_linhas):
            teste()
    finally:
        main.processo = guardado
    print("\n6 testes passaram.")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main_()
