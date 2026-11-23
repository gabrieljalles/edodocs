"""Testes offline do de-para (de_para.py) e do uso no Preencher sozinho (consultas.py).

  1. códigos comparados sem pontos, traços e zeros à esquerda; sem descrição fica o código;
  2. colar do Excel e pegar de uma tabela;
  3. sugestão do de-para pela cobertura e cara de código da coluna;
  4. completar: código → descrição → opção da lista (tradução); a busca mostra a descrição;
  5. juntar o que veio de fora sem trocar a descrição daqui;
  6. só uma parte do valor é o código (ex.: os 2 primeiros dígitos).

Rode com:  python testes/testar_de_para.py
"""
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import consultas  # noqa: E402
import de_para  # noqa: E402

LISTAS = {"Secretarias": {"0312": "Secretaria Municipal de Governo", "04.01": "Secretaria Municipal de Saúde",
                          "0500": ""},
          "Vínculos": {"1": "EFETIVO", "2": "COMISSIONADO"}}
salvo = {}
de_para.todos = lambda: dict(LISTAS)  # sem mexer no config de verdade
de_para.salvar = lambda nome, pares, antigo=None: salvo.update({nome: pares})


def teste_1_traduzir() -> None:
    assert de_para.traduzir("Secretarias", "312") == "Secretaria Municipal de Governo"
    assert de_para.traduzir("Secretarias", "0401") == "Secretaria Municipal de Saúde"
    assert de_para.traduzir("Secretarias", "0500") == "0500", "sem descrição: fica o código"
    assert de_para.traduzir("Secretarias", "9999") == "9999"
    assert de_para.traduzir("Não existe", "0312") == "0312"
    print("  [1] código vira descrição (03.12 = 312) ............... OK")


def teste_2_trazer() -> None:
    assert de_para.de_texto("0312\tGoverno\n0401\tSaúde\n0500\n") == {"0312": "Governo", "0401": "Saúde", "0500": ""}
    assert de_para.de_texto("A;Alfa\nB;Beta") == {"A": "Alfa", "B": "Beta"}
    tabela = consultas.Tabela(["COD", "DESC"], [{"COD": "1", "DESC": "Um"}, {"COD": "2"}, {"COD": "1", "DESC": "Um"}])
    assert de_para.de_tabela(tabela, "COD", "DESC") == {"1": "Um", "2": ""}
    assert de_para.de_tabela(tabela, "COD") == {"1": "", "2": ""}
    print("  [2] colar do Excel e pegar de uma tabela .............. OK")


def teste_3_sugerir() -> None:
    assert de_para.sugerir(["0312", "401", "0500"]) == "Secretarias"
    assert de_para.sugerir(["0312", "7777", "8888"]) is None
    assert de_para.parece_codigo(["0312", "0401"] * 3, 30)
    assert not de_para.parece_codigo(["Secretaria de Governo", "Secretaria de Saúde"], 30)
    assert not de_para.parece_codigo([f"38.{i:03d}-0" for i in range(30)], 30), "matrícula: uma por pessoa"
    assert not de_para.parece_codigo(["02/01/2025", "10/03/2010"], 30), "data não é código"
    print("  [3] sugere o de-para que conhece os códigos ........... OK")


def teste_4_completar() -> None:
    tabela = consultas.Tabela(["NOME", "LOT", "VINC"], [
        {"NOME": "ANA LIMA", "LOT": "0312", "VINC": "2"},
        {"NOME": "BETO REIS", "LOT": "0401", "VINC": "1"}])
    item = consultas.item_puxar("SECRETARIA", "LOT", tabela)
    assert item["de_para"] == "Secretarias", item
    vinculo = consultas.item_puxar("VÍNCULO", "VINC", tabela, ["EFETIVO", "COMISSÃO"])
    assert vinculo["de_para"] == "Vínculos" and vinculo["traducao"] == {"COMISSIONADO": "COMISSÃO"}, vinculo
    sem = consultas.item_puxar("SECRETARIA", "LOT", tabela, None, None)
    assert "de_para" not in sem
    consulta = {"chaves": [{"minha": "NOME", "dela": "NOME"}], "puxar": [item, vinculo]}
    assert consultas.completar(consulta, tabela.linhas[0]) == {
        "NOME": "ANA LIMA", "SECRETARIA": "Secretaria Municipal de Governo", "VÍNCULO": "COMISSÃO"}
    assert "Secretaria Municipal de Governo" in consultas.descricao(consulta, tabela.linhas[0])
    print("  [4] código → descrição → opção da lista ............... OK")


def teste_5_juntar() -> None:
    novos = de_para.juntar("Secretarias", {"312": "OUTRO NOME", "0500": "Secretaria de Obras", "0600": "Cultura"})
    assert novos == 2, novos
    assert salvo["Secretarias"]["0312"] == "Secretaria Municipal de Governo", "a daqui fica"
    assert salvo["Secretarias"]["0500"] == "Secretaria de Obras" and salvo["Secretarias"]["0600"] == "Cultura"
    print("  [5] juntar o que veio de fora sem trocar o daqui ...... OK")


def teste_6_so_uma_parte() -> None:
    dois = {"de": "inicio", "n": 2}
    assert de_para.parte("03.12.001", dois) == "03"
    assert de_para.parte("0312", {"de": "fim", "n": 2}) == "12"
    assert de_para.parte("0312", None) == "0312" and de_para.parte("", dois) == ""
    assert de_para.descrever_parte(dois) == "2 primeiros" and de_para.descrever_parte(None) == ""
    LISTAS["Secretarias por 2"] = {"03": "Secretaria Municipal de Governo", "04": "Secretaria Municipal de Saúde"}
    assert de_para.traduzir("Secretarias por 2", "0312", None, dois) == "Secretaria Municipal de Governo"
    assert de_para.traduzir("Secretarias por 2", "04.55", None, dois) == "Secretaria Municipal de Saúde"
    assert de_para.traduzir("Secretarias por 2", "9912", None, dois) == "9912", "sem o código: o valor inteiro"
    assert de_para.cobertura(LISTAS["Secretarias por 2"], ["0312", "0399", "0401"], dois) == 1.0
    tabela = consultas.Tabela(["NOME", "LOT"], [{"NOME": "ANA", "LOT": "0312"}, {"NOME": "BIA", "LOT": "0455"},
                                                {"NOME": "CAU", "LOT": "0399"}])
    assert de_para.de_tabela(tabela, "LOT", None, dois) == {"03": "", "04": ""}
    item = consultas.item_puxar("SECRETARIA", "LOT", tabela, None, "Secretarias por 2", dois)
    assert item["de_para_parte"] == dois, item
    consulta = {"chaves": [{"minha": "NOME", "dela": "NOME"}], "puxar": [item]}
    assert consultas.completar(consulta, tabela.linhas[1])["SECRETARIA"] == "Secretaria Municipal de Saúde"
    sem_parte = consultas.item_puxar("SECRETARIA", "LOT", tabela, None, "Secretarias por 2", None)
    assert "de_para_parte" not in sem_parte
    print("  [6] só uma parte do código (2 primeiros) ............. OK")


def teste_7_excel() -> None:
    import tempfile
    pasta = Path(tempfile.mkdtemp())
    lista = {"Secretarias": {"0312": "Governo", "0401": ""},
             "Nome muito comprido que não cabe na folha do Excel: lotações": {"1": "Um"}}
    arquivo = de_para.exportar_excel(list(lista), pasta / "de-para.xlsx", lista)
    assert de_para.ler_excel(arquivo) == lista, de_para.ler_excel(arquivo)
    import openpyxl  # planilha feita à mão: a folha vira o de-para
    wb = openpyxl.Workbook()
    wb.active.title = "Cargos"
    for linha in (["COD", "CARGO"], [10, "Assessor"], ["011", "Diretor"]):
        wb.active.append(linha)
    wb.save(pasta / "mao.xlsx")
    assert de_para.ler_excel(pasta / "mao.xlsx") == {"Cargos": {"10": "Assessor", "011": "Diretor"}}
    print("  [7] exportar e importar de-para em Excel ............. OK")


def teste_8_juntar_colunas() -> None:
    tabela = consultas.Tabela(["cd_matricula", "dv_matricula", "nm_pessoa"], [
        {"cd_matricula": "30244", "dv_matricula": "9", "nm_pessoa": "IVO TESTE"},
        {"cd_matricula": "38978", "dv_matricula": "1", "nm_pessoa": "IVO TESTE"},
        {"cd_matricula": "512", "nm_pessoa": "ANA SEM DV"}])
    matricula = {"minha": "MATRÍCULA", "dela": "cd_matricula",
                 "juntar": {"com": ["dv_matricula"], "separador": "-", "milhar": True}}
    consulta = {"chaves": [matricula, {"minha": "NOME", "dela": "nm_pessoa"}], "puxar": []}
    assert consultas.valor_juntado(matricula, tabela.linhas[0]) == "30.244-9"
    assert consultas.valor_juntado(matricula, tabela.linhas[2]) == "512", "sem o DV: só o código"
    assert consultas.completar(consulta, tabela.linhas[1]) == {"MATRÍCULA": "38.978-1", "NOME": "IVO TESTE"}
    # Procura pelo número inteiro (com ou sem ponto e traço) ou pelo começo
    assert consultas.buscar(consulta, tabela, "30.244-9") == [tabela.linhas[0]]
    assert consultas.buscar(consulta, tabela, "302449") == [tabela.linhas[0]]
    assert tabela.linhas[1] in consultas.buscar(consulta, tabela, "38978")
    assert "38.978-1" in consultas.descricao(consulta, tabela.linhas[1])
    sem_ponto = dict(matricula, juntar={"com": ["dv_matricula"], "separador": "", "milhar": False})
    assert consultas.valor_juntado(sem_ponto, tabela.linhas[0]) == "302449"
    print("  [8] juntar colunas (cd_matricula + dv → 30.244-9) ..... OK")


def teste_9_formato() -> None:
    f = consultas.formatar
    assert f("ASSESSOR TÉCNICO DAM-15", "titulo") == "Assessor Técnico DAM-15"
    assert f("Secretaria Municipal De Segurança Integrada", "titulo") == "Secretaria Municipal de Segurança Integrada"
    assert f("COORDENADOR II - DAM-6", "titulo") == "Coordenador II - DAM-6"
    assert f("ASSESSOR DE PROJETOS ESPECIAIS-CC-6", "titulo") == "Assessor de Projetos Especiais-CC-6"
    assert f("Ana", "maiusculas") == "ANA" and f("ANA", None) == "ANA"
    assert consultas.sugerir_formato("CARGO COMISSÃO", ["ASSESSOR TÉCNICO DAM-15", "DIRETOR GERAL"]) == "titulo"
    assert consultas.sugerir_formato("NOME", ["IVO TESTE MACHADO"]) == "", "o nome fica em maiúsculas"
    tabela = consultas.Tabela(["nm", "cargo", "sexo"], [{"nm": "IVO TESTE", "cargo": "ASSESSOR TÉCNICO DAM-15",
                                                         "sexo": "MASCULINO"}])
    consulta = {"chaves": [{"minha": "NOME", "dela": "nm"}], "puxar": [
        dict(consultas.item_puxar("CARGO COMISSÃO", "cargo", tabela), minha="CARGO COMISSÃO"),
        consultas.item_puxar("SEXO", "sexo", tabela, ["F", "M"])]}
    assert consultas.completar(consulta, tabela.linhas[0]) == {
        "NOME": "IVO TESTE", "CARGO COMISSÃO": "Assessor Técnico DAM-15", "SEXO": "M"}
    print("  [9] formato (ASSESSOR TÉCNICO → Assessor Técnico) .... OK")


def teste_10_so_se() -> None:
    tabela = consultas.Tabela(["nm", "cd_cargo_efetivo", "cd_cargo_comissao", "ds_cargo"], [
        {"nm": "ANA", "cd_cargo_efetivo": "4060", "cd_cargo_comissao": "0", "ds_cargo": "PROFESSOR"},
        {"nm": "BIA", "cd_cargo_efetivo": "0", "cd_cargo_comissao": "129", "ds_cargo": "ASSESSOR"},
        {"nm": "CAU", "ds_cargo": "DIRETOR"}])
    consulta = {"chaves": [{"minha": "NOME", "dela": "nm"}], "puxar": [
        {"minha": "CARGO EFETIVO", "dela": "ds_cargo", "se": [{"coluna": "cd_cargo_efetivo", "op": "tem_valor"}]},
        {"minha": "CARGO COMISSÃO", "dela": "ds_cargo", "se": [{"coluna": "cd_cargo_comissao", "op": "tem_valor"}]}]}
    assert consultas.completar(consulta, tabela.linhas[0]) == {"NOME": "ANA", "CARGO EFETIVO": "PROFESSOR"}
    assert consultas.completar(consulta, tabela.linhas[1]) == {"NOME": "BIA", "CARGO COMISSÃO": "ASSESSOR"}
    assert consultas.completar(consulta, tabela.linhas[2]) == {"NOME": "CAU"}, "vazio também não preenche"
    outra = {"chaves": [], "puxar": [{"minha": "X", "dela": "ds_cargo",
                                      "se": [{"coluna": "nm", "op": "diferente", "valor": "ANA"}]}]}
    assert consultas.completar(outra, tabela.linhas[0]) == {} and consultas.completar(outra, tabela.linhas[1])
    assert "tem valor" in consultas.descrever_se(consulta["puxar"][0]["se"])
    print("  [10] só preencher se (cargo efetivo só de quem tem) .. OK")


if __name__ == "__main__":
    teste_1_traduzir()
    teste_2_trazer()
    teste_3_sugerir()
    teste_4_completar()
    teste_5_juntar()
    teste_6_so_uma_parte()
    teste_7_excel()
    teste_8_juntar_colunas()
    teste_9_formato()
    teste_10_so_se()
    print("\n10 testes passaram.")
