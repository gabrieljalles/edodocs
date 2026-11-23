"""Testes offline das regras "Se... escrever..." e dos atos com várias pessoas.

  1. condições: é (vários valores), não é, vazio, maior que, igual para todos do ato;
  2. campo "Muda conforme a planilha": primeira regra que serve, senão, {{CAMPO}} no texto;
  3. juntar linhas pela coluna: um ato com várias pessoas, descrição e linhas do Excel;
  4. trecho {{REPETIR}}...{{FIM REPETIR}}: I, II..., ; e ., fora do trecho a 1ª pessoa;
  5. parágrafo que só tinha um campo vazio some (sem deixar duas linhas em branco);
  6. outro Word conforme a condição (1 pessoa / várias);
  7. marcar o trecho pelo assistente grava os marcadores no Word (e troca os antigos);
  8. o processo Exoneração de entradas/processos monta os decretos;
  9. linha de tabela que repete (nomeação por concurso): escolher a linha, células em
     branco, texto e HTML com uma linha por pessoa.

Rode com:  python testes/testar_varias_pessoas.py
"""
import json
import shutil
import sys
import tempfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

from docx import Document  # noqa: E402

import atos  # noqa: E402
import processos  # noqa: E402

TEMP = Path(tempfile.mkdtemp())
vale = processos.condicao_vale


def linha(n: int, **colunas) -> dict:
    return {atos.normalizar(c): v for c, v in colunas.items()} | {"_LINHA": n}


def word(nome: str, paragrafos: list[str]) -> Path:
    doc = Document()
    for texto in paragrafos:
        doc.add_paragraph(texto)
    caminho = TEMP / nome
    doc.save(caminho)
    return caminho


def teste_1_condicoes() -> None:
    campos = {"APEDIDO": "Sim", "VINCULO": "Comissão", "DATA": "", "QUANTIDADEDEPESSOAS": "2"}
    assert vale({"campo": "A PEDIDO", "op": "igual", "valor": "SIM"}, campos)  # sem diferenciar maiúscula
    assert vale({"campo": "VÍNCULO", "op": "igual", "valor": "EFETIVO; COMISSAO"}, campos)  # vários valores
    assert vale({"campo": "VÍNCULO", "op": "diferente", "valor": "EFETIVO"}, campos)
    assert vale({"campo": "DATA", "op": "vazio"}, campos)
    assert not vale({"campo": "DATA", "op": "preenchido"}, campos)
    assert vale({"campo": "QUANTIDADE DE PESSOAS", "op": "maior", "valor": "1"}, campos)
    assert not vale({"campo": "QUANTIDADE DE PESSOAS", "op": "menor", "valor": "2"}, campos)
    ato = {"_ITENS": [{"SECRETARIA": "Saúde"}, {"SECRETARIA": "SAUDE"}]}
    assert vale({"campo": "SECRETARIA", "op": "todos_iguais"}, ato)
    ato["_ITENS"].append({"SECRETARIA": "Obras"})
    assert vale({"campo": "SECRETARIA", "op": "nem_todos_iguais"}, ato)
    assert processos.descrever_condicao({"campo": "A PEDIDO", "op": "igual", "valor": "SIM;TALVEZ"}) == \
        "A PEDIDO é SIM ou TALVEZ"
    print("  [1] condições ........................................... OK")


def teste_2_texto_condicional() -> None:
    proc = processos.Processo(TEMP, {"ligacoes": {
        "FUNDAMENTO": {"tipo": "condicao", "casos": [
            {"se": [{"campo": "A PEDIDO", "op": "igual", "valor": "SIM"},
                    {"campo": "VÍNCULO", "op": "igual", "valor": "EFETIVO"}], "texto": "artigo 48"},
            {"se": [{"campo": "A PEDIDO", "op": "igual", "valor": "SIM"}], "texto": "artigo 49, II"},
        ], "senao": "artigo 49, inciso I"},
        "VIGÊNCIA": {"tipo": "condicao", "casos": [
            {"se": [{"campo": "DATA", "op": "preenchido"}], "texto": ", a partir de {{DATA POR EXTENSO}}"}],
            "senao": ""},
        "DATA POR EXTENSO": {"tipo": "data_extenso", "coluna": "DATA"},
    }})
    valores = proc.valores(linha(2, **{"A PEDIDO": "SIM", "VÍNCULO": "EFETIVO", "DATA": "01/09/2026"}))
    assert valores["FUNDAMENTO"] == "artigo 48"
    assert valores["VIGENCIA"] == ", a partir de 1º de setembro de 2026", valores["VIGENCIA"]
    valores = proc.valores(linha(2, **{"A PEDIDO": "SIM", "VÍNCULO": "COMISSÃO"}))
    assert valores["FUNDAMENTO"] == "artigo 49, II"
    assert valores["VIGENCIA"] == "" and "VIGENCIA" in valores["_VAZIOS"]  # vazio de propósito
    assert proc.valores(linha(2, **{"A PEDIDO": "NÃO"}))["FUNDAMENTO"] == "artigo 49, inciso I"
    # O campo citado só no texto da regra aparece para ser ligado no assistente
    assert "DATAPOREXTENSO" in dict(proc.campos())
    print("  [2] campo que muda conforme a planilha .................... OK")


def teste_3_juntar_linhas() -> None:
    proc = processos.Processo(TEMP, {"juntar": {"coluna": "DECRETO"}})
    registros = [linha(2, DECRETO="A", NOME="ANA"), linha(3, DECRETO="B", NOME="BRUNO"),
                 linha(4, DECRETO="a", NOME="CAIO"), linha(5, DECRETO="", NOME="DORA"),
                 linha(6, DECRETO="", NOME="EVA")]
    lista = proc.agrupar(registros)
    assert [processos.linhas(a) for a in lista] == [[2, 4], [3], [5], [6]]  # vazio fica sozinho
    assert proc.descricao(lista[0]) == "ANA e mais 1"
    assert proc.ato_da_linha(registros, 4)["NOME"] == "ANA"
    # Sem "juntar", cada linha é um ato, como sempre foi
    assert processos.Processo(TEMP, {}).agrupar(registros) == registros
    assert processos.linhas(registros[0]) == [2]
    print("  [3] linhas com o mesmo valor viram um ato ................. OK")


def teste_4_trecho_que_repete() -> None:
    modelo = word("repete.docx", [
        "Ficam exonerados{{EM COMUM}}, sendo:",
        "{{REPETIR}}",
        "{{ITEM}} – {{NOME}}, matrícula {{MATRÍCULA}}{{DO ITEM}}{{FIM DO ITEM}}",
        "{{FIM REPETIR}}",
        "Total: {{QUANTIDADE DE PESSOAS}}. Primeira: {{NOME}}.",
    ])
    assert [r for _, r in atos.campos_do_modelo(modelo)][:2] == ["EM COMUM", "ITEM"]  # sem os marcadores
    regra_comum = {"tipo": "condicao", "senao": "", "casos": [
        {"se": [{"campo": "SECRETARIA", "op": "todos_iguais"}], "texto": " da {{SECRETARIA}}"}]}
    regra_item = {"tipo": "condicao", "senao": "", "casos": [
        {"se": [{"campo": "SECRETARIA", "op": "nem_todos_iguais"}], "texto": ", da {{SECRETARIA}}"}]}
    proc = processos.Processo(TEMP, {"modelo": str(modelo), "juntar": {"coluna": "DECRETO"},
                                     "ligacoes": {"EM COMUM": regra_comum, "DO ITEM": regra_item}})
    diversas = proc.agrupar([
        linha(2, DECRETO="1", NOME="ANA", MATRÍCULA="1-1", SECRETARIA="Saúde"),
        linha(3, DECRETO="1", NOME="BRUNO", MATRÍCULA="2-2", SECRETARIA="Obras"),
        linha(4, DECRETO="1", NOME="CAIO", MATRÍCULA="3-3", SECRETARIA="Obras"),
    ])
    texto = proc.textos(diversas)[0]
    assert texto == ("Ficam exonerados, sendo:\n"
                     "I – ANA, matrícula 1-1, da Saúde;\n"
                     "II – BRUNO, matrícula 2-2, da Obras;\n"
                     "III – CAIO, matrícula 3-3, da Obras.\n"
                     "Total: 3. Primeira: ANA."), texto
    mesma = proc.agrupar([linha(2, DECRETO="1", NOME="ANA", MATRÍCULA="1", SECRETARIA="Saúde"),
                          linha(3, DECRETO="1", NOME="BIA", MATRÍCULA="2", SECRETARIA="Saúde")])
    texto = proc.textos(mesma)[0]
    assert texto.startswith("Ficam exonerados da Saúde, sendo:\nI – ANA, matrícula 1;\nII – BIA, matrícula 2."), texto
    assert "[" not in "".join(proc.htmls(mesma)), "HTML com campo sem valor"
    assert "BIA" in proc.folhas(mesma)[0]
    print("  [4] trecho que repete para cada pessoa .................... OK")


def teste_5_paragrafo_vazio_some() -> None:
    modelo = word("considerando.docx", ["O PREFEITO,", "", "{{CONSIDERANDO}}", "", "DECRETA:"])
    proc = processos.Processo(TEMP, {"modelo": str(modelo), "ligacoes": {"CONSIDERANDO": {
        "tipo": "condicao", "senao": "",
        "casos": [{"se": [{"campo": "A PEDIDO", "op": "igual", "valor": "SIM"}], "texto": "Considerando X,"}]}}})
    com, sem = proc.textos([linha(2, **{"A PEDIDO": "SIM"}), linha(3, **{"A PEDIDO": "NÃO"})])
    assert com == "O PREFEITO,\n\nConsiderando X,\n\nDECRETA:", com
    assert sem == "O PREFEITO,\n\nDECRETA:", sem  # uma linha em branco só
    # Coluna que falta na planilha continua aparecendo como [CAMPO] (não some calada)
    modelo = word("falta.docx", ["{{SEM COLUNA}}"])
    assert atos.textos(modelo, [{}])[0] == "[SEM COLUNA]"
    print("  [5] parágrafo de campo vazio some ......................... OK")


def teste_6_outro_word() -> None:
    um = word("um.docx", ["Fica exonerado {{NOME}}."])
    varios = word("varios.docx", ["Ficam exonerados:", "{{REPETIR}}", "{{NOME}}{{FIM DO ITEM}}", "{{FIM REPETIR}}"])
    proc = processos.Processo(TEMP, {
        "modelo": str(um), "juntar": {"coluna": "DECRETO"},
        "outros_modelos": [{"modelo": str(varios), "se": [
            {"campo": "QUANTIDADE DE PESSOAS", "op": "maior", "valor": "1"}]}]})
    lista = proc.agrupar([linha(2, DECRETO="1", NOME="ANA"), linha(3, DECRETO="2", NOME="BIA"),
                          linha(4, DECRETO="2", NOME="CAIO"), linha(5, DECRETO="3", NOME="DUDA")])
    assert [proc.modelo_do_ato(a).name for a in lista] == ["um.docx", "varios.docx", "um.docx"]
    assert proc.textos(lista) == ["Fica exonerado ANA.", "Ficam exonerados:\nBIA;\nCAIO.",
                                  "Fica exonerado DUDA."]
    assert proc.modelos() == [um, varios]
    print("  [6] outro Word conforme a condição ........................ OK")


def teste_7_marcar_trecho_no_word() -> None:
    modelo = word("marcar.docx", ["Abertura", "I – {{NOME}};", "", "Fim"])
    atos.marcar_repeticao(modelo, 1, 2)
    assert atos.paragrafos_do_modelo(modelo) == ["Abertura", "{{REPETIR}}", "I – {{NOME}};", "",
                                                 "{{FIM REPETIR}}", "Fim"]
    # Marcar de novo troca os marcadores antigos (não acumula)
    atos.marcar_repeticao(modelo, 2, 2)
    assert atos.paragrafos_do_modelo(modelo) == ["Abertura", "{{REPETIR}}", "I – {{NOME}};",
                                                 "{{FIM REPETIR}}", "", "Fim"]
    atos.marcar_repeticao(modelo, None)
    assert atos.paragrafos_do_modelo(modelo) == ["Abertura", "I – {{NOME}};", "", "Fim"]
    print("  [7] marcar o trecho grava os marcadores no Word ........... OK")


def teste_8_processo_exoneracao() -> None:
    origem = RAIZ / "entradas" / "processos" / "Exoneração"
    pasta = TEMP / "Exoneração"
    shutil.copytree(origem, pasta)
    proc = processos.carregar(pasta)
    assert json.loads((pasta / "processo.json").read_text(encoding="utf-8"))["juntar"]
    comum = {"A PEDIDO": "NÃO", "VÍNCULO": "COMISSÃO", "INICIAIS": "abc", "SEXO": "F"}
    registros = [
        linha(2, **comum, **{"Nº DO DECRETO": "10", "NOME": "MARIA TESTE", "MATRÍCULA": "1.111-1",
                             "CARGO COMISSÃO": "Assessora DAM-3", "UNIDADE": "Núcleo X",
                             "SECRETARIA": "Secretaria Municipal de Saúde"}),
        linha(3, **comum, **{"Nº DO DECRETO": "10", "NOME": "JOANA TESTE", "MATRÍCULA": "2.222-2",
                             "CARGO COMISSÃO": "Diretora DAM-8", "UNIDADE": "Diretoria Y",
                             "SECRETARIA": "Secretaria Municipal de Saúde"}),
        linha(4, **{"Nº DO DECRETO": "11", "NOME": "ANA EFETIVA", "MATRÍCULA": "3.333-3", "SEXO": "F",
                    "A PEDIDO": "SIM", "VÍNCULO": "EFETIVO", "CARGO EFETIVO": "Professora", "PADRÃO": "3",
                    "NÍVEL": "Mestrado", "SECRETARIA": "Secretaria Municipal de Educação",
                    "VIGÊNCIA": "RETROATIVO", "DATA": "03/08/2026", "INICIAIS": "pbs"}),
    ]
    varias, uma = proc.textos(proc.agrupar(registros))
    assert "Ficam exonerados os servidores ocupantes dos cargos de provimento em comissão abaixo " \
           "mencionados, da Secretaria Municipal de Saúde, sendo:" in varias, varias
    assert "I – MARIA TESTE, matrícula nº 1.111-1, do cargo de provimento em comissão de Assessora DAM-3, " \
           "Núcleo X;\n" in varias, varias
    assert "II – JOANA TESTE" in varias and "Diretoria Y.\n" in varias
    assert "artigo 49, inciso I" in varias and "HRB/abc" in varias and "[" not in varias
    assert "Fica exonerada a pedido, ANA EFETIVA, matrícula nº 3.333-3, do cargo de provimento efetivo de " \
           "Professora, Padrão 3, Nível de Qualificação Mestrado, da Secretaria Municipal de Educação, " \
           "retroativo a 3 de agosto de 2026." in uma, uma
    assert "artigo 48" in uma and "Considerando o Formulário de Desligamento," in uma and "[" not in uma
    print("  [8] processo Exoneração monta os decretos ................. OK")


def teste_9_linha_de_tabela() -> None:
    import marcar_ui
    import modelo_automatico
    doc = Document()
    doc.add_paragraph("DECRETO S/Nº")
    doc.add_paragraph("Art. 1º Ficam nomeados para o cargo de Professor os candidatos do anexo.")
    tabela = doc.add_table(rows=4, cols=4)
    for r, valores in enumerate([("ORDEM", "LISTA", "CLASS PCD", "CANDIDATO"),
                                     ("315", "AMPLA", "", "MARINA EXEMPLO FERREIRA"),
                                     ("316", "AMPLA", "", "ALICE FICTICIA RAMOS"),
                                     ("317", "PCD", "2", "CARLA MODELO SOUZA")]):
        for coluna, valor in enumerate(valores):
            tabela.cell(r, coluna).text = valor
    doc.add_paragraph("Art. 2º Este Decreto entra em vigor na data de sua publicação.")
    word = TEMP / "nomeia.docx"
    doc.save(word)

    # marcar na linha da MARCELLA, inclusive a célula em branco (CLASS PCD)
    pars = modelo_automatico.textos_do_word(word)
    n = pars.index("MARINA EXEMPLO FERREIRA")
    assert ("tabela", ) == tuple(t for t, _ in marcar_ui.estrutura(word) if t == "tabela")
    marcas = [marcar_ui.Marca("ORDEM", "texto", "315", [(n - 3, 0, 3)]),
              marcar_ui.Marca("LISTA", "texto", "AMPLA", [(n - 2, 0, 5)]),
              marcar_ui.Marca("CLASS PCD", "texto", "", [(n - 1, 0, 0)]),
              marcar_ui.Marca("NOME", "texto", "MARINA EXEMPLO FERREIRA", [(n, 0, len(pars[n]))])]
    modelo = TEMP / "nomeia_modelo.docx"
    ligacoes, _ = marcar_ui.criar_modelo(word, modelo, marcas)
    assert ligacoes["CLASS PCD"] == {"tipo": "coluna", "coluna": "CLASS PCD", "se_vazio": ""}

    # a linha escolhida repete; as dos outros candidatos saem
    indice = next(iter(atos.linhas_das_tabelas(modelo)))
    assert atos.preparar_linha_pessoa(modelo, indice, 1) == {"removidos": 2}
    assert atos.resumo_da_repeticao(modelo) == "{{ORDEM}} | {{LISTA}} | {{CLASS PCD}} | {{NOME}}"
    assert atos.trecho_que_repete(atos._abrir_modelo(modelo)[2]) is None  # não é trecho de parágrafos
    assert atos.campos_que_repetem(modelo) == {"ORDEM", "LISTA", "CLASSPCD", "NOME"}

    proc = processos.Processo(TEMP, {"modelo": str(modelo), "ligacoes": ligacoes, "juntar": {"coluna": "ATO"}})
    registros = [linha(2, ATO="1", ORDEM="315", LISTA="AMPLA", NOME="MARCELLA"),
                 linha(3, ATO="1", ORDEM="317", LISTA="PCD", **{"CLASS PCD": "2"}, NOME="CAMILA"),
                 linha(4, ATO="2", ORDEM="400", LISTA="AMPLA", NOME="OUTRO ATO")]
    atos_juntos = proc.agrupar(registros)
    texto = proc.textos(atos_juntos)[0]
    assert "ORDEM | LISTA | CLASS PCD | CANDIDATO\n315 | AMPLA |  | MARCELLA\n317 | PCD | 2 | CAMILA\n" \
           "Art. 2º" in texto, texto
    assert "[" not in texto and "OUTRO ATO" not in texto
    html = proc.htmls(atos_juntos)[0]
    assert html.count("<tr>") == 3 and "<table" in html and "CAMILA" in html
    assert "<table" in proc.folhas(atos_juntos)[0]
    # em um Word com o trecho de parágrafos, marcar a linha da tabela troca o trecho
    atos.marcar_repeticao(modelo, 1, 1)
    assert atos.resumo_da_repeticao(modelo).startswith("Art. 1º")
    atos.preparar_linha_pessoa(modelo, next(iter(atos.linhas_das_tabelas(modelo))), 1)
    assert atos.resumo_da_repeticao(modelo).startswith("{{ORDEM}}")
    assert "REPETIR" not in "".join(atos.paragrafos_do_modelo(modelo)[:2])  # só na tabela
    print("  [9] linha de tabela que repete ............................ OK")


def teste_10_trechos_que_mudam() -> None:
    """Um Word só para "Nomeia Andreia" (tem cargo efetivo) e "Nomeia Silvano" (não tem):
    trecho que só aparece às vezes (com campos dentro), sexo e uma pessoa / várias."""
    import marcar_ui
    texto = ("Art. 1º Fica nomeada ANDREIA LEAL, ocupante do cargo de provimento efetivo de Professor, "
             "Padrão 12, para o cargo de provimento em comissão de Supervisor DAM-8, a contar de 1º de "
             "outubro de 2026.")
    word_ = word("andreia.docx", ["DECRETO S/Nº", texto, "Considerando se tratar de reposição.", "FIM"])

    def onde(trecho: str, n: int = 1) -> tuple[int, int, int]:
        paragrafo = ["DECRETO S/Nº", texto, "Considerando se tratar de reposição.", "FIM"][n]
        i = paragrafo.index(trecho)
        return (n, i, i + len(trecho))

    efetivo = ", ocupante do cargo de provimento efetivo de Professor, Padrão 12"
    contar = ", a contar de 1º de outubro de 2026"
    marcas = [
        marcar_ui.Marca("FICA", "quantidade", "Fica", [onde("Fica")],
                        regra={"uma": "Fica", "uma_feminino": "", "varias": "Ficam"}),
        marcar_ui.Marca("NOMEADO OU NOMEADA", "genero", "nomeada", [onde("nomeada")],
                        regra={"masculino": "nomeado", "feminino": "nomeada"}),
        marcar_ui.Marca("NOME", "texto", "ANDREIA LEAL", [onde("ANDREIA LEAL")]),
        marcar_ui.Marca("CARGO EFETIVO", "texto", "Professor", [onde("Professor")]),
        marcar_ui.Marca("PADRÃO", "texto", "12", [onde("12")]),
        # o trecho do cargo efetivo, com os dois campos de cima dentro dele
        marcar_ui.Marca("OCUPANTE", "as_vezes", efetivo, [onde(efetivo)],
                        regra={"quando": "preenchido", "coluna": "CARGO EFETIVO"},
                        modelo=", ocupante do cargo de provimento efetivo de {{CARGO EFETIVO}}, Padrão {{PADRÃO}}"),
        marcar_ui.Marca("CARGO", "texto", "Supervisor DAM-8", [onde("Supervisor DAM-8")]),
        marcar_ui.Marca("A CONTAR DE", "data", "1º de outubro de 2026", [onde("1º de outubro de 2026")]),
        marcar_ui.Marca("VIGÊNCIA", "as_vezes", contar, [onde(contar)],
                        regra={"quando": "preenchido", "coluna": "A CONTAR DE"}, modelo=", a contar de {{A CONTAR DE}}"),
        marcar_ui.Marca("REPOSIÇÃO", "as_vezes", "Considerando se tratar de reposição.",
                        [onde("Considerando se tratar de reposição.", 2)], regra={"quando": "sim"},
                        modelo="Considerando se tratar de reposição."),
    ]
    modelo = TEMP / "andreia_modelo.docx"
    ligacoes, exemplo = marcar_ui.criar_modelo(word_, modelo, marcas)
    paragrafos = atos.paragrafos_do_modelo(modelo)
    assert paragrafos[1] == "Art. 1º {{FICA}} {{NOMEADO OU NOMEADA}} {{NOME}}{{OCUPANTE}}, para o cargo de " \
                            "provimento em comissão de {{CARGO}}{{VIGÊNCIA}}.", paragrafos[1]
    assert paragrafos[2] == "{{REPOSIÇÃO}}"
    assert exemplo["SEXO"] == "F" and exemplo["REPOSICAO"] == "SIM"

    proc = processos.Processo(TEMP, {"modelo": str(modelo), "ligacoes": ligacoes})
    # campos de dentro das regras também são colunas da planilha (e o SEXO, com F e M)
    colunas = {c["nome"]: c for c in proc.colunas_planilha_nova()}
    assert {"CARGO EFETIVO", "PADRÃO", "A CONTAR DE", "SEXO", "REPOSIÇÃO"} <= set(colunas), colunas
    assert colunas["SEXO"]["opcoes"] == ["F", "M"] and colunas["A CONTAR DE"]["tipo"] == "data"

    andreia, silvano = proc.textos([
        linha(2, NOME="ANDREIA LEAL", SEXO="F", **{"CARGO EFETIVO": "Professor", "PADRÃO": "12",
                                                  "A CONTAR DE": "01/10/2026", "REPOSIÇÃO": "SIM"},
              CARGO="Supervisor DAM-8"),
        linha(3, NOME="SILVANO COSTA", SEXO="M", CARGO="Subsecretário DAM-18", **{"REPOSIÇÃO": "NÃO"}),
    ])
    assert "Art. 1º Fica nomeada ANDREIA LEAL, ocupante do cargo de provimento efetivo de Professor, " \
           "Padrão 12, para o cargo de provimento em comissão de Supervisor DAM-8, a contar de 1º de " \
           "outubro de 2026.\nConsiderando se tratar de reposição.\nFIM" in andreia, andreia
    assert "Art. 1º Fica nomeado SILVANO COSTA, para o cargo de provimento em comissão de " \
           "Subsecretário DAM-18.\nFIM" in silvano, silvano  # sem efetivo, sem data, sem reposição
    juntos = dict(linha(2, NOME="A", SEXO="M", CARGO="X"), _ITENS=[{}, {}])
    assert proc.valores(juntos)["FICA"] == "Ficam"
    print("  [10] trechos às vezes, sexo e uma / várias .............. OK")


if __name__ == "__main__":
    testes = [teste_1_condicoes, teste_2_texto_condicional, teste_3_juntar_linhas, teste_4_trecho_que_repete,
              teste_5_paragrafo_vazio_some, teste_6_outro_word, teste_7_marcar_trecho_no_word,
              teste_8_processo_exoneracao, teste_9_linha_de_tabela, teste_10_trechos_que_mudam]
    for teste in testes:
        teste()
    print(f"\n{len(testes)} testes passaram.")
