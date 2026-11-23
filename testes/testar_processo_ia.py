"""Montar com IA (processo_ia.py), com uma IA de mentira: lê os Words, aplica o plano sem
perder a formatação, monta o trecho que repete, cria a planilha, refaz os exemplos e manda a
IA corrigir o que não bateu.

    python testes/testar_processo_ia.py
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
import openpyxl  # noqa: E402
import processo_ia  # noqa: E402
import processos  # noqa: E402

passou = 0


def ok(texto: str) -> None:
    global passou
    passou += 1
    print(f"  [{passou}] {texto} " + "." * max(2, 55 - len(texto)) + " OK")


def word(caminho: Path, paragrafos: list, tabela: list[list[str]] | None = None) -> Path:
    doc = Document()
    for texto in paragrafos:
        p = doc.add_paragraph()
        if isinstance(texto, tuple):  # (texto em negrito, resto)
            p.add_run(texto[0]).bold = True
            p.add_run(texto[1])
        else:
            p.add_run(texto)
    if tabela:
        t = doc.add_table(rows=0, cols=len(tabela[0]))
        for linha in tabela:
            celulas = t.add_row().cells
            for c, valor in zip(celulas, linha):
                c.text = valor
    doc.save(caminho)
    return caminho


TMP = Path(tempfile.mkdtemp(prefix="edodocs_ia_"))
try:
    # Dois atos de uma pessoa (o nome em negrito) e um de várias pessoas
    a1 = word(TMP / "Estabilidade João.docx", [
        "DECRETO S/Nº - DECLARA ESTÁVEL JOÃO DA SILVA",
        ("JOÃO DA SILVA", ", matrícula nº 12.345-6, fica declarado estável a partir de 19 de julho de 2026."),
        "Este Decreto entra em vigor na data de sua publicação."])
    a2 = word(TMP / "Estabilidade Maria.docx", [
        "DECRETO S/Nº - DECLARA ESTÁVEL MARIA SOUZA",
        ("MARIA SOUZA", ", matrícula nº 22.222-2, fica declarada estável a partir de 1º de agosto de 2026."),
        "Este Decreto entra em vigor na data de sua publicação."])
    a3 = word(TMP / "Estabilidade várias.docx", [
        "DECRETO S/Nº - DECLARA ESTÁVEIS",
        "Ficam declarados estáveis os servidores:",
        "I – ANA LIMA, matrícula nº 1.111-1;",
        "",
        "II – BETO REIS, matrícula nº 2.222-2;",
        "",
        "III – CAIO DIAS, matrícula nº 3.333-3.",
        "",
        "Este Decreto entra em vigor na data de sua publicação."])
    anexo = word(TMP / "Anexo.docx", ["ANEXO"], [["ORDEM", "NOME"], ["1", "ANA LIMA"], ["2", "BETO REIS"]])

    exemplos = processo_ia.ler_arquivos([a1, a2, a3, anexo], TMP / "trabalho")
    assert [e.nome for e in exemplos][:2] == ["Estabilidade João.docx", "Estabilidade Maria.docx"]
    assert "ORDEM | NOME" in exemplos[3].texto and "1 | ANA LIMA" in exemplos[3].texto
    pedido = processo_ia.pedido(exemplos, "Decretos de estabilidade")
    assert '=== ARQUIVO: "Estabilidade várias.docx" (Word, pode ser base) ===' in pedido
    assert "Decretos de estabilidade" in pedido
    try:
        processo_ia.ler_arquivos([TMP / "planilha.xlsx"], TMP / "x")
        raise AssertionError("devia recusar")
    except RuntimeError:
        pass
    ok("lê os Words (tabelas com |) e monta o pedido")

    genero = {"tipo": "genero", "coluna": "SEXO", "masculino": "declarado", "feminino": "declarada"}
    plano = {
        "versao_plano": 1, "nome": "Estabilidade IA", "explicacao": "Um Word para uma pessoa e um para várias.",
        "modelos": [
            {"arquivo": "Estabilidade - uma pessoa.docx", "base": "Estabilidade João.docx", "principal": True,
             "se": [], "repetir": None, "trocas": [
                 {"de": "JOÃO DA SILVA", "para": "{{NOME}}"},
                 {"de": "12.345-6", "para": "{{MATRÍCULA}}"},
                 {"de": "19 de julho de 2026", "para": "{{DATA POR EXTENSO}}"}]},
            {"arquivo": "Estabilidade - várias.docx", "base": "Estabilidade várias.docx", "principal": False,
             "se": [{"campo": "QUANTIDADE DE PESSOAS", "op": "maior", "valor": "1"}],
             "repetir": {"primeiro": "I – ANA LIMA, matrícula nº 1.111-1;",
                         "apagar": ["II – BETO REIS, matrícula nº 2.222-2;", "III – CAIO DIAS, matrícula nº 3.333-3."]},
             "trocas": [{"de": "I – ANA LIMA, matrícula nº 1.111-1;",
                         "para": "{{ITEM}} – {{NOME}}, matrícula nº {{MATRÍCULA}}{{FIM DO ITEM}}"}]}],
        "colunas": [{"nome": "Nº DO DECRETO", "tipo": "texto", "opcoes": []},
                    {"nome": "NOME", "tipo": "texto", "opcoes": []},
                    {"nome": "MATRÍCULA", "tipo": "texto", "opcoes": []},
                    {"nome": "SEXO", "tipo": "lista", "opcoes": ["F", "M"]},
                    {"nome": "DATA", "tipo": "data", "opcoes": []}],
        "processo": {"coluna_nome": "NOME", "coluna_numero": "Nº DO DECRETO",
                     "titulo_edocs": "DECRETO S/Nº - ESTABILIDADE {NOME}",
                     "titulo_diario": "DECRETO S/Nº - DECLARA ESTÁVEL {NOME}", "titulo_no_ato": "",
                     "titulo_na_primeira_linha": True, "juntar": {"coluna": "Nº DO DECRETO"},
                     "vai_ao_diario": True, "obrigatorios": ["NOME", "MATRÍCULA"],
                     "ligacoes": {"DECLARADO": genero,
                                  "DATA POR EXTENSO": {"tipo": "data_extenso", "coluna": "DATA"}}},
        "exemplos": [
            {"arquivo": "Estabilidade João.docx", "pessoas": [
                {"Nº DO DECRETO": "S/N", "NOME": "JOÃO DA SILVA", "MATRÍCULA": "12.345-6", "SEXO": "M",
                 "DATA": "19/07/2026"}]},
            {"arquivo": "Estabilidade Maria.docx", "pessoas": [
                {"Nº DO DECRETO": "S/N", "NOME": "MARIA SOUZA", "MATRÍCULA": "22.222-2", "SEXO": "F",
                 "DATA": "01/08/2026"}]},
            {"arquivo": "Estabilidade várias.docx", "pessoas": [
                {"Nº DO DECRETO": "S/N", "NOME": n, "MATRÍCULA": m, "SEXO": "F", "DATA": ""}
                for n, m in (("ANA LIMA", "1.111-1"), ("BETO REIS", "2.222-2"), ("CAIO DIAS", "3.333-3"))]}],
        "duvidas": ["O título do de várias pessoas ficou fixo."]}

    # Faltou a troca de "declarado": a conferência acha a diferença na Maria
    assert processo_ia.conferir_plano(plano, exemplos) == []
    pasta = TMP / "trabalho"
    config, planilha, problemas = processo_ia.montar(plano, exemplos, pasta)
    assert problemas == [], problemas
    principal = Document(pasta / "Estabilidade - uma pessoa.docx")
    p1 = principal.paragraphs[1]
    assert p1.runs[0].text == "{{NOME}}" and p1.runs[0].bold, [r.text for r in p1.runs]  # negrito ficou
    assert p1.text.endswith("declarado estável a partir de {{DATA POR EXTENSO}}.")
    assert principal.paragraphs[0].text == "DECRETO S/Nº - DECLARA ESTÁVEL {{NOME}}"  # troca todos
    ok("aplica as trocas sem perder a formatação")

    varias = [p.text for p in Document(pasta / "Estabilidade - várias.docx").paragraphs]
    assert varias[2:5] == ["{{REPETIR}}", "{{ITEM}} – {{NOME}}, matrícula nº {{MATRÍCULA}}{{FIM DO ITEM}}", ""], varias
    assert varias[5] == "{{FIM REPETIR}}" and not any("BETO" in t or "CAIO" in t for t in varias), varias
    ok("várias pessoas: repete a 1ª e tira as outras")

    cab = [c.value for c in openpyxl.load_workbook(planilha)["dados"][1]]
    assert cab == ["Nº DO DECRETO", "NOME", "MATRÍCULA", "SEXO", "DATA"], cab
    assert config["outros_modelos"] == [{"modelo": "Estabilidade - várias.docx", "se": plano["modelos"][1]["se"]}]
    assert config["_tipo"] == "varias" and config["planilha"] == planilha.name and config["aba"] == "dados", {k: config[k] for k in ("_tipo", "planilha", "aba")}
    ok("planilha sem cores e processo.json")

    comparacoes = processo_ia.refazer_exemplos(plano, exemplos, config, pasta)
    por_arquivo = {c.arquivo: c for c in comparacoes}
    assert por_arquivo["Estabilidade João.docx"].igual
    assert por_arquivo["Estabilidade várias.docx"].igual, por_arquivo["Estabilidade várias.docx"].diferencas
    maria = por_arquivo["Estabilidade Maria.docx"]
    assert not maria.igual and "declarada" in maria.diferencas[0][0], maria.diferencas
    problemas = processo_ia.problemas_das_comparacoes(comparacoes, exemplos)
    assert any("Maria" in p for p in problemas) and any("Anexo.docx" in p for p in problemas), problemas
    ok("refaz cada exemplo e acha o que não bate")

    # A IA de mentira: 1º plano sem o "declarado" e com um trecho que não existe; corrige depois
    errado = json.loads(json.dumps(plano))
    errado["modelos"][0]["trocas"].append({"de": "matrícula número 12.345-6", "para": "x"})
    certo = json.loads(json.dumps(plano))
    certo["modelos"][0]["trocas"].append({"de": "fica declarado", "para": "fica {{DECLARADO}}"})
    certo["exemplos"].append({"arquivo": "Anexo.docx", "pessoas": []})
    pedidos = []

    def pedir(sistema, texto):
        pedidos.append(texto)
        return errado if len(pedidos) == 1 else certo

    avisos = []
    resultado = processo_ia.montar_com_ia([e for e in exemplos if e.nome != "Anexo.docx"], pasta,
                                          aviso=avisos.append, pedir=pedir)
    assert len(pedidos) == 2 and "matrícula número 12.345-6" in pedidos[1] and "O SEU PLANO ANTERIOR" in pedidos[1]
    assert resultado.problemas == [] and resultado.rodadas == 1, resultado.problemas
    assert all(c.igual for c in resultado.comparacoes) and any("corrigindo" in a for a in avisos)
    ok("o que não bate volta para a IA corrigir")

    dados = processo_ia.rascunho_do_resultado(resultado, "Estabilidade IA")
    pessoas = dados["pessoas_exemplo"]
    assert len(pessoas) == 5 and dados["passo"] == "previa" and dados["modo_planilha"] == "criar"
    decretos = [p["NODODECRETO"] for p in pessoas]
    assert len(set(decretos)) == 3, decretos  # três exemplos "S/N" continuam três atos
    proc = processos.Processo(pasta, resultado.config)
    assert len(proc.agrupar(pessoas)) == 3
    assert all((Path(dados["_pasta"]) / nome).is_file() for nome in dados["arquivos"].values())
    ok("o resultado vira o rascunho do assistente")

    # Perguntas: limpa as malformadas; só volta para a IA se a resposta não for a sugestão
    com_perguntas = dict(resultado.plano, perguntas=[
        {"pergunta": "Pode ser de uma servidora?", "contexto": "Sempre \"o servidor\".", "por_que": "Só vi homens.",
         "opcoes": [{"texto": "Não, sempre \"o servidor\"", "como_fica": "o servidor"},
                    "Sim, muda conforme o sexo", {"texto": "Sim, muda conforme o sexo"}]},
        {"pergunta": "Só uma opção", "opcoes": ["Sim"]}, "texto solto"])
    perguntas = processo_ia.perguntas_do_plano(com_perguntas)
    assert len(perguntas) == 1 and len(perguntas[0]["opcoes"]) == 2, perguntas
    assert perguntas[0]["opcoes"][0] == {"texto": "Não, sempre \"o servidor\"", "como_fica": "o servidor"}
    assert perguntas[0]["opcoes"][1]["como_fica"] == "" and perguntas[0]["por_que"] == "Só vi homens."
    assert not processo_ia.precisa_ajustar([{"pergunta": "x", "resposta": "y", "sugerida": True}])
    respostas = [{"pergunta": perguntas[0]["pergunta"], "resposta": "Sim, muda conforme o sexo", "sugerida": False}]
    assert processo_ia.precisa_ajustar(respostas)
    pedidos.clear()

    def pedir_ajuste(sistema, texto):
        pedidos.append(texto)
        return dict(json.loads(json.dumps(certo)), perguntas=[{"pergunta": "de novo?", "opcoes": ["a", "b"]}])

    resultado_com = processo_ia.Resultado(**{**resultado.__dict__, "plano": com_perguntas})
    ajustado = processo_ia.ajustar_com_respostas(resultado_com, [e for e in exemplos if e.nome != "Anexo.docx"],
                                                 respostas, TMP / "ajustado", pedir=pedir_ajuste)
    assert "O QUE A PESSOA RESPONDEU" in pedidos[0] and "Sim, muda conforme o sexo" in pedidos[0]
    assert ajustado.plano["perguntas"] == [] and ajustado.problemas == [] and ajustado.pasta == TMP / "ajustado"
    dados = processo_ia.rascunho_do_resultado(ajustado, "Estabilidade IA", respostas)
    assert dados["ia"]["respostas"] == [(perguntas[0]["pergunta"], "Sim, muda conforme o sexo")]
    ok("perguntas da IA: respostas voltam para ela ajustar")

    # Plano ruim de vez: avisa em vez de montar
    try:
        processo_ia.montar_com_ia(exemplos, pasta, pedir=lambda s, t: {"nome": "X", "modelos": []})
        raise AssertionError("devia avisar")
    except processo_ia.ia.ErroIA as e:
        assert "modelos" in str(e)
    ok("plano sem modelo: avisa")
finally:
    shutil.rmtree(TMP, ignore_errors=True)

print(f"\n{passou} testes passaram.")
