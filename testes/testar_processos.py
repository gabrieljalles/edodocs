"""Testes offline dos processos configuráveis (processo.json) e dos campos {{...}} do Word.

  1. campo {{...}} quebrado pelo Word em vários pedaços é preenchido (e o vazio vira [CAMPO]);
  2. tipos de ligação: coluna (com "se vazio"), fixo, extenso, gênero e data por extenso;
  3. títulos com {COLUNA}, {número do ato} e {ano do ato}; coluna vazia vira erro claro;
  4. conferência: padrão do título igual ao antigo e padrão de um título diferente;
  5. processo completo com o Word e a planilha de exemplo do assistente.

Rode com:  python testes/testar_processos.py
"""
import sys
import tempfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

from docx import Document  # noqa: E402

import atos  # noqa: E402
import diario  # noqa: E402
import processos  # noqa: E402

TEMP = Path(tempfile.mkdtemp())


def teste_1_campo_quebrado_em_pedacos() -> None:
    doc = Document()
    p = doc.add_paragraph()
    # O Word costuma guardar "{{NOME}}" assim: "{{NO" + "ME}}", com formatações diferentes
    p.add_run("Servidor ").bold = False
    p.add_run("{{NO").bold = True
    p.add_run("ME}} matrícula {{MATRÍ")
    p.add_run("CULA}} e {{VAZIO}}.")
    modelo = TEMP / "quebrado.docx"
    doc.save(modelo)

    assert [r for _, r in atos.campos_do_modelo(modelo)] == ["NOME", "MATRÍCULA", "VAZIO"]
    texto = atos.textos(modelo, [{"NOME": "MARIA", "MATRICULA": "12.345-6"}])[0]
    assert texto == "Servidor MARIA matrícula 12.345-6 e [VAZIO].", texto
    html = atos.htmls(modelo, [{"NOME": "MARIA", "MATRICULA": "1"}])[0]
    assert "<strong>MARIA</strong>" in html, html  # valor fica com a formatação do início do campo
    print("  [1] {{campo}} quebrado em pedaços é preenchido ........ OK")


def teste_2_tipos_de_ligacao() -> None:
    dados = {"DIAS": "15", "SEXO": "F", "DATA": "05/10/2026", "REQ": ""}
    regra = processos.valor_da_regra
    assert regra({"tipo": "coluna", "coluna": "DIAS"}, dados) == "15"
    assert regra({"tipo": "coluna", "coluna": "REQ", "se_vazio": "padrão"}, dados) == "padrão"
    assert regra({"tipo": "fixo", "texto": "Texto"}, dados) == "Texto"
    assert regra({"tipo": "extenso", "coluna": "DIAS"}, dados) == "quinze"
    assert regra({"tipo": "genero", "coluna": "SEXO", "feminino": "lotada", "masculino": "lotado"},
                 dados) == "lotada"
    assert regra({"tipo": "genero", "coluna": "SEXO", "feminino": "a", "masculino": "o"},
                 {"SEXO": "Masculino"}) == "o"
    assert regra({"tipo": "data_extenso", "coluna": "DATA"}, dados) == "5 de outubro de 2026"
    print("  [2] tipos de ligação ................................... OK")


def teste_3_titulos() -> None:
    proc = processos.Processo(TEMP, {
        "coluna_numero": "PORTARIA SMA",
        "titulo_edocs": "ATO ADMINISTRATIVO SMA Nº {número do ato}",
        "titulo_diario": "ATO ADMINISTRATIVO SMA Nº {Nº NO DRIVE}/{ano do ato}",
    })
    dados = {"PORTARIASMA": "SMA 1.542/2026", "NONODRIVE": "1542", "NOME": "FULANA"}
    assert proc.titulo_edocs(dados) == "ATO ADMINISTRATIVO SMA Nº 1.542/2026"
    assert proc.titulo_diario(dados) == "ATO ADMINISTRATIVO SMA Nº 1542/2026"
    assert proc.descricao(dados) == "SMA 1.542/2026 · FULANA"
    # Um ato, várias pessoas: o título no e-docs termina com E OUTROS (o do Diário, não)
    varias = dict(dados, _ITENS=[dados, dict(dados, NOME="BELTRANA")])
    assert proc.titulo_edocs(varias) == "ATO ADMINISTRATIVO SMA Nº 1.542/2026 E OUTROS"
    assert proc.titulo_diario(varias) == "ATO ADMINISTRATIVO SMA Nº 1542/2026"
    assert proc.titulo_edocs(dict(dados, _ITENS=[dados])) == "ATO ADMINISTRATIVO SMA Nº 1.542/2026"
    # O e-docs recusa mais de 50 caracteres: os nomes do meio viram iniciais; se não couber, corta
    exo = processos.Processo(TEMP, {"titulo_edocs": "DECRETO S/Nº - EXONERAÇÃO {NOME}",
                                    "titulo_diario": "DECRETO S/Nº"})
    celia = {"NOME": "CELINA FICTICIOS TEIXEIRA"}
    assert exo.titulo_edocs(celia) == "DECRETO S/Nº - EXONERAÇÃO CELINA F. TEIXEIRA"
    curto = {"NOME": "ANA LIMA"}
    assert exo.titulo_edocs(curto) == "DECRETO S/Nº - EXONERAÇÃO ANA LIMA", "cabe: fica como está"
    dois = dict(celia, _ITENS=[celia, {"NOME": "BIA"}])
    assert exo.titulo_edocs(dois) == "DECRETO S/Nº - EXONERAÇÃO CELINA TEIXEIRA E OUTROS"
    for nome in ("MARIA APARECIDA DOS SANTOS DE OLIVEIRA FERREIRA DA SILVA", "PEDRO MODELO COSTA MARTINS"):
        for pessoas in (1, 2):
            d = {"NOME": nome, "_ITENS": [{"NOME": nome}] * pessoas}
            t = exo.titulo_edocs(d)
            assert len(t) <= 50, (t, len(t))
            assert t.endswith("E OUTROS") == (pessoas > 1), t
    try:
        proc.titulo_diario({"PORTARIASMA": "SMA 1/2026"})
    except RuntimeError as e:
        assert "Nº NO DRIVE" in str(e), e
    else:
        raise AssertionError("coluna vazia no título deveria dar erro")
    print("  [3] títulos com colunas e especiais ................... OK")


def teste_4_conferencia() -> None:
    padrao = processos.Processo(TEMP, {"titulo_diario": "ATO ADMINISTRATIVO SMA Nº {Nº NO DRIVE}/{ano do ato}"})
    assert padrao.padrao_conferencia().pattern == diario.PADRAO_ATO.pattern
    outro = processos.Processo(TEMP, {"coluna_numero": "Nº",
                                      "titulo_diario": "PORTARIA Nº {número do ato}"})
    pdf = "... PORTARIA N°\n 12/2026 concede ... PORTARIA Nº 1.013 / 2026 ..."
    achados = {diario.chave(n, a) for n, a in outro.padrao_conferencia().findall(pdf)}
    assert achados == {"12/2026", "1013/2026"}, achados
    assert outro.chave_conferencia({"NO": "12/2026"}) == "12/2026"
    print("  [4] conferência pelo título do processo ............... OK")


def teste_5_processo_de_exemplo() -> None:
    import assistente

    modelo, planilha = assistente.gerar_exemplos(TEMP / "exemplos")
    proc = processos.Processo(TEMP / "Exemplo", {
        "modelo": str(modelo), "planilha": str(planilha), "aba": "atos",
        "coluna_numero": "Nº DO ATO",
        "titulo_edocs": "ATO ADMINISTRATIVO SMA Nº {número do ato}",
        "titulo_diario": "ATO ADMINISTRATIVO SMA Nº {número do ato}",
        "ligacoes": {
            "DIAS POR EXTENSO": {"tipo": "extenso", "coluna": "DIAS"},
            "AO A SERVIDOR": {"tipo": "genero", "coluna": "SEXO",
                              "feminino": "à servidora", "masculino": "ao servidor"},
            "LOTADO": {"tipo": "genero", "coluna": "SEXO", "feminino": "lotada", "masculino": "lotado"},
        },
    })
    registros = proc.ler_planilha(proc.planilha_padrao())
    assert len(registros) == 3
    _, colunas = atos.cabecalho_planilha(planilha, "atos")
    assert proc.campos_sem_ligacao(colunas) == []
    texto = proc.textos(registros[1:2])[0]
    assert texto.startswith("ATO ADMINISTRATIVO SMA Nº 102/2026"), texto
    assert "30 (trinta) dia(s) de licença ao servidor JOÃO SOUZA" in texto, texto
    assert "lotado na Secretaria Municipal de Obras" in texto, texto
    sem_titulo = proc.textos(registros[1:2], sem_titulo=True)[0]
    assert not sem_titulo.startswith("ATO"), sem_titulo
    assert proc.titulo_diario(registros[0]) == "ATO ADMINISTRATIVO SMA Nº 101/2026"
    print("  [5] processo completo com os exemplos do assistente ... OK")


def teste_6_frase_sem_dado_some() -> None:
    proc = processos.Processo(TEMP, {"ligacoes": {
        "VIGÊNCIA": {"tipo": "opcoes", "coluna": "VIGÊNCIA", "opcoes": [
            {"valor": "A PARTIR DE", "texto": ", a partir de {{DATA}}"}]},
        "DATA": {"tipo": "data_extenso", "coluna": "DATA"},
        "CARGO": {"tipo": "fixo", "texto": "de {{CARGO COMISSÃO}}, {{UNIDADE}}"}}})
    valores = proc.valores({"VIGENCIA": "A PARTIR DE", "DATA": "", "CARGOCOMISSAO": "Assessor", "UNIDADE": ""})
    assert valores["VIGENCIA"] == "", "sem a data, a frase some (não fica ', a partir de ')"
    assert valores["CARGO"] == "de Assessor", "a frase fica, sem a vírgula do campo vazio"
    valores = proc.valores({"CARGOCOMISSAO": "Assessor", "UNIDADE": "Gabinete"})
    assert valores["CARGO"] == "de Assessor, Gabinete", valores["CARGO"]
    valores = proc.valores({"VIGENCIA": "A PARTIR DE", "DATA": "01/10/2026"})
    assert valores["VIGENCIA"] == ", a partir de 1º de outubro de 2026", valores["VIGENCIA"]
    print("  [6] frase sem nenhum dado preenchido some .............. OK")


def teste_7_titulo_no_ato() -> None:
    exo = processos.Processo(TEMP, {"titulo_no_ato": "DECRETO S/Nº - EXONERA {NOME}"})
    angela = {"NOME": "ANA MARIA EXEMPLO DA SILVA PEREIRA"}
    assert exo.titulo_no_ato(angela) == "DECRETO S/Nº - EXONERA ANA MARIA EXEMPLO DA SILVA PEREIRA", \
        "sem o limite do e-docs: nome inteiro"
    assert exo.titulo_no_ato(dict(angela, _ITENS=[angela, {"NOME": "ANA"}])).endswith("PEREIRA E OUTROS")
    assert exo.titulo_no_ato({"NOME": ""}) == "", "faltando a coluna, fica a linha do Word"
    assert processos.Processo(TEMP, {}).titulo_no_ato(angela) == ""
    print("  [7] título na 1ª linha do ato ........................... OK")


def teste_8_obrigatorios() -> None:
    proc = processos.Processo(TEMP, {"obrigatorios": ["MATRÍCULA", "A PEDIDO"]})
    assert proc.faltando({"NOME": "ANA", "MATRICULA": "1", "APEDIDO": "SIM"}) == []
    assert proc.faltando({"NOME": "ANA", "MATRICULA": " "}) == ["MATRÍCULA", "A PEDIDO"]
    ana, jose = {"NOME": "ANA", "MATRICULA": "1", "APEDIDO": "SIM"}, {"NOME": "JOSE", "APEDIDO": "NÃO"}
    assert proc.faltando(dict(ana, _ITENS=[ana, jose])) == ["MATRÍCULA (JOSE)"], "diz de quem falta"
    assert processos.Processo(TEMP, {}).faltando({}) == []
    print("  [8] campos obrigatórios .................................. OK")


def main() -> None:
    print("Testes dos processos configuráveis (offline)\n")
    testes = [teste_1_campo_quebrado_em_pedacos, teste_2_tipos_de_ligacao, teste_3_titulos,
              teste_4_conferencia, teste_5_processo_de_exemplo, teste_6_frase_sem_dado_some,
              teste_7_titulo_no_ato, teste_8_obrigatorios]
    for teste in testes:
        teste()
    print(f"\n{len(testes)} testes passaram.")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
