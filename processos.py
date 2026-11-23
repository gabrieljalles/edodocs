"""Configuração de cada processo (Licença Prêmio, Compensação, ou um criado pela pessoa).

Cada pasta em entradas/processos/ é um processo: o Word modelo, a planilha (ou o caminho
dela) e o arquivo processo.json, que diz como ligar o Word à planilha e como montar os
títulos. Exemplo de processo.json:

    {
      "modelo": "modelo.docx",
      "planilha": "C:/Users/.../planilha.xlsx",      (opcional: sem ele, a .xlsx da pasta)
      "aba": "processosPendentes",                    (opcional: sem ela, a primeira aba)
      "coluna_nome": "NOME",                          linha sem esta coluna é ignorada
      "coluna_numero": "PORTARIA SMA",                número do ato (bolinhas, avisos)
      "titulo_edocs": "ATO ADMINISTRATIVO SMA Nº {número do ato}",
      "titulo_diario": "ATO ADMINISTRATIVO SMA Nº {Nº NO DRIVE}/{ano do ato}",
      "titulo_na_primeira_linha": true,               o Diário recebe o ato sem ela
      "titulo_no_ato": "DECRETO S/Nº - EXONERA {NOME}",  (opcional) no lugar da 1ª linha do Word
      "obrigatorios": ["NOME", "MATRÍCULA"],          colunas que não podem ficar vazias
      "vai_ao_diario": true,                          false = só e-docs (sem as etapas do Diário)
      "ligacoes": {                                   campo do Word → de onde vem o valor
        "DIA EXT": {"tipo": "extenso", "coluna": "DIAS"}
      },
      "edocs": {"tipo": "ATOS ADMINISTRATIVOS", "setor_posicao": 2},
      "diario": {"secretaria": "Administração", "categoria": "Ato Administrativo",
                 "coluna": "Duas Colunas"},
      "confirmar_envio": true,                         mostra cada ato antes de enviar
      "precisa_anexo": false,                         true = a cada ato, pergunta o(s) arquivo(s)
                                                      do computador para anexar no e-docs
      "caixas": ["CONFERIDO"],                        colunas de marcar (caixinha) na tabela dos
                                                      atos; ficam na planilha (SIM ou vazio)
      "colunas_tabela": [{"nome": "NOME", "ver": true}, ...]
                                                      ordem e colunas visíveis na tabela dos atos
      "pessoas": {"fixas": true,                      false = pergunta a cada envio
                  "vistos": [{"cpf": "...", "nome": "..."}],    na ordem (quem envia só vista
                                                                se estiver aqui)
                  "assinantes": [{"cpf": "...", "nome": "..."}]}
    }

Campo do Word sem ligação puxa a coluna da planilha com o mesmo nome (sem diferenciar
acento, espaço ou maiúscula). Tipos de ligação:
    coluna        {"coluna": "NOME", "se_vazio": "texto opcional"}
    fixo          {"texto": "..."}
    extenso       {"coluna": "DIAS"}                 15 → quinze
    genero        {"coluna": "SEXO", "feminino": "lotada", "masculino": "lotado"}
    data_extenso  {"coluna": "DATA"}                 05/10/2026 → 5 de outubro de 2026
    periodo       {"coluna": "INÍCIO", "ate": "FIM"} 09/11/2026 e 28/11/2026 → 09 a 28 de novembro
                  de 2026 (mês e ano só se mudam: "23 de novembro a 07 de dezembro de 2026")
    opcoes        {"coluna": "A PEDIDO", "opcoes": [{"valor": "SIM", "texto": "a pedido, "},
                                                    {"valor": "NÃO", "texto": ""}]}
                  a coluna tem uma das opções (lista no Excel) e o Word recebe o texto dela
    condicao      {"casos": [{"se": [CONDIÇÃO, ...], "texto": "artigo 48"}, ...],
                   "senao": "inciso I do artigo 49"}
                  o texto do primeiro caso em que todas as condições valem; o texto pode ter
                  {{CAMPO}} de outras colunas ou campos. Texto vazio é permitido (o parágrafo
                  que só tinha esse campo some do ato).
Qualquer ligação pode ter "formato": "maiusculas", "titulo" (Primeiras Maiúsculas) ou
"minusculas" (ex.: "CARGO: {{CARGO MAIÚSCULO}}" com {"coluna": "CARGO EFETIVO",
"formato": "maiusculas"}).

CONDIÇÃO: {"campo": "A PEDIDO", "op": "igual", "valor": "SIM"}; os "op" estão em
OPERADORES ("igual" aceita vários valores separados por ;). Comparações ignoram acento,
espaço e maiúscula.

Várias pessoas no mesmo ato (opcional):
    "juntar": {"coluna": "Nº DO DECRETO"}          linhas com o mesmo valor viram um ato
    "outros_modelos": [{"modelo": "varias.docx", "se": [CONDIÇÃO, ...]}]
                                                    Word usado quando as condições valem
                                                    (senão, o "modelo" principal)
No Word, o trecho entre {{REPETIR}} e {{FIM REPETIR}} sai uma vez por pessoa; campos
automáticos em atos.AUTOMATICOS ({{ITEM}} = I, II...; {{QUANTIDADE DE PESSOAS}}...).

Nos títulos, {COLUNA} é o valor da coluna (ou do campo do Word) e há dois especiais:
{número do ato} (ex.: 1542/2026, tirado da coluna do número) e {ano do ato}.

Mudou algo aqui (tipo de regra, operador, chave do processo.json)? Atualize também o guia da
IA, assets/ia/guia_processos.md: é ele que ensina a IA a montar processos (testes/testar_guia_ia.py
falha enquanto o guia não citar o que é novo).
"""
import json
import re
from datetime import date
from pathlib import Path

import atos

ARQUIVO = "processo.json"
NUMERO_DO_ATO = "NUMERODOATO"
ANO_DO_ATO = "ANODOATO"
_NUMERO = re.compile(r"([\d.]+)\s*/\s*(\d{4})")
_TOKEN = re.compile(r"\{([^{}]+)\}")

PADRAO = {
    "modelo": "",
    "planilha": "",
    "aba": "",
    "coluna_nome": "NOME",
    "coluna_numero": "PORTARIA SMA",
    "titulo_edocs": "ATO ADMINISTRATIVO SMA Nº {número do ato}",
    "titulo_diario": "ATO ADMINISTRATIVO SMA Nº {número do ato}",
    "titulo_na_primeira_linha": True,
    "vai_ao_diario": True,
    "precisa_anexo": False,
    "ligacoes": {},
    "edocs": {"tipo": "ATOS ADMINISTRATIVOS", "setor_posicao": 2},
    "diario": {"secretaria": "Administração", "categoria": "Ato Administrativo",
               "coluna": "Duas Colunas"},
    # Vistos e assinantes fixos do processo; sem eles, o envio pergunta toda vez
    "pessoas": {"fixas": False, "vistos": [], "assinantes": []},
    # Caixinhas de marcar na tabela dos atos (cada uma é uma coluna da planilha) e a ordem /
    # visibilidade das colunas da tabela
    "caixas": [],
    "colunas_tabela": [],
}

TIPOS = {
    "coluna": "Coluna da planilha",
    "fixo": "Texto fixo",
    "extenso": "Número por extenso",
    "genero": "Feminino / masculino",
    "data_extenso": "Data por extenso",
    "periodo": "Período (de uma data a outra)",
    "opcoes": "Escolha entre opções",
    "condicao": "Muda conforme a planilha",
}

# Como cada condição aparece para a pessoa ("A PEDIDO" "é" "SIM")
OPERADORES = {
    "igual": "é",
    "diferente": "não é",
    "contem": "contém",
    "preenchido": "está preenchido",
    "vazio": "está vazio",
    "maior": "é maior que",
    "menor": "é menor que",
    "todos_iguais": "é igual para todas as pessoas do ato",
    "nem_todos_iguais": "muda entre as pessoas do ato",
}
SEM_VALOR = {"preenchido", "vazio", "todos_iguais", "nem_todos_iguais"}

# Datas de cada etapa, que a tabela dos atos mostra além dos campos do Word
COLUNAS_ETAPAS = [("DATAEDOCS", "Data e-docs"), ("DATADIARIO", "Data Diário"),
                  ("_DATAPUBLICACAO", "Data publicação")]
# Valores que contam como caixinha marcada (o Edodocs grava SIM)
MARCADO = {"SIM", "S", "X", "1", "TRUE", "VERDADEIRO", "OK"}

# "Nossa Identificação" no e-docs: a coluna NR_DOCUMENTO_INTERNO do banco deles aceita 50
LIMITE_TITULO_EDOCS = 50

MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto",
         "setembro", "outubro", "novembro", "dezembro"]


def _arquivos(pasta: Path, padrao: str) -> list[Path]:
    return sorted(p for p in pasta.glob(padrao) if not p.name.startswith("~$"))


def _caminho(pasta: Path, valor: str) -> Path | None:
    """Caminho salvo no processo.json: absoluto ou relativo à pasta do processo."""
    if not valor:
        return None
    caminho = Path(valor)
    return caminho if caminho.is_absolute() else pasta / caminho


class Processo:
    def __init__(self, pasta: Path, config: dict | None = None):
        self.pasta = pasta
        self.config = json.loads(json.dumps(PADRAO)) | (config or {})

    @property
    def nome(self) -> str:
        return self.pasta.name

    def __getitem__(self, chave: str):
        return self.config[chave]

    def salvar(self) -> None:
        import main
        main.guardar_versao(self.pasta)  # pasta compartilhada: dá para voltar se alguém errar
        self.pasta.mkdir(parents=True, exist_ok=True)
        (self.pasta / ARQUIVO).write_text(json.dumps(self.config, ensure_ascii=False, indent=2),
                                          encoding="utf-8")

    # --- Arquivos ---
    def modelo(self) -> Path | None:
        """O Word do processo (o configurado ou o único .docx da pasta)."""
        caminho = _caminho(self.pasta, self.config["modelo"])
        if caminho and caminho.is_file():
            return caminho
        modelos = _arquivos(self.pasta, "*.docx")
        return modelos[0] if modelos else None

    def planilha_padrao(self) -> Path | None:
        """A planilha do processo (a configurada ou a .xlsx da pasta). Cada pessoa ainda pode
        apontar outra em Configurações (vale só no computador dela)."""
        caminho = _caminho(self.pasta, self.config["planilha"])
        if caminho:
            return caminho if caminho.is_file() else None
        planilhas = _arquivos(self.pasta, "*.xlsx")
        return planilhas[0] if planilhas else None

    @property
    def aba(self) -> str | None:
        return self.config["aba"] or None

    def ler_planilha(self, planilha: Path) -> list[dict]:
        return atos.ler_planilha(planilha, self.aba, atos.normalizar(self.config["coluna_nome"]))

    # --- Valores dos campos ---
    def valores(self, dados: dict) -> dict:
        """Valor de cada campo do Word para uma linha da planilha (ou um ato com várias
        pessoas), seguindo as ligações. As ligações "condicao" vêm por último, para poderem
        olhar o valor das outras."""
        campos = dict(dados)
        campos.setdefault(atos.QUANTIDADE, str(len(atos.itens(dados))))
        vazios = set(campos.get("_VAZIOS", ()))
        condicoes = []
        for campo, regra in self.config["ligacoes"].items():
            nome = atos.normalizar(campo)
            if regra.get("tipo") == "condicao":
                condicoes.append((nome, regra))
                continue
            campos[nome] = valor_da_regra(regra, dados, manter_chaves=True)
            # texto vazio de propósito (ex.: célula de tabela que fica em branco: "se_vazio" "")
            if regra.get("tipo") in ("fixo", "opcoes") or "se_vazio" in regra:
                vazios.add(nome)
        # As regras olham o valor da planilha ("A PEDIDO é SIM"), não o texto que o campo de
        # mesmo nome escreve no ato ("a pedido, "); os outros campos (QUANTIDADE...) como estão
        planilha = {k: v for k, v in dados.items() if not str(k).startswith("_")}
        for nome, regra in condicoes:
            campos[nome] = escolher_texto(regra, {**campos, **planilha})
            vazios.add(nome)
        # {{CAMPO}} dentro do texto escolhido (", a contar de {{DATA POR EXTENSO}}"): duas
        # voltas, para um campo dentro de outro que também tem campo
        for _ in range(2):
            for campo in self.config["ligacoes"]:
                nome = atos.normalizar(campo)
                if "{{" in str(campos.get(nome, "")):
                    # Frase em que nenhum dado foi preenchido (", a partir de {{DATA}}" sem a
                    # data) some inteira, em vez de sair ", a partir de ."
                    dentro = [atos.normalizar(m.group(1)) for m in atos.CHAVES.finditer(campos[nome])]
                    pendentes = any("{{" in str(campos.get(c, "")) for c in dentro)
                    if dentro and not pendentes and not any(str(campos.get(c, "")).strip() for c in dentro):
                        campos[nome] = ""
                    else:
                        campos[nome] = trocar_chaves(campos[nome], campos)
        campos["_VAZIOS"] = vazios
        if self.config.get("titulo_no_ato") and self.config["titulo_na_primeira_linha"]:
            campos[atos.TITULO_NO_ATO] = self.titulo_no_ato(dados, campos)
        return campos

    # --- Várias pessoas no mesmo ato ---
    @property
    def coluna_juntar(self) -> str:
        """Coluna (normalizada) que junta linhas num ato só; "" = cada linha é um ato."""
        return atos.normalizar((self.config.get("juntar") or {}).get("coluna", ""))

    def agrupar(self, registros: list[dict]) -> list[dict]:
        """Os atos da planilha. Sem "juntar", cada linha é um ato (as linhas voltam como
        estão). Com ele, as linhas com o mesmo valor na coluna viram um ato só: os dados da
        1ª linha, mais _ITENS (todas as linhas) e _LINHAS (os números delas no Excel). Linha
        com a coluna vazia fica sozinha."""
        coluna = self.coluna_juntar
        if not coluna:
            return registros
        grupos: dict[str, list[dict]] = {}
        for r in registros:
            chave = atos.normalizar(r.get(coluna, "")) or f"_linha{r['_LINHA']}"
            grupos.setdefault(chave, []).append(r)
        saida = []
        for pessoas in grupos.values():
            ato = dict(pessoas[0])
            ato["_ITENS"] = pessoas
            ato["_LINHAS"] = [p["_LINHA"] for p in pessoas]
            saida.append(ato)
        return saida

    def colunas_da_pessoa(self) -> set[str]:
        """Colunas (normalizadas) que mudam de uma pessoa para outra do mesmo ato: as dos
        campos do trecho que repete no Word e as que as regras desses campos olham. As outras
        são do ato (iguais para todas as pessoas dele)."""
        colunas: set[str] = set()
        vistos: set[str] = set()
        pendentes = [c for modelo in self.modelos() for c in atos.campos_que_repetem(modelo)]
        while pendentes:
            nome = pendentes.pop()
            if nome in vistos or nome in atos.AUTOMATICOS:
                continue
            vistos.add(nome)
            regra = self.ligacao(nome)
            if regra.get("tipo") == "fixo":
                continue
            if regra.get("tipo") == "condicao":
                for caso in regra.get("casos") or []:
                    pendentes += [atos.normalizar(c.get("campo", "")) for c in caso.get("se") or []]
            elif regra.get("coluna"):
                colunas.add(atos.normalizar(regra["coluna"]))
            else:
                colunas.add(nome)
            textos = [c.get("texto", "") for c in regra.get("casos") or regra.get("opcoes") or []] \
                + [regra.get("senao", "")]
            pendentes += [atos.normalizar(m.group(1)) for m in atos.CHAVES.finditer(" ".join(textos))]
        return colunas

    def ato_da_linha(self, registros: list[dict], linha: int) -> dict | None:
        """O ato (com todas as pessoas dele) que contém a linha do Excel."""
        return next((a for a in self.agrupar(registros) if linha in linhas(a)), None)

    def modelos(self) -> list[Path]:
        """Todos os Words do processo: o principal e os outros (usados conforme condições)."""
        principal = self.modelo()
        saida = [principal] if principal else []
        for outro in self.config.get("outros_modelos") or []:
            caminho = _caminho(self.pasta, outro.get("modelo", ""))
            if caminho and caminho.is_file() and caminho not in saida:
                saida.append(caminho)
        return saida

    def modelo_do_ato(self, dados: dict) -> Path | None:
        """O Word deste ato: o primeiro dos outros_modelos cujas condições valem; senão o
        principal."""
        outros = self.config.get("outros_modelos") or []
        if outros:
            # Como nas regras: o valor da planilha ("TIPO é CONCURSO"), não o texto do campo
            campos = {**self.valores(dados), **{k: v for k, v in dados.items() if not str(k).startswith("_")}}
            for outro in outros:
                caminho = _caminho(self.pasta, outro.get("modelo", ""))
                se = outro.get("se") or []
                if caminho and caminho.is_file() and se and all(condicao_vale(c, campos) for c in se):
                    return caminho
        return self.modelo()

    def _por_modelo(self, registros: list[dict], montar) -> list:
        """montar(modelo, registros) para cada Word, devolvendo na ordem dos registros."""
        saida: list = [None] * len(registros)
        grupos: dict[Path, list[int]] = {}
        for i, dados in enumerate(registros):
            grupos.setdefault(self.modelo_do_ato(dados), []).append(i)
        for modelo, indices in grupos.items():
            for i, valor in zip(indices, montar(modelo, [registros[i] for i in indices])):
                saida[i] = valor
        return saida

    def textos(self, registros: list[dict], sem_titulo: bool = False) -> list[str]:
        sem_titulo = sem_titulo and self.config["titulo_na_primeira_linha"]
        return self._por_modelo(registros, lambda modelo, regs: atos.textos(
            modelo, regs, sem_titulo, self.valores))

    @property
    def vai_ao_diario(self) -> bool:
        """O processo é publicado no Diário Oficial (senão, só vai ao e-docs)."""
        return bool(self.config.get("vai_ao_diario", True))

    @property
    def precisa_anexo(self) -> bool:
        """Cada ato vai com anexo(s) diferente(s): o envio pergunta o arquivo a cada ato."""
        return bool(self.config.get("precisa_anexo"))

    def folhas(self, registros: list[dict], papel: bool = True, sem_titulo: bool = False) -> list[str]:
        """O ato como o papel do Word (cabeçalho, recuos...), para mostrar na tela. papel=False:
        só o texto, como fica no editor dos sites."""
        import folha
        sem_titulo = sem_titulo and self.config["titulo_na_primeira_linha"]
        return self._por_modelo(registros, lambda modelo, regs: folha.folhas(
            modelo, regs, self.valores, papel, sem_titulo))

    def salvar_word(self, dados: dict, destino: Path) -> Path:
        """O ato (uma linha ou um ato com várias pessoas) num Word preenchido, para imprimir."""
        return atos.salvar_ato(self.modelo_do_ato(dados), dados, destino, self.valores)

    def htmls(self, registros: list[dict], sem_titulo: bool = False) -> list[str]:
        sem_titulo = sem_titulo and self.config["titulo_na_primeira_linha"]
        return self._por_modelo(registros, lambda modelo, regs: atos.htmls(
            modelo, regs, sem_titulo, self.valores))

    # --- Campos obrigatórios ---
    @property
    def obrigatorios(self) -> list[str]:
        """Colunas da planilha que não podem ficar vazias, como a pessoa escreveu."""
        return [c for c in self.config.get("obrigatorios") or [] if str(c).strip()]

    def faltando(self, dados: dict) -> list[str]:
        """As obrigatórias vazias no ato: "MATRÍCULA" ou, num ato com várias pessoas,
        "MATRÍCULA (JOSE MARIA)"."""
        pessoas = atos.itens(dados)
        coluna_nome = atos.normalizar(self.config["coluna_nome"])
        saida = []
        for coluna in self.obrigatorios:
            nome = atos.normalizar(coluna)
            vazias = [p for p in pessoas if not str(p.get(nome, "")).strip()]
            if vazias and len(pessoas) > 1:
                quem = ", ".join(str(p.get(coluna_nome) or f"linha {p.get('_LINHA')}") for p in vazias)
                saida.append(f"{coluna} ({quem})")
            elif vazias:
                saida.append(coluna)
        return saida

    # --- Tabela dos atos ---
    @property
    def caixas(self) -> list[str]:
        """Nomes das colunas de marcar (caixinha), como a pessoa escreveu."""
        return [c for c in self.config.get("caixas") or [] if str(c).strip()]

    def colunas_tabela(self, campos: list[tuple[str, str]] | None = None) -> list[dict]:
        """Colunas da tabela dos atos, na ordem escolhida: [{"nome" (normalizado), "rotulo",
        "caixa", "ver"}]. Padrão: número do ato, datas das etapas, campos do Word e caixinhas.
        Coluna nova (campo novo no Word, caixinha nova) entra no fim, visível."""
        campos = self.campos() if campos is None else campos
        numero = self.config["coluna_numero"]
        base = [(atos.normalizar(numero), numero, False)] + [(n, r, False) for n, r in COLUNAS_ETAPAS]
        base += [(n, r, False) for n, r in campos]
        base += [(atos.normalizar(c), c, True) for c in self.caixas]
        disponiveis: dict[str, dict] = {}
        for nome, rotulo, caixa in base:
            if nome and nome not in disponiveis:
                disponiveis[nome] = {"nome": nome, "rotulo": rotulo, "caixa": caixa, "ver": True}
        saida = []
        for salva in self.config.get("colunas_tabela") or []:
            coluna = disponiveis.pop(salva.get("nome"), None)
            if coluna:
                saida.append(dict(coluna, ver=bool(salva.get("ver", True))))
        return saida + list(disponiveis.values())

    # --- Planilha criada pelo programa ---
    def colunas_planilha_nova(self, campos: list[tuple[str, str]] | None = None) -> list[dict]:
        """As colunas que uma planilha nova precisa para este processo: [{"nome", "tipo"
        ("texto", "numero", "data" ou "lista"), "opcoes"}]. A 1ª é a do número do ato (que
        também junta as pessoas de um mesmo ato); depois, uma por campo do Word, na ordem."""
        campos = self.campos() if campos is None else campos
        colunas: dict[str, dict] = {}

        def incluir(nome: str, tipo: str = "texto", opcoes: list[str] | None = None) -> None:
            nome = " ".join(str(nome or "").split())
            chave = atos.normalizar(nome)
            if not chave or chave in atos.AUTOMATICOS:
                return
            atual = colunas.setdefault(chave, {"nome": nome, "tipo": tipo, "opcoes": []})
            if opcoes:
                atual["tipo"] = "lista"
                atual["opcoes"] = list(dict.fromkeys(atual["opcoes"] + opcoes))
            elif atual["tipo"] == "texto":
                atual["tipo"] = tipo

        incluir(self.config["coluna_numero"])
        juntar = (self.config.get("juntar") or {}).get("coluna")
        if juntar:
            incluir(juntar)
        for _, rotulo in campos:
            regra = self.ligacao(rotulo)
            tipo = regra.get("tipo", "coluna")
            coluna = regra.get("coluna") or rotulo
            if tipo == "fixo":
                continue
            if tipo == "extenso":
                incluir(coluna, "numero")
            elif tipo == "data_extenso":
                incluir(coluna, "data")
            elif tipo == "periodo":
                incluir(coluna, "data")
                if regra.get("ate"):
                    incluir(regra["ate"], "data")
            elif tipo == "genero":
                incluir(coluna, opcoes=["F", "M"])
            elif tipo == "opcoes":
                incluir(coluna, opcoes=[o.get("valor", "") for o in regra.get("opcoes") or [] if o.get("valor")])
            elif tipo == "condicao":
                # as colunas que as regras olham (ex.: A PEDIDO, VÍNCULO), com os valores delas
                for caso in regra.get("casos") or []:
                    for cond in caso.get("se") or []:
                        valores = [v.strip() for v in str(cond.get("valor", "")).split(";") if v.strip()]
                        if atos.normalizar(cond.get("campo", "")) == "SEXO":  # a regra só cita F
                            valores = ["F", "M"]
                        incluir(cond.get("campo", ""), opcoes=valores if cond.get("op") == "igual" else None)
            else:
                incluir(coluna)
        for caixa in self.caixas:
            incluir(caixa, opcoes=["SIM"])
        return list(colunas.values())

    # --- Número e títulos ---
    def numero(self, dados: dict) -> str:
        """Valor da coluna do número do ato, como está na planilha (ex.: 'SMA 1542/2026')."""
        return dados.get(atos.normalizar(self.config["coluna_numero"]), "")

    def numero_do_ato(self, dados: dict) -> str:
        """Só o número/ano (ex.: '1542/2026'); sem esse formato, o valor como está."""
        bruto = self.numero(dados)
        m = _NUMERO.search(bruto)
        return re.sub(r"\s", "", m.group(0)) if m else bruto.strip()

    def descricao(self, dados: dict) -> str:
        """Como o ato aparece nas mensagens: 'SMA 1542/2026 · FULANA' (com várias pessoas:
        'SMA 1542/2026 · FULANA e mais 2')."""
        nome = dados.get(atos.normalizar(self.config["coluna_nome"]), "")
        outras = len(atos.itens(dados)) - 1
        if nome and outras:
            nome += f" e mais {outras}"
        partes = (self.numero(dados), nome)
        return " · ".join(p for p in partes if p) or "?"

    def montar_titulo(self, modelo: str, dados: dict, abreviar: int = 0, valores: dict | None = None) -> str:
        """Troca cada {COLUNA} do modelo do título pelo valor. Lança RuntimeError se faltar.
        abreviar (valores de 3 palavras ou mais): 1 = as do meio em iniciais ("CELIA
        EXEMPLO TEIXEIRA" → "CELIA G. TEIXEIRA"; "da", "de"... saem); 2 = só a primeira e
        a última ("CELIA TEIXEIRA")."""
        valores = self.valores(dados) if valores is None else valores

        def trocar(m) -> str:
            nome = atos.normalizar(m.group(1))
            if nome == NUMERO_DO_ATO:
                valor = self.numero_do_ato(dados)
            elif nome == ANO_DO_ATO:
                ano = _NUMERO.search(self.numero(dados))
                valor = ano.group(2) if ano else str(date.today().year)
            else:
                valor = str(valores.get(nome, "")).strip()
            if not valor:
                raise RuntimeError(f"coluna {m.group(1)} vazia na planilha (usada no título)")
            palavras = valor.split()
            if abreviar and len(palavras) >= 3:
                meio = [] if abreviar >= 2 else [
                    p[0] + "." for p in palavras[1:-1] if p.lower() not in ("da", "de", "do", "das", "dos", "e")]
                valor = " ".join([palavras[0], *meio, palavras[-1]])
            return valor

        return _TOKEN.sub(trocar, modelo).strip()

    def titulo_edocs(self, dados: dict) -> str:
        """A "Nossa identificação" no e-docs. Ato com várias pessoas: termina com "E OUTROS".
        O e-docs aceita até LIMITE_TITULO_EDOCS caracteres (mais que isso, ele recusa o
        documento): passando, os nomes do meio viram iniciais ("CELIA G. TEIXEIRA"), depois
        saem ("PEDRO MARTINS E OUTROS") e, se ainda não couber, o título é cortado numa
        palavra inteira."""
        sufixo = " E OUTROS" if len(atos.itens(dados)) > 1 else ""
        modelo = self.config["titulo_edocs"]
        titulo = self.montar_titulo(modelo, dados)
        if sufixo and titulo.upper().endswith(sufixo.strip()):
            sufixo = ""
        for nivel in (1, 2):  # 1: nomes do meio em iniciais; 2: só o primeiro e o último nome
            if len(titulo + sufixo) > LIMITE_TITULO_EDOCS:
                titulo = self.montar_titulo(modelo, dados, abreviar=nivel)
        if len(titulo + sufixo) > LIMITE_TITULO_EDOCS:
            cabe = LIMITE_TITULO_EDOCS - len(sufixo)
            titulo = titulo[:cabe].rsplit(" ", 1)[0] if " " in titulo[:cabe] else titulo[:cabe]
        return (titulo + sufixo).strip()

    def titulo_no_ato(self, dados: dict, valores: dict | None = None) -> str:
        """A 1ª linha do ato ("DECRETO S/Nº - EXONERA {NOME}"), sem o limite do e-docs: nome
        inteiro e, com várias pessoas, "E OUTROS". Vazio (fica a linha do Word) se o processo
        não tem esse título ou falta alguma coluna."""
        modelo = self.config.get("titulo_no_ato")
        if not modelo:
            return ""
        try:
            titulo = self.montar_titulo(modelo, dados, valores=valores)
        except RuntimeError:
            return ""
        if len(atos.itens(dados)) > 1 and not titulo.upper().endswith("E OUTROS"):
            titulo += " E OUTROS"
        return titulo

    def titulo_diario(self, dados: dict) -> str:
        return self.montar_titulo(self.config["titulo_diario"], dados)

    # --- Conferência no Diário Oficial ---
    def padrao_conferencia(self) -> re.Pattern:
        """Como o título do Diário aparece no PDF: o texto fixo antes do primeiro {campo},
        seguido de número/ano. Tolera Nº/N°/No, espaços e quebras de linha."""
        prefixo = self.config["titulo_diario"].split("{", 1)[0]
        partes = [r"N\s*[º°oO.]*" if re.fullmatch(r"N[º°oO.]*", p, re.IGNORECASE) else re.escape(p)
                  for p in prefixo.split()]
        inicio = r"\s+".join(partes)
        return re.compile((inicio + r"\s*" if inicio else "") + r"([\d.]+)\s*/\s*(\d{4})",
                          re.IGNORECASE)

    def chave_conferencia(self, dados: dict) -> str | None:
        """Número/ano do ato como será procurado no PDF do Diário (ex.: '1542/2026')."""
        import diario
        try:
            titulo = self.titulo_diario(dados)
        except RuntimeError:
            return None
        m = self.padrao_conferencia().search(titulo)
        return diario.chave(*m.groups()) if m else None

    # --- Ligações ---
    def ligacao(self, campo: str) -> dict:
        """A ligação do campo (a salva ou, sem ela, a coluna de mesmo nome)."""
        for nome, regra in self.config["ligacoes"].items():
            if atos.normalizar(nome) == atos.normalizar(campo):
                return regra
        return {"tipo": "coluna", "coluna": campo}

    def campos_sem_ligacao(self, colunas: list[str]) -> list[str]:
        """Campos do Word que não têm de onde tirar valor (para avisar antes de enviar)."""
        existentes = {atos.normalizar(c) for c in colunas}
        faltam = []
        for nome, rotulo in self.campos():
            regra = self.ligacao(rotulo)
            coluna = regra.get("coluna")
            if regra["tipo"] not in ("fixo", "condicao") and coluna and atos.normalizar(coluna) not in existentes:
                faltam.append(rotulo)
        return faltam

    def campos(self) -> list[tuple[str, str]]:
        """Campos de todos os Words do processo e os citados no texto das regras
        ({{DATA POR EXTENSO}} dentro de "retroativo a ..."), sem repetir e sem os automáticos."""
        vistos: dict[str, str] = {}
        for modelo in self.modelos():
            for nome, rotulo in atos.campos_do_modelo(modelo):
                vistos.setdefault(nome, rotulo)
        for regra in list(self.config["ligacoes"].values()):
            if regra.get("tipo") in ("condicao", "opcoes"):
                textos = [c.get("texto", "") for c in regra.get("casos") or regra.get("opcoes") or []] \
                    + [regra.get("senao", "")]
                for m in atos.CHAVES.finditer(" ".join(textos)):
                    vistos.setdefault(atos.normalizar(m.group(1)), atos._rotulo(m.group(1)))
        return [(n, r) for n, r in vistos.items() if n not in atos.AUTOMATICOS]


def linhas(dados: dict) -> list[int]:
    """As linhas do Excel de um ato (várias, se ele junta pessoas)."""
    return dados.get("_LINHAS") or [dados["_LINHA"]]


def _numero(texto: str) -> float | None:
    texto = str(texto or "").strip().replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None


def condicao_vale(cond: dict, campos: dict) -> bool:
    """A condição vale para estes campos? (ver CONDIÇÃO no início do arquivo)"""
    nome = atos.normalizar(cond.get("campo", ""))
    op = cond.get("op", "igual")
    valor = str(campos.get(nome, "")).strip()
    alvo = str(cond.get("valor", ""))
    if op in ("todos_iguais", "nem_todos_iguais"):
        diferentes = {atos.normalizar(p.get(nome, "")) for p in atos.itens(campos)}
        return (len(diferentes) <= 1) == (op == "todos_iguais")
    if op == "preenchido":
        return bool(valor)
    if op == "vazio":
        return not valor
    if op in ("maior", "menor"):
        a, b = _numero(valor), _numero(alvo)
        return a is not None and b is not None and (a > b if op == "maior" else a < b)
    if op == "contem":
        return atos.normalizar(alvo) in atos.normalizar(valor)
    iguais = atos.normalizar(valor) in {atos.normalizar(v) for v in alvo.split(";")}
    return iguais if op == "igual" else not iguais


def escolher_texto(regra: dict, campos: dict) -> str:
    """Ligação "condicao": o texto do primeiro caso que vale; senão o "senao"."""
    for caso in regra.get("casos") or []:
        if all(condicao_vale(c, campos) for c in caso.get("se") or []):
            return caso.get("texto", "")
    return regra.get("senao", "")


def trocar_chaves(texto: str, campos: dict, manter: bool = False) -> str:
    """{{CAMPO}} dentro de um texto vira o valor do campo (vazio se não houver). Com manter,
    os campos ficam como estão, para serem trocados depois que todos tiverem o valor final
    ({{DATA}} por extenso, e não como está na coluna)."""
    if manter:
        return texto

    def valor(m) -> str:
        bruto = str(campos.get(atos.normalizar(m.group(2)), ""))
        # Campo vazio leva junto a vírgula ou a barra de antes (", {{NÚCLEO}}" sem núcleo
        # não deixa ", , da Secretaria" no ato)
        return (m.group(1) or "") + bruto if bruto.strip() else ""
    return re.sub(r"(\s*[,/]\s*)?" + atos.CHAVES.pattern, valor, texto)


def descrever_condicao(cond: dict) -> str:
    """'A PEDIDO é SIM', para mostrar na tela."""
    op = cond.get("op", "igual")
    partes = [cond.get("campo", "?"), OPERADORES.get(op, op)]
    if op not in SEM_VALOR:
        partes.append(" ou ".join(v.strip() for v in str(cond.get("valor", "")).split(";")))
    return " ".join(partes)


def valor_da_regra(regra: dict, dados: dict, manter_chaves: bool = False) -> str:
    """Valor de um campo segundo a ligação (ver tipos no início do arquivo). manter_chaves:
    {{CAMPO}} que não é coluna da linha fica no texto (Processo.valores troca depois, quando
    já sabe o valor de todos os campos)."""
    if regra.get("formato"):  # MAIÚSCULAS, Primeiras Maiúsculas... (o mesmo do Preencher sozinho)
        from consultas import formatar
        return formatar(valor_da_regra(dict(regra, formato=""), dados, manter_chaves), regra["formato"])
    tipo = regra.get("tipo", "coluna")
    if tipo == "fixo":
        return regra.get("texto", "")
    if tipo == "condicao":
        campos = dict(dados)
        campos.setdefault(atos.QUANTIDADE, str(len(atos.itens(dados))))
        return trocar_chaves(escolher_texto(regra, campos), campos, manter_chaves)
    bruto = str(dados.get(atos.normalizar(regra.get("coluna", "")), "")).strip()
    if tipo == "opcoes":
        for opcao in regra.get("opcoes") or []:
            if atos.normalizar(opcao.get("valor", "")) == atos.normalizar(bruto):
                return trocar_chaves(opcao.get("texto", ""), dados, manter_chaves)
        return ""
    if tipo == "coluna":
        return bruto or regra.get("se_vazio", "")
    if tipo == "extenso":
        numero = bruto.replace(".", "")
        return atos.por_extenso(int(numero)) if numero.isdigit() else ""
    if tipo == "genero":
        letra = atos.normalizar(bruto)[:1]
        return regra.get("feminino", "") if letra == "F" else regra.get("masculino", "") if letra == "M" else ""
    if tipo == "data_extenso":
        dia = atos.ler_data(bruto)
        # Dia 1 é escrito "1º", como nos atos ("a partir de 1º de setembro de 2026")
        return f"{'1º' if dia.day == 1 else dia.day} de {MESES[dia.month - 1]} de {dia.year}" if dia else ""
    if tipo == "periodo":
        fim = atos.ler_data(str(dados.get(atos.normalizar(regra.get("ate", "")), "")).strip())
        return periodo(atos.ler_data(bruto), fim)
    return ""


def periodo(inicio, fim) -> str:
    """De uma data a outra, como nas portarias: "09 a 28 de novembro de 2026", "23 de novembro
    a 07 de dezembro de 2026", "21 de dezembro de 2026 a 04 de janeiro de 2027". Dia com dois
    algarismos; sem o fim (ou igual ao início), só a data."""
    if not inicio:
        return ""

    def extenso(d) -> str:
        return f"{d.day:02d} de {MESES[d.month - 1]} de {d.year}"

    if not fim or fim == inicio:
        return extenso(inicio)
    if (inicio.year, inicio.month) == (fim.year, fim.month):
        return f"{inicio.day:02d} a {extenso(fim)}"
    if inicio.year == fim.year:
        return f"{inicio.day:02d} de {MESES[inicio.month - 1]} a {extenso(fim)}"
    return f"{extenso(inicio)} a {extenso(fim)}"


def carregar(pasta: Path) -> Processo:
    """O processo da pasta (sem processo.json, valem os padrões)."""
    try:
        config = json.loads((pasta / ARQUIVO).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        config = {}
    return Processo(pasta, config)
