"""Acesso ao Edocs por requisições HTTP, sem navegador.

O Edocs é feito em JSF (Apache MyFaces Trinidad). Cada tela é um formulário
"corpo:formulario"; para "clicar" num botão, o navegador envia o formulário inteiro com
source=<id do botão> e o javax.faces.ViewState da tela, que muda a cada resposta. A classe
Tela faz exatamente isso: lê os campos da página, troca os que interessam e envia.

Fluxo de um ato (mapeado da gravação capturas/edocs_lancamento.har):
  1. login (j_security_check)
  2. instrumentosjuridicosman?param2=S  formulário novo → Setor → Gravar (F10): cria o
     documento e devolve o número dele (chPlc)
  3. vistosdocumentosman?param=<doc>        um visto por pessoa (busca o CPF em servidoressel)
  4. assinantesdocumentosman?param=<doc>    um assinante por pessoa, Remetente: Sim
  5. destinatariosdocumentosman?param=<doc> botão 4: grupo de destinatários
  6. Enviar para vistos                      botão do rodapé da tela do documento

Testes (usam o usuário salvo no programa, opção 4 → 1):
    python edocs_http.py              só o login
    python edocs_http.py rascunho     cria o 1º ato pendente como rascunho, SEM enviar
    python edocs_http.py botoes       lista os botões da tela do documento (acha o source)
    python edocs_http.py enviar       cria o rascunho e ENVIA (usa o source configurado)
"""
import html as html_lib
import re
import sys
import time
from pathlib import Path
from urllib.parse import urljoin

import lxml.html
# O Akamai do Edocs bloqueia (403 Reference #18) o fingerprint TLS do requests/OpenSSL.
# O curl_cffi faz o handshake TLS igual ao de um Chrome de verdade e passa.
from curl_cffi import requests

URL = "https://webedocspmu.uberlandia.mg.gov.br/webedocspmu/"
TELAS = URL + "f/n/"
TIMEOUT = 30
# Alguns servidores recusam requisições sem cara de navegador
NAVEGADOR = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
             "(KHTML, like Gecko) Chrome/130.0 Safari/537.36")
# Versão do Chrome que o curl_cffi imita no handshake TLS (JA3/JA4)
IMPERSONAR = "chrome124"
# Pausa entre uma requisição e a seguinte, para não parecer rajada para o Akamai
PAUSA_REQUISICAO = 0.25

# A sessão cai depois de 15 minutos parada ("Tempo sessão: 15.0m" no rodapé do menu)
TEMPO_SESSAO_MIN = 15

FORM = "corpo:formulario"
F = FORM + ":"  # prefixo dos campos: F + "setor" = "corpo:formulario:setor"
TIPO_ATO = "ATOS ADMINISTRATIVOS"
# A lista de Setor tem duas opções com o mesmo texto (NUCLEO DE REGISTRO DE PESSOAL);
# a correta é a 2ª. A 1ª está errada.
SETOR_POSICAO = 2
# Grupo de destinatários (Dest. Publicador) escolhido no processo, pelo texto; vazio = o 1º
GRUPO_DESTINATARIO = ""

# Botões do rodapé da tela do documento (instrumentosjuridicosman?chPlc=<doc>).
# O de envio é o 4º (xpath .../div[2]/div/button[4], texto "Enviar p/ Visto"); o onclick
# dele faz submitForm('corpo:formulario',1,{source:'corpo:formulario:botaoEnviarParaVistos'}).
BOTAO_ENVIAR = F + "botaoEnviarParaVistos"

# As duas telas de documento novo do e-docs. Cada conta tem uma, a outra ou as duas.
# Documentos Genéricos (capturas: caixa_entrada.har): Iniciais do Nome no Rodapé, "Incluir
# Visto?" (Não já vem marcado), sem destinatários, e o envio é "Enviar p/ Assinatura"
INSTRUMENTO, GENERICO = "instrumento", "generico"
MODELOS = {
    INSTRUMENTO: {"nome": "Instrumentos Jurídicos/Normativos", "pagina": "instrumentosjuridicosman",
                  "editor": "tinyEditorInstJurCorpo", "destinatarios": True},
    GENERICO: {"nome": "Documentos Genéricos", "pagina": "atestedocumentosman",
               "editor": "tinyEditorAtestDocCorpo", "destinatarios": False},
}


class LoginRecusado(Exception):
    """Usuário ou senha não aceitos pelo Edocs."""


class SessaoExpirada(Exception):
    """O Edocs voltou para a tela de login (sessão caiu)."""


class ErroEdocs(RuntimeError):
    """O Edocs não fez o que foi pedido (a mensagem diz o quê)."""


class DocumentoEmMontagem(ErroEdocs):
    """O e-docs abriu um documento em montagem no lugar de um novo (nada foi gravado)."""


def _no_menu(html: str) -> bool:
    """Página de quem está logado: tem o link Desconectar e não tem o campo de senha."""
    return "Desconectar" in html and 'name="j_password"' not in html


class _SessaoEducada(requests.Session):
    """Sessão que espera PAUSA_REQUISICAO antes de cada requisição, para o Edocs (e o
    Akamai na frente dele) não verem uma rajada de chamadas."""

    def request(self, *args, **kwargs):
        time.sleep(PAUSA_REQUISICAO)
        return super().request(*args, **kwargs)


def entrar(usuario: str, senha: str) -> requests.Session:
    """Faz o login e devolve a sessão (o cookie JSESSIONID fica guardado nela e vai junto
    em todas as próximas requisições). Lança LoginRecusado se o Edocs não aceitar."""
    sessao = _SessaoEducada(impersonate=IMPERSONAR)
    sessao.headers["User-Agent"] = NAVEGADOR

    # 1. Abre a página de login: o servidor cria a sessão. O login do Java
    #    (j_security_check) só é aceito dentro de uma sessão já aberta.
    sessao.get(URL, timeout=TIMEOUT).raise_for_status()

    # 2. Envia o formulário de login, como o botão ENTRAR faz
    resposta = sessao.post(URL + "j_security_check",
                           data={"j_username": usuario, "j_password": senha},
                           headers={"Referer": URL}, timeout=TIMEOUT)
    resposta.raise_for_status()
    if not _no_menu(resposta.text):
        raise LoginRecusado("o Edocs não abriu o menu depois do login")
    return sessao


def logado(sessao: requests.Session) -> bool:
    """A sessão ainda vale? (o servidor derruba depois de TEMPO_SESSAO_MIN parado)"""
    resposta = sessao.get(URL, timeout=TIMEOUT)
    return resposta.ok and _no_menu(resposta.text)


# --- Telas (formulários JSF) -------------------------------------------------

def _campos(form) -> dict[str, str]:
    """O que o navegador enviaria do formulário: campos com nome, sem botões, só os
    rádios/caixas marcados, a opção selecionada de cada lista. Desabilitados não vão."""
    dados = {}
    for el in form.xpath(".//input|.//select|.//textarea"):
        nome = el.get("name")
        if not nome or el.get("disabled") is not None:
            continue
        if el.tag == "select":
            opcoes = el.xpath(".//option")
            marcada = next((o for o in opcoes if o.get("selected") is not None),
                           opcoes[0] if opcoes else None)
            dados[nome] = marcada.get("value", "") if marcada is not None else ""
        elif el.tag == "textarea":
            dados[nome] = el.text or ""
        else:
            tipo = (el.get("type") or "text").lower()
            if tipo in ("submit", "button", "image", "reset", "file"):
                continue
            if tipo in ("checkbox", "radio") and el.get("checked") is None:
                continue
            dados[nome] = el.get("value", "")
    return dados


def _mensagens(doc) -> list[str]:
    """Mensagens de aviso/erro da tela (para explicar uma recusa)."""
    textos = []
    for el in doc.xpath('//*[contains(@class,"message") or contains(@class,"Message")'
                        ' or contains(@class,"erro") or contains(@class,"alerta")]'):
        t = " ".join(" ".join(el.itertext()).split())
        if t and t not in textos:
            textos.append(t)
    return textos[:3]


class Tela:
    """Uma tela do Edocs: os campos atuais do formulário e o ViewState dela."""

    def __init__(self, sessao: requests.Session, resposta: requests.Response):
        resposta.raise_for_status()
        if 'name="j_password"' in resposta.text:
            raise SessaoExpirada("o Edocs voltou para a tela de login")
        self.sessao = sessao
        self.status = resposta.status_code
        self.url = resposta.url
        self.html = resposta.text
        conteudo = resposta.content
        if b"<?Tr-XHR-Response-Type" in conteudo[:200]:
            # Resposta parcial (como o navegador recebe): o HTML vem dentro de CDATA
            conteudo = b"".join(re.findall(rb"<!\[CDATA\[(.*?)\]\]>", conteudo, re.S))
            self.doc = lxml.html.fromstring(conteudo, parser=lxml.html.HTMLParser(encoding="utf-8"))
        else:
            self.doc = lxml.html.fromstring(conteudo)
        forms = self.doc.xpath(f'//form[@id="{FORM}"]')
        if not forms:
            raise ErroEdocs(f"a tela {self.url.split('/f/n/')[-1]} veio sem o formulário")
        self.acao = urljoin(self.url, forms[0].get("action"))
        self.campos = _campos(forms[0])

    def opcoes(self, campo: str) -> list[tuple[str, str]]:
        """[(valor, texto)] da lista, sem o '[Selecione]'."""
        return [(o.get("value"), " ".join((o.text or "").split()))
                for o in self.doc.xpath(f'//select[@name="{campo}"]/option') if o.get("value")]

    def opcao(self, campo: str, texto: str | None = None, posicao: int | None = None) -> str:
        """Valor da opção pelo texto (exato ou contido) ou pela posição (1 = primeira)."""
        opcoes = self.opcoes(campo)
        if texto:
            alvo = texto.upper()
            achada = (next((v for v, t in opcoes if t.upper() == alvo), None)
                      or next((v for v, t in opcoes if alvo in t.upper()), None))
            if achada is not None:
                return achada
        if posicao and len(opcoes) >= posicao:
            return opcoes[posicao - 1][0]
        nome = campo.split(":")[-1]
        raise ErroEdocs(f"opção {texto or posicao} não existe em {nome} "
                        f"(opções: {', '.join(t for _, t in opcoes) or 'nenhuma'})")

    def enviar(self, botao: str, valores: dict | None = None, parcial: bool = False) -> "Tela":
        """'Clica' no botão: envia o formulário com os valores trocados. parcial = botões
        de janela (popup), que o site envia como atualização parcial (autosub)."""
        dados = dict(self.campos)
        dados.update(valores or {})
        dados["source"] = botao
        dados.setdefault(botao, "")
        if parcial:
            dados.update(event="autosub", partial="true")
        resposta = self.sessao.post(self.acao, data=dados, headers={"Referer": self.url},
                                    timeout=TIMEOUT)
        return Tela(self.sessao, resposta)

    def linhas(self) -> list[str]:
        """Texto das linhas da lista da tela (vistos, assinantes, destinatários)."""
        saida = []
        for caixa in self.doc.xpath('//input[contains(@name,"idExclusao")]'):
            tr = next(iter(caixa.iterancestors("tr")), None)
            if tr is not None:
                saida.append(" ".join(" ".join(tr.itertext()).split()))
        return saida

    def erro(self, o_que: str) -> ErroEdocs:
        msgs = _mensagens(self.doc)
        return ErroEdocs(o_que + (f" ({'; '.join(msgs)})" if msgs else ""))


def abrir(sessao: requests.Session, caminho: str) -> Tela:
    return Tela(sessao, sessao.get(TELAS + caminho, timeout=TIMEOUT))


# --- Etapas do ato -----------------------------------------------------------

def formatar_cpf(cpf: str) -> str:
    """'1234567890' → '012.345.678-90'. A busca do Edocs só acha com o CPF formatado
    (já para gravar visto/assinante ele vai só com os dígitos)."""
    d = re.sub(r"\D", "", str(cpf)).zfill(11)
    return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"


def _mesmo_cpf(a: str, b: str) -> bool:
    return re.sub(r"\D", "", a).lstrip("0") == re.sub(r"\D", "", b).lstrip("0")


def _selecao(clique: str) -> dict:
    """onclick="devolveSelecao('chave#valor,chave#valor,...')" → {chave: valor}."""
    m = re.search(r"devolveSelecao\('(.*)'\)", clique, re.S)
    if not m:
        return {}
    return dict(par.split("#", 1) for par in re.split(r",(?=\w+#)", m.group(1)) if "#" in par)


def _js_texto(valor: str) -> str:
    """Texto que veio dentro do JavaScript da tela, para mostrar: "S\u00E9rgio" → "Sérgio"."""
    valor = re.sub(r"\\u([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1), 16)), valor)
    return " ".join(valor.split())


def servidores_da_tela(tela: Tela) -> list[dict]:
    """Resultados da janela de seleção de servidores. Cada linha da tabela (Matrícula | CPF |
    Nome | Cargo | Lotação) tem os dados que as telas de visto/assinante precisam no onclick."""
    pessoas = []
    for el in tela.doc.xpath('//*[contains(@onclick, "devolveSelecao(")]'):
        dados = _selecao(el.get("onclick"))
        if not dados.get("oidServidoreGrh"):
            continue
        tr = el if el.tag == "tr" else next(iter(el.iterancestors("tr")), None)
        celulas = [" ".join(" ".join(td.itertext()).split()) for td in tr.xpath("./td")]             if tr is not None else []
        pessoas.append({
            "oid": dados["oidServidoreGrh"],
            # Para gravar, o CPF vai só com os 11 dígitos (com zero na frente, sem pontos)
            "cpf": re.sub(r"\D", "", dados.get("nrCpf", "")).zfill(11),
            "nome": _js_texto(dados.get("nmPessoa", "")),
            "matricula": celulas[0] if len(celulas) >= 5 else "",
            "cargo": _js_texto(dados.get("dsCargo", "")),
            "lotacao": _js_texto(dados.get("dsLotacao", "")),
            "lookup": dados.get("servidoresselLookup", ""),
        })
    return list({p["oid"]: p for p in pessoas}.values())


def pesquisar_servidores(sessao: requests.Session, termo: str) -> list[dict]:
    """Pesquisa pessoas como o botão de lupa dos vistos: por CPF (11 dígitos, com ou sem
    pontos), matrícula (outros números; precisa do dígito verificador) ou nome."""
    termo = " ".join(str(termo).split())
    digitos = re.sub(r"[\s.\-/]", "", termo)
    if not termo:
        return []
    if digitos.isdigit():
        tentativas = ([F + "nrCpf", F + "nrMatricula"] if len(digitos) == 11
                      else [F + "nrMatricula"])
    else:
        tentativas = [F + "nmPessoa"]
    for campo in tentativas:
        valor = formatar_cpf(digitos) if campo.endswith("nrCpf") else (
            digitos if digitos.isdigit() else termo)
        tela = abrir(sessao, "servidoressel?&modoJanelaPlc=popup&delimPropsPlc=,")
        tela = tela.enviar(F + "botaoAcaoPesquisar", {campo: valor}, parcial=True)
        if pessoas := servidores_da_tela(tela):
            return pessoas
    return []


def buscar_servidor(sessao: requests.Session, cpf: str) -> dict:
    """A pessoa do CPF (oid = id interno, cpf, lookup, nome...), para gravar visto/assinante."""
    for pessoa in pesquisar_servidores(sessao, formatar_cpf(cpf)):
        if _mesmo_cpf(pessoa["cpf"], cpf):
            return pessoa
    raise ErroEdocs(f"CPF {formatar_cpf(cpf)} não encontrado no Edocs")


def opcoes_do_documento(sessao: requests.Session, setor_posicao: int = 1, modelo: str = INSTRUMENTO) -> dict:
    """Setores e tipos que ESTE usuário vê na tela de novo documento (cada pessoa tem os
    seus). Os tipos podem mudar com o setor: o site recarrega a tela ao trocar o setor, e
    aqui também (sem gravar nada). Devolve {"setores": [texto], "tipos": [texto]}; o setor
    fica pela posição, porque pode haver dois com o mesmo nome. modelo: qual das duas telas
    (Instrumentos Jurídicos/Normativos ou Documentos Genéricos)."""
    tela = abrir(sessao, MODELOS[modelo]["pagina"] + "?param2=S")
    setores = [texto for _, texto in tela.opcoes(F + "setor")]
    if setores:
        valor = tela.opcao(F + "setor", posicao=min(max(setor_posicao, 1), len(setores)))
        tela = tela.enviar(F + "btRefreshNovoEditor", {F + "setor": valor})
    return {"setores": setores, "tipos": [texto for _, texto in tela.opcoes(F + "tipo")]}


def criar_documento(sessao: requests.Session, titulo: str, corpo_html: str,
                    avisar=lambda msg: None, data: str = "", modelo: str = INSTRUMENTO,
                    extras: dict | None = None) -> str:
    """Formulário novo → Setor → Gravar (o F10). Devolve o número do documento criado.
    data: dd/mm/aaaa do campo Data (vazio = a que o site traz, hoje). modelo: a tela do
    documento; extras: campos só dela ({"iniciais": ..., "incluir_visto": bool} nos Genéricos)."""
    tela = abrir(sessao, MODELOS[modelo]["pagina"] + "?param2=S")
    # Com um documento em montagem, o e-docs não abre um formulário novo: abre o em montagem
    # mais recente, já preenchido. Gravar aqui escreveria por cima dele: para antes
    aberto = tela.campos.get(F + "nrNossoNumero", "").strip()
    if aberto:
        raise DocumentoEmMontagem(
            f"o e-docs abriu o documento em montagem “{aberto}” ({tela.campos.get(F + 'dtDocumento', '')}) no "
            "lugar de um documento novo. Nada foi gravado. Envie ou cancele os documentos em montagem no "
            "e-docs (ou aceite cancelar quando o Edodocs perguntar) e envie de novo")
    corpo = {F + "dsCorpoDocumento": corpo_html, MODELOS[modelo]["editor"]: corpo_html}
    if modelo == GENERICO:
        extras = extras or {}
        # "Incluir Visto?": 0 = Sim, 1 = Não (o que o site já traz marcado)
        corpo |= {F + "dsIniciaisNomeRodape": str(extras.get("iniciais", "") or ""),
                  F + "idVistoAteste": "0" if extras.get("incluir_visto") else "1"}

    # O site recarrega a tela ao trocar o setor, como no navegador
    setor = tela.opcao(F + "setor", posicao=SETOR_POSICAO)
    tela = tela.enviar(F + "btRefreshNovoEditor", {F + "setor": setor, **corpo})
    avisar(f"Setor: {dict(tela.opcoes(F + 'setor')).get(setor, setor)}")

    tipo = tela.opcao(F + "tipo", texto=TIPO_ATO)
    tela = tela.enviar(F + "botaoGravarEdocs", {
        F + "setor": setor, F + "tipo": tipo, F + "nrNossoNumero": titulo, **corpo,
        **({F + "dtDocumento": data} if data else {})})
    # Gravado, o site redireciona para a tela do documento: ...?chPlc=<número>
    m = re.search(r"chPlc=(\d+)", tela.url)
    if not m:
        raise tela.erro("o Edocs não confirmou a gravação do documento")
    avisar(f"Documento {m.group(1)} gravado ({MODELOS[modelo]['nome']}) · Tipo: {TIPO_ATO} · Título: {titulo}")
    return m.group(1)


def _adicionar_pessoas(sessao, doc: str, pagina: str, campo_ordem: str, cpfs: list[str],
                       extras, nome: str, avisar) -> None:
    """Parte comum de vistos e assinantes: para cada CPF, busca a pessoa e grava na ordem
    (1ª pessoa = 1º, 2ª = 2º...). Confere na lista da tela se ficou gravada. Quem já
    estiver na lista não é adicionado de novo. extras(ordem): campos a mais da tela."""
    def na_lista(tela, cpf):
        # A lista mostra "1. 5750492605 - NOME - LOTAÇÃO | 1º | ...": CPF sem o zero da frente
        return any(re.sub(r"\D", "", cpf).lstrip("0") in linha for linha in tela.linhas())

    tela = abrir(sessao, f"{pagina}?param={doc}&modoJanelaPlc=popup")
    for posicao, cpf in enumerate(cpfs, 1):
        if na_lista(tela, cpf):
            avisar(f"{nome} {posicao}: {formatar_cpf(cpf)} já estava na lista")
            continue
        # A ordem é sempre a próxima depois de quem já está na lista (sem buraco: o e-docs
        # recusa o envio com "lacuna na sequência") e é escolhida pelo texto ("2º"), não pela
        # posição na lista da tela, que pode vir sem as ordens já usadas depois de gravar
        ordem = len(tela.linhas()) + 1
        valor = dict((t, v) for v, t in tela.opcoes(campo_ordem)).get(f"{ordem}º", str(ordem - 1))
        pessoa = buscar_servidor(sessao, cpf)
        tela = tela.enviar(F + "botaoAcaoGravar", {
            F + "oidServidoreGrh": pessoa["oid"], F + "nrCpf": pessoa["cpf"],
            "lookup_" + F + "nrCpf": pessoa["lookup"],
            campo_ordem: valor,
            F + "alertaAlteracaoPlc": "S", **extras(ordem)}, parcial=True)
        if not na_lista(tela, cpf):
            raise tela.erro(f"{nome} {pessoa['nome']} não apareceu na lista depois de gravar")
        avisar(f"{nome} {ordem}: {pessoa['nome']} · {ordem}º")


def adicionar_vistos(sessao, doc: str, cpfs: list[str], avisar=lambda msg: None) -> None:
    _adicionar_pessoas(sessao, doc, "vistosdocumentosman", F + "oidOrdemVisto", cpfs, lambda ordem: {},
                       "Visto", avisar)


def adicionar_assinantes(sessao, doc: str, cpfs: list[str], avisar=lambda msg: None) -> None:
    # Remetente (idRemetente 0 = "Sim", 1 = "Não"): só o 1º assinante; o e-docs recusa
    # dois ("É obrigatório cadastrar apenas um Remetente"). idAgrupado 1 é o que o site já
    # traz marcado
    _adicionar_pessoas(sessao, doc, "assinantesdocumentosman", F + "oidOrdemAssinatura", cpfs,
                       lambda ordem: {F + "idRemetente": "0" if ordem == 1 else "1", F + "idAgrupado": "1"},
                       "Assinante", avisar)


def grupos_destinatarios(sessao: requests.Session, tentativas: int = 4) -> list[str]:
    """Os grupos de destinatários (Dest. Publicador) que ESTE usuário tem: cada conta tem os
    seus. A lista só existe na janela de destinatários de um documento, então abre a de um
    documento da Caixa de Saída e marca "Grupo", como o navegador, SEM gravar nada."""
    docs = list(dict.fromkeys(d["doc"] for d in caixa_de_saida(sessao)))
    for doc in docs[:tentativas]:
        try:
            tela = abrir(sessao, f"destinatariosdocumentosman?param={doc}&modoJanelaPlc=popup")
            tela = tela.enviar(F + "idAgrupado", {F + "idAgrupado": "0"}, parcial=True)
        except ErroEdocs:
            continue  # documento que não deixa mexer nos destinatários: tenta o próximo
        grupos = [texto for _, texto in tela.opcoes(F + "gruposDestinatarios")]
        if grupos:
            return grupos
    return []


def configurar_destinatarios(sessao, doc: str, avisar=lambda msg: None) -> None:
    """Botão 4: marca 'agrupado', escolhe o grupo do processo (GRUPO_DESTINATARIO; vazio = o
    1º da lista) e grava."""
    tela = abrir(sessao, f"destinatariosdocumentosman?param={doc}&modoJanelaPlc=popup")
    if tela.linhas():
        avisar(f"Destinatário já definido: {tela.linhas()[0]}")
        return
    # Marcar o rádio recarrega a janela com a lista de grupos
    tela = tela.enviar(F + "idAgrupado", {F + "idAgrupado": "0"}, parcial=True)
    # Grupo escolhido que sumiu da lista da pessoa: para (não manda para o grupo errado)
    grupo = (tela.opcao(F + "gruposDestinatarios", texto=GRUPO_DESTINATARIO) if GRUPO_DESTINATARIO
             else tela.opcao(F + "gruposDestinatarios", posicao=1))
    texto = dict(tela.opcoes(F + "gruposDestinatarios")).get(grupo, grupo)
    tela = tela.enviar(F + "botaoAcaoGravar",
                       {F + "idAgrupado": "0", F + "gruposDestinatarios": grupo}, parcial=True)
    if not tela.linhas():
        raise tela.erro("o destinatário não apareceu na lista depois de gravar")
    avisar(f"Destinatário: {texto}")


def anexar_arquivo(sessao, doc: str, arquivo: Path, avisar=lambda msg: None) -> None:
    """Anexa um arquivo do computador ao documento (janela Anexos; mapeado da gravação
    capturas Doc_anexo.har). O site grava na hora em que o arquivo é escolhido: o formulário
    vai como multipart, com o arquivo em arquivoAnexo e source=adicionaArquivo. "Incluir
    Assinatura(s)?" fica no que o site já traz marcado (Não). Confere na lista se entrou."""
    from curl_cffi import CurlMime

    if not arquivo.is_file():
        raise ErroEdocs(f"o arquivo {arquivo.name} não foi encontrado no computador")
    if arquivo.suffix.lower() != ".pdf":
        raise ErroEdocs(f"o anexo {arquivo.name} não é PDF (só aceita PDF)")
    tela = abrir(sessao, f"anexosdocumentostab?param={doc}&modoJanelaPlc=popup")
    if any(arquivo.name in linha for linha in tela.linhas()):
        avisar(f"Anexo {arquivo.name} já estava no documento")
        return
    dados = dict(tela.campos)
    dados.update({"source": F + "adicionaArquivo", F + "alertaAlteracaoPlc": "S"})
    dados.pop(F + "arquivoAnexo", None)
    partes = CurlMime()
    for nome, valor in dados.items():
        partes.addpart(name=nome, data=str(valor).encode("utf-8"))
    partes.addpart(name=F + "arquivoAnexo", filename=arquivo.name, local_path=str(arquivo),
                   content_type=_tipo_do_arquivo(arquivo))
    try:
        resposta = sessao.post(tela.acao, multipart=partes, headers={"Referer": tela.url}, timeout=TIMEOUT * 4)
    finally:
        partes.close()
    tela = Tela(sessao, resposta)
    if not any(arquivo.name in linha for linha in tela.linhas()):
        raise tela.erro(f"o anexo {arquivo.name} não apareceu na lista depois de enviar")
    avisar(f"Anexo: {arquivo.name}")


def _tipo_do_arquivo(arquivo: Path) -> str:
    import mimetypes
    return mimetypes.guess_type(arquivo.name)[0] or "application/octet-stream"


# --- Documentos em montagem -------------------------------------------------------
# Se o usuário deixou um documento em montagem (criado no site e não enviado), o Edocs
# o reaproveita ao abrir um documento novo: o 1º ato é gravado por cima dele, que já tem
# vistos/assinantes, e o envio falha ("visto/assinante duplicado"). Antes do 1º envio o
# programa procura esses documentos na Caixa de Saída e, só com autorização, os cancela.

MOTIVO_CANCELAMENTO = "Cancelado pelo Edodocs: documento em montagem antes de novo envio"
_COLUNAS_CAIXA = {"processo de:": "papel", "tipo de documento": "tipo", "identificação sistema": "identificacao",
                  "nossa identificação": "titulo", "data": "data", "remetente": "remetente",
                  "destinatário": "destinatario", "situação": "situacao"}


def documentos_da_caixa(tela: Tela) -> list[dict]:
    """Linhas da Caixa de Saída (ou de Entrada): doc (número interno), papel ("Visto",
    "Responsável"...), tipo, identificação, título, data, remetente, destinatário, situação,
    linha (nrLinha da tabela), e se tem os botões Cancelar e Ciente."""
    docs = []
    for tabela in tela.doc.xpath('//table[.//th[normalize-space()="Situação"]]'):
        cabecalho = [" ".join(" ".join(th.itertext()).split()).lower()
                     for th in tabela.xpath('.//tr[th][last()]/th')]
        if "situação" not in cabecalho or "remetente" not in cabecalho:
            continue
        for tr in tabela.xpath('.//tr[td]'):
            tds = tr.xpath("./td")
            if len(tds) != len(cabecalho):
                continue
            linha = {chave: " ".join(" ".join(td.itertext()).split())
                     for nome, td in zip(cabecalho, tds) if (chave := _COLUNAS_CAIXA.get(nome))}
            cliques = " ".join(tr.xpath(".//@onclick"))
            m = (re.search(r"cancelardocumentoman\?chPlc=(\d+)", cliques)
                 or re.search(r"param=(\d+)_", cliques))
            if not m:
                continue
            nomes = " ".join(tr.xpath(".//@name"))
            n = re.search(r"plcLogicaItens:(\d+):", nomes)
            # Botão Tramitação: o endereço da janela vem pronto no clique (param com o número,
            # a identificação e o título do documento)
            t = re.search(r"(historicosituacoesdoccon\?[^'\\]+)", cliques)
            linha |= {"doc": m.group(1), "pode_cancelar": "cancelardocumentoman" in cliques,
                      "linha": int(n.group(1)) if n else None, "ciente": ":btLer" in nomes,
                      "tramitacao": t.group(1) if t else ""}
            docs.append(linha)
    return docs


def em_montagem(linha: dict) -> bool:
    """Documento nunca enviado: sem remetente e sem destinatário, situação vazia (ou "em
    montagem") e com o botão Cancelar. Qualquer documento já enviado tem remetente e
    situação ("Doc. enviado para...", "Vistado"...), então nunca entra aqui."""
    situacao = linha.get("situacao", "").lower()
    return (linha.get("pode_cancelar") and not linha.get("remetente")
            and not linha.get("destinatario") and (not situacao or "montagem" in situacao))


def caixa_de_saida(sessao) -> list[dict]:
    return _de_novo_se_demorar(lambda: _caixa_de_saida(sessao))


def _caixa_de_saida(sessao) -> list[dict]:
    tela = abrir(sessao, "caixasaidacon?evento=y")
    docs = documentos_da_caixa(tela)
    if not docs and tela.doc.xpath(f'//*[@name="{F}botaoAcaoPesquisar"]'):
        docs = documentos_da_caixa(tela.enviar(F + "botaoAcaoPesquisar"))
    return docs


def documentos_em_montagem(sessao) -> list[dict]:
    """TODOS os documentos em montagem do usuário, de qualquer data. Enquanto houver um, o
    e-docs não cria documento novo: abre o em montagem mais recente no lugar (e o 1º ato é
    gravado por cima dele). A Caixa de Saída só mostra os mais recentes, então usa a
    pesquisa com o filtro "Doc. em montagem" do site, desde 2000."""
    docs = [d for d in caixa_de_saida(sessao) + pesquisar_caixa(sessao, "montagem") if em_montagem(d)]
    return list({d["doc"]: d for d in docs}.values())


def _de_novo_se_demorar(ler, tentativas: int = 3):
    """ler() de novo quando o e-docs demora a responder (só para o que apenas lê: repetir
    não muda nada no site)."""
    for vez in range(1, tentativas + 1):
        try:
            return ler()
        except requests.exceptions.Timeout:
            if vez == tentativas:
                raise
            time.sleep(3 * vez)


def pesquisar_caixa(sessao, situacao: str) -> list[dict]:
    """Caixa de Saída filtrada pela situação (o texto contém, ex.: "montagem", "cancelado"),
    desde 2000: acha também os documentos antigos, que a lista normal não mostra."""
    return _de_novo_se_demorar(lambda: _pesquisar_caixa(sessao, situacao))


def _pesquisar_caixa(sessao, situacao: str) -> list[dict]:
    tela = abrir(sessao, "caixasaidacon?evento=y")
    campo = F + "oidSituacoesTramitacao"
    valor = next((v for v, t in tela.opcoes(campo) if situacao in t.lower()), None)
    if valor is None:
        return caixa_de_saida(sessao)
    return documentos_da_caixa(tela.enviar(F + "botaoAcaoPesquisar",
                                           {campo: valor, F + "dtCadastro": "01/01/2000"}))


def em_montagem_agora(sessao, doc: str) -> bool:
    """O documento continua em montagem (pela pesquisa do site, qualquer que seja a data)?"""
    return any(d["doc"] == doc and em_montagem(d) for d in pesquisar_caixa(sessao, "montagem"))


def cancelar_em_montagem(sessao, doc: str, avisar=lambda msg: None) -> None:
    """Cancela UM documento em montagem, como o botão Cancelar da Caixa de Saída. Confere
    na hora que ele continua em montagem (nunca cancela documento enviado) e, depois,
    que ele saiu dos em montagem."""
    if not em_montagem_agora(sessao, doc):
        raise ErroEdocs(f"o documento {doc} não está em montagem; nada foi cancelado")
    tela = _de_novo_se_demorar(lambda: abrir(sessao, f"cancelardocumentoman?chPlc={doc}&modoJanelaPlc=popup"))
    try:
        tela = tela.enviar(F + "botaoAcaoGravar", {F + "dsMotivoCancelamento": MOTIVO_CANCELAMENTO,
                                                    F + "alertaAlteracaoPlc": "S"}, parcial=True)
    except requests.exceptions.Timeout:
        pass  # o e-docs pode ter cancelado sem responder a tempo: a conferência abaixo diz
    # Confere pela lista dos em montagem (pequena) e não pela dos cancelados desde 2000,
    # que numa conta antiga passa do tempo de espera
    if em_montagem_agora(sessao, doc):
        raise tela.erro(f"o Edocs não confirmou o cancelamento do documento {doc}")
    avisar(f"Documento em montagem {doc} cancelado")


def preparar_documento(sessao, titulo: str, corpo_html: str, vistos: list[str],
                       assinantes: list[str], avisar=lambda msg: None, data: str = "",
                       modelo: str = INSTRUMENTO, extras: dict | None = None) -> str:
    """Tudo antes do envio: cria o documento, vistos, assinantes e destinatários.
    Devolve o número do documento (fica como rascunho no Edocs até ser enviado).
    Nos Documentos Genéricos não há destinatários, e vistos só com "Incluir Visto?" = Sim."""
    doc = criar_documento(sessao, titulo, corpo_html, avisar, data, modelo, extras)
    # Já havia um documento em montagem com a mesma Nossa Identificação: o e-docs devolve
    # ele (com os vistos de antes) em vez de criar outro, e misturar dá "lacuna na sequência"
    janela, quem = (("vistosdocumentosman", "visto(s)") if vistos
                    else ("assinantesdocumentosman", "assinante(s)"))
    antigos = abrir(sessao, f"{janela}?param={doc}&modoJanelaPlc=popup").linhas()
    if antigos:
        raise ErroEdocs(f"o e-docs reaproveitou o documento {doc}, que já estava em montagem com o título "
                        f"“{titulo}” e já tem {len(antigos)} {quem}. Cancele esse documento no e-docs (ou "
                        "aceite cancelar os em montagem quando o Edodocs perguntar) e envie de novo")
    try:
        if vistos:
            adicionar_vistos(sessao, doc, vistos, avisar)
        adicionar_assinantes(sessao, doc, assinantes, avisar)
        if MODELOS[modelo]["destinatarios"]:
            configurar_destinatarios(sessao, doc, avisar)
    except ErroEdocs as e:
        # O documento acabou de ser criado e ficou pela metade: em montagem, ele faria o
        # e-docs recusar todos os atos seguintes. Cancela (é só deste envio).
        raise ErroEdocs(f"{e}. {descartar_rascunho(sessao, doc, avisar)}") from e
    return doc


def descartar_rascunho(sessao, doc: str, avisar=lambda msg: None) -> str:
    """Cancela um documento que o próprio Edodocs acabou de criar e não enviou (pulado ou
    com erro), para ele não travar os próximos atos. Devolve o que aconteceu, para a
    mensagem."""
    try:
        cancelar_em_montagem(sessao, doc, avisar)
        return f"O rascunho {doc} foi cancelado no e-docs"
    except Exception as e:
        return f"Não deu para cancelar o rascunho {doc} ({e}): cancele no e-docs antes de enviar de novo"


def abrir_documento(sessao: requests.Session, doc: str, modelo: str = INSTRUMENTO) -> Tela:
    """Tela do documento já gravado (é nela que ficam os botões de visto/assinante/envio).

    O Edocs abre a tela do documento por `chPlc=<número>` (é o mesmo parâmetro que ele usa
    no redirecionamento após gravar). As janelas auxiliares (vistos, assinantes,
    destinatários) é que usam `param=<número>`; na tela principal `param` não abre o
    documento — por isso o envio era feito sobre a tela errada e o ato ficava em montagem."""
    return abrir(sessao, f"{MODELOS[modelo]['pagina']}?chPlc={doc}")


def botoes(tela: Tela) -> list[tuple[str, str]]:
    """[(id/source, texto)] dos botões e links que agem no formulário da tela.
    Serve para descobrir o id do botão de envio."""
    achados = []
    for el in tela.doc.xpath(f'//form[@id="{FORM}"]//button'
                             f' | //form[@id="{FORM}"]//input[@type="submit" or @type="button"]'
                             f' | //form[@id="{FORM}"]//a[@onclick]'):
        nome = el.get("name") or el.get("id") or ""
        texto = " ".join((el.get("value") or el.text or "").split())[:60]
        if nome or texto:
            achados.append((nome, texto))
    # tira repetidos mantendo a ordem
    vistos, saida = set(), []
    for par in achados:
        if par not in vistos:
            vistos.add(par)
            saida.append(par)
    return saida


def _confirmou_envio(tela: Tela) -> bool:
    """O Edocs confirmou o envio? A resposta tem que ser HTTP 200 e trazer a mensagem
    "Documento: ... enviado para visto(s) em paralelo" (ou "... para assinatura", nos
    Documentos Genéricos sem visto). Sem ela, o ato segue em montagem."""
    if getattr(tela, "status", 200) != 200:
        return False
    texto = " ".join(" ".join(tela.doc.itertext()).lower().split())
    return bool(re.search(r"enviad[oa]s? para (o |a )?(visto|assinatura)", texto))


def _botao_de_envio(tela: Tela) -> str:
    """O botão "Enviar p/ ..." da tela do documento: p/ Vistos no instrumento; nos Genéricos,
    p/ Assinatura (ou p/ Vistos, com "Incluir Visto?" = Sim). Achado na própria tela."""
    ids = [nome for nome, _ in botoes(tela) if nome.startswith(F + "botaoEnviarPara")]
    if BOTAO_ENVIAR in ids or not ids:
        return BOTAO_ENVIAR
    return next((i for i in ids if "Vistos" in i), ids[0])


def enviar_documento(sessao: requests.Session, doc: str, avisar=lambda msg: None,
                     modelo: str = INSTRUMENTO) -> None:
    """Envio (passo 6): abre a tela do documento e 'clica' no botão de envio.
    O envio é irreversível — chame só quando tiver certeza."""
    tela = abrir_documento(sessao, doc, modelo)
    tela = tela.enviar(_botao_de_envio(tela))
    # Depois de enviar, o documento sai do rascunho: o Edocs confirma com uma mensagem
    if "j_password" in tela.html:
        raise SessaoExpirada("a sessão caiu no momento do envio")

    # Só considera enviado quando o Edocs confirma de fato. Antes, bastava o site não
    # mostrar erro para o ato ser marcado como enviado — e ele ficava em montagem.
    if _confirmou_envio(tela):
        avisar(f"Documento {doc} enviado")
        return
    raise tela.erro("o Edocs não confirmou o envio (o documento continua em montagem)")


# --- Vistar o próprio ato logo depois de enviar ------------------------------------
# Como no site (capturas/vistar.har): na Caixa de Entrada, "Ciente" na linha do documento,
# a página recarrega e "Vistar documento" vale para a linha marcada (nrLinha).

CAIXA_ENTRADA = "caixaentradacon?evento=y"


def _linha_para_vistar(tela: Tela, doc: str) -> int | None:
    """nrLinha do documento na Caixa de Entrada (de preferência a linha de visto)."""
    linhas = [d for d in documentos_da_caixa(tela) if d["doc"] == doc and d["linha"] is not None]
    de_visto = [d for d in linhas if "visto" in d.get("papel", "").lower()]
    escolhidas = de_visto or (linhas if len(linhas) == 1 else [])
    return escolhidas[0]["linha"] if escolhidas else None


def vistar_documento(sessao, doc: str, avisar=lambda msg: None, tentativas: int = 3,
                     espera: float = 1.5) -> None:
    """Dá ciente e vista o documento na conta de quem está logado. Só clica na linha que
    é deste documento e só dá como vistado quando ele sai da Caixa de Entrada."""
    for tentativa in range(tentativas):  # recém-enviado: pode demorar a aparecer
        tela = abrir(sessao, CAIXA_ENTRADA)
        linha = _linha_para_vistar(tela, doc)
        if linha is not None:
            break
        time.sleep(espera)
    else:
        raise ErroEdocs(f"o documento {doc} não apareceu na Caixa de Entrada para vistar")

    try:  # "Ciente" da linha (atualização parcial, como no site)
        tela.enviar(F + f"plcLogicaItens:{linha}:btLer", {F + "nrLinha": str(linha)}, parcial=True)
    except ErroEdocs:
        pass  # a resposta parcial pode vir sem o formulário; o visto abaixo é conferido
    tela = abrir(sessao, CAIXA_ENTRADA + "&navDe=1")
    linha = _linha_para_vistar(tela, doc)  # a ordem pode ter mudado: confere de novo
    if linha is None:
        raise ErroEdocs(f"o documento {doc} sumiu da Caixa de Entrada antes do visto")
    tela = tela.enviar(F + "btVistarDocumento", {F + "nrLinha": str(linha)})

    if _linha_para_vistar(abrir(sessao, CAIXA_ENTRADA), doc) is not None:
        raise tela.erro(f"o Edocs não confirmou o visto do documento {doc}")
    avisar(f"Documento {doc} vistado (ciente e visto)")


# --- Recebidos (Caixa de Entrada): listar, baixar o PDF, ciente, vistar ----------------------
# Na Caixa de Entrada cada linha tem um papel: "Encaminhamento" (Visualizar, Ciente), "Visto"
# (Ciente e Vistar) ou "Despacho" (Encaminhar/Despachar, que ficam no site).

def recebidos(sessao) -> list[dict]:
    """Os documentos da Caixa de Entrada (os mais recentes), um por documento, com o que dá
    para fazer: ciente (tem o botão Ciente) e vistavel (é uma linha de visto)."""
    por_doc: dict[str, dict] = {}
    for d in documentos_da_caixa(abrir(sessao, CAIXA_ENTRADA)):
        atual = por_doc.setdefault(d["doc"], dict(d, ciente=False, vistavel=False, papeis=[]))
        atual["papeis"].append(d.get("papel", ""))
        atual["ciente"] = atual["ciente"] or d.get("ciente", False)
        atual["vistavel"] = atual["vistavel"] or "visto" in d.get("papel", "").lower()
        atual["pode_cancelar"] = atual.get("pode_cancelar") or d.get("pode_cancelar", False)
        atual["tramitacao"] = atual.get("tramitacao") or d.get("tramitacao", "")
    # Cancelar: quem vista pode jogar fora o documento (o site oferece o Cancelar da linha)
    for d in por_doc.values():
        d["cancelavel"] = bool(d["pode_cancelar"] or d["vistavel"])
    return list(por_doc.values())


def tramitacao(sessao, caminho: str) -> list[dict]:
    """Por onde o documento passou e quem ainda falta, como o botão Tramitação da Caixa de
    Entrada (gravação botao_tramitacao.har): a janela historicosituacoesdoccon traz uma linha
    por etapa (# | Tipo | Nome | Situação | Data). caminho é o endereço que vem no botão
    (recebidos()["tramitacao"]). Devolve [{"tipo", "nome", "situacao", "data", "feito"}];
    feito = a etapa já aconteceu (tem data e não está pendente)."""
    from urllib.parse import quote
    url = TELAS + quote(caminho, safe="/?=&_.-")
    tela = Tela(sessao, sessao.get(url, headers={"Referer": TELAS + CAIXA_ENTRADA}, timeout=TIMEOUT))
    passos = []
    for tr in tela.doc.xpath('//tr[contains(@id, ":linhaSel")]'):
        celulas = [" ".join(" ".join(td.itertext()).split()) for td in tr.xpath("./td")]
        if len(celulas) < 5:
            continue
        _, tipo, nome, situacao, data = celulas[:5]
        passos.append({"tipo": tipo, "nome": nome, "situacao": situacao, "data": data,
                       "feito": bool(data) and "PENDENTE" not in situacao.upper()})
    if not passos:
        raise tela.erro("o Edocs não mostrou a tramitação do documento")
    return passos


# Motivo do cancelamento: o e-docs exige no mínimo 5 palavras ("Motivo Cancelamento deve
# conter no mínimo 5 palavra(s)"). O texto padrão é a sugestão da tela
PALAVRAS_MOTIVO = 5
MOTIVO_CANCELAR_RECEBIDO = "Cancelado pelo sistema Edodocs Software Design"


def cancelar_recebido(sessao, doc: str, motivo: str = MOTIVO_CANCELAR_RECEBIDO,
                      avisar=lambda msg: None) -> None:
    """Cancela um documento recebido para visto, como o botão Cancelar da linha na Caixa de
    Entrada (gravação "processo de cancelamento de um documento na caixa de entrada.har"):
      1. abre a Caixa de Entrada (a janela de cancelar é aberta a partir dela);
      2. abre a janela cancelardocumentoman?chPlc=<doc>&param=V;
      3. Gravar com o motivo, como atualização parcial da janela (cabeçalho Tr-XHR-Message,
         que o navegador manda e sem o qual o site não grava);
      4. confere que o documento saiu da Caixa de Entrada.
    O documento é jogado fora sem ciente e sem visto."""
    motivo = " ".join(str(motivo or "").split())
    if len(motivo.split()) < PALAVRAS_MOTIVO:
        raise ErroEdocs(f"o motivo do cancelamento precisa de no mínimo {PALAVRAS_MOTIVO} palavras")
    caixa = abrir(sessao, CAIXA_ENTRADA)
    if doc not in {d["doc"] for d in documentos_da_caixa(caixa)}:
        raise ErroEdocs(f"o documento {doc} não está mais na Caixa de Entrada")
    janela = Tela(sessao, sessao.get(TELAS + f"cancelardocumentoman?chPlc={doc}&param=V&modoJanelaPlc=popup",
                                     headers={"Referer": caixa.url}, timeout=TIMEOUT))
    dados = dict(janela.campos)
    dados.update({F + "dsMotivoCancelamento": motivo, F + "alertaAlteracaoPlc": "S",
                  "source": F + "botaoAcaoGravar", "event": "autosub", "partial": "true"})
    resposta = sessao.post(janela.acao, data=dados, timeout=TIMEOUT, headers={
        "Referer": janela.url, "Origin": URL.split("/webedocspmu")[0], "Tr-XHR-Message": "true",
        "Accept": "*/*"})
    resposta.raise_for_status()
    if 'name="j_password"' in resposta.text:
        raise SessaoExpirada("o Edocs voltou para a tela de login")
    # A resposta é um XML com a janela atualizada; se o site recusou, a mensagem vem nela
    mensagens = []
    try:
        conteudo = b"".join(re.findall(rb"<!\[CDATA\[(.*?)\]\]>", resposta.content, re.S)) or resposta.content
        if conteudo.strip():
            mensagens = _mensagens(lxml.html.fromstring(conteudo, parser=lxml.html.HTMLParser(encoding="utf-8")))
    except Exception:
        pass
    if doc in {d["doc"] for d in documentos_da_caixa(abrir(sessao, CAIXA_ENTRADA + "&navDe=1"))}:
        detalhe = f" ({'; '.join(mensagens)})" if mensagens else ""
        raise ErroEdocs(f"o Edocs não cancelou o documento {doc}{detalhe}")
    avisar(f"Documento {doc} cancelado")


def _linha_do_doc(tela: Tela, doc: str, com_ciente: bool = False) -> int | None:
    for d in documentos_da_caixa(tela):
        if d["doc"] == doc and d["linha"] is not None and (d.get("ciente") or not com_ciente):
            return d["linha"]
    return None


def baixar_recebido(sessao, doc: str) -> tuple[str, bytes]:
    """O PDF do documento, como o botão Visualizar (sem o histórico de tramitação).
    Devolve (nome do arquivo, conteúdo)."""
    tela = abrir(sessao, CAIXA_ENTRADA)
    linha = _linha_do_doc(tela, doc)
    if linha is None:
        raise ErroEdocs(f"o documento {doc} não está mais na Caixa de Entrada")
    dados = dict(tela.campos)
    dados.update({F + "nrLinha": str(linha), F + "idImprimeHistoricoSituacoesDoc": "N",
                  "source": F + f"plcLogicaItens:{linha}:btVisualizar"})
    resposta = sessao.post(tela.acao, data=dados, headers={"Referer": tela.url}, timeout=TIMEOUT)
    resposta.raise_for_status()
    if not resposta.content.startswith(b"%PDF"):
        raise ErroEdocs(f"o Edocs não devolveu o PDF do documento {doc}")
    m = re.search(r'filename="?([^";]+)', resposta.headers.get("content-disposition", ""))
    return (m.group(1) if m else f"documento_{doc}.pdf"), resposta.content


def dar_ciente(sessao, doc: str, avisar=lambda msg: None) -> None:
    """Ciente no documento (botão Ciente da linha dele). Só dá como feito quando o Edocs
    deixa de oferecer o Ciente para ele."""
    tela = abrir(sessao, CAIXA_ENTRADA)
    linha = _linha_do_doc(tela, doc, com_ciente=True)
    if linha is None:
        avisar(f"Documento {doc}: já estava com ciente")
        return
    try:
        tela.enviar(F + f"plcLogicaItens:{linha}:btLer", {F + "nrLinha": str(linha)}, parcial=True)
    except ErroEdocs:
        pass  # a resposta parcial pode vir sem o formulário; a conferência abaixo decide
    if _linha_do_doc(abrir(sessao, CAIXA_ENTRADA + "&navDe=1"), doc, com_ciente=True) is not None:
        raise ErroEdocs(f"o Edocs não confirmou o ciente no documento {doc}")
    avisar(f"Ciente no documento {doc}")


def preparar_e_enviar(sessao, titulo: str, corpo_html: str, vistos: list[str],
                      assinantes: list[str], avisar=lambda msg: None) -> str:
    """Fluxo completo de um ato: cria o documento, vistos, assinantes, destinatário e envia.
    Devolve o número do documento enviado."""
    doc = preparar_documento(sessao, titulo, corpo_html, vistos, assinantes, avisar)
    enviar_documento(sessao, doc, avisar)
    return doc


# --- Testes ------------------------------------------------------------------

def _credenciais() -> tuple[str, str]:
    import credenciais

    dados = credenciais.carregar(Path(__file__).parent / "configs" / "credenciais.json", "edocs")
    if not dados:
        sys.exit("Nenhum usuário do Edocs salvo. Cadastre no programa (opção 4 → 1).")
    return dados


def _testar_login() -> None:
    usuario, senha = _credenciais()
    print(f"Entrando como {usuario}...")
    try:
        sessao = entrar(usuario, senha)
    except LoginRecusado:
        sys.exit("Login recusado: confira usuário e senha (opção 4 → 1 do programa).")

    resposta = sessao.get(URL, timeout=TIMEOUT)
    # "Usuário: XXX.178.69X-XX - Nome da Pessoa |": o nome vem depois do último " - "
    usuario_tela = re.search(r"Usu&aacute;rio:([^<|]+)", resposta.text)
    nome = html_lib.unescape(usuario_tela.group(1)).rsplit(" - ", 1)[-1].strip() if usuario_tela else ""
    print(f"No menu: {'sim' if _no_menu(resposta.text) else 'NÃO'}"
          + (f" (usuário: {nome})" if nome else ""))
    print(f"Sessão mantida na 2ª requisição: {'sim' if logado(sessao) else 'NÃO'}")


def _testar_rascunho() -> None:
    """Cria no Edocs, como rascunho, o 1º ato pendente do processo escolhido. NÃO envia."""
    import main
    import ui

    main.preparar_pastas()
    if not main.escolher_processo():
        return
    preparado = main.preparar_atos("edocs")
    if not preparado:
        return
    registros, textos, htmls, _ = preparado
    dados, texto, corpo = registros[0], textos[0], htmls[0]
    titulo = main.proc().titulo_edocs(dados)

    usuario, senha = _credenciais()
    with ui.processando("Entrando no Edocs"):
        sessao = entrar(usuario, senha)
    vistos, assinantes = main.pessoas_do_envio(
        usuario, lambda termo: pesquisar_servidores(sessao, termo))

    ui.titulo("Rascunho (não será enviado)")
    ui.info(f"Ato: {main.proc().descricao(dados)} (linha {dados['_LINHA']})")
    ui.info(f"Título: {titulo}")
    if "[" in texto:
        ui.aviso("O texto tem campos vazios entre colchetes: " + ", ".join(sorted(set(re.findall(r"\[[^\]]+\]", texto)))))
    if ui.console.input("Criar o rascunho no Edocs? (S/N): ").strip().upper() != "S":
        ui.aviso("Nada criado")
        return

    try:
        doc = preparar_documento(sessao, titulo, corpo, vistos, assinantes, avisar=ui.ok)
    except (ErroEdocs, SessaoExpirada) as e:
        ui.erro(f"Parou: {e}")
        return
    ui.ok(f"Rascunho pronto: documento {doc}. NÃO foi enviado.")
    ui.info("Abra no Edocs (Editar Instrumentos Jurídicos) e confira título, texto, vistos, "
            "assinantes e destinatário.")


def _preparar_primeiro_ato():
    """Lê o 1º ato pendente do processo escolhido e monta texto/html/título.
    Devolve (dados, titulo, corpo) ou None."""
    import main
    import ui

    main.preparar_pastas()
    if not main.escolher_processo():
        return None
    preparado = main.preparar_atos("edocs")
    if not preparado:
        return None
    registros, textos, htmls, _ = preparado
    dados, texto, corpo = registros[0], textos[0], htmls[0]
    titulo = main.proc().titulo_edocs(dados)
    ui.info(f"Ato: {main.proc().descricao(dados)} (linha {dados['_LINHA']})")
    ui.info(f"Título: {titulo}")
    return dados, titulo, texto, corpo


def _testar_botoes() -> None:
    """Cria o rascunho do 1º ato e lista os botões da tela do documento, para descobrir o
    id do botão de envio (a constante BOTAO_ENVIAR)."""
    import main
    import ui

    preparado = _preparar_primeiro_ato()
    if not preparado:
        return
    _, titulo, _, corpo = preparado

    usuario, senha = _credenciais()
    with ui.processando("Entrando no Edocs"):
        sessao = entrar(usuario, senha)
    vistos, assinantes = main.pessoas_do_envio(
        usuario, lambda termo: pesquisar_servidores(sessao, termo))

    ui.titulo("Botões da tela do documento")
    if ui.console.input("Criar o rascunho e listar os botões? (S/N): ").strip().upper() != "S":
        ui.aviso("Nada feito")
        return

    try:
        doc = preparar_documento(sessao, titulo, corpo, vistos, assinantes, avisar=ui.ok)
        tela = abrir_documento(sessao, doc)
    except (ErroEdocs, SessaoExpirada) as e:
        ui.erro(f"Parou: {e}")
        return
    ui.ok(f"Documento {doc} (rascunho). Botões encontrados:")
    for nome, texto in botoes(tela):
        ui.console.print(f"  [cyan]{nome or '(sem nome)'}[/]  →  {texto}")
    ui.info(f"Constante atual no código: BOTAO_ENVIAR = {BOTAO_ENVIAR!r}")
    ui.info("Ache o botão de envio na lista (texto tipo 'Enviar') e ajuste BOTAO_ENVIAR.")


def _testar_enviar() -> None:
    """Fluxo completo: cria o rascunho do 1º ato e ENVIA (usando BOTAO_ENVIAR)."""
    import main
    import ui

    preparado = _preparar_primeiro_ato()
    if not preparado:
        return
    dados, titulo, texto, corpo = preparado

    usuario, senha = _credenciais()
    with ui.processando("Entrando no Edocs"):
        sessao = entrar(usuario, senha)
    vistos, assinantes = main.pessoas_do_envio(
        usuario, lambda termo: pesquisar_servidores(sessao, termo))

    ui.titulo("Envio (IRREVERSÍVEL)")
    if "[" in texto:
        ui.aviso("O texto tem campos vazios entre colchetes: "
                 + ", ".join(sorted(set(re.findall(r"\[[^\]]+\]", texto)))))
    if ui.console.input("Criar e ENVIAR este ato no Edocs? (S/N): ").strip().upper() != "S":
        ui.aviso("Nada enviado")
        return

    try:
        doc = preparar_e_enviar(sessao, titulo, corpo, vistos, assinantes, avisar=ui.ok)
    except (ErroEdocs, SessaoExpirada) as e:
        ui.erro(f"Parou: {e}")
        return
    ui.ok(f"Documento {doc} enviado ({main.proc().numero(dados) or '?'}).")


def _testar_montagem() -> None:
    """Só LISTA os documentos em montagem (não cancela nada)."""
    usuario, senha = _credenciais()
    sessao = entrar(usuario, senha)
    caixa = caixa_de_saida(sessao)
    print(f"Caixa de Saída: {len(caixa)} linha(s) lida(s)")
    for d in caixa[:20]:
        marca = "EM MONTAGEM" if em_montagem(d) else ""
        print(f"  {d['doc']} · {d.get('titulo', '')[:45]:45} · rem: {d.get('remetente', '')[:15]:15}"
              f" · {d.get('situacao', '')[:30]:30} {marca}")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    modo = sys.argv[1:]
    if modo == ["rascunho"]:
        _testar_rascunho()
    elif modo == ["botoes"]:
        _testar_botoes()
    elif modo == ["enviar"]:
        _testar_enviar()
    elif modo == ["montagem"]:
        _testar_montagem()
    else:
        _testar_login()
