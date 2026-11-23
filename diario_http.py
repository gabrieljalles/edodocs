"""Acesso ao Diário Oficial por requisições HTTP, sem navegador.

O site é feito em Yii2 (PHP). Cada formulário carrega um token CSRF ("_csrf-backend"),
que muda a cada página; o login e o cadastro de publicação só são aceitos com o token da
tela certa. A classe _SessaoEducada espera PAUSA_REQUISICAO antes de cada requisição, para
o site não ver uma rajada de chamadas.

Fluxo de um ato (mapeado da gravação capturas/diario.uberlandia.mg.gov.br.har):
  1. login            GET  /admin/login  (token CSRF)
                      POST /admin/login  → 302 /admin/ (painel, com o link "Deslogar")
  2. formulário novo  GET  /admin/minhas-publicacoes/create  (token CSRF e listas)
  3. enviar           POST /admin/minhas-publicacoes/create  (multipart/form-data)
                      302 → /admin/minhas-publicacoes/index = publicado;
                      200 com o formulário de volta = recusado (mensagens de validação)
  4. número           GET  /admin/minhas-publicacoes/index → link /view/<id> do ato

O navegador ainda chama /admin/datecontrol/parse/convert só para converter a data da
tela; aqui a conversão é feita direto. A data vai em dois formatos, como o site espera:
  data_veiculacao-documento-data_veiculacao-disp  dd/mm/aaaa (o que aparece na tela)
  Documento[data_veiculacao]                       aaaa-mm-dd  (o que o site guarda)

Testes (usam o usuário do Diário salvo no programa, opção 4 → 1):
    python diario_http.py              só o login
    python diario_http.py campos       mostra os valores das listas da tela de cadastro
    python diario_http.py enviar       cadastra e ENVIA o 1º ato pendente
"""
import re
import sys
import time
from pathlib import Path

import lxml.html
# Alguns servidores recusam requisições sem cara de navegador / fingerprint TLS do requests
from curl_cffi import CurlMime, requests

URL = "https://diario.uberlandia.mg.gov.br/admin/"
LOGIN = URL + "login"
PUBLICACOES = URL + "minhas-publicacoes"
CRIAR = PUBLICACOES + "/create"
LISTA = PUBLICACOES + "/index"
TIMEOUT = 30
NAVEGADOR = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
             "(KHTML, like Gecko) Chrome/130.0 Safari/537.36")
IMPERSONAR = "chrome124"
# Pausa entre uma requisição e a seguinte, para não parecer rajada
PAUSA_REQUISICAO = 0.25

# Opções fixas do formulário: (nome do campo, texto da opção, valor visto na gravação).
# A opção é procurada pelo texto; o valor da gravação só é usado se o texto não for achado.
SECRETARIA = ("Documento[id_secretaria]", "Administração", "450")
CATEGORIA = ("Documento[id_documento_categoria]", "Ato Administrativo", "567")
COLUNA = ("Documento[coluna]", "Duas colunas", "B")
_GRAVADOS = {"secretaria": SECRETARIA, "categoria": CATEGORIA, "coluna": COLUNA}

# As listas da tela "Incluir" do Diário (lidas do site com "python diario_http.py campos"
# em 26/10/2026). O processo escolhe uma delas em "Como vai ficar → No Diário Oficial".
OPCOES = {
    "secretaria": ["Administração", "Atos do Prefeito"],
    "categoria": ["Ato Administrativo", "Aviso", "Boletim Informativo", "Comunicado", "Decisão",
                  "Decreto", "Decreto s/nº", "Demonstrativo", "Edital", "Extrato",
                  "Instrução Normativa", "Justificativa", "Lei Complementar", "Lei Ordinária",
                  "Mensagem", "Portaria", "Resolução", "Retificação"],
    "coluna": ["Página Inteira", "Duas Colunas"],
}


def usar_opcoes(secretaria: str, categoria: str, coluna: str) -> None:
    """Opções do processo em uso (textos como aparecem nas listas do site). O valor da
    gravação só serve de reserva quando o texto é o mesmo da gravação."""
    global SECRETARIA, CATEGORIA, COLUNA
    escolhidos = {"secretaria": secretaria, "categoria": categoria, "coluna": coluna}
    novos = {}
    for chave, (nome, texto, valor) in _GRAVADOS.items():
        novos[chave] = (nome, escolhidos[chave], valor if escolhidos[chave] == texto else "")
    SECRETARIA, CATEGORIA, COLUNA = novos["secretaria"], novos["categoria"], novos["coluna"]


class LoginRecusado(Exception):
    """Usuário ou senha não aceitos pelo Diário."""


class SessaoExpirada(Exception):
    """O Diário voltou para a tela de login (sessão caiu)."""


class ErroDiario(RuntimeError):
    """O Diário não fez o que foi pedido (a mensagem diz o quê)."""


class _SessaoEducada(requests.Session):
    """Sessão que espera PAUSA_REQUISICAO antes de cada requisição."""

    def request(self, *args, **kwargs):
        time.sleep(PAUSA_REQUISICAO)
        return super().request(*args, **kwargs)


# O site é todo em UTF-8; sem isso o lxml pode ler bytes como latin-1 ("Nº" vira "NÂº")
_PARSER = lxml.html.HTMLParser(encoding="utf-8")


def _html(resposta) -> lxml.html.HtmlElement:
    return lxml.html.fromstring(resposta.content, parser=_PARSER)


def _no_menu(html: str) -> bool:
    """Página de quem está logado: tem o link de sair e não tem o formulário de login."""
    return ("/admin/logout" in html or "Deslogar" in html) and 'name="EntidadeUsuarioForm' not in html


def _csrf(doc) -> str:
    """Token CSRF da página: o campo hidden "_csrf-backend" ou a meta "csrf-token"."""
    campo = doc.xpath('//input[@name="_csrf-backend"]/@value')
    if campo:
        return campo[0]
    meta = doc.xpath('//meta[@name="csrf-token"]/@content')
    return meta[0] if meta else ""


def entrar(usuario: str, senha: str) -> requests.Session:
    """Faz o login e devolve a sessão (o cookie de sessão fica guardado nela). Lança
    LoginRecusado se o Diário não aceitar. O POST precisa do token CSRF da tela de login."""
    sessao = _SessaoEducada(impersonate=IMPERSONAR)
    sessao.headers["User-Agent"] = NAVEGADOR

    # Abre a tela de login: o servidor cria a sessão e devolve o token CSRF
    resposta = sessao.get(LOGIN, timeout=TIMEOUT)
    resposta.raise_for_status()
    token = _csrf(_html(resposta))

    # Depois do login o site manda para http://.../admin/ e de lá para https (302 → 307)
    resposta = sessao.post(LOGIN, data={
        "_csrf-backend": token,
        "EntidadeUsuarioForm[username]": usuario,
        "EntidadeUsuarioForm[password]": senha,
        "EntidadeUsuarioForm[rememberMe]": "0",
        "login-button": "",
    }, headers={"Referer": LOGIN}, timeout=TIMEOUT)
    resposta.raise_for_status()
    if not _no_menu(resposta.text):
        raise LoginRecusado("o Diário não abriu o painel depois do login")
    return sessao


def logado(sessao: requests.Session) -> bool:
    """A sessão ainda vale?"""
    resposta = sessao.get(PUBLICACOES, timeout=TIMEOUT)
    return resposta.ok and _no_menu(resposta.text)


# --- Tela de cadastro ---------------------------------------------------------

def _opcoes(doc, nome: str) -> list[tuple[str, str]]:
    """[(valor, texto)] da lista 'nome', sem itens vazios."""
    return [(o.get("value"), " ".join(o.text_content().split()))
            for o in doc.xpath(f'//select[@name="{nome}"]/option') if o.get("value")]


def _opcao(doc, nome: str, texto: str, padrao: str) -> str:
    """Valor da opção pelo texto (exato, ou contido); se não achar, o valor da gravação,
    desde que ele ainda exista na lista."""
    opcoes = _opcoes(doc, nome)
    alvo = texto.upper()
    achada = (next((v for v, t in opcoes if t.upper() == alvo), None)
              or next((v for v, t in opcoes if alvo in t.upper()), None))
    if achada is not None:
        return achada
    if any(v == padrao for v, _ in opcoes):
        return padrao
    curto = nome.split("[")[-1].rstrip("]")
    raise ErroDiario(f"opção '{texto}' não existe em {curto} "
                     f"(opções: {', '.join(t for _, t in opcoes) or 'nenhuma'})")


def _mensagens(doc) -> list[str]:
    """Mensagens de erro/validação da tela (para explicar uma recusa)."""
    textos = []
    for el in doc.xpath('//*[contains(@class,"help-block-error") or contains(@class,"invalid-feedback")'
                        ' or contains(@class,"alert-danger") or contains(@class,"error-summary")]'):
        t = " ".join(el.text_content().split())
        if t and t not in textos:
            textos.append(t)
    return textos[:3]


def abrir_create(sessao: requests.Session) -> lxml.html.HtmlElement:
    """Abre a tela de cadastro de publicação e devolve o HTML dela."""
    resposta = sessao.get(CRIAR, headers={"Referer": PUBLICACOES}, timeout=TIMEOUT)
    resposta.raise_for_status()
    if not _no_menu(resposta.text):
        raise SessaoExpirada("o Diário voltou para a tela de login")
    return _html(resposta)


def _id_entidade(doc) -> str:
    """ID do usuário logado (Documento[id_entidade_usuario]), que vem no formulário."""
    campo = doc.xpath('//*[@name="Documento[id_entidade_usuario]"]/@value')
    if not campo:
        raise ErroDiario("não achei o id_entidade_usuario na tela de cadastro (mudou o layout?)")
    return campo[0]


def montar_campos(doc, titulo: str, corpo_html: str, data_veiculacao) -> list[tuple[str, str]]:
    """Campos do formulário na ordem em que o navegador manda. data_veiculacao é um
    datetime.date (a data de veiculação, já no próximo dia útil)."""
    token = _csrf(doc)
    if not token:
        raise ErroDiario("não achei o token CSRF na tela de cadastro")
    return [
        ("_csrf-backend", token),
        ("Documento[id_entidade_usuario]", _id_entidade(doc)),
        ("Documento[titulo]", titulo),
        ("data_veiculacao-documento-data_veiculacao-disp", data_veiculacao.strftime("%d/%m/%Y")),
        ("Documento[data_veiculacao]", data_veiculacao.strftime("%Y-%m-%d")),
        ("Documento[mostrar_anexo]", "0"),
        (SECRETARIA[0], _opcao(doc, *SECRETARIA)),
        (CATEGORIA[0], _opcao(doc, *CATEGORIA)),
        (COLUNA[0], _opcao(doc, *COLUNA)),
        ("Documento[descricao]", corpo_html),
        # O campo de anexo vai vazio duas vezes, como no navegador: o hidden do Yii e o
        # <input type=file> sem arquivo (este entra em _multipart, com filename vazio)
        ("Documento[file]", ""),
    ]


def _multipart(campos: list[tuple[str, str]]) -> CurlMime:
    """Corpo multipart/form-data igual ao do navegador (o curl_cffi não aceita files=)."""
    mime = CurlMime()
    for nome, valor in campos:
        mime.addpart(name=nome, data=valor.encode("utf-8"))
    mime.addpart(name="Documento[file]", filename="", content_type="application/octet-stream",
                 data=b"")
    return mime


def _id_do_documento(sessao: requests.Session, titulo: str) -> str:
    """Procura na lista de publicações o documento recém-criado (pelo título) e devolve o
    id dele (a lista vem do mais novo para o mais antigo). "" se não achar."""
    resposta = sessao.get(LISTA, timeout=TIMEOUT)
    if not resposta.ok or not resposta.content.strip():
        return ""
    doc = _html(resposta)
    alvo = titulo.strip().upper()
    for linha in doc.xpath('//tr'):
        if alvo in " ".join(linha.text_content().split()).upper():
            for href in linha.xpath('.//a[contains(@href,"/view/")]/@href'):
                m = re.search(r"/view/(\d+)", href)
                if m:
                    return m.group(1)
    return ""


def _confirmar_cadastro(sessao: requests.Session, titulo: str, tentativas: int = 3) -> str:
    """Confere no Diário que o ato foi gravado: ele tem que aparecer na lista de publicações
    e a página dele (/view/<id>) tem que abrir (HTTP 200) com o título. Devolve o número.
    Lança ErroDiario se não confirmar (aí o status não é marcado como enviado)."""
    for tentativa in range(tentativas):
        if tentativa:
            time.sleep(1.5)  # a lista pode demorar um pouco a mostrar o ato novo
        numero = _id_do_documento(sessao, titulo)
        if numero:
            pagina = sessao.get(f"{PUBLICACOES}/view/{numero}", timeout=TIMEOUT)
            if pagina.status_code == 200 and titulo.strip().upper() in " ".join(
                    _html(pagina).text_content().split()).upper():
                return numero
    raise ErroDiario("o Diário respondeu que cadastrou, mas o ato não apareceu na lista de "
                     "publicações (confira no Diário antes de enviar de novo)")


def enviar_ato(sessao: requests.Session, titulo: str, corpo_html: str, data_veiculacao,
               avisar=lambda msg: None) -> str:
    """Cadastra a publicação (título, data de veiculação, secretaria, categoria, coluna e
    corpo). Devolve o número do documento criado, depois de confirmar que ele existe.
    Lança ErroDiario se o site recusar ou não confirmar e SessaoExpirada se a sessão cair."""
    doc = abrir_create(sessao)
    mime = _multipart(montar_campos(doc, titulo, corpo_html, data_veiculacao))
    try:
        # Sem seguir o redirecionamento: o 302 para a lista é a confirmação do cadastro
        resposta = sessao.post(CRIAR, multipart=mime, headers={"Referer": CRIAR},
                               timeout=TIMEOUT, allow_redirects=False)
    finally:
        mime.close()

    destino = resposta.headers.get("Location", "")
    if resposta.status_code in (301, 302, 303) and "minhas-publicacoes" in destino:
        # O 302 diz que o site aceitou; só conta como enviado quando o ato aparece de fato
        numero = _confirmar_cadastro(sessao, titulo)
        avisar(f"Cadastrado: {titulo} · veiculação {data_veiculacao:%d/%m/%Y} · documento {numero}")
        return numero
    if resposta.status_code in (301, 302, 303) and "login" in destino:
        raise SessaoExpirada("a sessão caiu no momento do envio")
    if resposta.status_code != 200:
        raise ErroDiario(f"resposta inesperada do Diário (HTTP {resposta.status_code} {destino})".strip())

    # 200 = o site devolveu o formulário: recusou. Mostra o motivo.
    if not _no_menu(resposta.text):
        raise SessaoExpirada("a sessão caiu no momento do envio")
    msgs = _mensagens(_html(resposta))
    motivo = "; ".join(msgs) if msgs else "a página não saiu do formulário"
    raise ErroDiario(f"o Diário não confirmou o cadastro ({motivo})")


# --- Testes -------------------------------------------------------------------

def _credenciais() -> tuple[str, str]:
    import credenciais

    dados = credenciais.carregar(Path(__file__).parent / "configs" / "credenciais.json", "diario")
    if not dados:
        sys.exit("Nenhum usuário do Diário salvo. Cadastre no programa (opção 4 → 1).")
    return dados


def _testar_login() -> None:
    usuario, senha = _credenciais()
    print(f"Entrando como {usuario}...")
    try:
        sessao = entrar(usuario, senha)
    except LoginRecusado:
        sys.exit("Login recusado: confira usuário e senha (opção 4 → 1 do programa).")
    print(f"No painel: {'sim' if logado(sessao) else 'NÃO'}")


def _testar_campos() -> None:
    """Mostra os valores das listas da tela de cadastro e quais seriam escolhidos."""
    usuario, senha = _credenciais()
    try:
        sessao = entrar(usuario, senha)
        doc = abrir_create(sessao)
    except (ErroDiario, LoginRecusado, SessaoExpirada) as e:
        sys.exit(f"Parou: {e}")
    print(f"id_entidade_usuario: {_id_entidade(doc)}")
    for nome, texto, padrao in (SECRETARIA, CATEGORIA, COLUNA):
        try:
            escolhida = _opcao(doc, nome, texto, padrao)
        except ErroDiario as e:
            escolhida = f"ERRO: {e}"
        print(f"\n{nome} → {escolhida}")
        for valor, rotulo in _opcoes(doc, nome):
            print(f"  {valor}  {rotulo}")


def _testar_enviar() -> None:
    """Fluxo completo: cadastra e ENVIA o 1º ato pendente do Diário."""
    import main
    import ui

    main.preparar_pastas()
    if not main.escolher_processo():
        return
    preparado = main.preparar_atos("diario", depois_de="edocs", sem_titulo=True)
    if not preparado:
        return
    registros, textos, htmls, _ = preparado
    dados, texto, corpo = registros[0], textos[0], htmls[0]
    titulo = main.proc().titulo_diario(dados)
    data_pub = main.data_publicacao(dados)
    ui.info(f"Ato: {main.proc().descricao(dados)} (linha {dados['_LINHA']})")
    ui.info(f"Título: {titulo} · veiculação {data_pub:%d/%m/%Y}")

    usuario, senha = _credenciais()
    ui.titulo("Envio ao Diário (IRREVERSÍVEL)")
    if "[" in texto:
        ui.aviso("O texto tem campos vazios entre colchetes: "
                 + ", ".join(sorted(set(re.findall(r"\[[^\]]+\]", texto)))))
    if ui.console.input("Cadastrar e ENVIAR este ato no Diário? (S/N): ").strip().upper() != "S":
        ui.aviso("Nada enviado")
        return

    with ui.processando("Entrando no Diário"):
        sessao = entrar(usuario, senha)
    try:
        numero = enviar_ato(sessao, titulo, corpo, data_pub, avisar=ui.ok)
    except (ErroDiario, SessaoExpirada) as e:
        ui.erro(f"Parou: {e}")
        return
    ui.ok(f"Publicado {titulo}" + (f" (documento {numero})" if numero else ""))


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    modo = sys.argv[1:]
    if modo == ["campos"]:
        _testar_campos()
    elif modo == ["enviar"]:
        _testar_enviar()
    else:
        _testar_login()
