"""Testes offline do envio de atos no Diário (não acessa o site).

Sobem um servidor local que imita o Diário (login, tela de cadastro, 302 de sucesso ou
formulário de volta com erro) e conferem:
  1. o login manda o token CSRF e reconhece o painel;
  2. o cadastro manda os mesmos campos, na mesma ordem, que o navegador mandou na
     gravação (capturas/diario.uberlandia.mg.gov.br.har), inclusive o anexo vazio;
  3. o 302 para a lista é reconhecido como cadastrado e o número do documento é lido;
  4. formulário de volta (200) vira ErroDiario com a mensagem de validação;
  5. login recusado vira LoginRecusado;
  6. o fluxo por requisições não importa o Selenium.

Rode com:  python testes/testar_envio_diario.py
"""
import email.parser
import email.policy
import json
import subprocess
import sys
import threading
from datetime import date
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import diario_http  # noqa: E402

HAR = RAIZ / "capturas" / "diario.uberlandia.mg.gov.br.har"
TOKEN = "tok-123"
TITULO = "ATO ADMINISTRATIVO SMA Nº 2599/2026"

LOGIN_HTML = f"""<html><head><meta name="csrf-token" content="{TOKEN}"></head><body>
<form><input type="hidden" name="_csrf-backend" value="{TOKEN}">
<input name="EntidadeUsuarioForm[username]"></form></body></html>"""
MENU = '<a href="/admin/logout" data-method="post">Deslogar</a>'
CREATE_HTML = f"""<html><body>{MENU}<form enctype="multipart/form-data">
<input type="hidden" name="_csrf-backend" value="{TOKEN}">
<input type="hidden" name="Documento[id_entidade_usuario]" value="503">
<select name="Documento[id_secretaria]"><option value=""></option>
  <option value="449">Agricultura</option><option value="450">Administração</option></select>
<select name="Documento[id_documento_categoria]"><option value="">Selecione</option>
  <option value="566">Aviso</option><option value="567">Ato Administrativo</option></select>
<select name="Documento[coluna]"><option value="A">Uma coluna</option>
  <option value="B">Duas colunas</option></select>
</form></body></html>"""
INDEX_HTML = f"""<html><body>{MENU}<table>
<tr><td>{TITULO}</td><td><a href="/admin/minhas-publicacoes/view/52508">ver</a></td></tr>
<tr><td>Outro</td><td><a href="/admin/minhas-publicacoes/view/52507">ver</a></td></tr>
</table></body></html>"""
ERRO_HTML = f"""<html><body>{MENU}<form>
<div class="help-block help-block-error">Título não pode ficar em branco.</div>
</form></body></html>"""


class Servidor(BaseHTTPRequestHandler):
    recusar_envio = False
    sumir_da_lista = False
    senha_certa = "certa"
    envios: list = []

    def log_message(self, *args):
        pass

    def _responder(self, status: int, corpo: str = "", local: str | None = None):
        dados = corpo.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=UTF-8")
        self.send_header("Content-Length", str(len(dados)))
        if local:
            self.send_header("Location", local)
        self.end_headers()
        self.wfile.write(dados)

    def do_GET(self):
        logado = "sessao=ok" in (self.headers.get("Cookie") or "")
        if self.path == "/admin/login":
            self.send_response(200)
            self.send_header("Set-Cookie", "sessao=nova; Path=/")
            dados = LOGIN_HTML.encode()
            self.send_header("Content-Length", str(len(dados)))
            self.end_headers()
            self.wfile.write(dados)
        elif not logado:
            self._responder(302, local="/admin/login")
        elif self.path in ("/admin/", "/admin/minhas-publicacoes"):
            self._responder(200, f"<html><body>{MENU}</body></html>")
        elif self.path == "/admin/minhas-publicacoes/create":
            self._responder(200, CREATE_HTML)
        elif self.path == "/admin/minhas-publicacoes/index":
            self._responder(200, "" if Servidor.sumir_da_lista else INDEX_HTML)
        elif self.path == "/admin/minhas-publicacoes/view/52508":
            self._responder(200, f"<html><body>{MENU}<h4>Documento: {TITULO}</h4></body></html>")
        else:
            self._responder(404)

    def do_POST(self):
        corpo = self.rfile.read(int(self.headers.get("Content-Length", 0)))
        if self.path == "/admin/login":
            texto = corpo.decode()
            if f"_csrf-backend={TOKEN}" in texto and f"%5Bpassword%5D={self.senha_certa}" in texto:
                self.send_response(302)
                self.send_header("Set-Cookie", "sessao=ok; Path=/")
                self.send_header("Location", "/admin/")
                self.send_header("Content-Length", "0")
                self.end_headers()
            else:
                self._responder(200, LOGIN_HTML)
        elif self.path == "/admin/minhas-publicacoes/create":
            Servidor.envios.append((self.headers.get("Content-Type", ""), corpo))
            if Servidor.recusar_envio:
                self._responder(200, ERRO_HTML)
            else:
                self._responder(302, local="http://localhost/admin/minhas-publicacoes/index")
        else:
            self._responder(404)


def subir_servidor() -> ThreadingHTTPServer:
    servidor = ThreadingHTTPServer(("127.0.0.1", 0), Servidor)
    threading.Thread(target=servidor.serve_forever, daemon=True).start()
    base = f"http://127.0.0.1:{servidor.server_port}/admin/"
    diario_http.URL = base
    diario_http.LOGIN = base + "login"
    diario_http.PUBLICACOES = base + "minhas-publicacoes"
    diario_http.CRIAR = diario_http.PUBLICACOES + "/create"
    diario_http.LISTA = diario_http.PUBLICACOES + "/index"
    diario_http.PAUSA_REQUISICAO = 0
    diario_http.time.sleep = lambda s: None  # sem esperar nas novas tentativas
    return servidor


def partes(content_type: str, corpo: bytes) -> list[tuple[str, str | None, str]]:
    """[(nome, filename, valor)] de um corpo multipart, na ordem."""
    msg = email.parser.BytesParser(policy=email.policy.HTTP).parsebytes(
        f"Content-Type: {content_type}\r\n\r\n".encode() + corpo)
    saida = []
    for p in msg.iter_parts():
        nome = p.get_param("name", header="content-disposition")
        arquivo = p.get_param("filename", header="content-disposition")
        saida.append((nome, arquivo, p.get_payload(decode=True).decode("utf-8")))
    return saida


def partes_da_gravacao() -> list[tuple[str, str | None]]:
    """[(nome, filename)] do envio gravado no HAR (None se o HAR não estiver na pasta)."""
    if not HAR.exists():
        return None
    for e in json.loads(HAR.read_text(encoding="utf-8"))["log"]["entries"]:
        r = e["request"]
        if r["method"] == "POST" and r["url"].endswith("/minhas-publicacoes/create"):
            return [(p["name"], p.get("fileName")) for p in r["postData"]["params"]]
    raise RuntimeError("não achei o POST de cadastro no HAR")


# --- testes ------------------------------------------------------------------

def teste_1_login() -> object:
    sessao = diario_http.entrar("fulano", "certa")
    assert diario_http.logado(sessao)
    print("  [1] login com token CSRF e painel reconhecido ........ OK")
    return sessao


def teste_2_3_envio(sessao) -> None:
    Servidor.envios.clear()
    avisos = []
    corpo = '<p style="text-align: justify;">O DIRETOR… <strong>RESOLVE:</strong></p>'
    numero = diario_http.enviar_ato(sessao, TITULO, corpo, date(2026, 10, 9), avisar=avisos.append)
    assert numero == "52508", numero
    assert len(Servidor.envios) == 1, "o envio não pode ser repetido"
    tipo, bruto = Servidor.envios[0]
    assert tipo.startswith("multipart/form-data"), tipo
    enviados = partes(tipo, bruto)
    valores = {n: v for n, _, v in enviados}
    assert valores["_csrf-backend"] == TOKEN
    assert valores["Documento[id_entidade_usuario]"] == "503"
    assert valores["Documento[titulo]"] == TITULO
    assert valores["data_veiculacao-documento-data_veiculacao-disp"] == "09/10/2026"
    assert valores["Documento[data_veiculacao]"] == "2026-10-09"
    assert valores["Documento[id_secretaria]"] == "450"
    assert valores["Documento[id_documento_categoria]"] == "567"
    assert valores["Documento[coluna]"] == "B"
    assert valores["Documento[descricao]"] == corpo
    assert enviados[-2:] == [("Documento[file]", None, ""), ("Documento[file]", "", "")], enviados[-2:]
    gravado = partes_da_gravacao()
    if gravado is not None:
        assert [(n, f or None) for n, f, _ in enviados] == [(n, f or None) for n, f in gravado], \
            "campos diferentes da gravação"
    print("  [2] campos iguais aos do navegador (ordem e anexo) ... OK"
          + ("" if gravado is not None else "  (sem HAR: ordem não comparada)"))
    assert avisos and "52508" in avisos[0], avisos
    print("  [3] 302 → lista = cadastrado, número 52508 lido ...... OK")


def teste_4_recusa(sessao) -> None:
    Servidor.recusar_envio = True
    try:
        diario_http.enviar_ato(sessao, TITULO, "<p>x</p>", date(2026, 10, 9))
    except diario_http.ErroDiario as e:
        assert "não pode ficar em branco" in str(e), e
    else:
        raise AssertionError("deveria ter levantado ErroDiario")
    finally:
        Servidor.recusar_envio = False
    print("  [4] formulário de volta vira erro com o motivo ........ OK")


def teste_4b_cadastro_nao_confirmado(sessao) -> None:
    # O site responde 302 (aceitou), mas o ato não aparece na lista: não conta como enviado
    Servidor.sumir_da_lista = True
    Servidor.envios.clear()
    try:
        diario_http.enviar_ato(sessao, TITULO, "<p>x</p>", date(2026, 10, 9))
    except diario_http.ErroDiario as e:
        assert "não apareceu na lista" in str(e), e
    else:
        raise AssertionError("deveria ter levantado ErroDiario")
    finally:
        Servidor.sumir_da_lista = False
    assert len(Servidor.envios) == 1, "o envio não pode ser repetido"
    print("  [4b] 302 sem o ato na lista NÃO é enviado ............ OK")


def teste_5_login_recusado() -> None:
    try:
        diario_http.entrar("fulano", "errada")
    except diario_http.LoginRecusado:
        pass
    else:
        raise AssertionError("deveria ter levantado LoginRecusado")
    print("  [5] senha errada vira LoginRecusado ................... OK")


def teste_6_sem_selenium() -> None:
    # Bloqueia o selenium e importa o caminho HTTP inteiro
    codigo = ("import sys; sys.modules['selenium'] = None; "
              "import app, main, diario_http, edocs_http, ui, atos; "
              "assert callable(ui.confirmar_envio) and callable(main.lancar_diario_http)")
    r = subprocess.run([sys.executable, "-c", codigo], cwd=RAIZ, capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    print("  [6] fluxo por requisições não importa o Selenium ...... OK")


def main() -> None:
    print("Testes do envio de atos no Diário (offline)\n")
    servidor = subir_servidor()
    try:
        sessao = teste_1_login()
        teste_2_3_envio(sessao)
        teste_4_recusa(sessao)
        teste_4b_cadastro_nao_confirmado(sessao)
        teste_5_login_recusado()
        teste_6_sem_selenium()
    finally:
        servidor.shutdown()
    print("\n7 testes passaram.")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
