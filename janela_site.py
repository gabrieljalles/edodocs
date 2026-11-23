"""Abre o e-docs ou o Diário Oficial numa janela própria, já logado com o usuário salvo em
Configurações (botão do globo na página Atos).

Não dá para passar o login do Edodocs para o Chrome/Edge da pessoa: os dois sites só aceitam
o login dentro de uma sessão aberta pelo próprio navegador (o e-docs pela sessão Java, o
Diário pelo token CSRF da página). Então a janela abre a página de login, preenche usuário e
senha e clica em Entrar, como a pessoa faria. Roda num processo separado (Edodocs.exe --site
edocs), para a janela do site ficar independente da janela do Edodocs.
"""
import json
import subprocess
import sys
from pathlib import Path

SITES = {
    "edocs": ("e-docs", "https://webedocspmu.uberlandia.mg.gov.br/webedocspmu/",
              "j_username", "j_password"),
    "diario": ("Diário Oficial", "https://diario.uberlandia.mg.gov.br/admin/login",
               "EntidadeUsuarioForm[username]", "EntidadeUsuarioForm[password]"),
}


def _script_de_login(campo_usuario: str, campo_senha: str, usuario: str, senha: str) -> str:
    """Preenche o formulário de login e clica no botão de enviar (se ele estiver na página)."""
    return f"""(function () {{
        var u = document.querySelector('[name="{campo_usuario}"]');
        var p = document.querySelector('[name="{campo_senha}"]');
        if (!u || !p) return false;
        u.value = {json.dumps(usuario)};
        p.value = {json.dumps(senha)};
        var botao = u.form && u.form.querySelector('[type="submit"]');
        if (botao) botao.click(); else u.form.submit();
        return true;
    }})()"""


def abrir(site: str) -> None:
    """Abre a janela do site (processo separado: o Edodocs continua livre)."""
    if getattr(sys, "frozen", False):
        comando = [sys.executable, "--site", site]
    else:
        comando = [sys.executable, str(Path(__file__).parent / "app.py"), "--site", site]
    subprocess.Popen(comando, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))


def rodar(site: str) -> None:
    """Dentro do processo separado: a janela do site com o login automático."""
    import webview

    import credenciais
    import main

    titulo, url, campo_usuario, campo_senha = SITES[site]
    dados = credenciais.carregar(main.CREDENCIAIS, site)
    janela = webview.create_window(f"{titulo} · Edodocs", url, width=1366, height=860)
    tentou = {"login": False}

    def carregou() -> None:
        # Uma tentativa só: se a senha estiver errada, a pessoa vê a tela de login e entra à mão
        if dados and not tentou["login"]:
            tentou["login"] = bool(janela.evaluate_js(_script_de_login(campo_usuario, campo_senha, *dados)))

    janela.events.loaded += carregou
    webview.start()
