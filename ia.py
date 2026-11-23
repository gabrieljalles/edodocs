"""Resumo dos recebidos com inteligência artificial, de qualquer provedor.

A pessoa pode cadastrar vários provedores (Anthropic, OpenAI, DeepSeek, Gemini, um servidor
próprio...) e escolhe qual está em uso. Dois tipos de API:
  anthropic  a do Claude: o PDF vai inteiro (até 32 MB e 600 páginas por pedido)
  openai     o formato que quase todos aceitam (/chat/completions): vai só o texto do PDF
             (PDF digitalizado, só imagem, não dá para resumir)
A chave de cada provedor fica criptografada em configs/credenciais.json (site "ia:<nome>");
a lista dos provedores (tipo, endereço, modelos, o escolhido) e o em uso ficam em
configs/config.json ("ia_provedores", "ia_ativo"). O resumo fica guardado em
configs/resumos.json, para não pagar de novo pelo mesmo documento.
"""
import base64
import json
import sys
from pathlib import Path

import credenciais
import main

MODELO_PADRAO = "claude-opus-5-5"
# Modelos que aceitam o "plano B" do servidor quando o pedido é recusado (fallbacks)
COM_FALLBACK = {"claude-opus-5-5", "claude-opus-5", "claude-fable-5-1", "claude-sonnet-5-5"}
LIMITE_BYTES = 24 * 1024 * 1024  # o PDF em base64 cresce ~33%: fica abaixo dos 32 MB do pedido
LIMITE_TEXTO = 150_000           # caracteres do PDF mandados às IAs que só leem texto
TIMEOUT = 180
RESUMOS = main.BASE / "configs" / "resumos.json"

ANTHROPIC = "Anthropic (Claude)"
OUTRO = "Outro (compatível com OpenAI)"
# Provedores conhecidos: tipo da API, endereço e o modelo sugerido ("" = escolher da lista).
# "sem_chave": roda no próprio computador, não pede chave.
MODELOS_PRONTOS = {
    ANTHROPIC: {"tipo": "anthropic", "url": "", "modelo": MODELO_PADRAO},
    "OpenAI (ChatGPT)": {"tipo": "openai", "url": "https://api.openai.com/v1", "modelo": ""},
    "DeepSeek": {"tipo": "openai", "url": "https://api.deepseek.com", "modelo": "deepseek-chat"},
    "Google (Gemini)": {"tipo": "openai", "url": "https://generativelanguage.googleapis.com/v1beta/openai",
                        "modelo": ""},
    "Groq": {"tipo": "openai", "url": "https://api.groq.com/openai/v1", "modelo": ""},
    "OpenRouter": {"tipo": "openai", "url": "https://openrouter.ai/api/v1", "modelo": ""},
    "Mistral": {"tipo": "openai", "url": "https://api.mistral.ai/v1", "modelo": ""},
    "Ollama (no computador)": {"tipo": "openai", "url": "http://localhost:11434/v1", "modelo": "",
                               "sem_chave": True},
    OUTRO: {"tipo": "openai", "url": "", "modelo": ""},
}

INSTRUCOES = (
    "Você resume documentos oficiais recebidos no e-docs da Prefeitura de Uberlândia para um "
    "servidor que precisa decidir rápido o que fazer. Escreva em português do Brasil, em texto "
    "corrido e simples, sem markdown, com no máximo 5 frases curtas: do que se trata, quem pede, "
    "o que precisa ser feito (se precisar) e prazos ou datas importantes. Se houver anexos, diga "
    "quais são em uma frase.")


class SemChave(Exception):
    """A IA ainda não foi configurada (Configurações → Inteligência artificial)."""


class ErroIA(RuntimeError):
    """O provedor recusou o pedido (a mensagem diz o código e o motivo)."""


# --- Provedores --------------------------------------------------------------------------------
def _site(nome: str) -> str:
    return f"ia:{nome}"


def _migrar() -> None:
    """Versão antiga: uma chave só, da Anthropic, no site "ia" (e ia_modelos/ia_modelo soltos).
    Vira o provedor "Anthropic (Claude)", em uso."""
    antiga = credenciais.carregar(main.CREDENCIAIS, "ia")
    if not antiga:
        return
    cfg = main.ler_config()
    provedores = cfg.get("ia_provedores") or []
    if not any(p["nome"] == ANTHROPIC for p in provedores):
        modelos = cfg.get("ia_modelos") or [MODELO_PADRAO]
        provedores.append({"nome": ANTHROPIC, "tipo": "anthropic", "url": "", "modelos": modelos,
                           "modelo": cfg.get("ia_modelo") if cfg.get("ia_modelo") in modelos else modelos[0]})
        credenciais.salvar(main.CREDENCIAIS, ANTHROPIC, antiga[1], site=_site(ANTHROPIC))
    main.salvar_config(ia_provedores=provedores, ia_ativo=cfg.get("ia_ativo") or ANTHROPIC)
    credenciais.remover(main.CREDENCIAIS, "ia")


def provedores() -> list[dict]:
    """Os provedores cadastrados: {"nome", "tipo", "url", "modelos", "modelo"}."""
    _migrar()
    return list(main.ler_config().get("ia_provedores") or [])


def provedor(nome: str | None = None) -> dict | None:
    """O provedor pelo nome (sem nome: o em uso; se o em uso sumiu, o primeiro)."""
    lista = provedores()
    alvo = nome or main.ler_config().get("ia_ativo")
    achado = next((p for p in lista if p["nome"] == alvo), None)
    if achado or nome:
        return achado
    return lista[0] if lista else None


def chave(nome: str | None = None) -> str | None:
    """A chave do provedor ("" no que roda sem chave; None se não foi cadastrado)."""
    p = provedor(nome)
    if not p:
        return None
    dados = credenciais.carregar(main.CREDENCIAIS, _site(p["nome"]))
    return dados[1] if dados else None


def conectada() -> bool:
    """Tem um provedor em uso com a chave salva. Sem ele, nada de resumo aparece no programa
    (nem o do Diário, nem o dos recebidos): a pessoa só fica sabendo depois de conectar."""
    return chave() is not None


def salvar_provedor(nome: str, tipo: str, url: str, valor: str) -> None:
    """Cadastra (ou atualiza a chave/endereço de) um provedor e deixa ele em uso. Um novo vem
    com o modelo sugerido do provedor (se tiver)."""
    lista = provedores()
    atual = next((p for p in lista if p["nome"] == nome), None)
    if atual is None:
        sugerido = MODELOS_PRONTOS.get(nome, {}).get("modelo", "")
        atual = {"nome": nome, "modelos": [sugerido] if sugerido else [], "modelo": sugerido}
        lista.append(atual)
    atual.update(tipo=tipo, url=url.strip().rstrip("/"))
    credenciais.salvar(main.CREDENCIAIS, nome, valor.strip(), site=_site(nome))
    main.salvar_config(ia_provedores=lista, ia_ativo=nome)


def remover_provedor(nome: str) -> None:
    lista = [p for p in provedores() if p["nome"] != nome]
    credenciais.remover(main.CREDENCIAIS, _site(nome))
    ativo = main.ler_config().get("ia_ativo")
    main.salvar_config(ia_provedores=lista,
                       ia_ativo=(ativo if ativo != nome else (lista[0]["nome"] if lista else "")))


def usar_provedor(nome: str) -> None:
    main.salvar_config(ia_ativo=nome)


def _mudar(nome: str, **valores) -> None:
    lista = provedores()
    for p in lista:
        if p["nome"] == nome:
            p.update(valores)
    main.salvar_config(ia_provedores=lista)


def modelos_instalados(nome: str | None = None) -> list[str]:
    p = provedor(nome)
    return list(p.get("modelos") or []) if p else []


def modelo_escolhido(nome: str | None = None) -> str:
    """O modelo em uso do provedor ("" se nenhum foi adicionado ainda)."""
    p = provedor(nome)
    if not p:
        return ""
    modelos = p.get("modelos") or []
    return p.get("modelo") if p.get("modelo") in modelos else (modelos[0] if modelos else "")


def escolher_modelo(nome: str, modelo: str) -> None:
    _mudar(nome, modelo=modelo)


def adicionar_modelo(nome: str, modelo: str) -> None:
    modelos = modelos_instalados(nome)
    if modelo not in modelos:
        _mudar(nome, modelos=modelos + [modelo], **({} if modelos else {"modelo": modelo}))


def tirar_modelo(nome: str, modelo: str) -> None:
    _mudar(nome, modelos=[m for m in modelos_instalados(nome) if m != modelo])


def _cliente(valor: str | None = None):
    import anthropic  # só quem usa o Claude carrega a biblioteca

    valor = valor or chave()
    if not valor:
        raise SemChave("configure a chave da IA em Configurações")
    return anthropic.Anthropic(api_key=valor)


def _cabecalhos(valor: str | None) -> dict:
    return {"Authorization": f"Bearer {valor}"} if valor else {}


def _erro(resposta) -> ErroIA:
    try:
        corpo = resposta.json()
        motivo = (corpo.get("error") or {}).get("message") if isinstance(corpo.get("error"), dict) \
            else corpo.get("error") or corpo.get("message")
    except ValueError:
        motivo = resposta.text
    return ErroIA(f"HTTP {resposta.status_code}" + (f": {str(motivo)[:160]}" if motivo else ""))


def modelos_disponiveis(tipo: str, url: str, valor: str | None) -> list[tuple[str, str]]:
    """[(id, nome)] dos modelos que esta chave pode usar (também serve para testar a chave)."""
    if tipo == "anthropic":
        return [(m.id, m.display_name) for m in _cliente(valor).models.list()]
    import httpx

    if not url:
        raise ErroIA("falta o endereço da API")
    resposta = httpx.get(f"{url.rstrip('/')}/models", headers=_cabecalhos(valor), timeout=30)
    if resposta.status_code != 200:
        raise _erro(resposta)
    dados = resposta.json()
    itens = dados.get("data", dados.get("models", [])) if isinstance(dados, dict) else dados
    return sorted((m["id"], m.get("name") or m["id"]) for m in itens if isinstance(m, dict) and m.get("id"))


# --- Resumos -----------------------------------------------------------------------------------
def _resumos() -> dict:
    try:
        return json.loads(RESUMOS.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def resumo_guardado(doc: str) -> str | None:
    return _resumos().get(doc, {}).get("resumo")


def _guardar(doc: str, resumo: str, modelo: str) -> None:
    todos = _resumos()
    todos[doc] = {"resumo": resumo, "modelo": modelo}
    RESUMOS.parent.mkdir(parents=True, exist_ok=True)
    RESUMOS.write_text(json.dumps(todos, ensure_ascii=False, indent=1), encoding="utf-8")


def resumir_pdf(doc: str, pdf: bytes, titulo: str = "") -> str:
    """Resumo do PDF (o guardado, se já existir). Lança SemChave sem chave configurada."""
    guardado = resumo_guardado(doc)
    if guardado:
        return guardado
    if len(pdf) > LIMITE_BYTES:
        return ("O documento é grande demais para resumir de uma vez "
                f"({len(pdf) // (1024 * 1024)} MB). Abra o PDF para ler.")
    prov = provedor()
    if not prov or chave() is None:
        raise SemChave("configure a IA em Configurações")
    modelo = modelo_escolhido()
    if not modelo:
        return f"Escolha um modelo para {prov['nome']} em Configurações → Inteligência artificial."
    if prov.get("tipo") != "anthropic":
        texto = _resumir_texto(prov, modelo, pdf, titulo)
        if texto and not texto.startswith(SEM_TEXTO):
            _guardar(doc, texto, modelo)
        return texto or "A IA não devolveu um resumo. Abra o PDF para ler."
    cliente = _cliente()
    conteudo = [
        {"type": "document", "source": {"type": "base64", "media_type": "application/pdf",
                                        "data": base64.standard_b64encode(pdf).decode("ascii")}},
        {"type": "text", "text": f"Resuma este documento recebido{f' ({titulo})' if titulo else ''}."},
    ]
    pedido = dict(model=modelo, max_tokens=4000, system=INSTRUCOES,
                  messages=[{"role": "user", "content": conteudo}])
    if modelo in COM_FALLBACK:
        # Se o modelo recusar o pedido, o servidor tenta outro modelo na mesma chamada
        resposta = cliente.beta.messages.create(**pedido, betas=["server-side-fallback-2026-07-01"],
                                                fallbacks="default")
    else:
        resposta = cliente.messages.create(**pedido)
    if resposta.stop_reason == "refusal":
        return "A IA não quis resumir este documento. Abra o PDF para ler."
    texto = " ".join(b.text for b in resposta.content if b.type == "text").strip()
    if texto:
        _guardar(doc, texto, modelo)
    return texto or "A IA não devolveu um resumo. Abra o PDF para ler."


SEM_TEXTO = "Este PDF é uma imagem"


def texto_do_pdf(pdf: bytes) -> str:
    """O texto do PDF (para as IAs que não leem PDF). "" num PDF digitalizado (só imagem)."""
    import pypdfium2 as pdfium

    documento = pdfium.PdfDocument(pdf)
    try:
        return "\n".join(pagina.get_textpage().get_text_range() for pagina in documento).strip()
    finally:
        documento.close()


def _resumir_texto(prov: dict, modelo: str, pdf: bytes, titulo: str) -> str:
    """Resumo pela API no formato da OpenAI (/chat/completions), com o texto do PDF."""
    import httpx

    texto = texto_do_pdf(pdf)
    if not texto:
        return f"{SEM_TEXTO} (digitalizado) e {prov['nome']} só lê texto. Abra o PDF para ler."
    if len(texto) > LIMITE_TEXTO:
        texto = texto[:LIMITE_TEXTO] + "\n[... o resto do documento foi cortado]"
    pedido = {"model": modelo, "max_tokens": 4000, "messages": [
        {"role": "system", "content": INSTRUCOES},
        {"role": "user", "content": f"Resuma este documento recebido{f' ({titulo})' if titulo else ''}."
                                    f"\n\n{texto}"}]}
    resposta = httpx.post(f"{prov['url'].rstrip('/')}/chat/completions", json=pedido,
                          headers=_cabecalhos(chave()), timeout=TIMEOUT)
    if resposta.status_code != 200:
        raise _erro(resposta)
    escolhas = resposta.json().get("choices") or [{}]
    return str((escolhas[0].get("message") or {}).get("content") or "").strip()


# --- Montar processo com IA ----------------------------------------------------------------------
# O guia (assets/ia/guia_processos.md) ensina à IA como um processo funciona e o formato do plano
# que ela devolve. Mudou como o Edodocs monta os atos (processos.py, atos.py)? Atualize o guia:
# testes/testar_guia_ia.py confere se ele cita todos os tipos de regra, operadores e campos.
GUIA = Path(getattr(sys, "_MEIPASS", Path(__file__).parent)) / "assets" / "ia" / "guia_processos.md"
ARQUIVO_INSTRUCOES = "instruções para a IA.txt"  # o que a equipe aprendeu, sem compilar de novo
SAIDA_MAXIMA = 64000     # tokens da resposta (o plano de um processo grande passa de 20 mil)
TIMEOUT_PLANO = 900      # segundos: montar um processo demora bem mais que um resumo


def arquivo_instrucoes() -> Path:
    """As instruções da equipe: na pasta compartilhada (valem para todos) ou em configs/."""
    return (main.extras_compartilhados() or main.BASE / "configs") / ARQUIVO_INSTRUCOES


def guia_processos() -> str:
    """O guia, com as instruções da equipe no fim (quando houver)."""
    texto = GUIA.read_text(encoding="utf-8")
    try:
        extras = arquivo_instrucoes().read_text(encoding="utf-8").strip()
    except OSError:
        extras = ""
    if extras:
        texto += ("\n\n---\n\n## 9. Instruções da equipe da Prefeitura\n\n"
                  "Valem mais que o resto deste guia quando os dois discordarem.\n\n" + extras)
    return texto


def ler_json(texto: str) -> dict:
    """O objeto JSON da resposta, mesmo que a IA tenha posto ``` ou uma frase em volta."""
    inicio, fim = texto.find("{"), texto.rfind("}")
    if inicio < 0 or fim <= inicio:
        raise ErroIA("a IA não devolveu o plano no formato esperado")
    try:
        dados = json.loads(texto[inicio:fim + 1])
    except ValueError as e:
        raise ErroIA(f"a IA devolveu um plano com erro de formato ({e.msg}, linha {e.lineno})") from None
    if not isinstance(dados, dict):
        raise ErroIA("a IA não devolveu o plano no formato esperado")
    return dados


def pedir_json(sistema: str, pedido: str, max_tokens: int = SAIDA_MAXIMA) -> dict:
    """Um pedido à IA em uso (qualquer provedor) que responde com um objeto JSON. Cada pedido
    vai sozinho, sem conversa antes: numa correção, o pedido já leva o plano e os erros."""
    prov = provedor()
    if not prov or chave() is None:
        raise SemChave("configure a IA em Configurações")
    modelo = modelo_escolhido()
    if not modelo:
        raise ErroIA(f"escolha um modelo para {prov['nome']} em Configurações → Inteligência artificial")
    if prov.get("tipo") == "anthropic":
        texto = _json_anthropic(modelo, sistema, pedido, max_tokens)
    else:
        texto = _json_openai(prov, modelo, sistema, pedido, max_tokens)
    return ler_json(texto)


def _json_anthropic(modelo: str, sistema: str, pedido: str, max_tokens: int) -> str:
    cliente = _cliente().with_options(timeout=TIMEOUT_PLANO)
    # O guia é igual em todo pedido: fica guardado (cache) e a correção sai mais barata
    pedido_api = dict(model=modelo, max_tokens=max_tokens,
                      system=[{"type": "text", "text": sistema, "cache_control": {"type": "ephemeral"}}],
                      messages=[{"role": "user", "content": pedido}])
    if modelo in COM_FALLBACK:
        # Modelos atuais: pensam antes de responder (montar um processo é trabalho difícil) e,
        # se recusarem, o servidor tenta outro modelo na mesma chamada
        with cliente.beta.messages.stream(**pedido_api, output_config={"effort": "high"},
                                          betas=["server-side-fallback-2026-07-01"],
                                          fallbacks="default") as fluxo:
            resposta = fluxo.get_final_message()
    else:
        with cliente.messages.stream(**pedido_api) as fluxo:
            resposta = fluxo.get_final_message()
    if resposta.stop_reason == "refusal":
        raise ErroIA("a IA não quis montar este processo")
    if resposta.stop_reason == "max_tokens":
        raise ErroIA("o plano ficou grande demais para uma resposta; tente com menos exemplos")
    return "".join(b.text for b in resposta.content if b.type == "text")


def _json_openai(prov: dict, modelo: str, sistema: str, pedido: str, max_tokens: int) -> str:
    """Pela API no formato da OpenAI (ChatGPT, Gemini, DeepSeek, Groq, Ollama...)."""
    import httpx

    corpo = {"model": modelo, "max_tokens": max_tokens, "response_format": {"type": "json_object"},
             "messages": [{"role": "system", "content": sistema}, {"role": "user", "content": pedido}]}
    endereco = f"{prov['url'].rstrip('/')}/chat/completions"
    resposta = httpx.post(endereco, json=corpo, headers=_cabecalhos(chave()), timeout=TIMEOUT_PLANO)
    if resposta.status_code == 400:
        # Nem todo provedor aceita o modo JSON ou uma resposta tão longa: tenta sem os dois
        corpo.pop("response_format")
        corpo.pop("max_tokens")
        resposta = httpx.post(endereco, json=corpo, headers=_cabecalhos(chave()), timeout=TIMEOUT_PLANO)
    if resposta.status_code != 200:
        raise _erro(resposta)
    escolha = (resposta.json().get("choices") or [{}])[0]
    if escolha.get("finish_reason") == "length":
        raise ErroIA("o plano ficou grande demais para uma resposta; tente com menos exemplos ou outro modelo")
    return str((escolha.get("message") or {}).get("content") or "")
