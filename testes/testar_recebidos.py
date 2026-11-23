"""Testes offline da página Recebidos (não acessa o e-docs nem a IA).

Usam a Caixa de Entrada real gravada em capturas/vistar.har:
  1. junta as linhas por documento e diz o que dá para fazer (ciente, visto);
  2. baixar o PDF clica em Visualizar na linha certa, sem o histórico, e confere que é PDF;
  3. ciente clica no Ciente da linha do documento e só confirma quando o botão some;
  4. o resumo da IA manda o PDF como documento, usa o plano B (fallback) no Opus 5.5,
     guarda o resumo (não paga de novo) e trata recusa e PDF grande demais.

Rode com:  python testes/testar_recebidos.py
"""
import json
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import edocs_http  # noqa: E402
import ia  # noqa: E402

HAR = RAIZ / "capturas" / "vistar.har"
F = edocs_http.F


class _Resposta:
    def __init__(self, texto: str = "", conteudo: bytes | None = None, cabecalhos=None,
                 url: str = "https://x/f/n/caixaentradacon"):
        self.text = texto
        self.content = conteudo if conteudo is not None else texto.encode("utf-8")
        self.url = url
        self.status_code = 200
        self.headers = cabecalhos or {}

    def raise_for_status(self) -> None:
        pass


def caixa_real() -> str:
    for e in json.loads(HAR.read_text(encoding="utf-8"))["log"]["entries"]:
        texto = e["response"]["content"].get("text") or ""
        if "caixaentradacon" in e["request"]["url"] and "btVistarDocumento" in texto:
            return texto
    raise RuntimeError("não achei a Caixa de Entrada no HAR")


class _Site:
    """abrir() devolve as caixas em sequência; os cliques e os posts ficam registrados."""

    def __init__(self, caixas: list[str], pdf: bytes = b"%PDF-1.4 teste"):
        self.caixas, self.cliques, self.posts, self.pdf = caixas, [], [], pdf

    def __enter__(self):
        self.guardado = edocs_http.abrir
        edocs_http.abrir = self._abrir
        return self

    def __exit__(self, *erro):
        edocs_http.abrir = self.guardado

    def _abrir(self, sessao, caminho):
        site = self
        html = self.caixas.pop(0) if len(self.caixas) > 1 else self.caixas[0]
        t = edocs_http.Tela(None, _Resposta(html))

        def enviar(botao, valores=None, parcial=False):
            site.cliques.append((botao.split(":", 2)[-1], dict(valores or {}), parcial))
            return t
        t.enviar = enviar
        return t

    def post(self, url, data=None, headers=None, timeout=None):
        self.posts.append(data)
        return _Resposta(conteudo=self.pdf,
                         cabecalhos={"content-disposition": 'attachment;filename="OFICIO_1.pdf"'})


def teste_1_lista() -> None:
    with _Site([caixa_real()]):
        lista = edocs_http.recebidos(None)
    assert [(r["doc"], r["ciente"], r["vistavel"]) for r in lista] == [
        ("1554416", True, False), ("1411629", True, False)], lista
    assert lista[0]["papeis"] == ["Encaminhamento"] and lista[0]["titulo"] == "9804/2026/PAJ/PGM"
    print("  [1] lista os recebidos e o que dá para fazer ........... OK")


def teste_2_baixar() -> None:
    with _Site([caixa_real()]) as site:
        nome, pdf = edocs_http.baixar_recebido(site, "1411629")
    dados = site.posts[0]
    assert dados["source"] == F + "plcLogicaItens:1:btVisualizar" and dados[F + "nrLinha"] == "1"
    assert dados[F + "idImprimeHistoricoSituacoesDoc"] == "N"
    assert nome == "OFICIO_1.pdf" and pdf.startswith(b"%PDF")
    with _Site([caixa_real()], pdf=b"<html>erro</html>") as site:
        try:
            edocs_http.baixar_recebido(site, "1411629")
            raise AssertionError("aceitou algo que não é PDF")
        except edocs_http.ErroEdocs:
            pass
    print("  [2] baixa o PDF pela linha certa (e só aceita PDF) ..... OK")


def teste_3_ciente() -> None:
    real = caixa_real()
    sem_ciente = real.replace("plcLogicaItens:0:btLer", "plcLogicaItens:0:btOutro")  # o Ciente sumiu
    with _Site([real, sem_ciente]) as site:
        edocs_http.dar_ciente(None, "1554416")
    assert site.cliques == [("plcLogicaItens:0:btLer", {F + "nrLinha": "0"}, True)], site.cliques
    with _Site([real, real]):
        try:
            edocs_http.dar_ciente(None, "1554416")
            raise AssertionError("confirmou sem o Ciente sumir")
        except edocs_http.ErroEdocs:
            pass
    print("  [3] ciente na linha certa, confirmado pelo site ......... OK")


class _ClienteFalso:
    def __init__(self, stop="end_turn"):
        self.pedidos, self.stop = [], stop
        self.messages = SimpleNamespace(create=self._criar)
        self.beta = SimpleNamespace(messages=SimpleNamespace(create=self._criar))

    def _criar(self, **pedido):
        self.pedidos.append(pedido)
        return SimpleNamespace(stop_reason=self.stop,
                               content=[SimpleNamespace(type="text", text="Ofício pede informações.")])


def teste_4_resumo() -> None:
    pasta = Path(tempfile.mkdtemp())
    guardado = (ia.RESUMOS, ia._cliente, ia.modelo_escolhido, ia.provedor, ia.chave)
    ia.RESUMOS = pasta / "resumos.json"
    ia.provedor = lambda nome=None: {"nome": ia.ANTHROPIC, "tipo": "anthropic", "url": ""}
    ia.chave = lambda nome=None: "chave"
    cliente = _ClienteFalso()
    ia._cliente = lambda valor=None: cliente
    ia.modelo_escolhido = lambda: "claude-opus-5-5"
    try:
        assert ia.resumir_pdf("1", b"%PDF-1.4 x", "Ofício 1") == "Ofício pede informações."
        pedido = cliente.pedidos[0]
        bloco = pedido["messages"][0]["content"][0]
        assert bloco["type"] == "document" and bloco["source"]["media_type"] == "application/pdf"
        assert pedido["fallbacks"] == "default" and pedido["betas"] == ["server-side-fallback-2026-07-01"]
        ia.resumir_pdf("1", b"%PDF-1.4 x")  # já resumido: não chama a IA de novo
        assert len(cliente.pedidos) == 1
        assert "grande demais" in ia.resumir_pdf("2", b"%PDF" + b"0" * (ia.LIMITE_BYTES + 1))
        cliente.stop = "refusal"
        assert "não quis" in ia.resumir_pdf("3", b"%PDF-1.4 y")
        ia.modelo_escolhido = lambda: "claude-haiku-5-5"  # sem plano B: chamada normal
        cliente.stop = "end_turn"
        ia.resumir_pdf("4", b"%PDF-1.4 z")
        assert "fallbacks" not in cliente.pedidos[-1]
    finally:
        ia.RESUMOS, ia._cliente, ia.modelo_escolhido, ia.provedor, ia.chave = guardado
    print("  [4] resumo: PDF inteiro, plano B, guardado, recusa ...... OK")


def teste_5_outro_provedor() -> None:
    """Provedor no formato da OpenAI (DeepSeek...): vai o texto do PDF; PDF só imagem não."""
    import httpx

    pasta = Path(tempfile.mkdtemp())
    guardado = (ia.RESUMOS, ia.provedor, ia.chave, ia.modelo_escolhido, ia.texto_do_pdf, httpx.post)
    ia.RESUMOS = pasta / "resumos.json"
    ia.provedor = lambda nome=None: {"nome": "DeepSeek", "tipo": "openai", "url": "https://api.deepseek.com/"}
    ia.chave = lambda nome=None: "sk-1"
    ia.modelo_escolhido = lambda nome=None: "deepseek-chat"
    pedidos = []

    def post(url, json, headers, timeout):
        pedidos.append((url, json, headers))
        return SimpleNamespace(status_code=200, json=lambda: {
            "choices": [{"message": {"content": "Pede informações até sexta."}}]})

    httpx.post = post
    try:
        ia.texto_do_pdf = lambda pdf: "OFÍCIO 12 - pede informações"
        assert ia.resumir_pdf("10", b"%PDF") == "Pede informações até sexta."
        url, pedido, cabecalhos = pedidos[0]
        assert url == "https://api.deepseek.com/chat/completions" and cabecalhos == {"Authorization": "Bearer sk-1"}
        assert pedido["model"] == "deepseek-chat" and "OFÍCIO 12" in pedido["messages"][1]["content"]
        assert ia.resumo_guardado("10") == "Pede informações até sexta."
        ia.texto_do_pdf = lambda pdf: ""
        assert ia.resumir_pdf("11", b"%PDF").startswith(ia.SEM_TEXTO) and len(pedidos) == 1
        assert ia.resumo_guardado("11") is None, "o aviso de PDF imagem não fica guardado"
    finally:
        ia.RESUMOS, ia.provedor, ia.chave, ia.modelo_escolhido, ia.texto_do_pdf, httpx.post = guardado
    print("  [5] outro provedor: texto do PDF, PDF imagem ............ OK")


def main() -> None:
    print("Testes da página Recebidos (offline)\n")
    for teste in (teste_1_lista, teste_2_baixar, teste_3_ciente, teste_4_resumo,
                  teste_5_outro_provedor):
        teste()
    print("\n5 testes passaram.")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
