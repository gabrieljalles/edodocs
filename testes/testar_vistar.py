"""Testes offline do visto automático logo depois do envio (não acessa o site).

Usam a Caixa de Entrada real gravada em capturas/vistar.har:
  1. lê a Caixa de Entrada (documento, papel, nrLinha e botão Ciente de cada linha);
  2. vista só a linha do documento enviado: Ciente da linha, recarrega, Vistar com o nrLinha
     dela (mesmo que a ordem mude entre um e outro);
  3. documento que não aparece na Caixa de Entrada: não clica em nada e avisa;
  4. só dá como vistado se o documento sair da Caixa de Entrada.

Rode com:  python testes/testar_vistar.py
"""
import json
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import edocs_http  # noqa: E402

HAR = RAIZ / "capturas" / "vistar.har"
F = edocs_http.F


class _RespostaFalsa:
    def __init__(self, texto: str, url: str = "https://x/f/n/caixaentradacon"):
        self.text = texto
        self.content = texto.encode("utf-8")
        self.url = url
        self.status_code = 200

    def raise_for_status(self) -> None:
        pass


def caixa_real() -> str:
    for e in json.loads(HAR.read_text(encoding="utf-8"))["log"]["entries"]:
        texto = e["response"]["content"].get("text") or ""
        if "caixaentradacon" in e["request"]["url"] and "btVistarDocumento" in texto:
            return texto
    raise RuntimeError("não achei a Caixa de Entrada no HAR")


def com_documento(html: str, doc: str, linha_de: str = "1554416") -> str:
    """A Caixa de Entrada real com o documento 'doc' no lugar de uma das linhas."""
    return html.replace(f"param={linha_de}_", f"param={doc}_")


def tela(html: str) -> edocs_http.Tela:
    return edocs_http.Tela(None, _RespostaFalsa(html))


class _Site:
    """abrir() devolve as caixas em sequência; enviar() registra os cliques."""

    def __init__(self, caixas: list[str]):
        self.caixas, self.cliques, self.abertos = caixas, [], []

    def __enter__(self):
        self.guardado = edocs_http.abrir
        edocs_http.abrir = self._abrir
        return self

    def __exit__(self, *erro):
        edocs_http.abrir = self.guardado

    def _abrir(self, sessao, caminho):
        self.abertos.append(caminho)
        html = self.caixas.pop(0)
        site = self
        t = tela(html)

        def enviar(botao, valores=None, parcial=False):
            site.cliques.append((botao.split(":", 2)[-1], dict(valores or {}), parcial))
            return t
        t.enviar = enviar
        return t


def teste_1_le_caixa_de_entrada() -> None:
    docs = edocs_http.documentos_da_caixa(tela(caixa_real()))
    assert [(d["doc"], d["linha"], d["ciente"]) for d in docs] == [("1554416", 0, True),
                                                                    ("1411629", 1, True)], docs
    assert docs[0]["papel"] == "Encaminhamento" and docs[0]["titulo"] == "9804/2026/PAJ/PGM"
    print("  [1] lê a Caixa de Entrada real (linha, papel, Ciente) ... OK")


def teste_2_vista_a_linha_certa() -> None:
    real = caixa_real()
    antes = com_documento(real, "999", "1554416")            # o enviado está na linha 0
    depois_do_ciente = com_documento(real, "999", "1411629")  # e passou para a linha 1
    vistado = real                                           # depois do visto, saiu da caixa
    with _Site([antes, depois_do_ciente, vistado]) as site:
        edocs_http.vistar_documento(None, "999", espera=0)
    assert site.cliques == [
        ("plcLogicaItens:0:btLer", {F + "nrLinha": "0"}, True),
        ("btVistarDocumento", {F + "nrLinha": "1"}, False)], site.cliques
    print("  [2] Ciente e Vistar só na linha do documento enviado .. OK")


def teste_3_sem_documento_nao_clica() -> None:
    real = caixa_real()
    with _Site([real, real, real]) as site:
        try:
            edocs_http.vistar_documento(None, "999", espera=0)
            raise AssertionError("vistou sem achar o documento")
        except edocs_http.ErroEdocs:
            pass
    assert site.cliques == [] and len(site.abertos) == 3
    print("  [3] documento fora da Caixa de Entrada: nenhum clique .. OK")


def teste_4_confere_que_saiu() -> None:
    caixa = com_documento(caixa_real(), "999")
    with _Site([caixa, caixa, caixa]):
        try:
            edocs_http.vistar_documento(None, "999", espera=0)
            raise AssertionError("deu como vistado com o documento ainda na caixa")
        except edocs_http.ErroEdocs:
            pass
    print("  [4] só confirma o visto quando o documento sai da caixa  OK")


def main() -> None:
    print("Testes do visto automático (offline)\n")
    for teste in (teste_1_le_caixa_de_entrada, teste_2_vista_a_linha_certa,
                  teste_3_sem_documento_nao_clica, teste_4_confere_que_saiu):
        teste()
    print("\n4 testes passaram.")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
