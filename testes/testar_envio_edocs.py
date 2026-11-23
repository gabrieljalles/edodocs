"""Testes offline do fluxo de envio do Edocs (não acessa o site).

Usam a gravação real de uma sessão (capturas/edocs_lancamento.har) para conferir:
  1. abrir_documento usa chPlc=<doc> (e não param=<doc>);
  2. a confirmação de envio reconhece a mensagem de sucesso do Edocs;
  3. enviar_documento só reporta "enviado" quando o site confirma;
     se o documento continuar em montagem, levanta erro.

Rode com:  python testes/testar_envio_edocs.py
"""
import json
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import edocs_http  # noqa: E402

HAR = RAIZ / "capturas" / "edocs_lancamento.har"

FORM = "corpo:formulario"
FORM_VAZIO = (
    '<html><body><form id="corpo:formulario" action="instrumentosjuridicosman">'
    '<input type="hidden" name="javax.faces.ViewState" value="x">'
    "</form></body></html>"
)


class _RespostaFalsa:
    """Imita o Response do curl_cffi para montar uma Tela sem rede."""

    def __init__(self, texto: str, url: str = "https://x/f/n/instrumentosjuridicosman",
                 status_code: int = 200):
        self.text = texto
        self.content = texto.encode("utf-8")
        self.url = url
        self.status_code = status_code

    def raise_for_status(self) -> None:
        pass


def tela_de_html(texto: str) -> edocs_http.Tela:
    return edocs_http.Tela(None, _RespostaFalsa(texto))


def texto_resposta_envio_real() -> str:
    """HTML real que o Edocs devolve DEPOIS de enviar um documento (do HAR)."""
    dados = json.loads(HAR.read_text(encoding="utf-8"))
    for entrada in dados["log"]["entries"]:
        corpo = entrada["request"].get("postData", {}).get("text", "")
        if "botaoEnviarParaVistos" in corpo:
            return entrada["response"]["content"]["text"]
    raise RuntimeError("não achei o envio (botaoEnviarParaVistos) no HAR")


class TelaComResposta(edocs_http.Tela):
    """Tela de documento cujo 'enviar' devolve uma resposta fixa."""

    def __init__(self, html_resposta: str):
        super().__init__(None, _RespostaFalsa(FORM_VAZIO))
        self._html_resposta = html_resposta

    def enviar(self, botao, valores=None, parcial=False):
        return tela_de_html(self._html_resposta)


# --- testes ------------------------------------------------------------------

def teste_1_abre_documento_com_chplc() -> None:
    chamadas = []

    def abrir_falso(sessao, caminho):
        chamadas.append(caminho)
        return tela_de_html(FORM_VAZIO)

    edocs_http.abrir, guardado = abrir_falso, edocs_http.abrir
    try:
        edocs_http.abrir_documento(object(), "1556634")
    finally:
        edocs_http.abrir = guardado

    assert chamadas == ["instrumentosjuridicosman?chPlc=1556634"], chamadas
    print("  [1] abrir_documento abre por chPlc=<doc> ................ OK")


def teste_2_confirma_envio_real() -> None:
    tela = tela_de_html(texto_resposta_envio_real())
    assert edocs_http._confirmou_envio(tela) is True
    print("  [2] resposta real de envio é reconhecida como enviada . OK")


def teste_3_montagem_nao_confirma() -> None:
    # Página sem a mensagem de sucesso = documento ficou em montagem
    tela = tela_de_html(FORM_VAZIO)
    assert edocs_http._confirmou_envio(tela) is False
    # Mensagem de sucesso, mas resposta HTTP diferente de 200: também não conta
    estranha = edocs_http.Tela(None, _RespostaFalsa(texto_resposta_envio_real(), status_code=202))
    assert edocs_http._confirmou_envio(estranha) is False
    print("  [3] sem confirmação (ou HTTP != 200) NÃO é enviado ... OK")


def teste_4_enviar_documento_sucesso() -> None:
    resposta = texto_resposta_envio_real()
    guardado = edocs_http.abrir_documento
    edocs_http.abrir_documento = lambda sessao, doc, *_: TelaComResposta(resposta)
    try:
        avisos = []
        edocs_http.enviar_documento(object(), "1556634", avisar=avisos.append)
    finally:
        edocs_http.abrir_documento = guardado
    assert avisos == ["Documento 1556634 enviado"], avisos
    print("  [4] enviar_documento reporta enviado no caso de sucesso . OK")


def teste_5_enviar_documento_montagem_leva_erro() -> None:
    guardado = edocs_http.abrir_documento
    edocs_http.abrir_documento = lambda sessao, doc, *_: TelaComResposta(FORM_VAZIO)
    try:
        try:
            edocs_http.enviar_documento(object(), "1556634")
        except edocs_http.ErroEdocs as e:
            assert "montagem" in str(e), e
        else:
            raise AssertionError("deveria ter levantado ErroEdocs (documento em montagem)")
    finally:
        edocs_http.abrir_documento = guardado
    print("  [5] enviar_documento levanta erro se ficar em montagem . OK")


def main() -> None:
    print("Testes do envio de atos no Edocs (offline)\n")
    testes = [
        teste_1_abre_documento_com_chplc,
        teste_2_confirma_envio_real,
        teste_3_montagem_nao_confirma,
        teste_4_enviar_documento_sucesso,
        teste_5_enviar_documento_montagem_leva_erro,
    ]
    for teste in testes:
        teste()
    print(f"\n{len(testes)} testes passaram.")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
