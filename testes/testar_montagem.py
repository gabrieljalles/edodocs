"""Testes offline do cancelamento de documentos em montagem (não acessa o site).

Usam a Caixa de Saída real gravada em capturas/Exclusão de um ato em montagem.har:
  1. lê as linhas da Caixa de Saída (número interno, título, remetente, situação, Cancelar);
  2. documento já enviado NUNCA conta como em montagem, mesmo tendo o botão Cancelar;
     só conta o que não tem remetente/destinatário nem situação;
  3. cancelar_em_montagem recusa documento que não está em montagem (não envia nada);
  4. cancela com o motivo e só dá certo se o Edocs mostrar o documento como cancelado.

Rode com:  python testes/testar_montagem.py
"""
import json
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import edocs_http  # noqa: E402

HAR = RAIZ / "capturas" / "Exclusão de um ato em montagem.har"
F = edocs_http.F


class _RespostaFalsa:
    def __init__(self, texto: str, url: str = "https://x/f/n/caixasaidacon"):
        self.text = texto
        self.content = texto.encode("utf-8")
        self.url = url
        self.status_code = 200

    def raise_for_status(self) -> None:
        pass


def caixa_real() -> list[dict]:
    for e in json.loads(HAR.read_text(encoding="utf-8"))["log"]["entries"]:
        texto = e["response"]["content"].get("text") or ""
        if "caixasaidacon" in e["request"]["url"] and "botaoAcaoPesquisarPaginado" in texto:
            return edocs_http.documentos_da_caixa(edocs_http.Tela(None, _RespostaFalsa(texto)))
    raise RuntimeError("não achei a Caixa de Saída no HAR")


def teste_1_le_caixa() -> None:
    docs = caixa_real()
    assert len(docs) == 20, len(docs)
    primeiro = docs[0]
    assert primeiro["doc"] == "1557272" and primeiro["pode_cancelar"], primeiro
    assert primeiro["remetente"] == "Servidora De Exemplo"
    assert primeiro["situacao"] == "Doc. enviado para Assinatura"
    cancelado = next(d for d in docs if d["doc"] == "1557208")
    assert cancelado["titulo"] == "ATO ADMINISTRATIVO SMA Nº 1551/2026" and not cancelado["remetente"]
    print("  [1] lê a Caixa de Saída real (20 linhas) ................ OK")


def teste_2_so_montagem() -> None:
    docs = caixa_real()
    assert not any(edocs_http.em_montagem(d) for d in docs)  # nenhum em montagem nessa lista
    enviado = docs[0]  # tem Cancelar, mas já foi enviado
    assert not edocs_http.em_montagem(enviado)
    montagem = {"doc": "1", "remetente": "", "destinatario": "", "situacao": "", "pode_cancelar": True}
    assert edocs_http.em_montagem(montagem)
    assert edocs_http.em_montagem(montagem | {"situacao": "Em montagem"})
    assert not edocs_http.em_montagem(montagem | {"situacao": "Doc. cancelado"})
    assert not edocs_http.em_montagem(montagem | {"remetente": "Fulano"})
    assert not edocs_http.em_montagem(montagem | {"pode_cancelar": False})
    print("  [2] enviado nunca conta como em montagem ................ OK")


class _Simulado:
    """Caixa de Saída e janela de cancelamento falsas."""

    def __init__(self, caixas: list[list[dict]]):
        self.caixas, self.abertos, self.enviados = caixas, [], []

    def __enter__(self):
        self.guardado = edocs_http.caixa_de_saida, edocs_http.abrir, edocs_http.pesquisar_caixa
        # A pesquisa filtrada (que acha também os antigos) é que diz a situação de cada vez
        edocs_http.caixa_de_saida = lambda sessao: []
        edocs_http.pesquisar_caixa = lambda sessao, situacao: self.caixas.pop(0)
        edocs_http.abrir = self._abrir
        return self

    def __exit__(self, *erro):
        edocs_http.caixa_de_saida, edocs_http.abrir, edocs_http.pesquisar_caixa = self.guardado

    def _abrir(self, sessao, caminho):
        self.abertos.append(caminho)
        simulado = self

        class Janela:
            doc = None

            def enviar(self, botao, valores=None, parcial=False):
                simulado.enviados.append((botao, valores, parcial))
                return self

            def erro(self, o_que):
                return edocs_http.ErroEdocs(o_que)
        return Janela()


MONTAGEM = {"doc": "999", "titulo": "X", "remetente": "", "destinatario": "", "situacao": "",
            "pode_cancelar": True}


def teste_3_recusa_o_que_nao_esta_em_montagem() -> None:
    with _Simulado([caixa_real()]) as sim:
        try:
            edocs_http.cancelar_em_montagem(None, "1557272")
            raise AssertionError("cancelou um documento enviado")
        except edocs_http.ErroEdocs:
            pass
    assert sim.abertos == [] and sim.enviados == []
    print("  [3] recusa cancelar documento que não está em montagem .. OK")


def teste_4_cancela_e_confere() -> None:
    depois = [MONTAGEM | {"situacao": "Doc. cancelado", "pode_cancelar": False}]
    with _Simulado([[MONTAGEM], depois]) as sim:
        edocs_http.cancelar_em_montagem(None, "999")
    assert sim.abertos == ["cancelardocumentoman?chPlc=999&modoJanelaPlc=popup"]
    botao, valores, parcial = sim.enviados[0]
    assert botao == F + "botaoAcaoGravar" and parcial
    assert valores[F + "dsMotivoCancelamento"] == edocs_http.MOTIVO_CANCELAMENTO
    with _Simulado([[MONTAGEM], [MONTAGEM]]):  # o site não cancelou: tem de dar erro
        try:
            edocs_http.cancelar_em_montagem(None, "999")
            raise AssertionError("aceitou sem confirmação do Edocs")
        except edocs_http.ErroEdocs:
            pass
    print("  [4] cancela com motivo e confere no Edocs ............... OK")


def teste_5_todos_os_em_montagem() -> None:
    """Os em montagem antigos (fora dos mais recentes da caixa) também entram: enquanto houver
    um, o e-docs abre ele no lugar de um documento novo."""
    recente = MONTAGEM | {"doc": "1", "titulo": "PORTARIA 1"}
    velho = MONTAGEM | {"doc": "265543", "titulo": "538/2022 SMA/DAP/NCP"}
    enviado = MONTAGEM | {"doc": "7", "remetente": "Fulano", "situacao": "Doc. em visto"}
    guardado = edocs_http.caixa_de_saida, edocs_http.pesquisar_caixa
    edocs_http.caixa_de_saida = lambda sessao: [recente, enviado]
    edocs_http.pesquisar_caixa = lambda sessao, situacao: [recente, velho, enviado]
    try:
        assert [d["doc"] for d in edocs_http.documentos_em_montagem(None)] == ["1", "265543"]
    finally:
        edocs_http.caixa_de_saida, edocs_http.pesquisar_caixa = guardado
    print("  [5] acha os em montagem de qualquer data .................. OK")


def teste_6_nao_grava_por_cima() -> None:
    """A tela de documento novo veio preenchida (o e-docs abriu um em montagem): para antes
    de gravar qualquer coisa."""
    enviados = []

    class TelaNova:
        campos = {F + "nrNossoNumero": "TORNA S/ EFEITO NOMEIA VICTORIA", F + "dtDocumento": "21/08/2026"}

        def enviar(self, *args, **kwargs):
            enviados.append(args)
            return self

    guardado = edocs_http.abrir
    edocs_http.abrir = lambda sessao, caminho: TelaNova()
    try:
        edocs_http.criar_documento(None, "DECRETO S/Nº - EXONERAÇÃO CELIA G. TEIXEIRA", "<p>x</p>")
        raise AssertionError("gravou por cima do documento em montagem")
    except edocs_http.DocumentoEmMontagem as e:
        assert "TORNA S/ EFEITO" in str(e) and "Nada foi gravado" in str(e)
    finally:
        edocs_http.abrir = guardado
    assert enviados == [], "não pode clicar em nada"
    print("  [6] não grava por cima de um documento em montagem ....... OK")


def teste_7_um_remetente_so() -> None:
    """Só o 1º assinante é Remetente: o e-docs recusa dois ("É obrigatório cadastrar apenas
    um Remetente do Documento")."""
    gravados = []

    class Janela:
        def __init__(self):
            self.lista = []

        def linhas(self):
            return self.lista

        def opcoes(self, campo):
            return [("0", "1º"), ("1", "2º")]

        def enviar(self, botao, valores=None, parcial=False):
            gravados.append(valores)
            self.lista = self.lista + [valores[F + "nrCpf"]]
            return self

    janela = Janela()
    guardado = edocs_http.abrir, edocs_http.buscar_servidor
    edocs_http.abrir = lambda sessao, caminho: janela
    edocs_http.buscar_servidor = lambda sessao, cpf: {"oid": cpf, "cpf": cpf, "lookup": cpf, "nome": cpf}
    try:
        edocs_http.adicionar_assinantes(None, "1", ["11111111111", "22222222222"])
    finally:
        edocs_http.abrir, edocs_http.buscar_servidor = guardado
    assert [g[F + "idRemetente"] for g in gravados] == ["0", "1"], gravados  # Sim, Não
    print("  [7] só o 1º assinante é remetente ......................... OK")


def teste_8_demora_do_edocs() -> None:
    """O e-docs não respondeu ao Cancelar a tempo, mas cancelou: confere e segue."""
    from curl_cffi import requests

    class Lenta:
        def enviar(self, *args, **kwargs):
            raise requests.exceptions.Timeout("curl: (28) Operation timed out")

    depois = [MONTAGEM | {"situacao": "Doc. cancelado", "pode_cancelar": False}]
    guardado = edocs_http.abrir, edocs_http.pesquisar_caixa
    caixas = [[MONTAGEM], depois]
    edocs_http.abrir = lambda sessao, caminho: Lenta()
    edocs_http.pesquisar_caixa = lambda sessao, situacao: caixas.pop(0)
    try:
        edocs_http.cancelar_em_montagem(None, "999")
    finally:
        edocs_http.abrir, edocs_http.pesquisar_caixa = guardado
    # leitura que demora é tentada de novo
    vezes = []

    def ler():
        vezes.append(1)
        if len(vezes) < 2:
            raise requests.exceptions.Timeout("demorou")
        return "ok"
    guardado_sleep = edocs_http.time.sleep
    edocs_http.time.sleep = lambda s: None
    try:
        assert edocs_http._de_novo_se_demorar(ler) == "ok" and len(vezes) == 2
    finally:
        edocs_http.time.sleep = guardado_sleep
    print("  [8] e-docs demorando: confere e tenta de novo ............. OK")


def teste_9_rascunho_pela_metade_e_cancelado() -> None:
    """Erro depois de criar o documento (ex.: assinante recusado): o rascunho é cancelado,
    para não travar os próximos atos."""
    cancelados = []
    guardado = (edocs_http.criar_documento, edocs_http.abrir, edocs_http.adicionar_vistos,
                edocs_http.adicionar_assinantes, edocs_http.cancelar_em_montagem)

    class SemVistos:
        def linhas(self):
            return []

    def recusa(*args, **kwargs):
        raise edocs_http.ErroEdocs("Assinante não apareceu na lista")
    edocs_http.criar_documento = lambda *a, **k: "555"
    edocs_http.abrir = lambda sessao, caminho: SemVistos()
    edocs_http.adicionar_vistos = lambda *a, **k: None
    edocs_http.adicionar_assinantes = recusa
    edocs_http.cancelar_em_montagem = lambda sessao, doc, avisar=None: cancelados.append(doc)
    try:
        edocs_http.preparar_documento(None, "T", "<p/>", ["1"], ["2"])
        raise AssertionError("devia dar erro")
    except edocs_http.ErroEdocs as e:
        assert "rascunho 555 foi cancelado" in str(e), e
    finally:
        (edocs_http.criar_documento, edocs_http.abrir, edocs_http.adicionar_vistos,
         edocs_http.adicionar_assinantes, edocs_http.cancelar_em_montagem) = guardado
    assert cancelados == ["555"]
    print("  [9] rascunho pela metade é cancelado ...................... OK")


def main() -> None:
    print("Testes do cancelamento de documentos em montagem (offline)\n")
    for teste in (teste_1_le_caixa, teste_2_so_montagem, teste_3_recusa_o_que_nao_esta_em_montagem,
                  teste_4_cancela_e_confere, teste_5_todos_os_em_montagem,
                  teste_6_nao_grava_por_cima, teste_7_um_remetente_so, teste_8_demora_do_edocs,
                  teste_9_rascunho_pela_metade_e_cancelado):
        teste()
    print("\n9 testes passaram.")


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
