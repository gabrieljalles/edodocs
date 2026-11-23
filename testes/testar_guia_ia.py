"""O guia da IA (assets/ia/guia_processos.md) acompanha o programa, e o pedido em JSON funciona
com qualquer provedor.

Mudou como o Edodocs monta os atos (tipo de regra, operador, campo automático, chave do
processo.json)? Este teste falha até o guia ser atualizado.

    python testes/testar_guia_ia.py
"""
import json
import re
import shutil
import sys
import tempfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import atos  # noqa: E402
import ia  # noqa: E402
import main  # noqa: E402
import processos  # noqa: E402

GUIA = (RAIZ / "assets" / "ia" / "guia_processos.md").read_text(encoding="utf-8")
passou = 0


def ok(texto: str) -> None:
    global passou
    passou += 1
    print(f"  [{passou}] {texto} " + "." * max(2, 55 - len(texto)) + " OK")


def blocos_json(texto: str) -> list[str]:
    # também os recuados (dentro de listas)
    return [re.sub(r"(?m)^" + m.group(1), "", m.group(2))
            for m in re.finditer(r"(?ms)^( *)```json\n(.*?)\n *```", texto)]


# 1. Tudo o que o motor sabe fazer está no guia
faltam = [f"tipo {t}" for t in processos.TIPOS if f"**{t}**" not in GUIA]
faltam += [f"operador {o}" for o in processos.OPERADORES if f"`{o}`" not in GUIA]
faltam += [f"campo {r}" for r, _ in atos.AUTOMATICOS.values() if f"{{{{{r}}}}}" not in GUIA]
faltam += [f"formato {f}" for f in ("maiusculas", "minusculas", "titulo") if f'"{f}"' not in GUIA]
CHAVES_DO_PLANO = {"modelo", "planilha", "outros_modelos", "aba",  # o Edodocs preenche
                   "pessoas", "caixas", "colunas_tabela", "confirmar_envio", "consulta"}  # a pessoa escolhe depois
faltam += [f"chave {c}" for c in processos.PADRAO if c not in CHAVES_DO_PLANO and f"`{c}`" not in GUIA]
faltam += [f"automático {c}" for c in ("REPETIR", "FIM REPETIR") if f"{{{{{c}}}}}" not in GUIA]
faltam += [f"especial {c}" for c in ("{número do ato}", "{ano do ato}") if c not in GUIA]
assert not faltam, "Atualize assets/ia/guia_processos.md: falta " + ", ".join(faltam)
ok("o guia cita todos os tipos, operadores e campos")

# 2. Os exemplos de regra do guia são JSON válido e dão o texto que o guia promete
for bloco in blocos_json(GUIA):
    json.loads(bloco if bloco.lstrip().startswith("{") else "{" + bloco + "}")
plano = json.loads(blocos_json(GUIA)[-1])
ok("todo bloco JSON do guia é válido")

# 3. O exemplo completo (Estabilidade) funciona no motor de verdade
config = dict(plano["processo"], outros_modelos=[])
proc = processos.Processo(Path(tempfile.gettempdir()), config)
pessoa = {atos.normalizar(k): v for k, v in plano["exemplos"][0]["pessoas"][0].items()}
valores = proc.valores(pessoa)
assert valores["DECLARADO"] == "declarado" and valores["OSERVIDOR"] == "o servidor"
assert valores["DATAPOREXTENSO"] == "19 de julho de 2026", valores["DATAPOREXTENSO"]
assert valores["CONSIDERANDOPROCESSO"] == "Considerando a decisão proferida no Processo Nº 1234/2026,"
assert valores["CARGOMAIUSCULO"] == "AGENTE ADMINISTRATIVO"
sem_processo = proc.valores(dict(pessoa, PROCESSO=""))
assert sem_processo["CONSIDERANDOPROCESSO"] == ""
ok("o exemplo do guia dá o texto certo no motor")

# 4. No exemplo, todo campo das trocas tem regra, coluna ou é automático, e toda coluna citada existe
colunas = {atos.normalizar(c["nome"]) for c in plano["colunas"]}
ligacoes = {atos.normalizar(c) for c in plano["processo"]["ligacoes"]}
for modelo in plano["modelos"]:
    for troca in modelo["trocas"]:
        for campo in atos.CHAVES.findall(troca["para"]):
            n = atos.normalizar(campo)
            assert n in colunas | ligacoes | set(atos.AUTOMATICOS), campo
citadas = [r.get("coluna") for r in plano["processo"]["ligacoes"].values() if r.get("coluna")]
citadas += plano["processo"]["obrigatorios"] + [plano["processo"]["juntar"]["coluna"]]
assert all(atos.normalizar(c) in colunas for c in citadas), citadas
ok("o exemplo do guia é coerente (campos e colunas)")

# 5. JSON com ``` ou frase em volta
assert ia.ler_json('Aqui está:\n```json\n{"a": 1}\n```') == {"a": 1}
try:
    ia.ler_json("não sei")
    raise AssertionError("devia recusar")
except ia.ErroIA:
    pass
ok("lê o JSON mesmo com ``` em volta")

# 6. Instruções da equipe vão no fim do guia
TMP = Path(tempfile.mkdtemp(prefix="edodocs_guia_"))
try:
    original = main.BASE
    main.BASE = TMP
    (TMP / "configs").mkdir()
    assert "Instruções da equipe" not in ia.guia_processos()
    ia.arquivo_instrucoes().write_text("Sempre escreva Lei Orgânica Municipal.", encoding="utf-8")
    assert ia.guia_processos().endswith("Sempre escreva Lei Orgânica Municipal.")
    main.BASE = original
    ok("instruções da equipe no fim do guia")

    # 7. Provedor no formato da OpenAI (ChatGPT, Gemini, DeepSeek...): tenta sem o modo JSON se recusar
    import httpx

    class Resposta:
        def __init__(self, status, corpo):
            self.status_code, self._corpo, self.text = status, corpo, json.dumps(corpo)

        def json(self):
            return self._corpo

    pedidos = []

    def post(url, json=None, headers=None, timeout=None):
        pedidos.append(dict(json))
        if "response_format" in json:
            return Resposta(400, {"error": {"message": "response_format not supported"}})
        return Resposta(200, {"choices": [{"finish_reason": "stop",
                                           "message": {"content": '```json\n{"nome": "Teste"}\n```'}}]})

    ia.provedor = lambda nome=None: {"nome": "DeepSeek", "tipo": "openai", "url": "https://api.deepseek.com/",
                                     "modelos": ["deepseek-chat"], "modelo": "deepseek-chat"}
    ia.chave = lambda nome=None: "sk-teste"
    ia.modelo_escolhido = lambda nome=None: "deepseek-chat"
    httpx.post = post
    assert ia.pedir_json("guia", "pedido") == {"nome": "Teste"}
    assert len(pedidos) == 2 and pedidos[0]["messages"][0] == {"role": "system", "content": "guia"}
    ok("provedor compatível com OpenAI, com e sem modo JSON")

    # 8. Claude: streaming, guia em cache, fallback nos modelos atuais; resposta cortada avisa
    class Bloco:
        type, text = "text", '{"nome": "Claude"}'

    class Final:
        stop_reason, content = "end_turn", [Bloco()]

    chamadas = []

    class Fluxo:
        def __init__(self, **kw):
            chamadas.append(kw)

        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

        def get_final_message(self):
            return Final()

    class Mensagens:
        stream = staticmethod(lambda **kw: Fluxo(**kw))

    class Cliente:
        messages = Mensagens()
        beta = type("Beta", (), {"messages": Mensagens()})()

        def with_options(self, **kw):
            return self

    ia._cliente = lambda valor=None: Cliente()
    ia.provedor = lambda nome=None: {"nome": ia.ANTHROPIC, "tipo": "anthropic", "url": "",
                                     "modelos": ["claude-opus-5-5"], "modelo": "claude-opus-5-5"}
    ia.modelo_escolhido = lambda nome=None: "claude-opus-5-5"
    assert ia.pedir_json("guia", "pedido") == {"nome": "Claude"}
    assert chamadas[0]["system"][0]["cache_control"] == {"type": "ephemeral"}
    assert chamadas[0]["fallbacks"] == "default" and chamadas[0]["output_config"] == {"effort": "high"}
    Final.stop_reason = "max_tokens"
    try:
        ia.pedir_json("guia", "pedido")
        raise AssertionError("devia avisar")
    except ia.ErroIA as e:
        assert "grande demais" in str(e)
    ok("Claude: streaming, cache do guia e fallback")
finally:
    shutil.rmtree(TMP, ignore_errors=True)

print(f"\n{passou} testes passaram.")
