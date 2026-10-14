import csv
import os

import pytest

from app.llm import (
    LLMClient, LLMConfig, LLMError, ConfigError, ProjectContext, LANGUAGES, render, load_prompt, MODES,
)
from fake_anthropic import FakeAnthropic


@pytest.fixture(scope="module")
def pctx(ctrl):
    return ProjectContext(ctrl)


def make(tmp_path, responder=None, **overrides):
    cfg = LLMConfig(model_fast="fast-model", model_deep="deep-model", log_path=str(tmp_path / "log.csv"), **overrides)
    fake = FakeAnthropic(responder)
    return LLMClient(cfg, client=fake), fake


# --------------------------------------------------------------------------
# config
# --------------------------------------------------------------------------
def test_config_from_env_and_secrets(monkeypatch):
    for k in ("MODEL_FAST", "MODEL_DEEP", "ANTHROPIC_API_KEY", "APP_LANGUAGE", "LLM_EFFORT_DEEP", "MAX_CONTEXT_TOKENS"):
        monkeypatch.delenv(k, raising=False)
    with pytest.raises(ConfigError):
        LLMConfig.from_env(dotenv=False)
    monkeypatch.setenv("MODEL_FAST", "env-fast")
    monkeypatch.setenv("MAX_CONTEXT_TOKENS", "12345")
    cfg = LLMConfig.from_env(dotenv=False)
    assert cfg.model_fast == "env-fast" and cfg.model_deep == "env-fast"  # deep cai no fast
    assert cfg.max_context_tokens == 12345 and cfg.language == "pt-BR"
    cfg = LLMConfig.from_env(secrets={"MODEL_DEEP": "secret-deep", "APP_LANGUAGE": "en", "LLM_EFFORT_DEEP": "high", "ANTHROPIC_API_KEY": "k"}, dotenv=False)
    assert cfg.model_fast == "env-fast" and cfg.model_deep == "secret-deep"
    assert cfg.language == "en" and cfg.effort_deep == "high" and cfg.api_key == "k"
    assert cfg.model_for("explain") == "env-fast" and cfg.model_for("review") == "secret-deep"
    assert cfg.effort_for("document") == "high" and cfg.effort_for("chat") is None


def test_prompts_exist_and_render():
    for name in list(MODES) + ["system", "consolidate"]:
        assert load_prompt(name).strip()
    assert "{language}" in load_prompt("system")
    out = render("a {x} b {missing} {not_a_key}", x="1")
    assert out == "a 1 b {missing} {not_a_key}"
    # texto de rung com chaves não quebra
    assert render("{content}", content="MOV(A,B) {x}") == "MOV(A,B) {x}"


# --------------------------------------------------------------------------
# system block / caching
# --------------------------------------------------------------------------
def test_system_blocks_rules_then_cached_overview(pctx, tmp_path):
    llm, fake = make(tmp_path)
    blocks = llm.system_blocks(pctx)
    assert len(blocks) == 2
    assert "cache_control" not in blocks[0]
    assert blocks[1]["cache_control"] == {"type": "ephemeral"}
    assert "português do Brasil" in blocks[0]["text"]
    assert "# Projeto MINIMAL" in blocks[1]["text"]
    assert "validar com o engenheiro responsável" in blocks[0]["text"]
    en = llm.system_blocks(pctx, language="en")
    assert "English" in en[0]["text"] and en[1] == blocks[1]  # overview idêntico -> prefixo cacheável
    # nada de XML cru
    assert "<Rung" not in blocks[1]["text"] and "<![CDATA[" not in blocks[1]["text"]


def test_explain_single_routine(pctx, tmp_path):
    llm, fake = make(tmp_path)
    chunks = []
    r = llm.explain(pctx, "MainProgram", "MainRoutine", on_text=chunks.append)
    assert r.text == "resposta 1" and "".join(chunks) == "resposta 1"
    assert r.mode == "explain" and r.model == "fast-model" and r.calls == 1
    call = fake.calls[0]
    assert call["model"] == "fast-model"
    assert call["system"][1]["cache_control"] == {"type": "ephemeral"}
    assert "output_config" not in call
    user = call["messages"][0]["content"]
    assert "Alvo: MainProgram / MainRoutine" in user
    assert "Rung 0: XIC(Start)XIO(Stop)OTE(Motor);" in user
    assert "## Tags usadas nesta rotina" in user
    assert r.usage.cache_write_tokens == 500 and r.context_tokens > 0


def test_effort_and_deep_model_for_review(pctx, tmp_path):
    llm, fake = make(tmp_path, effort_fast="low", effort_deep="high")
    llm.explain(pctx, "MainProgram", "MainRoutine")
    assert fake.calls[0]["output_config"] == {"effort": "low"}
    r = llm.review(pctx, "MainProgram", "MainRoutine")
    assert fake.calls[1]["model"] == "deep-model" and fake.calls[1]["output_config"] == {"effort": "high"}
    user = fake.calls[1]["messages"][0]["content"]
    assert "jsr_missing_routine" in user and "output_multiple_writes" in user
    assert "Rung 6: JSR(Missing,0);" in user
    # findings de outra rotina não entram
    assert "Orphan" not in user.split("### Rotinas")[0]
    assert r.model == "deep-model"


def test_review_whole_project_fits_in_one_call(pctx, tmp_path):
    llm, fake = make(tmp_path)
    r = llm.review(pctx)
    assert r.calls == 1
    user = fake.calls[0]["messages"][0]["content"]
    assert "# MainProgram / MainRoutine [RLL]" in user and "# SafetyProg / Main [RLL]" in user
    assert "# MainProgram / Protected" not in user  # protegida fica fora
    assert "program_not_scheduled" in user


def test_map_reduce_when_over_budget(pctx, tmp_path):
    llm, fake = make(tmp_path, max_context_tokens=4500)  # regras+overview ~3.3k -> orçamento mínimo
    r = llm.review(pctx)
    assert r.calls >= 3  # partes + consolidação
    last = fake.calls[-1]["messages"][0]["content"]
    assert "resultados parciais do modo \"review\"" in last
    assert "<!-- parte 1/" in last
    assert len(r.partial_results) == r.calls - 1
    assert r.text == f"resposta {len(fake.calls)}"
    # parte 1 só recebe findings das suas rotinas
    first = fake.calls[0]["messages"][0]["content"]
    assert "— parte 1/" in first


def test_explain_program_concatenates_parts(pctx, tmp_path):
    llm, fake = make(tmp_path, max_context_tokens=3800)
    r = llm.explain(pctx, "MainProgram")
    assert r.calls > 1 and "---" in r.text
    assert all("parte" in c["messages"][0]["content"] for c in fake.calls)


def test_document_includes_tables(pctx, tmp_path):
    llm, fake = make(tmp_path)
    r = llm.document(pctx, program="MainProgram")
    user = fake.calls[0]["messages"][0]["content"]
    assert "T1 | TON | 0 (valor da tag) | MainProgram/MainRoutine rung 2" in user
    assert "Tank1.FillTimer | TON | 0 (valor da tag)" in user
    assert "T2 | TOF | 0 |" in user
    assert "Counter1 | CTU | 10 (valor da tag)" in user
    assert "## Lista de I/O" in user
    assert r.mode == "document" and fake.calls[0]["model"] == "deep-model"


def test_chat_uses_retrieval_and_history(pctx, tmp_path):
    llm, fake = make(tmp_path)
    history = [{"role": "user", "content": "oi"}, {"role": "assistant", "content": "olá"}]
    r = llm.chat(pctx, history, "O que liga a tag Motor?")
    msgs = fake.calls[0]["messages"]
    assert [m["role"] for m in msgs] == ["user", "assistant", "user"]
    assert msgs[0]["content"] == "oi"  # histórico sem rotinas injetadas
    assert "O que liga a tag Motor?" in msgs[2]["content"]
    assert "# MainProgram / MainRoutine [RLL]" in msgs[2]["content"]
    assert "MainProgram/MainRoutine" in r.partial_results
    assert fake.calls[0]["model"] == "fast-model"


def test_language_override_per_call(pctx, tmp_path):
    llm, fake = make(tmp_path)
    llm.explain(pctx, "MainProgram", "Called", language="es")
    assert "español" in fake.calls[0]["system"][0]["text"]


# --------------------------------------------------------------------------
# logging, refusal, truncation, errors
# --------------------------------------------------------------------------
def test_log_has_metrics_and_no_program_content(pctx, tmp_path):
    llm, fake = make(tmp_path)
    llm.explain(pctx, "MainProgram", "MainRoutine")
    llm.chat(pctx, [], "qual o preset de T2?")
    with open(llm.config.log_path, encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))
    assert [r["mode"] for r in rows] == ["explain", "chat"]
    assert rows[0]["model"] == "fast-model" and rows[0]["status"] == "end_turn"
    assert int(rows[1]["cache_read_tokens"]) == 500
    raw = open(llm.config.log_path, encoding="utf-8").read()
    for forbidden in ("XIC", "Motor", "MainRoutine", "preset de T2"):
        assert forbidden not in raw


def test_refusal_and_truncation(pctx, tmp_path):
    def responder(params, n):
        if n == 1:
            return {"text": "parcial", "stop_reason": "refusal", "explanation": "motivo"}
        return {"text": "cortado", "stop_reason": "max_tokens"}

    llm, fake = make(tmp_path, responder=responder)
    r = llm.explain(pctx, "MainProgram", "MainRoutine")
    assert r.refusal and "recusou" in r.text and "motivo" in r.text
    r = llm.explain(pctx, "MainProgram", "MainRoutine")
    assert r.truncated and "limite de tokens" in r.text
    assert r.calls == 3  # 1 chamada + 2 continuações (padrão), todas cortadas


def test_hybrid_routing_caps_and_deep_option(pctx, tmp_path, monkeypatch):
    llm, fake = make(tmp_path, effort_fast="low", effort_deep="high")
    llm.explain(pctx, "MainProgram", "MainRoutine")
    llm.explain(pctx, "MainProgram", "MainRoutine", deep=True)
    llm.chat(pctx, [], "o que liga o motor?", deep=True)
    llm.aoi(pctx, "MyAoi")
    models = [(c["model"], c["max_tokens"], c.get("output_config", {}).get("effort")) for c in fake.calls]
    assert models == [("fast-model", 6_000, "low"), ("deep-model", 6_000, "high"), ("deep-model", 4_000, "high"), ("deep-model", 12_000, "high")]
    cfg = llm.config
    assert cfg.has_deep_option and cfg.model_for("review") == "deep-model" and cfg.model_for("explain") == "fast-model" and cfg.model_for("explain", deep=True) == "deep-model"
    assert cfg.output_cap("review") == 16_000 and cfg.price_for("claude-fable-5-1") == (10.0, 50.0, 0.25) and cfg.price_for("fast-model") == (10.0, 50.0, 1.0)
    # configuração por ambiente
    for k in ("MODEL_FAST", "MODEL_DEEP", "OUTPUT_TOKENS_AOI", "CREDIT_USD", "CREDIT_MARKUP", "MODEL_PRICES", "MAX_OUTPUT_TOKENS"):
        monkeypatch.delenv(k, raising=False)
    monkeypatch.setenv("MODEL_FAST", "a")
    monkeypatch.setenv("MODEL_DEEP", "a")
    monkeypatch.setenv("OUTPUT_TOKENS_AOI", "3000")
    monkeypatch.setenv("MAX_OUTPUT_TOKENS", "5000")
    monkeypatch.setenv("CREDIT_USD", "0.25")
    monkeypatch.setenv("CREDIT_MARKUP", "2")
    monkeypatch.setenv("MODEL_PRICES", "a:1:2:0.1,ruim")
    env_cfg = LLMConfig.from_env(dotenv=False)
    assert not env_cfg.has_deep_option and env_cfg.output_cap("aoi") == 3000 and env_cfg.output_cap("review") == 5000  # teto global prevalece
    assert env_cfg.credit_usd == 0.25 and env_cfg.markup == 2 and env_cfg.price_for("a") == (1.0, 2.0, 0.1)


def test_cost_estimates(pctx, tmp_path):
    llm, fake = make(tmp_path, prices={"fast-model": (2.0, 10.0, 0.2), "deep-model": (10.0, 50.0, 0.25)})
    e = llm.estimate("explain", pctx, program="MainProgram", routine="MainRoutine")
    assert e.model == "fast-model" and e.calls == 1 and e.output_tokens == 2500 and e.input_tokens > llm.system_tokens(pctx)
    assert abs(e.usd - (e.input_tokens * 2 + 2500 * 10) / 1e6) < 1e-4 and e.credits == llm.credits_for_usd(e.usd) >= 1  # usd arredondado a 4 casas
    deep = llm.estimate("explain", pctx, program="MainProgram", routine="MainRoutine", deep=True)
    assert deep.model == "deep-model" and deep.usd > e.usd and deep.credits >= e.credits
    assert "deep-model" in deep.summary() and "US$" in deep.summary()
    # mais caro quanto maior o escopo; map-reduce aparece nas chamadas
    prog = llm.estimate("review", pctx, program="MainProgram")
    proj = llm.estimate("review", pctx)
    assert proj.input_tokens >= prog.input_tokens and proj.model == "deep-model"
    small, _ = make(tmp_path, max_context_tokens=2_500, prices={"deep-model": (10.0, 50.0, 0.25)})  # orçamento mínimo (500 tokens) força map-reduce
    assert small.estimate("review", pctx).calls > 1
    assert "entrada, ~" in e.summary()  # separador de milhar não corrompe o texto
    assert llm.estimate("document", pctx, program="MainProgram").output_tokens == 10_000
    assert llm.estimate("aoi", pctx, aoi="MyAoi").model == "deep-model"
    chat = llm.estimate("chat", pctx, question="o que liga o Motor?", history=[{"role": "user", "content": "oi"}])
    assert chat.model == "fast-model" and chat.calls == 1 and chat.output_tokens == 3500
    assert llm.credits_for_usd(0.0) == 1 and llm.credits_for_usd(0.10) == 4 and llm.credits_for_usd(0.1001) == 5  # 1 crédito = US$ 0,10 × margem 4
    assert fake.calls == []  # estimar nunca chama o modelo
    with pytest.raises(LLMError):
        llm.estimate("explain", pctx, program="Nope")
    with pytest.raises(LLMError):
        llm.estimate("aoi", pctx, aoi="Nope")


def test_truncated_answer_is_continued(pctx, tmp_path):
    def responder(params, n):
        if n == 1:
            return {"text": "## Resumo\nprimeira parte", "stop_reason": "max_tokens", "usage": {"output_tokens": 10}}
        return {"text": " e o final.", "stop_reason": "end_turn", "usage": {"output_tokens": 5}}

    llm, fake = make(tmp_path, responder=responder)
    chunks = []
    r = llm.explain(pctx, "MainProgram", "MainRoutine", on_text=chunks.append)
    assert r.text == "## Resumo\nprimeira parte e o final." and "".join(chunks) == r.text
    assert not r.truncated and r.calls == 2 and r.usage.output_tokens == 15
    msgs = fake.calls[1]["messages"]
    assert [m["role"] for m in msgs] == ["user", "assistant", "user"]
    assert msgs[1]["content"] == "## Resumo\nprimeira parte" and "Continue exatamente de onde parou" in msgs[2]["content"]
    assert fake.calls[0]["max_tokens"] == 6_000  # teto de saída do modo explain (MAX_OUTPUT_TOKENS é o teto global)
    # continuações desligadas: devolve o corte com o aviso
    llm, fake = make(tmp_path, responder=responder, max_continuations=0)
    r = llm.explain(pctx, "MainProgram", "MainRoutine")
    assert r.truncated and r.calls == 1 and "limite de tokens" in r.text


def test_api_errors_are_mapped(pctx, tmp_path):
    import anthropic
    import httpx2 as httpx

    def mk(cls, status, msg):
        req = httpx.Request("POST", "https://api.anthropic.com/v1/messages")
        resp = httpx.Response(status, request=req, headers={"retry-after": "7"})
        return cls(msg, response=resp, body={"error": {"message": msg}})

    cases = [
        (mk(anthropic.AuthenticationError, 401, "bad key"), "auth", "ANTHROPIC_API_KEY"),
        (mk(anthropic.NotFoundError, 404, "model not found"), "model", "MODEL_FAST"),
        (mk(anthropic.RateLimitError, 429, "slow down"), "rate_limit", "7s"),
        (mk(anthropic.BadRequestError, 400, "prompt is too long: 250000 tokens"), "context", "MAX_CONTEXT_TOKENS"),
        (mk(anthropic.InternalServerError, 500, "boom"), "api", "temporário"),
        (anthropic.APIConnectionError(request=httpx.Request("POST", "https://x")), "connection", "conexão"),
    ]
    for exc, kind, needle in cases:
        llm, fake = make(tmp_path, responder=lambda p, n, e=exc: e)
        with pytest.raises(LLMError) as ei:
            llm.explain(pctx, "MainProgram", "MainRoutine")
        assert ei.value.kind == kind, (kind, ei.value.message)
        assert needle in ei.value.message
    with open(llm.config.log_path, encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))
    assert rows[-1]["status"] == "APIConnectionError"


def test_scope_errors(pctx, tmp_path):
    llm, fake = make(tmp_path)
    with pytest.raises(LLMError):
        llm.explain(pctx, "NaoExiste")
    with pytest.raises(KeyError):
        llm.explain(pctx, "MainProgram", "NaoExiste")
    with pytest.raises(LLMError):
        llm.review(pctx, routine="MainRoutine")
    assert fake.calls == []


def test_split_text_respects_rung_boundaries():
    text = "# P / R [RLL]\n" + "".join(f"Rung {i}: XIC(A{i})OTE(B{i});\n" for i in range(200))
    parts = LLMClient._split_text(text, budget=400)
    assert len(parts) > 1
    for p in parts[1:]:
        assert p.startswith("# P / R [RLL] (continuação)\n")
        assert p.splitlines()[1].startswith("Rung ")
    assert "".join(l for p in parts for l in p.splitlines(keepends=True) if l.startswith("Rung ")) == "".join(f"Rung {i}: XIC(A{i})OTE(B{i});\n" for i in range(200))


def test_context_tables(pctx):
    io = pctx.io_table()
    assert "nenhuma referência direta a I/O" in io
    assert pctx.findings_text([]) == "(nenhum finding do parser para este escopo)"
    f = pctx.findings_for("MainProgram", "MainRoutine", min_severity="warning")
    assert f and all(x.severity != "info" for x in f)
    assert all((x.routine or "").lower() == "mainroutine" for x in f)
    assert pctx.is_safety("SafetyProg") and not pctx.is_safety("MainProgram")
