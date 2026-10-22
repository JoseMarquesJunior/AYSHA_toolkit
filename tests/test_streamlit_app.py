"""Testes da interface com streamlit.testing.v1.AppTest (sem navegador, sem API).

O file_uploader não é suportado pelo AppTest, então o fluxo de upload é coberto
chamando as funções auxiliares diretamente (parse_upload, action_cost, result_markdown).
"""

import os

import pytest
from streamlit.testing.v1 import AppTest

from app.credits import CreditStore
from app.streamlit_app import action_cost, parse_upload, result_markdown, file_hash, UI
from conftest import MINIMAL, ROOT

APP = os.path.join(ROOT, "app", "streamlit_app.py")


@pytest.fixture
def env(tmp_path, monkeypatch):
    db = str(tmp_path / "credits.db")
    monkeypatch.setenv("CREDITS_DB", db)
    monkeypatch.setenv("ACCESS_CODES", "DEMO:Demo:5")
    monkeypatch.setenv("MODEL_FAST", "fake-fast")
    monkeypatch.setenv("MODEL_DEEP", "fake-deep")
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-test")
    monkeypatch.setenv("LLM_LOG_PATH", str(tmp_path / "log.csv"))
    monkeypatch.delenv("APP_LANGUAGE", raising=False)
    monkeypatch.delenv("APP_DEMO_FILE", raising=False)
    import streamlit as st

    st.cache_resource.clear()
    return db


def run_app():
    at = AppTest.from_file(APP, default_timeout=30)
    return at.run()


def test_login_required_and_invalid_code(env):
    at = run_app()
    assert not at.exception
    assert at.title[0].value == UI["pt-BR"]["title"]
    assert len(at.text_input) == 1  # só o campo de código
    assert at.file_uploader == [] if hasattr(at, "file_uploader") else True
    at.text_input[0].input("ERRADO")
    at.button[0].click().run()
    assert any(UI["pt-BR"]["code_invalid"] in e.value for e in at.error)
    assert "code" not in at.session_state


def test_login_screen_shows_demo_gif(env):
    from app.streamlit_app import DEMO_GIF

    assert DEMO_GIF.is_file() and DEMO_GIF.stat().st_size < 8_000_000  # leve o bastante para a página inicial
    at = run_app()
    assert any(UI["pt-BR"]["demo_gif_caption"] in c.value for c in at.caption)
    assert len(at.get("image")) == 1


def test_login_ok_shows_upload_sidebar_and_balance(env):
    at = run_app()
    at.text_input[0].input("DEMO")
    at.button[0].click().run()
    assert not at.exception
    assert at.session_state["code"] == "DEMO" and at.session_state["owner"] == "Demo"
    sidebar_md = " ".join(m.value for m in at.sidebar.markdown)
    assert "Demo" in sidebar_md and "**5** de 5" in sidebar_md
    assert any(UI["pt-BR"]["no_project"] in i.value for i in at.info)
    assert at.sidebar.selectbox[0].value == "pt-BR"
    # troca de idioma da interface
    at.sidebar.selectbox[0].set_value("en").run()
    assert at.title[0].value == UI["en"]["title"]
    # logout
    at.sidebar.button[0].click().run()
    assert "code" not in at.session_state and len(at.text_input) == 1


def test_seed_codes_from_env(env):
    store = CreditStore(env)
    assert store.get("DEMO") is None  # só é semeado quando o app roda
    run_app()
    assert CreditStore(env).get("DEMO").quota == 5


def test_llm_missing_config_is_visible_but_not_fatal(env, monkeypatch):
    monkeypatch.delenv("MODEL_FAST")
    monkeypatch.delenv("MODEL_DEEP")
    import streamlit as st

    st.cache_resource.clear()
    at = run_app()
    at.text_input[0].input("DEMO")
    at.button[0].click().run()
    assert not at.exception  # sem projeto carregado não há aviso ainda; a config é lida no upload


def test_full_app_with_demo_file(env, monkeypatch):
    """APP_DEMO_FILE carrega um L5X sem upload: renderiza as quatro áreas sem exceção."""
    monkeypatch.setenv("APP_DEMO_FILE", MINIMAL)
    at = run_app()
    at.text_input[0].input("DEMO")
    at.button[0].click().run()
    assert not at.exception
    assert at.session_state["project_hash"] and at.session_state["file_meta"]["name"] == "minimal.l5x"
    assert any("minimal.l5x" in s.value for s in at.success)
    sidebar_md = " ".join(m.value for m in at.sidebar.markdown)
    assert "MINIMAL" in sidebar_md and "1756-L85E" in sidebar_md
    assert [m.label for m in at.sidebar.metric][:2] == [UI["pt-BR"]["programs"], UI["pt-BR"]["routines"]]
    # área Programas (padrão): seletores com rótulos ricos e os três botões com custo
    labels = [s.label for s in at.selectbox]
    assert UI["pt-BR"]["program"] in labels and UI["pt-BR"]["routine"] in labels
    routine_box = next(s for s in at.selectbox if s.label == UI["pt-BR"]["routine"])
    assert routine_box.format_func("MainRoutine").startswith("MainRoutine · RLL · 12 rungs · main")
    buttons = [b.label for b in at.button]
    assert any(UI["pt-BR"]["explain"] in b and "crédito" in b for b in buttons)
    explain_btn = next(b for b in at.button if UI["pt-BR"]["explain"] in b.label)
    assert int(explain_btn.label.split("custo: ")[1].split(" ")[0]) >= 1  # custo estimado pelo contexto (modelo fake = preço máximo)
    assert any(UI["pt-BR"]["deep_toggle"].split(" {")[0] in tg.label for tg in at.toggle)  # opção de análise profunda
    assert any(UI["pt-BR"]["review"] in b for b in buttons) and any(UI["pt-BR"]["document"] in b for b in buttons)
    assert any(UI["pt-BR"]["results_empty"] in c.value for c in at.caption)
    # demais áreas, via navegação persistida na sessão
    monkeypatch.setenv("APP_DEMO_OTHER", os.path.join(ROOT, "tests", "data", "plant.l5x"))
    for nav, expected in (("aois", UI["pt-BR"]["aoi_select"]), ("diag", UI["pt-BR"]["findings_sev"]), ("compare", UI["pt-BR"]["library_title"]), ("chat", UI["pt-BR"]["chat_title"])):
        at.session_state["nav"] = nav
        at.run()
        assert not at.exception, nav
        texts = " ".join(e.label for e in at.selectbox) + " ".join(e.value for e in at.subheader) + " ".join(getattr(e, "label", "") for e in at.get("pills"))
        assert expected in texts, nav
    at.session_state["nav"] = "aois"
    at.run()
    aoi_box = next(s for s in at.selectbox if s.label == UI["pt-BR"]["aoi_select"])
    assert "MyAoi rev. 1.0 · 1 inst. · 2 chamadas" in aoi_box.format_func("MyAoi")
    assert "SecretAoi" in aoi_box.format_func("SecretAoi") and "protegida" in aoi_box.format_func("SecretAoi")
    assert any(UI["pt-BR"]["aoi_analyze"] in b.label for b in at.button)


def test_demo_button_loads_bundled_project(env):
    """Sem APP_DEMO_FILE, o botão 'Testar com programa de demonstração' abre o L5X sintético que acompanha o app."""
    from app.streamlit_app import DEMO_FILE

    assert DEMO_FILE.is_file()
    at = run_app()
    at.text_input[0].input("DEMO")
    at.button[0].click().run()
    assert "project_hash" not in at.session_state
    assert any(UI["pt-BR"]["demo_hint"] in c.value for c in at.caption)
    demo_btn = next(b for b in at.button if UI["pt-BR"]["demo_button"] in b.label)
    demo_btn.click().run()
    assert not at.exception
    assert at.session_state["use_demo"] is True
    assert at.session_state["file_meta"]["name"] == DEMO_FILE.name
    assert any(DEMO_FILE.name in s.value for s in at.success)
    assert any(UI["pt-BR"]["demo_loaded"] in c.value for c in at.caption)
    sidebar_md = " ".join(m.value for m in at.sidebar.markdown)
    assert "EB01_DEMO" in sidebar_md and "1756-L83E" in sidebar_md
    routine_box = next(s for s in at.selectbox if s.label == UI["pt-BR"]["routine"])
    assert any(o.startswith("Partida_Bombas · RLL · 10 rungs") for o in routine_box.options)
    assert any(UI["pt-BR"]["explain"] in b.label for b in at.button)


def test_helpers():
    assert action_cost("explain", "routine") == 1
    assert action_cost("review", "project") == 1
    assert action_cost("document", "program") == 1
    assert action_cost("document", "project") == 3
    # custo estimado: usa o LLMClient quando disponível, senão o custo fixo
    from app.llm import LLMClient, LLMConfig
    from app.streamlit_app import estimated_cost
    from fake_anthropic import FakeAnthropic

    ctx = parse_upload(open(MINIMAL, "rb").read(), "minimal.l5x")
    assert estimated_cost(None, ctx, "review", "project") == (1, None)
    llm = LLMClient(LLMConfig(model_fast="fast", model_deep="deep", log_path=None, prices={"fast": (2, 10, 0.2), "deep": (10, 50, 0.25)}), client=FakeAnthropic())
    credits, est = estimated_cost(llm, ctx, "explain", "routine", "MainProgram", "MainRoutine")
    assert est is not None and est.model == "fast" and credits == est.credits >= 1
    credits_deep, est_deep = estimated_cost(llm, ctx, "explain", "routine", "MainProgram", "MainRoutine", deep=True)
    assert est_deep.model == "deep" and credits_deep >= credits
    assert estimated_cost(llm, ctx, "explain", "routine", "Nope", "x") == (1, None)  # erro na estimativa -> custo fixo
    with open(MINIMAL, "rb") as fh:
        data = fh.read()
    ctx = parse_upload(data, "minimal.l5x")
    assert ctx.ctrl.name == "MINIMAL" and ctx.metrics.programs == 3
    assert len(file_hash(data)) == 64
    md = result_markdown("Explicação: MainProgram/MainRoutine", "corpo", "MINIMAL", "modelo-x", "pt-BR")
    assert md.startswith("# Explicação: MainProgram/MainRoutine\n") and "corpo" in md and "modelo-x" not in md  # o modelo não aparece para o usuário
    assert "{model}" not in UI["pt-BR"]["deep_toggle"] and "{inp}" not in UI["pt-BR"]["usage"] and "cost_detail" not in UI["pt-BR"]
    assert UI["pt-BR"]["disclaimer"] in md
    with pytest.raises(Exception):
        parse_upload(b"<html/>", "x.l5x")


def test_ui_dictionaries_complete():
    from app.llm import LANGUAGES

    keys = set(UI["pt-BR"])
    assert set(UI) == set(LANGUAGES) == {"pt-BR", "en", "es", "zh-CN", "hi"}
    for lang in UI:
        assert set(UI[lang]) == keys, lang
        for k, v in UI[lang].items():
            assert {p for p in __import__("re").findall(r"\{(\w+)", v)} == {p for p in __import__("re").findall(r"\{(\w+)", UI["pt-BR"][k])}, (lang, k)  # mesmos placeholders
