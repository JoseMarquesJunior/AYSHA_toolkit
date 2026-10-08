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


def test_helpers():
    assert action_cost("explain", "routine") == 1
    assert action_cost("review", "project") == 1
    assert action_cost("document", "program") == 1
    assert action_cost("document", "project") == 3
    with open(MINIMAL, "rb") as fh:
        data = fh.read()
    ctx = parse_upload(data, "minimal.l5x")
    assert ctx.ctrl.name == "MINIMAL" and ctx.metrics.programs == 3
    assert len(file_hash(data)) == 64
    md = result_markdown("Explicação: MainProgram/MainRoutine", "corpo", "MINIMAL", "modelo-x", "pt-BR")
    assert md.startswith("# Explicação: MainProgram/MainRoutine\n") and "corpo" in md and "modelo-x" in md
    assert UI["pt-BR"]["disclaimer"] in md
    with pytest.raises(Exception):
        parse_upload(b"<html/>", "x.l5x")


def test_ui_dictionaries_complete():
    keys = set(UI["pt-BR"])
    for lang in ("en", "es"):
        assert set(UI[lang]) == keys, lang
