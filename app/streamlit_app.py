"""Interface Streamlit do analisador de programas Rockwell (L5X).

    streamlit run app/streamlit_app.py

Fluxo: código de acesso -> upload do .L5X -> resumo + findings determinísticos ->
Explicar / Revisar / Documentar / Conversar (LLM) -> download em Markdown.
O arquivo enviado é gravado em um diretório temporário só durante o parsing e
apagado em seguida; o modelo só recebe o índice produzido pelo parser.
"""

from __future__ import annotations

import hashlib
import os
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import streamlit as st  # noqa: E402

from app.credits import CreditStore  # noqa: E402
from app.llm import LANGUAGES, ConfigError, LLMClient, LLMConfig, LLMError, ProjectContext  # noqa: E402
from l5x_core.model import SEVERITIES  # noqa: E402
from l5x_core.parser import L5XError, parse_file  # noqa: E402

# ----------------------------------------------------------------------------
# Textos da interface
# ----------------------------------------------------------------------------
UI: Dict[str, Dict[str, str]] = {
    "pt-BR": {
        "title": "Analisador de programas Rockwell (L5X)",
        "subtitle": "Explica, revisa e documenta programas Logix exportados em L5X. Somente leitura: nada é enviado ao controlador.",
        "code_label": "Código de acesso",
        "code_button": "Entrar",
        "code_invalid": "Código inválido.",
        "code_hint": "Sem um código válido nada funciona. Peça um código ao administrador.",
        "logout": "Sair",
        "welcome": "Olá, {owner}.",
        "balance": "Créditos: **{balance}** de {quota}",
        "language": "Idioma das respostas",
        "upload_label": "Envie o arquivo .L5X (File → Save As → L5X no Studio 5000)",
        "parse_error": "Arquivo inválido: {error}",
        "parsing": "Lendo e analisando o programa…",
        "summary": "Resumo do projeto",
        "controller": "Controlador",
        "revision": "Revisão",
        "programs": "Programas",
        "routines": "Rotinas",
        "rungs": "Rungs",
        "tags": "Tags",
        "findings": "Findings por severidade",
        "protected_warning": "{n} item(ns) com source protection: conteúdo não lido nem analisado.",
        "safety_warning": "Este projeto contém conteúdo de segurança (Class=\"Safety\"). Ele é somente leitura: o app explica e documenta, mas nunca sugere alterações nessa lógica.",
        "no_project": "Envie um arquivo .L5X para começar.",
        "program": "Programa",
        "routine": "Rotina",
        "scope": "Escopo da ação",
        "scope_routine": "Rotina selecionada",
        "scope_program": "Programa selecionado",
        "scope_project": "Projeto inteiro",
        "explain": "Explicar",
        "review": "Revisar riscos e qualidade",
        "document": "Gerar documentação",
        "cost": "custo: {n} crédito(s)",
        "insufficient": "Créditos insuficientes: esta ação custa {cost} e você tem {balance}.",
        "routine_protected": "A rotina selecionada tem source protection: não há conteúdo para explicar ou revisar.",
        "routine_not_analyzed": "Rotina {type}: o parser não analisa a lógica interna; a explicação usará só nome, descrição e tags referenciadas.",
        "deterministic": "Findings do parser para o escopo selecionado ({n})",
        "no_findings": "Nenhum finding para este escopo.",
        "results": "Resultados",
        "download": "Baixar Markdown",
        "chat_title": "Conversar sobre o programa",
        "chat_placeholder": "Pergunte sobre o programa (1 crédito por pergunta)",
        "chat_clear": "Limpar conversa",
        "chat_sources": "Rotinas consultadas: {sources}",
        "llm_missing": "Camada LLM não configurada: {error}. Upload, resumo e findings continuam funcionando.",
        "working": "Consultando o modelo…",
        "usage": "{model} · {calls} chamada(s) · entrada {inp} · saída {out} · cache {cache} · {ms} ms",
        "refusal": "O modelo recusou parte do pedido; veja a resposta.",
        "truncated": "A resposta foi cortada pelo limite de tokens de saída.",
        "mode_explain": "Explicação",
        "mode_review": "Revisão",
        "mode_document": "Documentação",
        "mode_chat": "Chat",
        "disclaimer": "Análise automática: validar com o engenheiro responsável antes de qualquer alteração.",
    },
    "en": {
        "title": "Rockwell program analyzer (L5X)",
        "subtitle": "Explains, reviews and documents Logix programs exported as L5X. Read-only: nothing is sent to the controller.",
        "code_label": "Access code",
        "code_button": "Enter",
        "code_invalid": "Invalid code.",
        "code_hint": "Nothing works without a valid code. Ask the administrator for one.",
        "logout": "Log out",
        "welcome": "Hello, {owner}.",
        "balance": "Credits: **{balance}** of {quota}",
        "language": "Answer language",
        "upload_label": "Upload the .L5X file (File → Save As → L5X in Studio 5000)",
        "parse_error": "Invalid file: {error}",
        "parsing": "Parsing and analyzing the program…",
        "summary": "Project summary",
        "controller": "Controller",
        "revision": "Revision",
        "programs": "Programs",
        "routines": "Routines",
        "rungs": "Rungs",
        "tags": "Tags",
        "findings": "Findings by severity",
        "protected_warning": "{n} item(s) with source protection: content not read nor analyzed.",
        "safety_warning": "This project contains safety content (Class=\"Safety\"). It is read-only: the app explains and documents it but never suggests changes.",
        "no_project": "Upload a .L5X file to start.",
        "program": "Program",
        "routine": "Routine",
        "scope": "Action scope",
        "scope_routine": "Selected routine",
        "scope_program": "Selected program",
        "scope_project": "Whole project",
        "explain": "Explain",
        "review": "Review risks and quality",
        "document": "Generate documentation",
        "cost": "cost: {n} credit(s)",
        "insufficient": "Not enough credits: this action costs {cost} and you have {balance}.",
        "routine_protected": "The selected routine is source-protected: there is no content to explain or review.",
        "routine_not_analyzed": "{type} routine: the parser does not analyze its internal logic; the explanation uses only name, description and referenced tags.",
        "deterministic": "Parser findings for the selected scope ({n})",
        "no_findings": "No findings for this scope.",
        "results": "Results",
        "download": "Download Markdown",
        "chat_title": "Chat about the program",
        "chat_placeholder": "Ask about the program (1 credit per question)",
        "chat_clear": "Clear conversation",
        "chat_sources": "Routines consulted: {sources}",
        "llm_missing": "LLM layer not configured: {error}. Upload, summary and findings still work.",
        "working": "Querying the model…",
        "usage": "{model} · {calls} call(s) · input {inp} · output {out} · cache {cache} · {ms} ms",
        "refusal": "The model declined part of the request; see the answer.",
        "truncated": "The answer was cut by the output token limit.",
        "mode_explain": "Explanation",
        "mode_review": "Review",
        "mode_document": "Documentation",
        "mode_chat": "Chat",
        "disclaimer": "Automated analysis: validate with the responsible engineer before any change.",
    },
    "es": {
        "title": "Analizador de programas Rockwell (L5X)",
        "subtitle": "Explica, revisa y documenta programas Logix exportados en L5X. Solo lectura: nada se envía al controlador.",
        "code_label": "Código de acceso",
        "code_button": "Entrar",
        "code_invalid": "Código inválido.",
        "code_hint": "Sin un código válido nada funciona. Pida un código al administrador.",
        "logout": "Salir",
        "welcome": "Hola, {owner}.",
        "balance": "Créditos: **{balance}** de {quota}",
        "language": "Idioma de las respuestas",
        "upload_label": "Suba el archivo .L5X (File → Save As → L5X en Studio 5000)",
        "parse_error": "Archivo inválido: {error}",
        "parsing": "Leyendo y analizando el programa…",
        "summary": "Resumen del proyecto",
        "controller": "Controlador",
        "revision": "Revisión",
        "programs": "Programas",
        "routines": "Rutinas",
        "rungs": "Rungs",
        "tags": "Tags",
        "findings": "Findings por severidad",
        "protected_warning": "{n} elemento(s) con source protection: contenido no leído ni analizado.",
        "safety_warning": "Este proyecto contiene contenido de seguridad (Class=\"Safety\"). Es solo lectura: la app explica y documenta, pero nunca sugiere cambios.",
        "no_project": "Suba un archivo .L5X para empezar.",
        "program": "Programa",
        "routine": "Rutina",
        "scope": "Alcance de la acción",
        "scope_routine": "Rutina seleccionada",
        "scope_program": "Programa seleccionado",
        "scope_project": "Proyecto completo",
        "explain": "Explicar",
        "review": "Revisar riesgos y calidad",
        "document": "Generar documentación",
        "cost": "costo: {n} crédito(s)",
        "insufficient": "Créditos insuficientes: esta acción cuesta {cost} y usted tiene {balance}.",
        "routine_protected": "La rutina seleccionada tiene source protection: no hay contenido para explicar o revisar.",
        "routine_not_analyzed": "Rutina {type}: el parser no analiza la lógica interna; la explicación usará solo nombre, descripción y tags referenciadas.",
        "deterministic": "Findings del parser para el alcance seleccionado ({n})",
        "no_findings": "Ningún finding para este alcance.",
        "results": "Resultados",
        "download": "Descargar Markdown",
        "chat_title": "Conversar sobre el programa",
        "chat_placeholder": "Pregunte sobre el programa (1 crédito por pregunta)",
        "chat_clear": "Limpiar conversación",
        "chat_sources": "Rutinas consultadas: {sources}",
        "llm_missing": "Capa LLM no configurada: {error}. Carga, resumen y findings siguen funcionando.",
        "working": "Consultando el modelo…",
        "usage": "{model} · {calls} llamada(s) · entrada {inp} · salida {out} · caché {cache} · {ms} ms",
        "refusal": "El modelo rechazó parte del pedido; vea la respuesta.",
        "truncated": "La respuesta fue cortada por el límite de tokens de salida.",
        "mode_explain": "Explicación",
        "mode_review": "Revisión",
        "mode_document": "Documentación",
        "mode_chat": "Chat",
        "disclaimer": "Análisis automático: validar con el ingeniero responsable antes de cualquier cambio.",
    },
}

LANG_LABELS = {"pt-BR": "Português (BR)", "en": "English", "es": "Español"}
DOC_PROJECT_COST = 3


def t(key: str, **kw: Any) -> str:
    lang = st.session_state.get("lang", "pt-BR")
    text = UI.get(lang, UI["pt-BR"]).get(key) or UI["pt-BR"][key]
    return text.format(**kw) if kw else text


# ----------------------------------------------------------------------------
# Helpers (testáveis sem Streamlit)
# ----------------------------------------------------------------------------
def action_cost(mode: str, scope: str) -> int:
    """Créditos por ação: documentação do projeto inteiro custa 3; o resto 1."""
    if mode == "document" and scope == "project":
        return DOC_PROJECT_COST
    return 1


def parse_upload(data: bytes, name: str = "upload.L5X") -> ProjectContext:
    """Grava o upload em diretório temporário, faz o parse e apaga o diretório."""
    with tempfile.TemporaryDirectory(prefix="l5x_") as tmp:
        path = os.path.join(tmp, os.path.basename(name) or "upload.L5X")
        with open(path, "wb") as fh:
            fh.write(data)
        ctrl = parse_file(path)
    return ProjectContext(ctrl)


def file_hash(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def result_markdown(title: str, text: str, project: str, model: str, lang: str) -> str:
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    disclaimer = UI.get(lang, UI["pt-BR"])["disclaimer"]
    return f"# {title}\n\n_{project} · {stamp} · {model}_\n\n{text}\n\n---\n_{disclaimer}_\n"


def secrets_mapping() -> Optional[Any]:
    try:
        if st.secrets.to_dict():  # lança se não há arquivo de secrets
            return st.secrets
    except Exception:  # noqa: BLE001
        return None
    return None


# ----------------------------------------------------------------------------
# Recursos cacheados
# ----------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def get_store() -> CreditStore:
    store = CreditStore()
    spec = os.environ.get("ACCESS_CODES")
    sec = secrets_mapping()
    if not spec and sec is not None:
        try:
            spec = sec.get("ACCESS_CODES")
        except Exception:  # noqa: BLE001
            spec = None
    if spec:
        store.seed(str(spec))
    return store


@st.cache_resource(show_spinner=False)
def get_llm() -> LLMClient:
    cfg = LLMConfig.from_env(secrets=secrets_mapping())
    return LLMClient(cfg)


@st.cache_resource(show_spinner=False, max_entries=4, ttl=3600)
def load_project(digest: str, _data: bytes, name: str) -> ProjectContext:
    """Chaveado pelo hash do arquivo: o mesmo upload não é analisado duas vezes."""
    return parse_upload(_data, name)


# ----------------------------------------------------------------------------
# Telas
# ----------------------------------------------------------------------------
def login_screen(store: CreditStore) -> None:
    st.title(t("title"))
    st.caption(t("subtitle"))
    with st.form("login"):
        code = st.text_input(t("code_label"), type="password")
        ok = st.form_submit_button(t("code_button"))
    if ok:
        entry = store.validate(code)
        if entry is None:
            st.error(t("code_invalid"))
        else:
            st.session_state["code"] = entry.code
            st.session_state["owner"] = entry.owner
            st.rerun()
    st.info(t("code_hint"))


def sidebar(store: CreditStore, ctx: Optional[ProjectContext]) -> None:
    with st.sidebar:
        entry = store.get(st.session_state["code"])
        st.markdown(t("welcome", owner=st.session_state.get("owner", "")))
        if entry:
            st.markdown(t("balance", balance=entry.balance, quota=entry.quota))
        if st.button(t("logout"), key="logout"):
            for k in ("code", "owner", "history", "results", "project_hash"):
                st.session_state.pop(k, None)
            st.rerun()
        st.selectbox(t("language"), options=list(LANG_LABELS), format_func=lambda k: LANG_LABELS[k], key="lang")
        st.divider()
        if ctx is None:
            st.caption(t("no_project"))
            return
        m = ctx.metrics
        c = ctx.ctrl
        st.subheader(t("summary"))
        st.markdown(f"**{t('controller')}:** {c.name} ({c.processor_type or '?'})  \n**{t('revision')}:** {c.major_rev or '?'}.{c.minor_rev or '?'} · Studio 5000 {c.software_revision or '?'}")
        cols = st.columns(2)
        cols[0].metric(t("programs"), m.programs)
        cols[1].metric(t("routines"), m.routines)
        cols = st.columns(2)
        cols[0].metric(t("rungs"), m.rungs)
        cols[1].metric(t("tags"), m.tags_controller + m.tags_program)
        st.markdown(f"**{t('findings')}**")
        sev = m.findings_by_severity
        st.markdown(" · ".join(f"{s}: **{sev.get(s, 0)}**" for s in SEVERITIES))
        protected = m.protected_routines + m.aois_protected
        if protected:
            st.warning(t("protected_warning", n=protected))
        if m.safety_programs or c.safety_controller:
            st.error(t("safety_warning"))


def findings_table(ctx: ProjectContext, program: Optional[str], routine: Optional[str]) -> None:
    rows = ctx.findings_for(program, routine)
    with st.expander(t("deterministic", n=len(rows)), expanded=False):
        if not rows:
            st.caption(t("no_findings"))
            return
        st.dataframe(
            [{"sev": f.severity, "regra": f.rule, "local": f.location(), "tag": f.tag or "", "mensagem": f.message, "safety": f.safety} for f in rows[:500]],
            width="stretch", hide_index=True,
        )


def run_action(mode: str, scope: str, ctx: ProjectContext, store: CreditStore, program: Optional[str], routine: Optional[str]) -> None:
    cost = action_cost(mode, scope)
    code = st.session_state["code"]
    balance = store.balance(code)
    if balance < cost:
        st.error(t("insufficient", cost=cost, balance=balance))
        return
    try:
        llm = get_llm()
    except LLMError as exc:
        st.error(exc.message)
        return
    lang = st.session_state.get("lang", "pt-BR")
    placeholder = st.empty()
    buf: List[str] = []

    def on_text(delta: str) -> None:
        buf.append(delta)
        if len(buf) % 4 == 0:
            placeholder.markdown("".join(buf) + " ▌")

    try:
        with st.spinner(t("working")):
            if mode == "explain":
                result = llm.explain(ctx, program, routine if scope == "routine" else None, lang, on_text=on_text)  # type: ignore[arg-type]
                title = f"{t('mode_explain')}: {program}" + (f"/{routine}" if scope == "routine" else "")
            elif mode == "review":
                p = None if scope == "project" else program
                r = routine if scope == "routine" else None
                result = llm.review(ctx, p, r, lang, on_text=on_text)
                title = f"{t('mode_review')}: " + (ctx.ctrl.name if scope == "project" else f"{program}" + (f"/{routine}" if r else ""))
            else:
                p = None if scope == "project" else program
                result = llm.document(ctx, p, lang, on_text=on_text)
                title = f"{t('mode_document')}: " + (ctx.ctrl.name if scope == "project" else str(program))
    except LLMError as exc:
        placeholder.empty()
        st.error(exc.message)
        return
    except KeyError as exc:
        placeholder.empty()
        st.error(str(exc))
        return
    placeholder.empty()
    if not store.debit(code, cost):
        st.error(t("insufficient", cost=cost, balance=store.balance(code)))
    st.session_state.setdefault("results", []).insert(0, {
        "title": title, "text": result.text, "mode": mode, "model": result.model,
        "usage": t("usage", model=result.model, calls=result.calls, inp=result.usage.input_tokens, out=result.usage.output_tokens, cache=result.usage.cache_read_tokens, ms=result.latency_ms),
        "refusal": result.refusal, "truncated": result.truncated,
    })
    st.rerun()


def results_area(ctx: ProjectContext) -> None:
    results = st.session_state.get("results", [])
    if not results:
        return
    st.subheader(t("results"))
    lang = st.session_state.get("lang", "pt-BR")
    for i, r in enumerate(results):
        with st.expander(r["title"], expanded=(i == 0)):
            if r.get("refusal"):
                st.warning(t("refusal"))
            if r.get("truncated"):
                st.warning(t("truncated"))
            st.markdown(r["text"])
            st.caption(r["usage"])
            md = result_markdown(r["title"], r["text"], ctx.ctrl.name, r["model"], lang)
            fname = f"{ctx.ctrl.name}_{r['mode']}_{i}.md".replace(" ", "_")
            st.download_button(t("download"), data=md.encode("utf-8"), file_name=fname, mime="text/markdown", key=f"dl_{i}")


def chat_area(ctx: ProjectContext, store: CreditStore) -> None:
    st.subheader(t("chat_title"))
    history: List[Dict[str, str]] = st.session_state.setdefault("history", [])
    if history and st.button(t("chat_clear"), key="chat_clear"):
        st.session_state["history"] = []
        st.rerun()
    for msg in history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
            if msg.get("sources"):
                st.caption(t("chat_sources", sources=msg["sources"]))
    question = st.chat_input(t("chat_placeholder"))
    if not question:
        return
    code = st.session_state["code"]
    if store.balance(code) < 1:
        st.error(t("insufficient", cost=1, balance=store.balance(code)))
        return
    try:
        llm = get_llm()
    except LLMError as exc:
        st.error(exc.message)
        return
    with st.chat_message("user"):
        st.markdown(question)
    with st.chat_message("assistant"):
        placeholder = st.empty()
        buf: List[str] = []

        def on_text(delta: str) -> None:
            buf.append(delta)
            if len(buf) % 4 == 0:
                placeholder.markdown("".join(buf) + " ▌")

        try:
            result = llm.chat(ctx, history, question, st.session_state.get("lang", "pt-BR"), on_text=on_text)
        except LLMError as exc:
            placeholder.empty()
            st.error(exc.message)
            return
        placeholder.markdown(result.text)
        sources = ", ".join(result.partial_results)
        if sources:
            st.caption(t("chat_sources", sources=sources))
    store.debit(code, 1)
    history.append({"role": "user", "content": question})
    history.append({"role": "assistant", "content": result.text, "sources": sources})
    st.rerun()


def main_area(store: CreditStore) -> Optional[ProjectContext]:
    st.title(t("title"))
    st.caption(t("subtitle"))
    uploaded = st.file_uploader(t("upload_label"), type=["l5x", "L5X"], accept_multiple_files=False)
    ctx: Optional[ProjectContext] = None
    if uploaded is not None:
        data = uploaded.getvalue()
        digest = file_hash(data)
        if st.session_state.get("project_hash") != digest:
            # novo arquivo: zera histórico e resultados do anterior
            st.session_state["history"] = []
            st.session_state["results"] = []
            st.session_state["project_hash"] = digest
        try:
            with st.spinner(t("parsing")):
                ctx = load_project(digest, data, uploaded.name)
        except L5XError as exc:
            st.error(t("parse_error", error=exc))
            return None
    if ctx is None:
        st.info(t("no_project"))
        return None

    # --- camada LLM disponível? ---
    llm_ok = True
    try:
        get_llm()
    except LLMError as exc:
        llm_ok = False
        st.warning(t("llm_missing", error=exc.message))

    # --- seleção de escopo ---
    programs = [p.name for p in ctx.ctrl.programs]
    c1, c2, c3 = st.columns([2, 2, 2])
    program = c1.selectbox(t("program"), programs, key="sel_program") if programs else None
    prog = ctx.ctrl.program(program) if program else None
    routines = [r.name for r in prog.routines] if prog else []
    routine = c2.selectbox(t("routine"), routines, key="sel_routine") if routines else None
    scope_labels = {"routine": t("scope_routine"), "program": t("scope_program"), "project": t("scope_project")}
    scope = c3.radio(t("scope"), list(scope_labels), format_func=lambda k: scope_labels[k], key="scope", horizontal=True)

    rt = prog.routine(routine) if (prog and routine) else None
    if scope == "routine" and rt is not None:
        if rt.protected:
            st.warning(t("routine_protected"))
        elif not rt.analyzed:
            st.info(t("routine_not_analyzed", type=rt.type))

    findings_table(ctx, None if scope == "project" else program, routine if scope == "routine" else None)

    # --- ações ---
    b1, b2, b3 = st.columns(3)
    can_act = llm_ok and not (scope == "routine" and (rt is None or rt.protected))
    explain_disabled = not can_act or scope == "project"
    cost_explain = action_cost("explain", scope)
    cost_review = action_cost("review", scope)
    cost_doc = action_cost("document", scope if scope != "routine" else "program")
    if b1.button(f"{t('explain')} · {t('cost', n=cost_explain)}", disabled=explain_disabled, width="stretch"):
        run_action("explain", scope, ctx, store, program, routine)
    if b2.button(f"{t('review')} · {t('cost', n=cost_review)}", disabled=not can_act, width="stretch"):
        run_action("review", scope, ctx, store, program, routine)
    if b3.button(f"{t('document')} · {t('cost', n=cost_doc)}", disabled=not llm_ok, width="stretch"):
        run_action("document", "project" if scope == "project" else "program", ctx, store, program, routine)

    results_area(ctx)
    st.divider()
    if llm_ok:
        chat_area(ctx, store)
    return ctx


def main() -> None:
    st.set_page_config(page_title="Analisador L5X", page_icon="🔎", layout="wide")
    st.session_state.setdefault("lang", os.environ.get("APP_LANGUAGE", "pt-BR") if os.environ.get("APP_LANGUAGE") in LANGUAGES else "pt-BR")
    store = get_store()
    if not st.session_state.get("code"):
        login_screen(store)
        return
    # a sidebar precisa do ctx, que vem do upload na área principal
    ctx = main_area(store)
    sidebar(store, ctx)


if __name__ == "__main__" or st.runtime.exists():  # pragma: no cover - executado pelo Streamlit
    main()
