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
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

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
        "mode_aoi": "AOI",
        "aoi_section": "AOIs (Add-On Instructions)",
        "aoi_select": "AOI",
        "aoi_none": "Este projeto não tem AOIs.",
        "aoi_option": "{name} rev. {rev} · {instances} inst. · {calls} chamadas{protected}",
        "aoi_protected_tag": " · protegida",
        "aoi_summary": "{params} parâmetros · {locals} tags locais · rotinas: {routines} · instâncias: {instances} · chamadas: {calls} · findings: {findings}",
        "aoi_report": "Relatório determinístico (grátis): documentação + roteiro de debug",
        "aoi_download_report": "Baixar relatório (Markdown)",
        "aoi_findings": "Findings do parser para esta AOI ({n})",
        "aoi_analyze": "Analisar AOI com o modelo",
        "aoi_protected_msg": "AOI protegida (source protection): o parser lista instâncias e chamadas, mas não há lógica para analisar.",
        "aoi_help": "A análise pelo modelo usa a interface, os findings e as rotinas internas da AOI e devolve finalidade, comportamento, uso no projeto, riscos e um roteiro de debug.",
        "unconnected_section": "Tags sem fio em rotinas FBD ({tags} tag(s) em {routines} rotina(s))",
        "unconnected_help": "Referências de tag (IRef/ORef) e conectores colocados na folha sem nenhum fio: a tag não alimenta nem recebe nada. Blocos sem nenhum fio também aparecem. Análise determinística, sem custo.",
        "unconnected_none": "Nenhum elemento FBD sem fio neste projeto.",
        "unconnected_download": "Baixar lista (CSV)",
        "nav_programs": "Programas e rotinas",
        "nav_aois": "AOIs",
        "nav_diag": "Diagnóstico",
        "nav_chat": "Chat",
        "how_title": "Como funciona",
        "how_steps": "1. Exporte o projeto do Studio 5000 em L5X (File → Save As → L5X).\n2. Envie o arquivo: o parser lê programas, rotinas, tags, AOIs e gera findings na hora, sem custo.\n3. Peça ao modelo para explicar, revisar ou documentar; cada ação debita créditos só se concluir.",
        "privacy_note": "O arquivo fica só nesta sessão e é apagado após a leitura. O modelo recebe apenas o índice estruturado, nunca o XML; conteúdo protegido não é decodificado.",
        "file_section": "Arquivo do projeto",
        "file_loaded": "**{name}** · {size:.1f} MB · lido em {secs:.1f} s · {programs} programas · {routines} rotinas",
        "demo_loaded": "Arquivo de demonstração carregado de APP_DEMO_FILE.",
        "empty_title": "Nenhum projeto carregado",
        "target": "Alvo",
        "target_hint": "Explicar e Revisar aceitam rotina ou programa; Documentar e Revisar aceitam o projeto inteiro.",
        "size_rungs": "{n} rungs",
        "size_lines": "{n} linhas",
        "size_sheets": "{n} folha(s)",
        "size_na": "não analisada",
        "badge_protected": "protegida",
        "badge_safety": "SAFETY",
        "badge_main": "main",
        "badge_fault": "fault",
        "badge_unscheduled": "não agendado",
        "badge_disabled": "desabilitado",
        "called_by": "Chamada por: {names}",
        "calls_to": "Chama: {names}",
        "never_called": "Nunca chamada por JSR.",
        "no_description": "Sem descrição no projeto.",
        "explain_help": "Explica a lógica para um técnico de manutenção, rung a rung ou por bloco, citando rung/folha. Não disponível para o projeto inteiro.",
        "review_help": "Prioriza os findings do parser, explica o impacto e sugere correções genéricas; separa o confirmado das hipóteses.",
        "document_help": "Gera documentação em Markdown: visão geral, descrição por rotina, I/O, timers, intertravamentos e observações do parser.",
        "disabled_llm": "Camada LLM não configurada.",
        "disabled_protected": "Rotina protegida: sem conteúdo.",
        "disabled_project": "Escolha um programa ou uma rotina.",
        "results_clear": "Limpar resultados",
        "results_empty": "Os resultados das ações desta área aparecem aqui, do mais recente para o mais antigo.",
        "findings_title": "Findings do parser",
        "findings_sev": "Severidade",
        "findings_search": "Buscar por regra, local, tag ou mensagem",
        "findings_shown": "{shown} de {total} findings",
        "findings_download": "Baixar findings (CSV)",
        "col_sev": "Sev.",
        "col_rule": "Regra",
        "col_where": "Local",
        "col_tag": "Tag",
        "col_msg": "Mensagem",
        "col_evidence": "Evidência",
        "col_safety": "Safety",
        "credits_low": "Créditos acabando: {balance} restante(s).",
        "credits_none": "Sem créditos. Peça mais ao administrador.",
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
        "mode_aoi": "AOI",
        "aoi_section": "AOIs (Add-On Instructions)",
        "aoi_select": "AOI",
        "aoi_none": "This project has no AOIs.",
        "aoi_option": "{name} rev. {rev} · {instances} inst. · {calls} calls{protected}",
        "aoi_protected_tag": " · protected",
        "aoi_summary": "{params} parameters · {locals} local tags · routines: {routines} · instances: {instances} · calls: {calls} · findings: {findings}",
        "aoi_report": "Deterministic report (free): documentation + debug guide",
        "aoi_download_report": "Download report (Markdown)",
        "aoi_findings": "Parser findings for this AOI ({n})",
        "aoi_analyze": "Analyze AOI with the model",
        "aoi_protected_msg": "Source-protected AOI: the parser lists instances and calls, but there is no logic to analyze.",
        "aoi_help": "The model analysis uses the AOI interface, findings and internal routines and returns purpose, behaviour, usage in the project, risks and a debug guide.",
        "unconnected_section": "Unwired tags in FBD routines ({tags} tag(s) in {routines} routine(s))",
        "unconnected_help": "Tag references (IRef/ORef) and connectors placed on the sheet without any wire: the tag feeds nothing and is fed by nothing. Blocks without any wire are listed too. Deterministic, free.",
        "unconnected_none": "No unwired FBD element in this project.",
        "unconnected_download": "Download list (CSV)",
        "nav_programs": "Programs and routines",
        "nav_aois": "AOIs",
        "nav_diag": "Diagnostics",
        "nav_chat": "Chat",
        "how_title": "How it works",
        "how_steps": "1. Export the project from Studio 5000 as L5X (File → Save As → L5X).\n2. Upload it: the parser reads programs, routines, tags and AOIs and produces findings instantly, free of charge.\n3. Ask the model to explain, review or document; each action is charged only when it completes.",
        "privacy_note": "The file lives only in this session and is deleted after parsing. The model receives the structured index only, never the XML; protected content is not decoded.",
        "file_section": "Project file",
        "file_loaded": "**{name}** · {size:.1f} MB · parsed in {secs:.1f} s · {programs} programs · {routines} routines",
        "demo_loaded": "Demo file loaded from APP_DEMO_FILE.",
        "empty_title": "No project loaded",
        "target": "Target",
        "target_hint": "Explain and Review take a routine or a program; Document and Review also take the whole project.",
        "size_rungs": "{n} rungs",
        "size_lines": "{n} lines",
        "size_sheets": "{n} sheet(s)",
        "size_na": "not analyzed",
        "badge_protected": "protected",
        "badge_safety": "SAFETY",
        "badge_main": "main",
        "badge_fault": "fault",
        "badge_unscheduled": "unscheduled",
        "badge_disabled": "disabled",
        "called_by": "Called by: {names}",
        "calls_to": "Calls: {names}",
        "never_called": "Never called by a JSR.",
        "no_description": "No description in the project.",
        "explain_help": "Explains the logic for a maintenance technician, rung by rung or by block, citing rung/sheet. Not available for the whole project.",
        "review_help": "Prioritizes parser findings, explains impact and suggests generic fixes; separates confirmed facts from hypotheses.",
        "document_help": "Generates Markdown documentation: overview, per-routine description, I/O, timers, interlocks and parser notes.",
        "disabled_llm": "LLM layer not configured.",
        "disabled_protected": "Protected routine: no content.",
        "disabled_project": "Pick a program or a routine.",
        "results_clear": "Clear results",
        "results_empty": "Results of this area's actions appear here, newest first.",
        "findings_title": "Parser findings",
        "findings_sev": "Severity",
        "findings_search": "Search by rule, location, tag or message",
        "findings_shown": "{shown} of {total} findings",
        "findings_download": "Download findings (CSV)",
        "col_sev": "Sev.",
        "col_rule": "Rule",
        "col_where": "Location",
        "col_tag": "Tag",
        "col_msg": "Message",
        "col_evidence": "Evidence",
        "col_safety": "Safety",
        "credits_low": "Running low on credits: {balance} left.",
        "credits_none": "No credits left. Ask the administrator for more.",
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
        "mode_aoi": "AOI",
        "aoi_section": "AOIs (Add-On Instructions)",
        "aoi_select": "AOI",
        "aoi_none": "Este proyecto no tiene AOIs.",
        "aoi_option": "{name} rev. {rev} · {instances} inst. · {calls} llamadas{protected}",
        "aoi_protected_tag": " · protegida",
        "aoi_summary": "{params} parámetros · {locals} tags locales · rutinas: {routines} · instancias: {instances} · llamadas: {calls} · findings: {findings}",
        "aoi_report": "Informe determinístico (gratis): documentación + guía de debug",
        "aoi_download_report": "Descargar informe (Markdown)",
        "aoi_findings": "Findings del parser para esta AOI ({n})",
        "aoi_analyze": "Analizar AOI con el modelo",
        "aoi_protected_msg": "AOI protegida (source protection): el parser lista instancias y llamadas, pero no hay lógica para analizar.",
        "aoi_help": "El análisis del modelo usa la interfaz, los findings y las rutinas internas de la AOI y devuelve finalidad, comportamiento, uso en el proyecto, riesgos y una guía de debug.",
        "unconnected_section": "Tags sin cable en rutinas FBD ({tags} tag(s) en {routines} rutina(s))",
        "unconnected_help": "Referencias de tag (IRef/ORef) y conectores colocados en la hoja sin ningún cable: la tag no alimenta ni recibe nada. Los bloques sin ningún cable también aparecen. Análisis determinístico, sin costo.",
        "unconnected_none": "Ningún elemento FBD sin cable en este proyecto.",
        "unconnected_download": "Descargar lista (CSV)",
        "nav_programs": "Programas y rutinas",
        "nav_aois": "AOIs",
        "nav_diag": "Diagnóstico",
        "nav_chat": "Chat",
        "how_title": "Cómo funciona",
        "how_steps": "1. Exporte el proyecto desde Studio 5000 en L5X (File → Save As → L5X).\n2. Suba el archivo: el parser lee programas, rutinas, tags y AOIs y genera findings al instante, sin costo.\n3. Pida al modelo explicar, revisar o documentar; cada acción se cobra solo si concluye.",
        "privacy_note": "El archivo vive solo en esta sesión y se borra tras la lectura. El modelo recibe únicamente el índice estructurado, nunca el XML; el contenido protegido no se decodifica.",
        "file_section": "Archivo del proyecto",
        "file_loaded": "**{name}** · {size:.1f} MB · leído en {secs:.1f} s · {programs} programas · {routines} rutinas",
        "demo_loaded": "Archivo de demostración cargado desde APP_DEMO_FILE.",
        "empty_title": "Ningún proyecto cargado",
        "target": "Objetivo",
        "target_hint": "Explicar y Revisar aceptan rutina o programa; Documentar y Revisar aceptan el proyecto completo.",
        "size_rungs": "{n} rungs",
        "size_lines": "{n} líneas",
        "size_sheets": "{n} hoja(s)",
        "size_na": "no analizada",
        "badge_protected": "protegida",
        "badge_safety": "SAFETY",
        "badge_main": "main",
        "badge_fault": "fault",
        "badge_unscheduled": "no programado",
        "badge_disabled": "deshabilitado",
        "called_by": "Llamada por: {names}",
        "calls_to": "Llama: {names}",
        "never_called": "Nunca llamada por JSR.",
        "no_description": "Sin descripción en el proyecto.",
        "explain_help": "Explica la lógica para un técnico de mantenimiento, rung a rung o por bloque, citando rung/hoja. No disponible para el proyecto completo.",
        "review_help": "Prioriza los findings del parser, explica el impacto y sugiere correcciones genéricas; separa lo confirmado de las hipótesis.",
        "document_help": "Genera documentación en Markdown: visión general, descripción por rutina, I/O, timers, enclavamientos y observaciones del parser.",
        "disabled_llm": "Capa LLM no configurada.",
        "disabled_protected": "Rutina protegida: sin contenido.",
        "disabled_project": "Elija un programa o una rutina.",
        "results_clear": "Limpiar resultados",
        "results_empty": "Los resultados de las acciones de esta área aparecen aquí, del más reciente al más antiguo.",
        "findings_title": "Findings del parser",
        "findings_sev": "Severidad",
        "findings_search": "Buscar por regla, lugar, tag o mensaje",
        "findings_shown": "{shown} de {total} findings",
        "findings_download": "Descargar findings (CSV)",
        "col_sev": "Sev.",
        "col_rule": "Regla",
        "col_where": "Lugar",
        "col_tag": "Tag",
        "col_msg": "Mensaje",
        "col_evidence": "Evidencia",
        "col_safety": "Safety",
        "credits_low": "Se acaban los créditos: quedan {balance}.",
        "credits_none": "Sin créditos. Pida más al administrador.",
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
    """Créditos por ação: documentação do projeto inteiro custa 3; o resto (inclusive AOI) 1."""
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
# Apresentação
# ----------------------------------------------------------------------------
SEV_ICON = {"error": "🔴", "warning": "🟠", "info": "🔵"}
NAV = ("programs", "aois", "diag", "chat")
NAV_ICON = {"programs": "📘", "aois": "🧩", "diag": "🩺", "chat": "💬"}
AREA_MODES = {"programs": ("explain", "review", "document"), "aois": ("aoi",)}


def routine_size(r: Any) -> str:
    if r.protected:
        return t("badge_protected")
    if r.type == "RLL":
        return t("size_rungs", n=len(r.rungs))
    if r.type == "ST":
        return t("size_lines", n=len(r.st_lines))
    if r.type == "FBD":
        return t("size_sheets", n=r.sheets)
    return t("size_na")


def routine_label(prog: Any, r: Any) -> str:
    flags = []
    if prog.main_routine and r.name.lower() == prog.main_routine.lower():
        flags.append(t("badge_main"))
    if prog.fault_routine and r.name.lower() == prog.fault_routine.lower():
        flags.append(t("badge_fault"))
    if r.protected:
        flags.append("🔒")
    tail = f" · {' '.join(flags)}" if flags else ""
    return f"{r.name} · {r.type} · {routine_size(r)}{tail}"


def program_label(p: Any) -> str:
    flags = []
    if p.task is None:
        flags.append(t("badge_unscheduled"))
    if p.disabled:
        flags.append(t("badge_disabled"))
    if p.is_safety:
        flags.append(t("badge_safety"))
    tail = f" · {', '.join(flags)}" if flags else ""
    return f"{p.name} · {len(p.routines)} {t('routines').lower()}{tail}"


def findings_rows(findings: List[Any]) -> List[Dict[str, Any]]:
    return [{
        "sev": f"{SEV_ICON.get(f.severity, '')} {f.severity}", "rule": f.rule, "where": f.location(), "tag": f.tag or "",
        "message": f.message, "evidence": f.evidence or "", "safety": f.safety,
    } for f in findings]


def findings_csv(findings: List[Any]) -> str:
    def cell(v: Any) -> str:
        return str(v).replace(";", ",").replace("\n", " ")

    lines = ["severidade;regra;local;tag;mensagem;evidencia;safety"]
    for f in findings:
        lines.append(";".join(cell(v) for v in (f.severity, f.rule, f.location(), f.tag or "", f.message, f.evidence or "", f.safety)))
    return "\n".join(lines) + "\n"


def filter_findings(findings: List[Any], severities: List[str], query: str) -> List[Any]:
    q = (query or "").strip().lower()
    out = []
    for f in findings:
        if severities and f.severity not in severities:
            continue
        if q and q not in " ".join((f.rule, f.location(), f.tag or "", f.message, f.evidence or "")).lower():
            continue
        out.append(f)
    return out


def unconnected_csv(rows: List[Dict[str, Any]]) -> str:
    lines = ["programa;rotina;folha;tipo;elemento"]
    for r in rows:
        lines.append(";".join(str(r[k]) for k in ("programa", "rotina", "folha", "tipo", "elemento")))
    return "\n".join(lines) + "\n"


# ----------------------------------------------------------------------------
# Telas
# ----------------------------------------------------------------------------
def login_screen(store: CreditStore) -> None:
    st.title(t("title"))
    st.caption(t("subtitle"))
    left, right = st.columns([1, 1], gap="large")
    with left:
        with st.form("login"):
            code = st.text_input(t("code_label"), type="password")
            ok = st.form_submit_button(t("code_button"), type="primary")
        if ok:
            entry = store.validate(code)
            if entry is None:
                st.error(t("code_invalid"))
            else:
                st.session_state["code"] = entry.code
                st.session_state["owner"] = entry.owner
                st.rerun()
        st.info(t("code_hint"))
    with right:
        st.subheader(t("how_title"))
        st.markdown(t("how_steps").replace("\n", "  \n"))
        st.caption(t("privacy_note"))


def sidebar(store: CreditStore, ctx: Optional[ProjectContext]) -> None:
    with st.sidebar:
        entry = store.get(st.session_state["code"])
        st.markdown(t("welcome", owner=st.session_state.get("owner", "")))
        if entry:
            st.markdown(t("balance", balance=entry.balance, quota=entry.quota))
            st.progress(min(1.0, entry.balance / entry.quota) if entry.quota else 0.0)
            if entry.balance <= 0:
                st.error(t("credits_none"))
            elif entry.balance <= max(2, entry.quota // 10):
                st.warning(t("credits_low", balance=entry.balance))
        if st.button(t("logout"), key="logout", width="stretch"):
            for k in ("code", "owner", "history", "results", "project_hash", "file_meta", "nav"):
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
        st.markdown(f"**{c.name}**  \n{c.processor_type or '?'} · {t('revision')} {c.major_rev or '?'}.{c.minor_rev or '?'} · Studio 5000 {c.software_revision or '?'}")
        meta = st.session_state.get("file_meta")
        if meta:
            st.caption(f"{meta['name']} · {meta['size_mb']:.1f} MB · {meta['secs']:.1f} s")
        cols = st.columns(2)
        cols[0].metric(t("programs"), m.programs)
        cols[1].metric(t("routines"), m.routines)
        cols = st.columns(2)
        cols[0].metric(t("rungs"), m.rungs)
        cols[1].metric(t("tags"), m.tags_controller + m.tags_program)
        cols = st.columns(2)
        cols[0].metric("AOIs", m.aois)
        cols[1].metric("FBD", m.routines_by_type.get("FBD", 0))
        st.markdown(f"**{t('findings')}**")
        sev = m.findings_by_severity
        st.markdown(" · ".join(f"{SEV_ICON[s]} **{sev.get(s, 0)}** {s}" for s in SEVERITIES))
        protected = m.protected_routines + m.aois_protected
        if protected:
            st.warning(t("protected_warning", n=protected))
        if m.safety_programs or c.safety_controller:
            st.error(t("safety_warning"))


def findings_panel(findings: List[Any], key: str, title: Optional[str] = None, expanded: bool = True) -> None:
    """Tabela de findings com filtro por severidade, busca e download."""
    total = len(findings)
    with st.expander(title or t("deterministic", n=total), expanded=expanded):
        if not findings:
            st.caption(t("no_findings"))
            return
        f1, f2 = st.columns([1, 2])
        with f1:
            severities = st.pills(t("findings_sev"), list(SEVERITIES), selection_mode="multi", default=list(SEVERITIES),
                                  format_func=lambda s: f"{SEV_ICON[s]} {s}", key=f"{key}_sev") or []
        with f2:
            query = st.text_input(t("findings_search"), key=f"{key}_q", placeholder="output_multiple_writes, Motor, rung 3…")
        rows = filter_findings(findings, severities, query)
        st.caption(t("findings_shown", shown=len(rows), total=total))
        st.dataframe(
            findings_rows(rows[:1000]), width="stretch", hide_index=True,
            column_config={
                "sev": st.column_config.TextColumn(t("col_sev"), width="small"),
                "rule": st.column_config.TextColumn(t("col_rule"), width="medium"),
                "where": st.column_config.TextColumn(t("col_where"), width="medium"),
                "tag": st.column_config.TextColumn(t("col_tag"), width="small"),
                "message": st.column_config.TextColumn(t("col_msg"), width="large"),
                "evidence": st.column_config.TextColumn(t("col_evidence"), width="medium"),
                "safety": st.column_config.CheckboxColumn(t("col_safety"), width="small"),
            },
        )
        st.download_button(t("findings_download"), data=findings_csv(rows).encode("utf-8-sig"), file_name=f"findings_{key}.csv", mime="text/csv", key=f"{key}_dl")


def findings_table(ctx: ProjectContext, program: Optional[str], routine: Optional[str], rows: Optional[List[Any]] = None, title: Optional[str] = None) -> None:
    if rows is None:
        rows = ctx.findings_for(program, routine)
    findings_panel(rows, key=f"f_{program or 'all'}_{routine or 'all'}".replace(" ", "_"), title=title, expanded=False)


def run_action(mode: str, scope: str, ctx: ProjectContext, store: CreditStore, program: Optional[str], routine: Optional[str], aoi: Optional[str] = None) -> None:
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
            elif mode == "aoi":
                result = llm.aoi(ctx, aoi or "", lang, on_text=on_text)
                title = f"{t('mode_aoi')}: {aoi}"
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
    st.toast(title, icon="✅")
    st.rerun()


def results_panel(ctx: ProjectContext, modes: Tuple[str, ...]) -> None:
    results = [(i, r) for i, r in enumerate(st.session_state.get("results", [])) if r["mode"] in modes]
    head, action = st.columns([4, 1])
    head.subheader(t("results"))
    if results and action.button(t("results_clear"), key=f"clear_{modes[0]}", width="stretch"):
        st.session_state["results"] = [r for r in st.session_state.get("results", []) if r["mode"] not in modes]
        st.rerun()
    if not results:
        st.caption(t("results_empty"))
        return
    lang = st.session_state.get("lang", "pt-BR")
    for n, (i, r) in enumerate(results):
        with st.expander(r["title"], expanded=(n == 0)):
            if r.get("refusal"):
                st.warning(t("refusal"))
            if r.get("truncated"):
                st.warning(t("truncated"))
            st.markdown(r["text"])
            st.caption(r["usage"])
            md = result_markdown(r["title"], r["text"], ctx.ctrl.name, r["model"], lang)
            fname = f"{ctx.ctrl.name}_{r['mode']}_{i}.md".replace(" ", "_")
            st.download_button(t("download"), data=md.encode("utf-8"), file_name=fname, mime="text/markdown", key=f"dl_{i}")


def target_card(ctx: ProjectContext, scope: str, prog: Any, rt: Any) -> None:
    """Cartão com o contexto do alvo escolhido (descrição, chamadas, avisos)."""
    if scope == "project":
        c = ctx.ctrl
        st.caption(f"{c.name} · {ctx.metrics.programs} {t('programs').lower()} · {ctx.metrics.routines} {t('routines').lower()} · " + ", ".join(f"{k}: {v}" for k, v in sorted(ctx.metrics.routines_by_type.items())))
        return
    if prog is None:
        return
    badges = []
    if prog.is_safety:
        badges.append(f":red-badge[{t('badge_safety')}]")
    if prog.task is None:
        badges.append(f":orange-badge[{t('badge_unscheduled')}]")
    if prog.disabled:
        badges.append(f":orange-badge[{t('badge_disabled')}]")
    if scope == "routine" and rt is not None:
        if rt.protected:
            badges.append(f":gray-badge[🔒 {t('badge_protected')}]")
        badges.append(f":blue-badge[{rt.type} · {routine_size(rt)}]")
        desc = " ".join((rt.description or "").split()) or t("no_description")
        callers = ctx.xref.callers_of(prog.name, rt.name)
        is_root = rt.name.lower() in {(prog.main_routine or '').lower(), (prog.fault_routine or '').lower()}
        lines = [desc]
        if callers:
            lines.append(t("called_by", names=", ".join(f"{p}/{r}" for p, r in callers[:6]) + (" …" if len(callers) > 6 else "")))
        elif not is_root:
            lines.append(t("never_called"))
        if rt.calls:
            lines.append(t("calls_to", names=", ".join(dict.fromkeys(rt.calls))))
        if rt.aoi_calls:
            lines.append("AOIs: " + ", ".join(dict.fromkeys(c.aoi for c in rt.aoi_calls)))
    else:
        badges.append(f":blue-badge[{len(prog.routines)} {t('routines').lower()}]")
        desc = " ".join((prog.description or "").split()) or t("no_description")
        lines = [desc, f"main: {prog.main_routine or '-'} · fault: {prog.fault_routine or '-'} · task: {prog.task or '-'}"]
    st.markdown(" ".join(badges))
    st.caption("  \n".join(lines))
    if scope == "routine" and rt is not None:
        if rt.protected:
            st.warning(t("routine_protected"))
        elif not rt.has_content:
            st.info(t("routine_not_analyzed", type=rt.type))


def programs_area(ctx: ProjectContext, store: CreditStore, llm_ok: bool) -> None:
    scope_labels = {"project": t("scope_project"), "program": t("scope_program"), "routine": t("scope_routine")}
    st.session_state.setdefault("scope", "routine")
    scope = st.segmented_control(t("target"), list(scope_labels), format_func=lambda k: scope_labels[k], key="scope", width="stretch") or "routine"
    programs = ctx.ctrl.programs
    prog = rt = None
    if scope != "project" and programs:
        c1, c2 = st.columns([1, 1])
        by_name = {p.name: p for p in programs}
        program = c1.selectbox(t("program"), list(by_name), format_func=lambda n: program_label(by_name[n]), key="sel_program")
        prog = by_name[program]
        if scope == "routine" and prog.routines:
            by_r = {r.name: r for r in prog.routines}
            routine = c2.selectbox(t("routine"), list(by_r), format_func=lambda n: routine_label(prog, by_r[n]), key="sel_routine")
            rt = by_r[routine]
    target_card(ctx, scope, prog, rt)
    st.caption(t("target_hint"))

    program = prog.name if prog else None
    routine = rt.name if rt else None
    reason = None
    if not llm_ok:
        reason = t("disabled_llm")
    elif scope == "routine" and (rt is None or rt.protected):
        reason = t("disabled_protected")
    b1, b2, b3 = st.columns(3)
    explain_reason = reason or (t("disabled_project") if scope == "project" else None)
    cost_explain = action_cost("explain", scope)
    cost_review = action_cost("review", scope)
    cost_doc = action_cost("document", scope if scope != "routine" else "program")
    if b1.button(f"✨ {t('explain')} · {t('cost', n=cost_explain)}", disabled=explain_reason is not None, width="stretch", type="primary",
                 help=explain_reason or t("explain_help")):
        run_action("explain", scope, ctx, store, program, routine)
    if b2.button(f"🔍 {t('review')} · {t('cost', n=cost_review)}", disabled=reason is not None, width="stretch", help=reason or t("review_help")):
        run_action("review", scope, ctx, store, program, routine)
    doc_reason = t("disabled_llm") if not llm_ok else None
    if b3.button(f"📄 {t('document')} · {t('cost', n=cost_doc)}", disabled=doc_reason is not None, width="stretch", help=doc_reason or t("document_help")):
        run_action("document", "project" if scope == "project" else "program", ctx, store, program, routine)

    findings_table(ctx, None if scope == "project" else program, routine if scope == "routine" else None)
    st.divider()
    results_panel(ctx, AREA_MODES["programs"])


def unconnected_area(ctx: ProjectContext) -> None:
    """Elementos FBD sem fio (tags soltas na folha), agrupados por rotina. Grátis."""
    rows = ctx.unconnected_rows()
    tags = sum(1 for r in rows if r["tag"])
    routines = len({(r["programa"], r["rotina"]) for r in rows})
    with st.expander(t("unconnected_section", tags=tags, routines=routines), expanded=bool(rows)):
        st.caption(t("unconnected_help"))
        if not rows:
            st.caption(t("unconnected_none"))
            return
        st.dataframe([{k: r[k] for k in ("programa", "rotina", "folha", "tipo", "elemento")} for r in rows[:1000]], width="stretch", hide_index=True)
        st.download_button(t("unconnected_download"), data=unconnected_csv(rows).encode("utf-8-sig"), file_name=f"{ctx.ctrl.name}_tags_sem_fio.csv".replace(" ", "_"), mime="text/csv", key="dl_unconnected")


def diag_area(ctx: ProjectContext) -> None:
    sev = ctx.metrics.findings_by_severity
    cols = st.columns(3)
    for col, s in zip(cols, SEVERITIES):
        col.metric(f"{SEV_ICON[s]} {s}", sev.get(s, 0))
    findings_panel(ctx.findings, key="project", title=t("findings_title") + f" ({len(ctx.findings)})", expanded=True)
    unconnected_area(ctx)


def aoi_area(ctx: ProjectContext, store: CreditStore, llm_ok: bool) -> None:
    """Seção de AOIs: relatório determinístico (grátis) + análise pelo modelo (1 crédito)."""
    rows = ctx.aoi_rows()
    if not rows:
        st.caption(t("aoi_none"))
        return
    by_name = {r["aoi"]: r for r in rows}

    def label(name: str) -> str:
        r = by_name[name]
        return t("aoi_option", name=name, rev=r["revision"] or "?", instances=r["instances"], calls=r["calls"], protected=t("aoi_protected_tag") if r["protected"] else "")

    name = st.selectbox(t("aoi_select"), list(by_name), format_func=label, key="sel_aoi")
    row = by_name[name]
    f = row["findings"]
    badges = [f":blue-badge[{row['parameters']} par.]", f":blue-badge[{row['instances']} inst.]", f":blue-badge[{row['calls']} calls]"]
    if row["protected"]:
        badges.append(f":gray-badge[🔒 {t('badge_protected')}]")
    for s in SEVERITIES:
        if f.get(s, 0):
            color = {"error": "red", "warning": "orange", "info": "gray"}[s]
            badges.append(f":{color}-badge[{SEV_ICON[s]} {f[s]} {s}]")
    st.markdown(" ".join(badges))
    st.caption(t("aoi_summary", params=row["parameters"], locals=row["locals"], routines=row["routines"] or "-", instances=row["instances"], calls=row["calls"],
                  findings=" · ".join(f"{s}: {f.get(s, 0)}" for s in SEVERITIES)))
    if row["description"]:
        st.caption(row["description"])
    if row["protected"]:
        st.info(t("aoi_protected_msg"))
    cost = action_cost("aoi", "aoi")
    reason = t("disabled_llm") if not llm_ok else (t("disabled_protected") if row["protected"] else None)
    if st.button(f"✨ {t('aoi_analyze')} · {t('cost', n=cost)}", disabled=reason is not None, help=reason or t("aoi_help"), key="btn_aoi", type="primary"):
        run_action("aoi", "aoi", ctx, store, None, None, aoi=name)
    with st.expander(t("aoi_report"), expanded=False):
        md = ctx.aoi_markdown(name)
        st.markdown(md)
        st.download_button(t("aoi_download_report"), data=md.encode("utf-8"), file_name=f"{ctx.ctrl.name}_AOI_{name}.md".replace(" ", "_"), mime="text/markdown", key="dl_aoi_report")
    aoi_findings = ctx.aoi_findings(name)
    findings_panel(aoi_findings, key=f"aoi_{name}", title=t("aoi_findings", n=len(aoi_findings)), expanded=False)
    st.divider()
    results_panel(ctx, AREA_MODES["aois"])


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


def load_from_upload(data: bytes, name: str) -> Optional[ProjectContext]:
    """Carrega (com cache por hash) e registra os metadados do arquivo na sessão."""
    digest = file_hash(data)
    if st.session_state.get("project_hash") != digest:
        st.session_state["history"] = []
        st.session_state["results"] = []
        st.session_state["project_hash"] = digest
        st.session_state.pop("file_meta", None)
    t0 = time.perf_counter()
    try:
        with st.spinner(t("parsing")):
            ctx = load_project(digest, data, name)
    except L5XError as exc:
        st.error(t("parse_error", error=exc))
        return None
    if "file_meta" not in st.session_state:
        st.session_state["file_meta"] = {"name": name, "size_mb": len(data) / 1e6, "secs": time.perf_counter() - t0}
    return ctx


def main_area(store: CreditStore) -> Optional[ProjectContext]:
    st.title(t("title"))
    st.caption(t("subtitle"))
    loaded = bool(st.session_state.get("project_hash"))
    ctx: Optional[ProjectContext] = None
    with st.expander(t("file_section"), expanded=not loaded):
        uploaded = st.file_uploader(t("upload_label"), type=["l5x", "L5X"], accept_multiple_files=False)
        if uploaded is not None:
            ctx = load_from_upload(uploaded.getvalue(), uploaded.name)
        else:
            demo = os.environ.get("APP_DEMO_FILE")
            if demo and os.path.isfile(demo):
                with open(demo, "rb") as fh:
                    ctx = load_from_upload(fh.read(), os.path.basename(demo))
                if ctx is not None:
                    st.caption(t("demo_loaded"))
        if ctx is not None:
            meta = st.session_state.get("file_meta") or {"name": "", "size_mb": 0.0, "secs": 0.0}
            st.success(t("file_loaded", name=meta["name"], size=meta["size_mb"], secs=meta["secs"], programs=ctx.metrics.programs, routines=ctx.metrics.routines))
    if ctx is None:
        st.info(t("no_project"))
        st.subheader(t("how_title"))
        st.markdown(t("how_steps").replace("\n", "  \n"))
        st.caption(t("privacy_note"))
        return None

    llm_ok = True
    try:
        get_llm()
    except LLMError as exc:
        llm_ok = False
        st.warning(t("llm_missing", error=exc.message))

    nav_labels = {"programs": t("nav_programs"), "aois": t("nav_aois"), "diag": t("nav_diag"), "chat": t("nav_chat")}
    st.session_state.setdefault("nav", "programs")
    nav = st.segmented_control("nav", list(NAV), format_func=lambda k: f"{NAV_ICON[k]} {nav_labels[k]}", key="nav", label_visibility="collapsed") or "programs"
    if nav == "programs":
        programs_area(ctx, store, llm_ok)
    elif nav == "aois":
        aoi_area(ctx, store, llm_ok)
    elif nav == "diag":
        diag_area(ctx)
    else:
        if llm_ok:
            chat_area(ctx, store)
        else:
            st.warning(t("llm_missing", error=t("disabled_llm")))
    return ctx


def main() -> None:
    st.set_page_config(page_title="Analisador L5X", page_icon="🔎", layout="wide", initial_sidebar_state="expanded")
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
