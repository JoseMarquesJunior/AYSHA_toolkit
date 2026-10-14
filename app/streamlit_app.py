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
from l5x_core.aoi_config import table_csv, tables_xlsx_bytes, to_csv  # noqa: E402
from l5x_core.alarms import alarm_summary  # noqa: E402
from l5x_core.diff import DIFF_HEADER, diff_markdown, diff_rows  # noqa: E402
from l5x_core.drift import DRIFT_HEADER, LIBRARY_HEADER  # noqa: E402
from l5x_core.io_list import channel_summary  # noqa: E402
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
        "export_title": "Exportar configuração das instâncias",
        "export_help": "Lê os valores guardados em cada tag de instância (ranges, setpoints, atrasos, flags de HMI) e monta uma tabela por AOI: uma linha por instância, uma coluna por parâmetro. Sem custo.",
        "export_mode": "Membros",
        "export_mode_config": "Configuração (entradas não obrigatórias + locais editáveis pelo HMI)",
        "export_mode_all": "Todos os parâmetros e tags locais",
        "export_xlsx": "Baixar Excel com todas as AOIs",
        "export_csv": "Baixar CSV desta AOI",
        "export_preview": "{rows} instância(s) × {cols} coluna(s) nesta AOI",
        "export_none": "Nenhuma instância com dados para exportar.",
        "nav_compare": "Comparar",
        "compare_help": "Envie uma segunda exportação do mesmo projeto para ver o que mudou (lógica, tags, AOIs, configuração armazenada, findings), ou um L5X de outro PLC para comparar a biblioteca de AOIs e UDTs. Sem custo.",
        "compare_upload": "Segundo arquivo .L5X",
        "compare_none": "Envie o segundo arquivo para comparar.",
        "compare_direction": "O segundo arquivo é",
        "dir_other_old": "a versão antiga (o projeto carregado é o novo)",
        "dir_other_new": "a versão nova (o projeto carregado é o antigo)",
        "compare_summary": "{n} diferença(s): {detail}",
        "compare_empty": "Nenhuma diferença semântica entre os dois arquivos.",
        "compare_download_md": "Baixar diff (Markdown)",
        "compare_download_csv": "Baixar diff (CSV)",
        "library_title": "Drift de biblioteca entre PLCs (AOIs e UDTs)",
        "library_same": "Biblioteca idêntica nos dois arquivos.",
        "drift_title": "Drift entre instâncias ({n} membro(s) fora do padrão)",
        "drift_help": "Para cada AOI, membros de configuração em que a grande maioria das instâncias compartilha um valor e algumas divergem (ex.: uma válvula com tempo de curso diferente das demais).",
        "drift_none": "Nenhum desvio em relação à maioria.",
        "drift_download": "Baixar drift (CSV)",
        "alarms_title": "Lista de alarmes ({n})",
        "alarms_help": "Alarmes nativos (ALMD/ALMA) e instâncias de AOIs de alarme, inclusive aninhadas em outras AOIs, com severidade e configuração, para racionalização (ISA 18.2).",
        "alarm_total": "Alarmes",
        "alarm_disabled": "Desabilitados",
        "alarm_nodesc": "Sem descrição",
        "alarm_unused": "Sem uso na lógica",
        "alarms_download_xlsx": "Baixar alarmes (Excel)",
        "alarms_download_csv": "Baixar alarmes (CSV)",
        "io_title": "Instrumentos e canais de I/O ({inst} instrumentos, {ch} canais)",
        "io_help": "Instrumentos = instâncias de AOI com unidade, range ou canal de I/O ligado; canais = pontos comentados nos módulos, com uso na lógica. Ligação instrumento-canal é aproximada.",
        "io_instruments": "Instrumentos",
        "io_channels": "Canais de I/O",
        "io_summary": "{spare} spare · {unused} sem uso · {nodesc} sem descrição",
        "io_download": "Baixar instrumentos e I/O (Excel)",
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
        "export_title": "Export instance configuration",
        "export_help": "Reads the values stored in each instance tag (ranges, setpoints, delays, HMI flags) and builds one table per AOI: one row per instance, one column per parameter. Free.",
        "export_mode": "Members",
        "export_mode_config": "Configuration (non-required inputs + HMI-editable locals)",
        "export_mode_all": "All parameters and local tags",
        "export_xlsx": "Download Excel with all AOIs",
        "export_csv": "Download CSV for this AOI",
        "export_preview": "{rows} instance(s) × {cols} column(s) in this AOI",
        "export_none": "No instance with data to export.",
        "nav_compare": "Compare",
        "compare_help": "Upload a second export of the same project to see what changed (logic, tags, AOIs, stored configuration, findings), or another PLC's L5X to compare the AOI/UDT library. Free.",
        "compare_upload": "Second .L5X file",
        "compare_none": "Upload the second file to compare.",
        "compare_direction": "The second file is",
        "dir_other_old": "the old version (the loaded project is the new one)",
        "dir_other_new": "the new version (the loaded project is the old one)",
        "compare_summary": "{n} difference(s): {detail}",
        "compare_empty": "No semantic difference between the two files.",
        "compare_download_md": "Download diff (Markdown)",
        "compare_download_csv": "Download diff (CSV)",
        "library_title": "Library drift between PLCs (AOIs and UDTs)",
        "library_same": "Identical library in both files.",
        "drift_title": "Instance drift ({n} member(s) off the common value)",
        "drift_help": "For each AOI, configuration members where the large majority of instances share one value and a few deviate (e.g. one valve with a different travel time).",
        "drift_none": "No deviation from the majority.",
        "drift_download": "Download drift (CSV)",
        "alarms_title": "Alarm list ({n})",
        "alarms_help": "Native alarms (ALMD/ALMA) and alarm AOI instances, including those nested in other AOIs, with severity and configuration, for rationalization (ISA 18.2).",
        "alarm_total": "Alarms",
        "alarm_disabled": "Disabled",
        "alarm_nodesc": "No description",
        "alarm_unused": "Unused in logic",
        "alarms_download_xlsx": "Download alarms (Excel)",
        "alarms_download_csv": "Download alarms (CSV)",
        "io_title": "Instruments and I/O channels ({inst} instruments, {ch} channels)",
        "io_help": "Instruments = AOI instances with a unit, a range or a linked I/O channel; channels = commented module points with their usage. Instrument-to-channel linking is approximate.",
        "io_instruments": "Instruments",
        "io_channels": "I/O channels",
        "io_summary": "{spare} spare · {unused} unused · {nodesc} without description",
        "io_download": "Download instruments and I/O (Excel)",
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
        "export_title": "Exportar configuración de las instancias",
        "export_help": "Lee los valores guardados en cada tag de instancia (rangos, setpoints, retardos, flags de HMI) y arma una tabla por AOI: una fila por instancia, una columna por parámetro. Sin costo.",
        "export_mode": "Miembros",
        "export_mode_config": "Configuración (entradas no obligatorias + locales editables por el HMI)",
        "export_mode_all": "Todos los parámetros y tags locales",
        "export_xlsx": "Descargar Excel con todas las AOIs",
        "export_csv": "Descargar CSV de esta AOI",
        "export_preview": "{rows} instancia(s) × {cols} columna(s) en esta AOI",
        "export_none": "Ninguna instancia con datos para exportar.",
        "nav_compare": "Comparar",
        "compare_help": "Suba una segunda exportación del mismo proyecto para ver qué cambió (lógica, tags, AOIs, configuración guardada, findings), o un L5X de otro PLC para comparar la biblioteca de AOIs y UDTs. Sin costo.",
        "compare_upload": "Segundo archivo .L5X",
        "compare_none": "Suba el segundo archivo para comparar.",
        "compare_direction": "El segundo archivo es",
        "dir_other_old": "la versión antigua (el proyecto cargado es el nuevo)",
        "dir_other_new": "la versión nueva (el proyecto cargado es el antiguo)",
        "compare_summary": "{n} diferencia(s): {detail}",
        "compare_empty": "Ninguna diferencia semántica entre los dos archivos.",
        "compare_download_md": "Descargar diff (Markdown)",
        "compare_download_csv": "Descargar diff (CSV)",
        "library_title": "Drift de biblioteca entre PLCs (AOIs y UDTs)",
        "library_same": "Biblioteca idéntica en los dos archivos.",
        "drift_title": "Drift entre instancias ({n} miembro(s) fuera del patrón)",
        "drift_help": "Para cada AOI, miembros de configuración en que la gran mayoría de las instancias comparte un valor y algunas divergen (p. ej. una válvula con tiempo de carrera distinto).",
        "drift_none": "Ningún desvío respecto a la mayoría.",
        "drift_download": "Descargar drift (CSV)",
        "alarms_title": "Lista de alarmas ({n})",
        "alarms_help": "Alarmas nativas (ALMD/ALMA) e instancias de AOIs de alarma, incluso anidadas en otras AOIs, con severidad y configuración, para racionalización (ISA 18.2).",
        "alarm_total": "Alarmas",
        "alarm_disabled": "Deshabilitadas",
        "alarm_nodesc": "Sin descripción",
        "alarm_unused": "Sin uso en la lógica",
        "alarms_download_xlsx": "Descargar alarmas (Excel)",
        "alarms_download_csv": "Descargar alarmas (CSV)",
        "io_title": "Instrumentos y canales de I/O ({inst} instrumentos, {ch} canales)",
        "io_help": "Instrumentos = instancias de AOI con unidad, rango o canal de I/O ligado; canales = puntos comentados en los módulos con su uso. La ligación instrumento-canal es aproximada.",
        "io_instruments": "Instrumentos",
        "io_channels": "Canales de I/O",
        "io_summary": "{spare} spare · {unused} sin uso · {nodesc} sin descripción",
        "io_download": "Descargar instrumentos e I/O (Excel)",
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

UI["zh-CN"] = {
    "title": "罗克韦尔程序分析器 (L5X)",
    "subtitle": "解释、审查并记录以 L5X 导出的 Logix 程序。只读：不会向控制器发送任何内容。",
    "code_label": "访问码",
    "code_button": "进入",
    "code_invalid": "访问码无效。",
    "code_hint": "没有有效的访问码无法使用。请向管理员索取。",
    "logout": "退出",
    "welcome": "你好，{owner}。",
    "balance": "额度：**{balance}** / {quota}",
    "language": "回答语言",
    "upload_label": "上传 .L5X 文件（Studio 5000 中 File → Save As → L5X）",
    "parse_error": "文件无效：{error}",
    "parsing": "正在读取并分析程序…",
    "summary": "项目概览",
    "controller": "控制器",
    "revision": "版本",
    "programs": "程序",
    "routines": "例程",
    "rungs": "梯级",
    "tags": "标签",
    "findings": "按严重度统计的发现项",
    "protected_warning": "{n} 个项目受源代码保护：内容未读取也未分析。",
    "safety_warning": "该项目包含安全内容（Class=\"Safety\"）。仅限只读：应用会解释和记录，但绝不建议修改。",
    "no_project": "上传 .L5X 文件开始。",
    "program": "程序",
    "routine": "例程",
    "scope": "操作范围",
    "scope_routine": "所选例程",
    "scope_program": "所选程序",
    "scope_project": "整个项目",
    "explain": "解释",
    "review": "审查风险与质量",
    "document": "生成文档",
    "cost": "费用：{n} 额度",
    "insufficient": "额度不足：此操作需要 {cost}，你有 {balance}。",
    "routine_protected": "所选例程受源代码保护：没有可解释或审查的内容。",
    "routine_not_analyzed": "{type} 例程：解析器不分析其内部逻辑；解释仅使用名称、描述和引用的标签。",
    "deterministic": "所选范围的解析器发现项（{n}）",
    "no_findings": "此范围没有发现项。",
    "results": "结果",
    "download": "下载 Markdown",
    "chat_title": "关于程序的对话",
    "chat_placeholder": "就程序提问（每个问题 1 额度）",
    "chat_clear": "清除对话",
    "chat_sources": "查阅的例程：{sources}",
    "llm_missing": "未配置 LLM 层：{error}。上传、概览和发现项仍可使用。",
    "working": "正在查询模型…",
    "usage": "{model} · {calls} 次调用 · 输入 {inp} · 输出 {out} · 缓存 {cache} · {ms} ms",
    "refusal": "模型拒绝了部分请求；请查看回答。",
    "truncated": "回答因输出令牌上限被截断。",
    "mode_explain": "解释",
    "mode_review": "审查",
    "mode_document": "文档",
    "mode_chat": "对话",
    "mode_aoi": "AOI",
    "aoi_section": "AOI（附加指令）",
    "aoi_select": "AOI",
    "aoi_none": "该项目没有 AOI。",
    "aoi_option": "{name} 版本 {rev} · {instances} 个实例 · {calls} 次调用{protected}",
    "aoi_protected_tag": " · 受保护",
    "aoi_summary": "{params} 个参数 · {locals} 个本地标签 · 例程：{routines} · 实例：{instances} · 调用：{calls} · 发现项：{findings}",
    "aoi_report": "确定性报告（免费）：文档 + 调试指南",
    "aoi_download_report": "下载报告（Markdown）",
    "aoi_findings": "该 AOI 的解析器发现项（{n}）",
    "aoi_analyze": "用模型分析 AOI",
    "aoi_protected_msg": "受源代码保护的 AOI：解析器列出实例和调用，但没有可分析的逻辑。",
    "aoi_help": "模型分析使用 AOI 的接口、发现项和内部例程，返回用途、行为、项目中的使用、风险和调试指南。",
    "unconnected_section": "FBD 例程中未接线的标签（{routines} 个例程中的 {tags} 个标签）",
    "unconnected_help": "放在图纸上但没有任何连线的标签引用（IRef/ORef）和连接器：该标签既不驱动任何东西，也不被任何东西驱动。没有任何连线的功能块也会列出。确定性分析，免费。",
    "unconnected_none": "该项目没有未接线的 FBD 元素。",
    "unconnected_download": "下载列表（CSV）",
    "export_title": "导出实例配置",
    "export_help": "读取每个实例标签中存储的值（量程、设定值、延时、HMI 标志），为每个 AOI 生成一张表：每个实例一行，每个参数一列。免费。",
    "export_mode": "成员",
    "export_mode_config": "配置（非必需输入 + HMI 可编辑的本地标签）",
    "export_mode_all": "全部参数和本地标签",
    "export_xlsx": "下载包含所有 AOI 的 Excel",
    "export_csv": "下载该 AOI 的 CSV",
    "export_preview": "该 AOI 有 {rows} 个实例 × {cols} 列",
    "export_none": "没有可导出的实例数据。",
    "nav_compare": "比较",
    "compare_help": "上传同一项目的第二份导出以查看变化（逻辑、标签、AOI、存储的配置、发现项），或上传另一台 PLC 的 L5X 以比较 AOI/UDT 库。免费。",
    "compare_upload": "第二个 .L5X 文件",
    "compare_none": "上传第二个文件以进行比较。",
    "compare_direction": "第二个文件是",
    "dir_other_old": "旧版本（已加载的项目是新版本）",
    "dir_other_new": "新版本（已加载的项目是旧版本）",
    "compare_summary": "{n} 处差异：{detail}",
    "compare_empty": "两个文件之间没有语义差异。",
    "compare_download_md": "下载差异（Markdown）",
    "compare_download_csv": "下载差异（CSV）",
    "library_title": "PLC 之间的库漂移（AOI 和 UDT）",
    "library_same": "两个文件中的库完全相同。",
    "drift_title": "实例漂移（{n} 个成员偏离常见值）",
    "drift_help": "对每个 AOI，列出绝大多数实例共享同一值而少数实例偏离的配置成员（例如某个阀门的行程时间与其他不同）。",
    "drift_none": "没有偏离多数值的情况。",
    "drift_download": "下载漂移（CSV）",
    "alarms_title": "报警清单（{n}）",
    "alarms_help": "原生报警（ALMD/ALMA）和报警 AOI 实例（包括嵌套在其他 AOI 中的），含严重度和配置，用于报警合理化（ISA 18.2）。",
    "alarm_total": "报警",
    "alarm_disabled": "已禁用",
    "alarm_nodesc": "无描述",
    "alarm_unused": "逻辑中未使用",
    "alarms_download_xlsx": "下载报警（Excel）",
    "alarms_download_csv": "下载报警（CSV）",
    "io_title": "仪表与 I/O 通道（{inst} 个仪表，{ch} 个通道）",
    "io_help": "仪表 = 带单位、量程或已连接 I/O 通道的 AOI 实例；通道 = 模块中带注释的点及其使用情况。仪表与通道的关联为近似结果。",
    "io_instruments": "仪表",
    "io_channels": "I/O 通道",
    "io_summary": "{spare} 个备用 · {unused} 个未使用 · {nodesc} 个无描述",
    "io_download": "下载仪表与 I/O（Excel）",
    "nav_programs": "程序与例程",
    "nav_aois": "AOI",
    "nav_diag": "诊断",
    "nav_chat": "对话",
    "how_title": "工作原理",
    "how_steps": "1. 在 Studio 5000 中将项目导出为 L5X（File → Save As → L5X）。\n2. 上传文件：解析器读取程序、例程、标签和 AOI，并立即免费生成发现项。\n3. 让模型解释、审查或生成文档；每项操作仅在完成后计费。",
    "privacy_note": "文件仅存在于本次会话，解析后即删除。模型只接收结构化索引，绝不接收 XML；受保护内容不会被解码。",
    "file_section": "项目文件",
    "file_loaded": "**{name}** · {size:.1f} MB · 解析用时 {secs:.1f} 秒 · {programs} 个程序 · {routines} 个例程",
    "demo_loaded": "已从 APP_DEMO_FILE 加载演示文件。",
    "empty_title": "未加载项目",
    "target": "目标",
    "target_hint": "解释和审查接受例程或程序；文档和审查也接受整个项目。",
    "size_rungs": "{n} 个梯级",
    "size_lines": "{n} 行",
    "size_sheets": "{n} 张图纸",
    "size_na": "未分析",
    "badge_protected": "受保护",
    "badge_safety": "SAFETY",
    "badge_main": "main",
    "badge_fault": "fault",
    "badge_unscheduled": "未调度",
    "badge_disabled": "已禁用",
    "called_by": "被调用于：{names}",
    "calls_to": "调用：{names}",
    "never_called": "从未被 JSR 调用。",
    "no_description": "项目中没有描述。",
    "explain_help": "面向维护技术人员逐梯级或按功能块解释逻辑，并引用梯级/图纸。不适用于整个项目。",
    "review_help": "对解析器发现项排序，解释影响并提出通用修正建议；区分已确认事实和假设。",
    "document_help": "生成 Markdown 文档：概览、各例程描述、I/O、定时器、联锁和解析器备注。",
    "disabled_llm": "未配置 LLM 层。",
    "disabled_protected": "受保护的例程：没有内容。",
    "disabled_project": "请选择一个程序或例程。",
    "results_clear": "清除结果",
    "results_empty": "此区域操作的结果显示在这里，最新的在前。",
    "findings_title": "解析器发现项",
    "findings_sev": "严重度",
    "findings_search": "按规则、位置、标签或消息搜索",
    "findings_shown": "{shown} / {total} 个发现项",
    "findings_download": "下载发现项（CSV）",
    "col_sev": "严重度",
    "col_rule": "规则",
    "col_where": "位置",
    "col_tag": "标签",
    "col_msg": "消息",
    "col_evidence": "证据",
    "col_safety": "Safety",
    "credits_low": "额度即将用完：剩余 {balance}。",
    "credits_none": "额度已用完。请向管理员申请。",
    "disclaimer": "自动分析：任何更改前请与负责工程师核实。",
}

UI["hi"] = {
    "title": "Rockwell प्रोग्राम विश्लेषक (L5X)",
    "subtitle": "L5X में निर्यात किए गए Logix प्रोग्राम को समझाता, समीक्षा करता और दस्तावेज़ बनाता है। केवल पढ़ने के लिए: कंट्रोलर को कुछ नहीं भेजा जाता।",
    "code_label": "एक्सेस कोड",
    "code_button": "प्रवेश करें",
    "code_invalid": "कोड अमान्य है।",
    "code_hint": "मान्य कोड के बिना कुछ काम नहीं करेगा। प्रशासक से कोड माँगें।",
    "logout": "लॉग आउट",
    "welcome": "नमस्ते, {owner}।",
    "balance": "क्रेडिट: **{balance}** / {quota}",
    "language": "उत्तर की भाषा",
    "upload_label": ".L5X फ़ाइल अपलोड करें (Studio 5000 में File → Save As → L5X)",
    "parse_error": "अमान्य फ़ाइल: {error}",
    "parsing": "प्रोग्राम पढ़ा और विश्लेषित किया जा रहा है…",
    "summary": "प्रोजेक्ट सारांश",
    "controller": "कंट्रोलर",
    "revision": "रिवीज़न",
    "programs": "प्रोग्राम",
    "routines": "रूटीन",
    "rungs": "रंग (rungs)",
    "tags": "टैग",
    "findings": "गंभीरता के अनुसार निष्कर्ष",
    "protected_warning": "{n} आइटम सोर्स प्रोटेक्शन वाले हैं: सामग्री न पढ़ी गई न विश्लेषित।",
    "safety_warning": "इस प्रोजेक्ट में सुरक्षा सामग्री है (Class=\"Safety\")। यह केवल पढ़ने के लिए है: ऐप समझाता और दस्तावेज़ बनाता है, पर कभी बदलाव का सुझाव नहीं देता।",
    "no_project": "शुरू करने के लिए .L5X फ़ाइल अपलोड करें।",
    "program": "प्रोग्राम",
    "routine": "रूटीन",
    "scope": "कार्रवाई का दायरा",
    "scope_routine": "चयनित रूटीन",
    "scope_program": "चयनित प्रोग्राम",
    "scope_project": "पूरा प्रोजेक्ट",
    "explain": "समझाएँ",
    "review": "जोखिम और गुणवत्ता की समीक्षा",
    "document": "दस्तावेज़ बनाएँ",
    "cost": "लागत: {n} क्रेडिट",
    "insufficient": "क्रेडिट अपर्याप्त: इस कार्रवाई की लागत {cost} है और आपके पास {balance} हैं।",
    "routine_protected": "चयनित रूटीन सोर्स प्रोटेक्टेड है: समझाने या समीक्षा के लिए कोई सामग्री नहीं।",
    "routine_not_analyzed": "{type} रूटीन: पार्सर इसके आंतरिक लॉजिक का विश्लेषण नहीं करता; व्याख्या में केवल नाम, विवरण और संदर्भित टैग उपयोग होंगे।",
    "deterministic": "चयनित दायरे के लिए पार्सर निष्कर्ष ({n})",
    "no_findings": "इस दायरे में कोई निष्कर्ष नहीं।",
    "results": "परिणाम",
    "download": "Markdown डाउनलोड करें",
    "chat_title": "प्रोग्राम के बारे में चैट",
    "chat_placeholder": "प्रोग्राम के बारे में पूछें (प्रति प्रश्न 1 क्रेडिट)",
    "chat_clear": "बातचीत साफ़ करें",
    "chat_sources": "देखी गई रूटीन: {sources}",
    "llm_missing": "LLM परत कॉन्फ़िगर नहीं है: {error}। अपलोड, सारांश और निष्कर्ष फिर भी काम करते हैं।",
    "working": "मॉडल से पूछा जा रहा है…",
    "usage": "{model} · {calls} कॉल · इनपुट {inp} · आउटपुट {out} · कैश {cache} · {ms} ms",
    "refusal": "मॉडल ने अनुरोध का कुछ हिस्सा अस्वीकार किया; उत्तर देखें।",
    "truncated": "आउटपुट टोकन सीमा के कारण उत्तर कट गया।",
    "mode_explain": "व्याख्या",
    "mode_review": "समीक्षा",
    "mode_document": "दस्तावेज़",
    "mode_chat": "चैट",
    "mode_aoi": "AOI",
    "aoi_section": "AOI (Add-On Instructions)",
    "aoi_select": "AOI",
    "aoi_none": "इस प्रोजेक्ट में कोई AOI नहीं है।",
    "aoi_option": "{name} rev. {rev} · {instances} इंस्टेंस · {calls} कॉल{protected}",
    "aoi_protected_tag": " · प्रोटेक्टेड",
    "aoi_summary": "{params} पैरामीटर · {locals} लोकल टैग · रूटीन: {routines} · इंस्टेंस: {instances} · कॉल: {calls} · निष्कर्ष: {findings}",
    "aoi_report": "निर्धारक रिपोर्ट (निःशुल्क): दस्तावेज़ + डीबग गाइड",
    "aoi_download_report": "रिपोर्ट डाउनलोड करें (Markdown)",
    "aoi_findings": "इस AOI के पार्सर निष्कर्ष ({n})",
    "aoi_analyze": "मॉडल से AOI का विश्लेषण करें",
    "aoi_protected_msg": "सोर्स प्रोटेक्टेड AOI: पार्सर इंस्टेंस और कॉल सूचीबद्ध करता है, पर विश्लेषण के लिए कोई लॉजिक नहीं।",
    "aoi_help": "मॉडल विश्लेषण AOI के इंटरफ़ेस, निष्कर्षों और आंतरिक रूटीन का उपयोग करता है और उद्देश्य, व्यवहार, प्रोजेक्ट में उपयोग, जोखिम और डीबग गाइड देता है।",
    "unconnected_section": "FBD रूटीन में बिना तार वाले टैग ({routines} रूटीन में {tags} टैग)",
    "unconnected_help": "शीट पर रखे गए टैग संदर्भ (IRef/ORef) और कनेक्टर जिनमें कोई तार नहीं: टैग न कुछ चलाता है न किसी से चलता है। बिना तार वाले ब्लॉक भी सूचीबद्ध होते हैं। निर्धारक, निःशुल्क।",
    "unconnected_none": "इस प्रोजेक्ट में बिना तार वाला कोई FBD तत्व नहीं।",
    "unconnected_download": "सूची डाउनलोड करें (CSV)",
    "export_title": "इंस्टेंस कॉन्फ़िगरेशन निर्यात करें",
    "export_help": "हर इंस्टेंस टैग में संग्रहीत मान (रेंज, सेटपॉइंट, विलंब, HMI फ़्लैग) पढ़कर प्रति AOI एक तालिका बनाता है: प्रति इंस्टेंस एक पंक्ति, प्रति पैरामीटर एक कॉलम। निःशुल्क।",
    "export_mode": "सदस्य",
    "export_mode_config": "कॉन्फ़िगरेशन (गैर-अनिवार्य इनपुट + HMI से संपादन योग्य लोकल)",
    "export_mode_all": "सभी पैरामीटर और लोकल टैग",
    "export_xlsx": "सभी AOI के साथ Excel डाउनलोड करें",
    "export_csv": "इस AOI का CSV डाउनलोड करें",
    "export_preview": "इस AOI में {rows} इंस्टेंस × {cols} कॉलम",
    "export_none": "निर्यात के लिए डेटा वाला कोई इंस्टेंस नहीं।",
    "nav_compare": "तुलना",
    "compare_help": "बदलाव देखने के लिए उसी प्रोजेक्ट का दूसरा निर्यात अपलोड करें (लॉजिक, टैग, AOI, संग्रहीत कॉन्फ़िगरेशन, निष्कर्ष), या AOI/UDT लाइब्रेरी की तुलना के लिए दूसरे PLC का L5X। निःशुल्क।",
    "compare_upload": "दूसरी .L5X फ़ाइल",
    "compare_none": "तुलना के लिए दूसरी फ़ाइल अपलोड करें।",
    "compare_direction": "दूसरी फ़ाइल है",
    "dir_other_old": "पुराना संस्करण (लोड किया गया प्रोजेक्ट नया है)",
    "dir_other_new": "नया संस्करण (लोड किया गया प्रोजेक्ट पुराना है)",
    "compare_summary": "{n} अंतर: {detail}",
    "compare_empty": "दोनों फ़ाइलों में कोई सिमैंटिक अंतर नहीं।",
    "compare_download_md": "diff डाउनलोड करें (Markdown)",
    "compare_download_csv": "diff डाउनलोड करें (CSV)",
    "library_title": "PLC के बीच लाइब्रेरी ड्रिफ़्ट (AOI और UDT)",
    "library_same": "दोनों फ़ाइलों में लाइब्रेरी समान है।",
    "drift_title": "इंस्टेंस ड्रिफ़्ट ({n} सदस्य सामान्य मान से भिन्न)",
    "drift_help": "हर AOI के लिए, वे कॉन्फ़िगरेशन सदस्य जहाँ अधिकांश इंस्टेंस एक ही मान साझा करते हैं और कुछ भिन्न हैं (जैसे एक वाल्व का ट्रैवल टाइम बाकी से अलग)।",
    "drift_none": "बहुमत से कोई विचलन नहीं।",
    "drift_download": "ड्रिफ़्ट डाउनलोड करें (CSV)",
    "alarms_title": "अलार्म सूची ({n})",
    "alarms_help": "नेटिव अलार्म (ALMD/ALMA) और अलार्म AOI इंस्टेंस, जिनमें अन्य AOI के भीतर नेस्टेड भी शामिल हैं, गंभीरता और कॉन्फ़िगरेशन के साथ, युक्तिकरण (ISA 18.2) के लिए।",
    "alarm_total": "अलार्म",
    "alarm_disabled": "अक्षम",
    "alarm_nodesc": "बिना विवरण",
    "alarm_unused": "लॉजिक में अप्रयुक्त",
    "alarms_download_xlsx": "अलार्म डाउनलोड करें (Excel)",
    "alarms_download_csv": "अलार्म डाउनलोड करें (CSV)",
    "io_title": "उपकरण और I/O चैनल ({inst} उपकरण, {ch} चैनल)",
    "io_help": "उपकरण = यूनिट, रेंज या जुड़े I/O चैनल वाले AOI इंस्टेंस; चैनल = मॉड्यूल के टिप्पणी वाले बिंदु और उनका उपयोग। उपकरण-चैनल लिंकिंग अनुमानित है।",
    "io_instruments": "उपकरण",
    "io_channels": "I/O चैनल",
    "io_summary": "{spare} स्पेयर · {unused} अप्रयुक्त · {nodesc} बिना विवरण",
    "io_download": "उपकरण और I/O डाउनलोड करें (Excel)",
    "nav_programs": "प्रोग्राम और रूटीन",
    "nav_aois": "AOI",
    "nav_diag": "निदान",
    "nav_chat": "चैट",
    "how_title": "यह कैसे काम करता है",
    "how_steps": "1. Studio 5000 से प्रोजेक्ट को L5X में निर्यात करें (File → Save As → L5X)।\n2. फ़ाइल अपलोड करें: पार्सर प्रोग्राम, रूटीन, टैग और AOI पढ़ता है और तुरंत, निःशुल्क निष्कर्ष बनाता है।\n3. मॉडल से समझाने, समीक्षा या दस्तावेज़ बनाने को कहें; हर कार्रवाई पूरी होने पर ही क्रेडिट कटता है।",
    "privacy_note": "फ़ाइल केवल इस सत्र में रहती है और पढ़ने के बाद हटा दी जाती है। मॉडल को केवल संरचित इंडेक्स मिलता है, कभी XML नहीं; प्रोटेक्टेड सामग्री डिकोड नहीं होती।",
    "file_section": "प्रोजेक्ट फ़ाइल",
    "file_loaded": "**{name}** · {size:.1f} MB · {secs:.1f} s में पढ़ी गई · {programs} प्रोग्राम · {routines} रूटीन",
    "demo_loaded": "APP_DEMO_FILE से डेमो फ़ाइल लोड हुई।",
    "empty_title": "कोई प्रोजेक्ट लोड नहीं",
    "target": "लक्ष्य",
    "target_hint": "समझाएँ और समीक्षा रूटीन या प्रोग्राम लेते हैं; दस्तावेज़ और समीक्षा पूरा प्रोजेक्ट भी लेते हैं।",
    "size_rungs": "{n} rungs",
    "size_lines": "{n} पंक्तियाँ",
    "size_sheets": "{n} शीट",
    "size_na": "विश्लेषित नहीं",
    "badge_protected": "प्रोटेक्टेड",
    "badge_safety": "SAFETY",
    "badge_main": "main",
    "badge_fault": "fault",
    "badge_unscheduled": "अनिर्धारित",
    "badge_disabled": "अक्षम",
    "called_by": "इनके द्वारा कॉल: {names}",
    "calls_to": "कॉल करता है: {names}",
    "never_called": "किसी JSR द्वारा कभी कॉल नहीं।",
    "no_description": "प्रोजेक्ट में कोई विवरण नहीं।",
    "explain_help": "रखरखाव तकनीशियन के लिए लॉजिक को rung दर rung या ब्लॉक के अनुसार समझाता है, rung/शीट का हवाला देते हुए। पूरे प्रोजेक्ट के लिए उपलब्ध नहीं।",
    "review_help": "पार्सर निष्कर्षों को प्राथमिकता देता है, प्रभाव समझाता है और सामान्य सुधार सुझाता है; पुष्ट तथ्यों को परिकल्पनाओं से अलग रखता है।",
    "document_help": "Markdown दस्तावेज़ बनाता है: अवलोकन, प्रति रूटीन विवरण, I/O, टाइमर, इंटरलॉक और पार्सर टिप्पणियाँ।",
    "disabled_llm": "LLM परत कॉन्फ़िगर नहीं है।",
    "disabled_protected": "प्रोटेक्टेड रूटीन: कोई सामग्री नहीं।",
    "disabled_project": "कोई प्रोग्राम या रूटीन चुनें।",
    "results_clear": "परिणाम साफ़ करें",
    "results_empty": "इस क्षेत्र की कार्रवाइयों के परिणाम यहाँ दिखेंगे, सबसे नया पहले।",
    "findings_title": "पार्सर निष्कर्ष",
    "findings_sev": "गंभीरता",
    "findings_search": "नियम, स्थान, टैग या संदेश से खोजें",
    "findings_shown": "{total} में से {shown} निष्कर्ष",
    "findings_download": "निष्कर्ष डाउनलोड करें (CSV)",
    "col_sev": "गंभ.",
    "col_rule": "नियम",
    "col_where": "स्थान",
    "col_tag": "टैग",
    "col_msg": "संदेश",
    "col_evidence": "प्रमाण",
    "col_safety": "Safety",
    "credits_low": "क्रेडिट कम हो रहे हैं: {balance} शेष।",
    "credits_none": "क्रेडिट समाप्त। प्रशासक से और माँगें।",
    "disclaimer": "स्वचालित विश्लेषण: किसी भी बदलाव से पहले ज़िम्मेदार इंजीनियर से सत्यापित करें।",
}

LANG_LABELS = {"pt-BR": "Português (BR)", "en": "English", "es": "Español", "zh-CN": "中文（简体）", "hi": "हिन्दी"}
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
NAV = ("programs", "aois", "diag", "compare", "chat")
NAV_ICON = {"programs": "📘", "aois": "🧩", "diag": "🩺", "compare": "🔀", "chat": "💬"}
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
    drift_area(ctx)
    alarms_area(ctx)
    io_area(ctx)


def drift_area(ctx: ProjectContext) -> None:
    rows = ctx.drift_rows()
    with st.expander(t("drift_title", n=len(rows)), expanded=False):
        st.caption(t("drift_help"))
        if not rows:
            st.caption(t("drift_none"))
            return
        st.dataframe(rows[:500], width="stretch", hide_index=True)
        st.download_button(t("drift_download"), data=table_csv(DRIFT_HEADER, rows).encode("utf-8-sig"), file_name=f"{ctx.ctrl.name}_drift.csv".replace(" ", "_"), mime="text/csv", key="dl_drift")


def alarms_area(ctx: ProjectContext) -> None:
    header, rows = ctx.alarm_table()
    with st.expander(t("alarms_title", n=len(rows)), expanded=False):
        st.caption(t("alarms_help"))
        s = alarm_summary(rows)
        c = st.columns(4)
        c[0].metric(t("alarm_total"), s["total"])
        c[1].metric(t("alarm_disabled"), s["desabilitados"])
        c[2].metric(t("alarm_nodesc"), s["sem_descricao"])
        c[3].metric(t("alarm_unused"), s["sem_uso"])
        if s["por_severidade"]:
            st.caption(" · ".join(f"{k}: {v}" for k, v in sorted(s["por_severidade"].items())))
        if rows:
            st.dataframe([{k: r.get(k, "") for k in header} for r in rows[:1000]], width="stretch", hide_index=True)
            c1, c2 = st.columns(2)
            c1.download_button(t("alarms_download_csv"), data=table_csv(header, rows).encode("utf-8-sig"), file_name=f"{ctx.ctrl.name}_alarmes.csv".replace(" ", "_"), mime="text/csv", key="dl_alm_csv")
            c2.download_button(t("alarms_download_xlsx"), data=tables_xlsx_bytes([("Alarmes", header, rows)]), file_name=f"{ctx.ctrl.name}_alarmes.xlsx".replace(" ", "_"),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_alm_xlsx")


def io_area(ctx: ProjectContext) -> None:
    tables = ctx.io_tables()
    inst, chans = tables[0][2], tables[1][2]
    with st.expander(t("io_title", inst=len(inst), ch=len(chans)), expanded=False):
        st.caption(t("io_help"))
        s = channel_summary(chans)
        st.caption(t("io_summary", spare=s["spare"], unused=s["sem_uso"], nodesc=s["sem_descricao"]))
        st.markdown(f"**{t('io_instruments')}**")
        st.dataframe(inst[:1000], width="stretch", hide_index=True)
        st.markdown(f"**{t('io_channels')}**")
        st.dataframe(chans[:1000], width="stretch", hide_index=True)
        st.download_button(t("io_download"), data=tables_xlsx_bytes(tables), file_name=f"{ctx.ctrl.name}_instrumentos_io.xlsx".replace(" ", "_"),
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_io")


def compare_area(ctx: ProjectContext) -> None:
    st.caption(t("compare_help"))
    other_file = st.file_uploader(t("compare_upload"), type=["l5x", "L5X"], accept_multiple_files=False, key="upload_other")
    other: Optional[ProjectContext] = None
    if other_file is not None:
        data = other_file.getvalue()
        try:
            with st.spinner(t("parsing")):
                other = load_project(file_hash(data), data, other_file.name)
        except L5XError as exc:
            st.error(t("parse_error", error=exc))
            return
    elif os.environ.get("APP_DEMO_OTHER") and os.path.isfile(os.environ["APP_DEMO_OTHER"]):
        with open(os.environ["APP_DEMO_OTHER"], "rb") as fh:
            data = fh.read()
        other = load_project(file_hash(data), data, os.path.basename(os.environ["APP_DEMO_OTHER"]))
    if other is None:
        st.info(t("compare_none"))
        return
    dirs = {"other_old": t("dir_other_old"), "other_new": t("dir_other_new")}
    direction = st.radio(t("compare_direction"), list(dirs), format_func=lambda k: dirs[k], horizontal=True, key="cmp_dir")
    old, new = (other, ctx) if direction == "other_old" else (ctx, other)
    key = (st.session_state.get("project_hash"), other.ctrl.name, other.metrics.routines, other.metrics.tags_controller, direction)
    cache = st.session_state.setdefault("diff_cache", {})
    if cache.get("key") != key:
        with st.spinner(t("parsing")):
            cache["key"], cache["diff"] = key, old.diff_to(new)
    d = cache["diff"]
    st.markdown(f"**{old.ctrl.name}** ({old.ctrl.export_date or '?'}) → **{new.ctrl.name}** ({new.ctrl.export_date or '?'})")
    if d.empty:
        st.success(t("compare_empty"))
    else:
        counts = d.counts()
        st.caption(t("compare_summary", n=len(d.changes), detail=", ".join(f"{k}: {v}" for k, v in counts.items())))
        rows = diff_rows(d)
        cats = st.pills("categorias", list(counts), selection_mode="multi", default=list(counts), key="cmp_cats", label_visibility="collapsed") or list(counts)
        st.dataframe([r for r in rows if r["Categoria"] in cats][:2000], width="stretch", hide_index=True)
        c1, c2 = st.columns(2)
        c1.download_button(t("compare_download_md"), data=diff_markdown(d).encode("utf-8"), file_name=f"diff_{old.ctrl.name}_{new.ctrl.name}.md".replace(" ", "_"), mime="text/markdown", key="dl_diff_md")
        c2.download_button(t("compare_download_csv"), data=table_csv(DIFF_HEADER, rows).encode("utf-8-sig"), file_name=f"diff_{old.ctrl.name}_{new.ctrl.name}.csv".replace(" ", "_"), mime="text/csv", key="dl_diff_csv")
    st.subheader(t("library_title"))
    lib = [r for r in ctx.library_drift_with([other]) if r["Status"] != "igual"]
    if lib:
        st.dataframe(lib, width="stretch", hide_index=True)
        st.download_button(t("drift_download"), data=table_csv(LIBRARY_HEADER, lib).encode("utf-8-sig"), file_name="library_drift.csv", mime="text/csv", key="dl_lib")
    else:
        st.caption(t("library_same"))


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
    export_area(ctx, name)
    st.divider()
    results_panel(ctx, AREA_MODES["aois"])


def export_area(ctx: ProjectContext, name: str) -> None:
    """Configuração das instâncias de AOI: Excel do projeto inteiro ou CSV da AOI escolhida."""
    with st.expander(t("export_title"), expanded=False):
        st.caption(t("export_help"))
        modes = {"config": t("export_mode_config"), "all": t("export_mode_all")}
        mode = st.radio(t("export_mode"), list(modes), format_func=lambda k: modes[k], key="export_mode", horizontal=True)
        sheet = ctx.config_sheet(name, mode)
        if sheet is None:
            st.caption(t("export_none"))
        else:
            st.caption(t("export_preview", rows=len(sheet.rows), cols=len(sheet.columns)))
            st.dataframe([{k: r[k] for k in sheet.header} for r in sheet.rows[:200]], width="stretch", hide_index=True)
        c1, c2 = st.columns(2)
        if sheet is not None:
            c1.download_button(t("export_csv"), data=to_csv(sheet).encode("utf-8-sig"), file_name=f"{ctx.ctrl.name}_{name}_{mode}.csv".replace(" ", "_"), mime="text/csv", key="dl_cfg_csv")
        c2.download_button(t("export_xlsx"), data=ctx.config_xlsx(mode), file_name=f"{ctx.ctrl.name}_config_AOIs_{mode}.xlsx".replace(" ", "_"),
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_cfg_xlsx")


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

    nav_labels = {"programs": t("nav_programs"), "aois": t("nav_aois"), "diag": t("nav_diag"), "compare": t("nav_compare"), "chat": t("nav_chat")}
    st.session_state.setdefault("nav", "programs")
    nav = st.segmented_control("nav", list(NAV), format_func=lambda k: f"{NAV_ICON[k]} {nav_labels[k]}", key="nav", label_visibility="collapsed") or "programs"
    if nav == "programs":
        programs_area(ctx, store, llm_ok)
    elif nav == "aois":
        aoi_area(ctx, store, llm_ok)
    elif nav == "diag":
        diag_area(ctx)
    elif nav == "compare":
        compare_area(ctx)
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
