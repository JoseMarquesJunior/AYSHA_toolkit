"""Command line front-end.

    python -m l5x_core.cli file.L5X --summary --findings
    python -m l5x_core.cli file.L5X --dump-routine Programa/Rotina
    python -m l5x_core.cli file.L5X --dump-routine aoi:NomeDaAoi/Logic
    python -m l5x_core.cli file.L5X --aois                       # tabela de AOIs
    python -m l5x_core.cli file.L5X --aoi NomeDaAoi              # relatório Markdown (documentação + debug)
    python -m l5x_core.cli file.L5X --aoi NomeDaAoi --aoi-text   # texto da AOI como o LLM recebe
    python -m l5x_core.cli file.L5X --json saida.json
"""

from __future__ import annotations

import argparse
import sys
import time
from typing import List

from .analysis import analyze
from .aoi import aoi_markdown, aoi_summary_rows, build_aoi_report
from .model import Finding, SEVERITIES
from .parser import parse_file, L5XError
from .serialize import aoi_text, project_overview, routine_text, tag_sheet, to_json
from .xref import build_xref


def _print_findings(findings: List[Finding], min_severity: str, rules: List[str]) -> None:
    order = {s: i for i, s in enumerate(SEVERITIES)}
    limit = order.get(min_severity, 2)
    rows = [f for f in findings if order.get(f.severity, 9) <= limit and (not rules or f.rule in rules)]
    print(f"\n{len(rows)} findings (min severity: {min_severity})")
    print(f"{'sev':<8} {'regra':<24} {'local':<60} mensagem")
    print("-" * 120)
    for f in rows:
        loc = f.location()
        if len(loc) > 58:
            loc = loc[:57] + "…"
        safety = " [SAFETY]" if f.safety else ""
        print(f"{f.severity:<8} {f.rule:<24} {loc:<60} {f.message}{safety}")
        if f.evidence:
            print(f"{'':<8} {'':<24} {'':<60}   -> {f.evidence}")


def main(argv: List[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="l5x_core", description="Analisador determinístico de L5X (Rockwell/Allen-Bradley).")
    ap.add_argument("file", help="arquivo .L5X")
    ap.add_argument("--summary", action="store_true", help="visão geral do projeto (texto para o LLM)")
    ap.add_argument("--findings", action="store_true", help="tabela de findings")
    ap.add_argument("--min-severity", default="info", choices=SEVERITIES, help="severidade mínima para --findings")
    ap.add_argument("--rule", action="append", default=[], help="filtra --findings por regra (repetível)")
    ap.add_argument("--dump-routine", metavar="PROGRAMA/ROTINA", help="texto de uma rotina (use aoi:NOME/ROTINA para rotinas de AOI)")
    ap.add_argument("--tags", action="store_true", help="tabela de tags com contagem de leituras/escritas")
    ap.add_argument("--aois", action="store_true", help="tabela das AOIs: parâmetros, rotinas, instâncias, chamadas e findings")
    ap.add_argument("--aoi", metavar="NOME", help="relatório Markdown de uma AOI (documentação + roteiro de debug)")
    ap.add_argument("--aoi-text", action="store_true", help="com --aoi: imprime o texto da AOI como o LLM recebe, em vez do relatório")
    ap.add_argument("--json", metavar="SAIDA.json", help="grava modelo + findings + métricas em JSON")
    args = ap.parse_args(argv)
    for stream in (sys.stdout, sys.stderr):  # Windows consoles default to cp1252
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):  # pragma: no cover
            pass

    t0 = time.perf_counter()
    try:
        ctrl = parse_file(args.file)
    except L5XError as exc:
        print(f"erro: {exc}", file=sys.stderr)
        return 2
    t1 = time.perf_counter()
    xref = build_xref(ctrl)
    findings, metrics = analyze(ctrl, xref)
    t2 = time.perf_counter()
    print(f"{ctrl.name}: parse {t1 - t0:.2f}s, xref+análise {t2 - t1:.2f}s", file=sys.stderr)

    if args.summary:
        ov = project_overview(ctrl, findings, metrics)
        print(ov.text)
        print(f"\n[~{ov.tokens} tokens{', truncado' if ov.truncated else ''}]", file=sys.stderr)
    if args.findings:
        _print_findings(findings, args.min_severity, args.rule)
    if args.dump_routine:
        if "/" not in args.dump_routine:
            print("erro: use PROGRAMA/ROTINA", file=sys.stderr)
            return 2
        prog, rt = args.dump_routine.split("/", 1)
        try:
            sr = routine_text(ctrl, prog, rt, xref)
        except KeyError as exc:
            print(f"erro: {exc}", file=sys.stderr)
            return 2
        print(sr.text)
        print(f"\n[~{sr.tokens} tokens]", file=sys.stderr)
    if args.tags:
        print(tag_sheet(ctrl, xref).text)
    if args.aois:
        rows = aoi_summary_rows(ctrl, xref, findings)
        print(f"\n{len(rows)} AOIs")
        print(f"{'aoi':<28} {'rev':<7} {'prot':<5} {'par':>4} {'loc':>4} {'inst':>5} {'cham':>5} {'err':>4} {'warn':>4} {'info':>4}  rotinas")
        print("-" * 120)
        for r in rows:
            f = r["findings"]
            print(f"{r['aoi'][:28]:<28} {r['revision'][:7]:<7} {'sim' if r['protected'] else '':<5} {r['parameters']:>4} {r['locals']:>4} {r['instances']:>5} {r['calls']:>5} {f.get('error', 0):>4} {f.get('warning', 0):>4} {f.get('info', 0):>4}  {r['routines']}")
    if args.aoi:
        try:
            report = build_aoi_report(ctrl, args.aoi, xref, findings)
        except KeyError as exc:
            print(f"erro: {exc}", file=sys.stderr)
            return 2
        if args.aoi_text:
            sr = aoi_text(ctrl, args.aoi, xref, report)
            print(sr.text)
            print(f"\n[~{sr.tokens} tokens]", file=sys.stderr)
        else:
            print(aoi_markdown(report))
    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            fh.write(to_json(ctrl, findings, metrics))
        print(f"JSON gravado em {args.json}", file=sys.stderr)
    if not any((args.summary, args.findings, args.dump_routine, args.tags, args.aois, args.aoi, args.json)):
        ov = project_overview(ctrl, findings, metrics)
        print(ov.text)
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
