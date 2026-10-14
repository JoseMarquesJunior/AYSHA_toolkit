"""Command line front-end.

    python -m l5x_core.cli file.L5X --summary --findings
    python -m l5x_core.cli file.L5X --dump-routine Programa/Rotina
    python -m l5x_core.cli file.L5X --dump-routine aoi:NomeDaAoi/Logic
    python -m l5x_core.cli file.L5X --aois                       # tabela de AOIs
    python -m l5x_core.cli file.L5X --aoi NomeDaAoi              # relatório Markdown (documentação + debug)
    python -m l5x_core.cli file.L5X --aoi NomeDaAoi --aoi-text   # texto da AOI como o LLM recebe
    python -m l5x_core.cli file.L5X --unconnected                # tags soltas (sem fio) em rotinas FBD
    python -m l5x_core.cli file.L5X --export-config cfg.xlsx     # configuração das instâncias de AOI (aba por AOI)
    python -m l5x_core.cli file.L5X --aoi AI_TYPE4 --export-config ai4.csv --config-mode all
    python -m l5x_core.cli file.L5X --alarms --export-alarms alarmes.xlsx   # lista de alarmes (ISA 18.2)
    python -m l5x_core.cli file.L5X --io-list --export-io io.xlsx           # instrumentos e canais de I/O
    python -m l5x_core.cli file.L5X --drift                                  # instâncias fora do padrão
    python -m l5x_core.cli file.L5X --library-drift outro_plc.L5X            # AOIs/UDTs entre PLCs
    python -m l5x_core.cli antigo.L5X --diff novo.L5X --export-diff diff.csv # diff semântico
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
    ap.add_argument("--unconnected", action="store_true", help="tabela de elementos FBD sem fio (tags soltas na folha), por rotina")
    ap.add_argument("--export-config", metavar="SAIDA.xlsx", help="exporta a configuração das instâncias de AOI (uma aba por AOI; .csv exporta só a AOI de --aoi)")
    ap.add_argument("--config-mode", default="config", choices=("config", "all"), help="config = entradas não obrigatórias + locais editáveis pelo HMI; all = todos os membros")
    ap.add_argument("--alarms", action="store_true", help="lista de alarmes (ALMD/ALMA e AOIs de alarme, inclusive aninhadas) com resumo ISA 18.2")
    ap.add_argument("--export-alarms", metavar="SAIDA.xlsx|csv", help="grava a lista de alarmes")
    ap.add_argument("--io-list", action="store_true", help="lista de instrumentos (instâncias de AOI com unidade/range/canal) e canais de I/O comentados")
    ap.add_argument("--export-io", metavar="SAIDA.xlsx", help="grava instrumentos e canais de I/O (duas abas)")
    ap.add_argument("--drift", action="store_true", help="drift entre instâncias: membros de configuração fora do valor da maioria")
    ap.add_argument("--library-drift", metavar="OUTRO.L5X", nargs="+", help="compara AOIs e UDTs deste projeto com outros L5X (revisão, interface, ausência)")
    ap.add_argument("--diff", metavar="NOVO.L5X", help="diff semântico: este arquivo (antigo) -> NOVO.L5X")
    ap.add_argument("--export-diff", metavar="SAIDA.csv|md", help="com --diff: grava as diferenças")
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
    if args.unconnected:
        rows = [(scope, r.name, d) for scope, routines in [(p.name, p.routines) for p in ctrl.programs] + [(f"aoi:{a.name}", a.routines) for a in ctrl.aois] for r in routines for d in r.dangling]
        routines_hit = {(s, r) for s, r, _ in rows}
        print(f"\n{len(rows)} elemento(s) FBD sem fio em {len(routines_hit)} rotina(s)")
        print(f"{'programa':<28} {'rotina':<28} {'folha':>5} {'tipo':<18} elemento")
        print("-" * 110)
        for scope, name, d in rows:
            print(f"{scope[:28]:<28} {name[:28]:<28} {d.sheet:>5} {d.kind:<18} {d.operand}")
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
    if args.export_config:
        from .aoi_config import build_sheets, to_csv, write_xlsx

        sheets = build_sheets(ctrl, xref, mode=args.config_mode, aoi=args.aoi)
        if not sheets:
            print("erro: nenhuma instância de AOI com dados para exportar", file=sys.stderr)
            return 2
        if args.export_config.lower().endswith(".csv"):
            if len(sheets) != 1:
                print("erro: .csv exporta uma única AOI; use --aoi NOME ou grave .xlsx", file=sys.stderr)
                return 2
            with open(args.export_config, "w", encoding="utf-8-sig", newline="") as fh:
                fh.write(to_csv(sheets[0]))
        else:
            write_xlsx(sheets, args.export_config, args.config_mode)
        print(f"Configuração exportada em {args.export_config}: " + ", ".join(f"{s.aoi} ({len(s.rows)} instâncias, {len(s.columns)} colunas)" for s in sheets), file=sys.stderr)
    if args.alarms or args.export_alarms:
        from .alarms import alarm_markdown, alarm_rows
        from .aoi_config import table_csv, write_tables_xlsx

        header, rows = alarm_rows(ctrl, xref)
        if args.alarms:
            print(alarm_markdown(header, rows))
        if args.export_alarms:
            if args.export_alarms.lower().endswith(".csv"):
                with open(args.export_alarms, "w", encoding="utf-8-sig", newline="") as fh:
                    fh.write(table_csv(header, rows))
            else:
                write_tables_xlsx([("Alarmes", header, rows)], args.export_alarms)
            print(f"Alarmes gravados em {args.export_alarms} ({len(rows)} linhas)", file=sys.stderr)
    if args.io_list or args.export_io:
        from .aoi_config import write_tables_xlsx
        from .io_list import channel_summary, io_tables

        tables = io_tables(ctrl, xref)
        inst, chans = tables[0][2], tables[1][2]
        if args.io_list:
            s = channel_summary(chans)
            print(f"\n{len(inst)} instrumento(s); {s['canais']} canal(is) de I/O em {s['modulos']} módulo(s): {s['spare']} spare, {s['sem_uso']} sem uso, {s['sem_descricao']} sem descrição")
            print(f"{'tag':<24} {'tipo':<16} {'unid':<6} {'range':<16} {'entrada I/O':<32} descrição")
            print("-" * 120)
            for r in inst:
                print(f"{r['Tag'][:24]:<24} {r['Tipo'][:16]:<16} {r['Unidade'][:6]:<6} {str(r['RangeMin']) + '..' + str(r['RangeMax']):<16} {r['EntradaIO'][:32]:<32} {r['Descricao'][:40]}")
            print(f"\n{'canal':<34} {'módulo':<24} {'uso':>3}  descrição")
            print("-" * 120)
            for r in chans:
                print(f"{r['Canal'][:34]:<34} {r['Modulo'][:24]:<24} {r['Usos']:>3}  {r['Descricao'][:50]}{' [SPARE]' if r['Spare'] else ''}")
        if args.export_io:
            write_tables_xlsx(tables, args.export_io)
            print(f"Instrumentos e I/O gravados em {args.export_io}", file=sys.stderr)
    if args.drift:
        from .drift import drift_markdown, instance_drift

        print(drift_markdown(instance_drift(ctrl, xref)))
    if args.library_drift:
        from .drift import library_drift, library_markdown

        others = []
        for path in args.library_drift:
            try:
                others.append(parse_file(path))
            except L5XError as exc:
                print(f"erro: {exc}", file=sys.stderr)
                return 2
        print(library_markdown(library_drift([ctrl] + others), [ctrl.name] + [o.name for o in others]))
    if args.diff:
        from .aoi_config import table_csv
        from .diff import DIFF_HEADER, diff_markdown, diff_projects, diff_rows

        try:
            new_ctrl = parse_file(args.diff)
        except L5XError as exc:
            print(f"erro: {exc}", file=sys.stderr)
            return 2
        new_findings, _ = analyze(new_ctrl, build_xref(new_ctrl))
        d = diff_projects(ctrl, new_ctrl, findings, new_findings)
        md = diff_markdown(d)
        if args.export_diff:
            with open(args.export_diff, "w", encoding="utf-8-sig" if args.export_diff.lower().endswith(".csv") else "utf-8", newline="") as fh:
                fh.write(table_csv(DIFF_HEADER, diff_rows(d)) if args.export_diff.lower().endswith(".csv") else md)
            print(f"Diff gravado em {args.export_diff} ({len(d.changes)} diferenças)", file=sys.stderr)
        else:
            print(md)
    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            fh.write(to_json(ctrl, findings, metrics))
        print(f"JSON gravado em {args.json}", file=sys.stderr)
    if not any((args.summary, args.findings, args.dump_routine, args.tags, args.aois, args.aoi, args.unconnected, args.export_config, args.alarms, args.export_alarms, args.io_list, args.export_io, args.drift, args.library_drift, args.diff, args.json)):
        ov = project_overview(ctrl, findings, metrics)
        print(ov.text)
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
