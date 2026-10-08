import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from l5x_core import parse_file, build_xref, analyze  # noqa: E402

MINIMAL = os.path.join(ROOT, "tests", "data", "minimal.l5x")
SAMPLE = os.path.join(ROOT, "samples", "P80_HULL_HCS01.L5X")


@pytest.fixture(scope="session")
def ctrl():
    return parse_file(MINIMAL)


@pytest.fixture(scope="session")
def xref(ctrl):
    return build_xref(ctrl)


@pytest.fixture(scope="session")
def analysis(ctrl, xref):
    findings, metrics = analyze(ctrl, xref)
    return findings, metrics


@pytest.fixture(scope="session")
def findings(analysis):
    return analysis[0]


@pytest.fixture(scope="session")
def metrics(analysis):
    return analysis[1]


def by_rule(findings, rule):
    return [f for f in findings if f.rule == rule]
