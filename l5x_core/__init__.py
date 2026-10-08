"""l5x_core: deterministic parser and analysis for Rockwell L5X exports.

UI-independent. Nothing here talks to an LLM; see serialize.py for the
compact index handed to the model layer.
"""

from .model import Controller, Finding, Metrics
from .parser import parse_file, parse_bytes, L5XError
from .xref import build_xref, CrossRef
from .analysis import analyze

__all__ = [
    "Controller",
    "Finding",
    "Metrics",
    "parse_file",
    "parse_bytes",
    "L5XError",
    "build_xref",
    "CrossRef",
    "analyze",
]
