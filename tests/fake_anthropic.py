"""Cliente Anthropic falso: grava os parâmetros de cada chamada e devolve respostas
canônicas, sem rede. Imita só o que app.llm usa: messages.stream(...) como
context manager com .text_stream e .get_final_message()."""

from __future__ import annotations

from types import SimpleNamespace
from typing import Any, Callable, Dict, List, Optional


class _Stream:
    def __init__(self, text: str, stop_reason: str, usage: Dict[str, int], explanation: Optional[str] = None):
        self._text = text
        self._stop = stop_reason
        self._usage = usage
        self._explanation = explanation

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    @property
    def text_stream(self):
        # entrega em pedaços para exercitar o on_text
        step = max(1, len(self._text) // 3)
        for i in range(0, len(self._text), step):
            yield self._text[i:i + step]

    def get_final_message(self):
        details = SimpleNamespace(type="refusal", category="other", explanation=self._explanation) if self._stop == "refusal" else None
        return SimpleNamespace(
            content=[SimpleNamespace(type="text", text=self._text)],
            stop_reason=self._stop,
            stop_details=details,
            usage=SimpleNamespace(
                input_tokens=self._usage.get("input_tokens", 100),
                output_tokens=self._usage.get("output_tokens", 20),
                cache_read_input_tokens=self._usage.get("cache_read_input_tokens", 0),
                cache_creation_input_tokens=self._usage.get("cache_creation_input_tokens", 0),
            ),
            model="fake-model",
        )


class FakeMessages:
    def __init__(self, responder: Optional[Callable[[Dict[str, Any], int], Any]] = None):
        self.calls: List[Dict[str, Any]] = []
        self.responder = responder

    def stream(self, **params):
        self.calls.append(params)
        n = len(self.calls)
        if self.responder:
            out = self.responder(params, n)
            if isinstance(out, Exception):
                raise out
            if isinstance(out, _Stream):
                return out
            if isinstance(out, dict):
                return _Stream(out.get("text", ""), out.get("stop_reason", "end_turn"), out.get("usage", {}), out.get("explanation"))
            return _Stream(str(out), "end_turn", {})
        return _Stream(f"resposta {n}", "end_turn", {"cache_read_input_tokens": 0 if n == 1 else 500, "cache_creation_input_tokens": 500 if n == 1 else 0})


class FakeAnthropic:
    def __init__(self, responder=None):
        self.messages = FakeMessages(responder)

    @property
    def calls(self):
        return self.messages.calls
