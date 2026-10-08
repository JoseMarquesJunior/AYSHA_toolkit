"""Códigos de acesso e cotas em SQLite.

Tabela: codes(code TEXT PRIMARY KEY, owner TEXT, quota INTEGER, used INTEGER).

    python -m app.credits add CODIGO --owner nome --quota 20
    python -m app.credits list
    python -m app.credits reset CODIGO
    python -m app.credits remove CODIGO

Caminho do banco: variável CREDITS_DB (padrão ``credits.db``). Em hospedagem
efêmera (Streamlit Community Cloud) use a variável/secret ACCESS_CODES com
``codigo:dono:cota,codigo2:dono2:cota2`` para semear o banco na inicialização.
"""

from __future__ import annotations

import argparse
import os
import sqlite3
import sys
from dataclasses import dataclass
from typing import Iterable, List, Optional

DEFAULT_DB = "credits.db"


@dataclass
class Code:
    code: str
    owner: str
    quota: int
    used: int

    @property
    def balance(self) -> int:
        return max(0, self.quota - self.used)


class CreditStore:
    def __init__(self, path: Optional[str] = None) -> None:
        self.path = path or os.environ.get("CREDITS_DB") or DEFAULT_DB
        self._init()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.path, timeout=10)
        conn.row_factory = sqlite3.Row
        return conn

    def _init(self) -> None:
        with self._connect() as conn:
            conn.execute(
                "CREATE TABLE IF NOT EXISTS codes (code TEXT PRIMARY KEY, owner TEXT NOT NULL, quota INTEGER NOT NULL, used INTEGER NOT NULL DEFAULT 0)"
            )

    @staticmethod
    def _norm(code: str) -> str:
        return (code or "").strip()

    # ---- admin ------------------------------------------------------------
    def add(self, code: str, owner: str, quota: int, replace: bool = False) -> Code:
        code = self._norm(code)
        if not code:
            raise ValueError("código vazio")
        if quota < 0:
            raise ValueError("cota negativa")
        with self._connect() as conn:
            if replace:
                conn.execute("INSERT INTO codes(code, owner, quota, used) VALUES (?, ?, ?, 0) ON CONFLICT(code) DO UPDATE SET owner=excluded.owner, quota=excluded.quota", (code, owner, quota))
            else:
                try:
                    conn.execute("INSERT INTO codes(code, owner, quota, used) VALUES (?, ?, ?, 0)", (code, owner, quota))
                except sqlite3.IntegrityError as exc:
                    raise ValueError(f"código '{code}' já existe") from exc
        return self.get(code)  # type: ignore[return-value]

    def remove(self, code: str) -> bool:
        with self._connect() as conn:
            cur = conn.execute("DELETE FROM codes WHERE code = ?", (self._norm(code),))
            return cur.rowcount > 0

    def reset(self, code: str) -> bool:
        with self._connect() as conn:
            cur = conn.execute("UPDATE codes SET used = 0 WHERE code = ?", (self._norm(code),))
            return cur.rowcount > 0

    def list(self) -> List[Code]:
        with self._connect() as conn:
            rows = conn.execute("SELECT code, owner, quota, used FROM codes ORDER BY owner, code").fetchall()
        return [Code(r["code"], r["owner"], r["quota"], r["used"]) for r in rows]

    # ---- runtime ----------------------------------------------------------
    def get(self, code: str) -> Optional[Code]:
        with self._connect() as conn:
            r = conn.execute("SELECT code, owner, quota, used FROM codes WHERE code = ?", (self._norm(code),)).fetchone()
        return Code(r["code"], r["owner"], r["quota"], r["used"]) if r else None

    def validate(self, code: str) -> Optional[Code]:
        """Código existente (mesmo com saldo zero); None se inválido."""
        return self.get(code)

    def balance(self, code: str) -> int:
        c = self.get(code)
        return c.balance if c else 0

    def debit(self, code: str, amount: int = 1) -> bool:
        """Debita atomicamente; False se o saldo for insuficiente ou o código não existir."""
        if amount < 0:
            raise ValueError("valor negativo")
        with self._connect() as conn:
            cur = conn.execute(
                "UPDATE codes SET used = used + ? WHERE code = ? AND used + ? <= quota",
                (amount, self._norm(code), amount),
            )
            return cur.rowcount > 0

    # ---- seeding (hospedagem efêmera) ---------------------------------------
    def seed(self, spec: Optional[str]) -> int:
        """``codigo:dono:cota,...`` -> cria os códigos que ainda não existem. Retorna quantos criou."""
        if not spec:
            return 0
        created = 0
        for item in spec.split(","):
            item = item.strip()
            if not item:
                continue
            parts = item.split(":")
            if len(parts) != 3:
                raise ValueError(f"entrada inválida em ACCESS_CODES: '{item}' (use codigo:dono:cota)")
            code, owner, quota = parts[0].strip(), parts[1].strip(), int(parts[2])
            if self.get(code) is None:
                self.add(code, owner, quota)
                created += 1
        return created


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(prog="app.credits", description="Gerencia códigos de acesso (SQLite).")
    ap.add_argument("--db", default=None, help=f"arquivo SQLite (padrão: $CREDITS_DB ou {DEFAULT_DB})")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p_add = sub.add_parser("add", help="cria um código")
    p_add.add_argument("code")
    p_add.add_argument("--owner", required=True)
    p_add.add_argument("--quota", type=int, required=True)
    p_add.add_argument("--replace", action="store_true", help="atualiza dono/cota se já existir")
    sub.add_parser("list", help="lista códigos e saldos")
    p_rm = sub.add_parser("remove", help="apaga um código")
    p_rm.add_argument("code")
    p_rs = sub.add_parser("reset", help="zera o uso de um código")
    p_rs.add_argument("code")
    args = ap.parse_args(argv)

    store = CreditStore(args.db)
    try:
        if args.cmd == "add":
            c = store.add(args.code, args.owner, args.quota, replace=args.replace)
            print(f"ok: {c.code} ({c.owner}) cota={c.quota} usado={c.used}")
        elif args.cmd == "list":
            rows = store.list()
            if not rows:
                print("(nenhum código)")
            for c in rows:
                print(f"{c.code:<20} {c.owner:<20} cota={c.quota:<5} usado={c.used:<5} saldo={c.balance}")
        elif args.cmd == "remove":
            ok = store.remove(args.code)
            print("ok" if ok else "código não encontrado")
            return 0 if ok else 1
        elif args.cmd == "reset":
            ok = store.reset(args.code)
            print("ok" if ok else "código não encontrado")
            return 0 if ok else 1
    except ValueError as exc:
        print(f"erro: {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
