import pytest

from app.credits import CreditStore, main


@pytest.fixture
def store(tmp_path):
    return CreditStore(str(tmp_path / "credits.db"))


def test_add_get_balance_debit(store):
    c = store.add("ABC123", "Maria", 3)
    assert (c.code, c.owner, c.quota, c.used, c.balance) == ("ABC123", "Maria", 3, 0, 3)
    assert store.get("  abc123 ") is None  # case sensitive, espaços ignorados
    assert store.get(" ABC123 ").owner == "Maria"
    assert store.validate("ABC123") is not None and store.validate("nope") is None
    assert store.debit("ABC123") and store.balance("ABC123") == 2
    assert store.debit("ABC123", 2) and store.balance("ABC123") == 0
    assert not store.debit("ABC123")  # sem saldo
    assert store.balance("ABC123") == 0
    assert not store.debit("nope")
    assert store.validate("ABC123") is not None  # código continua válido, só sem saldo


def test_add_duplicate_replace_reset_remove(store):
    store.add("X", "a", 5)
    with pytest.raises(ValueError):
        store.add("X", "b", 1)
    store.debit("X", 4)
    c = store.add("X", "b", 10, replace=True)
    assert (c.owner, c.quota, c.used) == ("b", 10, 4)  # uso preservado
    assert store.reset("X") and store.get("X").used == 0
    assert store.remove("X") and store.get("X") is None
    assert not store.remove("X") and not store.reset("X")
    with pytest.raises(ValueError):
        store.add("", "a", 1)
    with pytest.raises(ValueError):
        store.add("Y", "a", -1)


def test_seed_from_spec(store):
    assert store.seed(None) == 0
    assert store.seed("A:ana:5, B:bruno:2,") == 2
    assert store.seed("A:ana:99,C:carla:1") == 1  # A já existe: mantém cota 5
    assert store.get("A").quota == 5 and store.get("C").quota == 1
    with pytest.raises(ValueError):
        store.seed("semcota")
    assert [c.code for c in store.list()] == ["A", "B", "C"]


def test_cli(tmp_path, capsys):
    db = str(tmp_path / "c.db")
    assert main(["--db", db, "add", "CODIGO", "--owner", "nome", "--quota", "20"]) == 0
    assert main(["--db", db, "add", "CODIGO", "--owner", "x", "--quota", "1"]) == 2
    assert main(["--db", db, "list"]) == 0
    out = capsys.readouterr().out
    assert "CODIGO" in out and "saldo=20" in out
    assert main(["--db", db, "reset", "CODIGO"]) == 0
    assert main(["--db", db, "remove", "CODIGO"]) == 0
    assert main(["--db", db, "remove", "CODIGO"]) == 1
    assert main(["--db", db, "list"]) == 0
    assert "(nenhum código)" in capsys.readouterr().out


def test_env_default_path(tmp_path, monkeypatch):
    monkeypatch.setenv("CREDITS_DB", str(tmp_path / "env.db"))
    s = CreditStore()
    s.add("Z", "z", 1)
    assert CreditStore().get("Z") is not None
    assert (tmp_path / "env.db").exists()
