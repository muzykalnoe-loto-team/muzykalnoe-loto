import pytest

from scripts.secret_store import MissingSecret, get_secret, read_secrets


def test_reads_values_and_skips_comments(tmp_path):
    env = tmp_path / ".env"
    env.write_text('# пояснение\n\nTOKEN=abc=123\nQUOTED="x y"\nEMPTY=\n', encoding="utf-8")

    assert read_secrets(str(env)) == {"TOKEN": "abc=123", "QUOTED": "x y", "EMPTY": ""}
    assert get_secret("TOKEN", str(env)) == "abc=123"


def test_empty_or_missing_value_is_reported_without_leaking(tmp_path):
    env = tmp_path / ".env"
    env.write_text("EMPTY=\nOTHER=secret-value\n", encoding="utf-8")

    with pytest.raises(MissingSecret, match="EMPTY") as err:
        get_secret("EMPTY", str(env))
    assert "secret-value" not in str(err.value)
    with pytest.raises(MissingSecret, match="Нет файла"):
        get_secret("X", str(tmp_path / "нет.env"))
