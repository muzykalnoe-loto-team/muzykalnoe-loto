import pytest

from scripts import yandex_disk
from scripts.yandex_disk import DiskError, check_space, disk_path, ensure_folder, event_folder_name


def test_event_folder_name_uses_dashes_in_date():
    assert event_folder_name("events/31.10.2026 Готический бал") == "31-10-2026 Готический бал"
    assert event_folder_name("events/31.10.2026 Готический бал/") == "31-10-2026 Готический бал"


def test_disk_path_joins_parts():
    assert disk_path("Музыкальное лото", "31-10-2026 Готический бал") == "disk:/Музыкальное лото/31-10-2026 Готический бал"


def test_ensure_folder_creates_every_parent(monkeypatch):
    calls = []
    monkeypatch.setattr(yandex_disk, "_api", lambda method, endpoint, token, ok=(), **p: calls.append((method, p["path"], ok)) or (201, {}))

    ensure_folder("disk:/Музыкальное лото/31-10-2026 Готический бал", "t")

    assert [c[1] for c in calls] == ["disk:/Музыкальное лото", "disk:/Музыкальное лото/31-10-2026 Готический бал"]
    assert all(c[0] == "PUT" and 409 in c[2] for c in calls)  # «уже есть» — не ошибка


def _fake_disk(monkeypatch, total, used, max_file):
    info = {"total_space": total, "used_space": used, "max_file_size": max_file}
    monkeypatch.setattr(yandex_disk, "_api", lambda *a, **k: (200, info))


def test_check_space_rejects_too_big_file(monkeypatch, tmp_path):
    f = tmp_path / "Презентация.pptx"
    f.write_bytes(b"x" * 200)
    _fake_disk(monkeypatch, total=10_000, used=0, max_file=100)

    with pytest.raises(DiskError, match="Презентация.pptx"):
        check_space([str(f)], "t")


def test_check_space_rejects_when_disk_is_full(monkeypatch, tmp_path):
    f = tmp_path / "Бланки.xlsm"
    f.write_bytes(b"x" * 50)
    _fake_disk(monkeypatch, total=1000, used=980, max_file=100)

    with pytest.raises(DiskError, match="не хватает места"):
        check_space([str(f)], "t")


def test_write_shortcut_makes_windows_url_file(tmp_path):
    path = tmp_path / "Ссылка.url"

    yandex_disk.write_shortcut(str(path), "https://disk.yandex.ru/d/abc")

    assert path.read_text(encoding="utf-8") == "[InternetShortcut]\nURL=https://disk.yandex.ru/d/abc\n"
