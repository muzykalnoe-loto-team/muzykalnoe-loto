# Загрузка готовых файлов мероприятия на Яндекс Диск через REST API
# (https://yandex.ru/dev/disk-api/doc/ru/). Ключ — YANDEX_DISK_TOKEN из .env.
import json
import os
import urllib.error
import urllib.parse
import urllib.request

from scripts.secret_store import get_secret

API = "https://cloud-api.yandex.net/v1/disk"
ROOT = "Музыкальное лото"
UPLOAD_TIMEOUT = 60 * 60  # файл около 1 ГБ на медленном интернете


class DiskError(RuntimeError):
    pass


def _call(method: str, url: str, token: str, ok=(200, 201, 202)) -> tuple[int, dict]:
    req = urllib.request.Request(url, method=method, headers={"Authorization": "OAuth " + token})
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            body = resp.read()
            return resp.status, json.loads(body) if body else {}
    except urllib.error.HTTPError as e:
        body = e.read()
        data = json.loads(body) if body else {}
        if e.code in ok:
            return e.code, data
        raise DiskError(f"Яндекс Диск ответил {e.code}: {data.get('description') or data.get('message') or e.reason}") from None


def _api(method: str, endpoint: str, token: str, ok=(200, 201, 202), **params) -> tuple[int, dict]:
    query = urllib.parse.urlencode(params)
    return _call(method, f"{API}{endpoint}?{query}" if query else f"{API}{endpoint}", token, ok)


def disk_path(*parts: str) -> str:
    return "disk:/" + "/".join(p.strip("/") for p in parts if p)


def event_folder_name(local_folder: str) -> str:
    # «31.10.2026 Готический бал» → «31-10-2026 Готический бал»
    name = os.path.basename(os.path.normpath(local_folder))
    date, _, rest = name.partition(" ")
    return f"{date.replace('.', '-')} {rest}".strip()


def ensure_folder(path: str, token: str) -> None:
    # Создаём папку и все родительские; 409 — «уже есть», это нормально
    parts = path.removeprefix("disk:/").split("/")
    for i in range(1, len(parts) + 1):
        _api("PUT", "/resources", token, ok=(201, 409), path=disk_path(*parts[:i]))


def _remote_size(path: str, token: str) -> int | None:
    status, data = _api("GET", "/resources", token, ok=(200, 404), path=path, fields="size")
    return data.get("size") if status == 200 else None


def check_space(local_files: list[str], token: str) -> None:
    _, info = _api("GET", "", token)
    max_file = info.get("max_file_size")
    for f in local_files:
        if max_file and os.path.getsize(f) > max_file:
            raise DiskError(
                f"Файл «{os.path.basename(f)}» ({os.path.getsize(f) / 1024**3:.2f} ГБ) больше, "
                f"чем разрешает Диск ({max_file / 1024**3:.2f} ГБ)"
            )
    need = sum(os.path.getsize(f) for f in local_files)
    free = info["total_space"] - info["used_space"]
    if need > free:
        raise DiskError(f"На Диске не хватает места: нужно {need / 1024**3:.2f} ГБ, свободно {free / 1024**3:.2f} ГБ")


def upload_file(local_file: str, remote_path: str, token: str) -> bool:
    """Загружает файл; если на Диске уже лежит файл того же размера — пропускает. True, если загружал."""
    size = os.path.getsize(local_file)
    if _remote_size(remote_path, token) == size:
        return False
    _, link = _api("GET", "/resources/upload", token, path=remote_path, overwrite="true")
    with open(local_file, "rb") as f:
        req = urllib.request.Request(
            link["href"], data=f, method=link.get("method", "PUT"),
            headers={"Content-Length": str(size)},
        )
        try:
            with urllib.request.urlopen(req, timeout=UPLOAD_TIMEOUT) as resp:
                if resp.status not in (200, 201, 202):
                    raise DiskError(f"Загрузка «{os.path.basename(local_file)}» не удалась: {resp.status}")
        except urllib.error.HTTPError as e:
            raise DiskError(f"Загрузка «{os.path.basename(local_file)}» не удалась: {e.code} {e.reason}") from None
    if _remote_size(remote_path, token) != size:
        raise DiskError(f"После загрузки размер «{os.path.basename(local_file)}» на Диске не совпал")
    return True


def publish(path: str, token: str) -> str:
    _api("PUT", "/resources/publish", token, path=path)
    _, data = _api("GET", "/resources", token, path=path, fields="public_url")
    if not data.get("public_url"):
        raise DiskError("Диск не выдал публичную ссылку на папку")
    return data["public_url"]


def write_shortcut(path: str, url: str) -> None:
    """Ярлык Windows (.url): двойной щелчок в проводнике открывает ссылку в браузере."""
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"[InternetShortcut]\nURL={url}\n")


SHORTCUT_NAME = "Файлы на Яндекс Диске.url"


def upload_event(local_folder: str, file_names: list[str], token: str | None = None, log=print) -> str:
    """Кладёт файлы мероприятия в «Музыкальное лото/<дата> <тема>» и возвращает публичную ссылку на папку."""
    token = token or get_secret("YANDEX_DISK_TOKEN")
    local_files = [os.path.join(local_folder, n) for n in file_names]
    missing = [n for n, f in zip(file_names, local_files) if not os.path.exists(f)]
    if missing:
        raise DiskError("Нет файлов для загрузки: " + ", ".join(missing))
    remote = disk_path(ROOT, event_folder_name(local_folder))
    # место проверяем только под файлы, которых на Диске ещё нет (повторный запуск их пропустит)
    pending = [f for n, f in zip(file_names, local_files) if _remote_size(f"{remote}/{n}", token) != os.path.getsize(f)]
    check_space(pending, token)
    ensure_folder(remote, token)
    for name, f in zip(file_names, local_files):
        log(f"{name}: загружаю…")
        done = upload_file(f, f"{remote}/{name}", token)
        log(f"{name}: {'загружен' if done else 'уже был на Диске, пропущен'}")
    url = publish(remote, token)
    write_shortcut(os.path.join(local_folder, SHORTCUT_NAME), url)
    return url
