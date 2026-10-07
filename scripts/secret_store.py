# Ключи и пароли хранятся в файле .env в корне проекта (см. paths.SECRETS_FILE), он не попадает в git.
# Значения никогда не печатаются и не пишутся в логи — только передаются туда, где нужны.
from scripts import paths


class MissingSecret(RuntimeError):
    pass


def read_secrets(path: str = paths.SECRETS_FILE) -> dict:
    secrets = {}
    with open(path, encoding="utf-8-sig") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            secrets[key.strip()] = value.strip().strip('"').strip("'")
    return secrets


def get_secret(name: str, path: str = paths.SECRETS_FILE) -> str:
    try:
        value = read_secrets(path).get(name, "")
    except FileNotFoundError:
        raise MissingSecret(f"Нет файла с ключами: {path}") from None
    if not value:
        raise MissingSecret(f"В файле {path} не заполнено значение {name}")
    return value
