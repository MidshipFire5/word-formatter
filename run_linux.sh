#!/usr/bin/env bash
# Запуск Word Formatter в Linux.
#
# Задача этого файла — найти Python 3 и передать работу tools/bootstrap.py:
# он проверит версию, создаст локальное окружение .venv, установит
# зависимости и запустит приложение.
#
# Если система не позволяет запустить файл напрямую (нет права на исполнение),
# выполните:  bash run_linux.sh
#
# Права администратора не нужны, ничего не устанавливается в систему.

set -u

cd "$(dirname "$0")" || {
    echo "Не удалось перейти в папку программы." >&2
    exit 1
}

if [ ! -f "tools/bootstrap.py" ]; then
    echo "Не найден файл tools/bootstrap.py." >&2
    echo "Похоже, архив с программой распакован не полностью." >&2
    echo "Распакуйте архив заново целиком и повторите запуск." >&2
    exit 1
fi

echo "Word Formatter — оформление документов Word"
echo "Подготовка окружения, при первом запуске это может занять несколько минут."
echo

PYTHON=""
for candidate in python3 python; do
    if command -v "$candidate" >/dev/null 2>&1 && "$candidate" -c 'print("ok")' >/dev/null 2>&1; then
        PYTHON="$candidate"
        break
    fi
done

if [ -z "$PYTHON" ]; then
    echo "Не найден Python 3.10 или новее." >&2
    echo >&2
    echo "Как установить Python:" >&2
    echo "  Ubuntu и Debian:  sudo apt update && sudo apt install -y python3 python3-venv python3-pip" >&2
    echo "  Fedora:           sudo dnf install -y python3 python3-pip" >&2
    echo "  Arch:             sudo pacman -S python" >&2
    echo >&2
    echo "После установки запустите приложение снова: bash run_linux.sh" >&2
    exit 2
fi

exec "$PYTHON" tools/bootstrap.py "$@"
