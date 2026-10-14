"""Точка входа приложения Word Formatter.

Запуск: python main.py
Скрипты run_windows.bat и run_linux.sh сами создают окружение и вызывают этот файл.
"""

from __future__ import annotations

import sys

from word_formatter.app import main

if __name__ == "__main__":
    sys.exit(main())
