"""Фоновая обработка документа.

Обработка большого документа занимает время, и интерфейс не должен
«зависать». Поэтому работа идёт в отдельном потоке, а результат передаётся
в окно сигналами Qt.
"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QObject, QThread, Signal

from word_formatter.core.errors import FormatterError
from word_formatter.core.profiles import Profile
from word_formatter.core.processing import ProcessingResult, process_document

UNEXPECTED_ERROR = (
    "Во время обработки документа произошла непредвиденная ошибка. "
    "Исходный документ не изменялся. Попробуйте другой файл или сообщите о проблеме."
)


class ProcessingWorker(QObject):
    """Выполняет обработку документа в отдельном потоке."""

    progress = Signal(str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, source_path: Path, output_path: Path, profile: Profile) -> None:
        super().__init__()
        self._source_path = source_path
        self._output_path = output_path
        self._profile = profile

    def run(self) -> None:
        try:
            result: ProcessingResult = process_document(
                self._source_path,
                self._output_path,
                self._profile,
                progress=self.progress.emit,
            )
        except FormatterError as error:
            self.failed.emit(error.user_message)
        except Exception:  # непредвиденная ошибка не должна закрывать приложение
            self.failed.emit(UNEXPECTED_ERROR)
        else:
            self.finished.emit(result)


class ProcessingTask(QObject):
    """Управляет потоком обработки и живёт столько же, сколько сама работа."""

    progress = Signal(str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, source_path: Path, output_path: Path, profile: Profile) -> None:
        super().__init__()
        self._thread = QThread()
        self._worker = ProcessingWorker(source_path, output_path, profile)
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self.progress.emit)
        self._worker.finished.connect(self.finished.emit)
        self._worker.failed.connect(self.failed.emit)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)

    def start(self) -> None:
        self._thread.start()

    def wait(self, timeout_ms: int = 30_000) -> bool:
        """Дожидается завершения работы — нужно тестам и закрытию приложения."""
        return self._thread.wait(timeout_ms)
