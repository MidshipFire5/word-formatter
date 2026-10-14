"""Логика приложения, не зависящая от графического интерфейса."""

from __future__ import annotations

from word_formatter.core.errors import FormatterError, ProfileError
from word_formatter.core.report import Report, ReportSection, build_report
from word_formatter.core.verification import (
    Difference,
    VerificationResult,
    verify_documents,
    verify_files,
)

__all__ = [
    "Difference",
    "FormatterError",
    "ProfileError",
    "Report",
    "ReportSection",
    "VerificationResult",
    "build_report",
    "verify_documents",
    "verify_files",
]
