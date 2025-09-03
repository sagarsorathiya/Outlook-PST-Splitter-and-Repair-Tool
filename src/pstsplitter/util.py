"""Utility helpers for the PST splitter.

Thread-safe logging queue, size formatting, and simple timing helpers.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from queue import Queue, Empty
from threading import Event
from time import perf_counter
import logging
from typing import Iterable
import json

LOG_QUEUE: "Queue[logging.LogRecord]" = Queue()
__all__: list[str] = []


class QueueHandler(logging.Handler):
    """Logging handler that pushes records into a thread-safe queue."""

    def emit(self, record: logging.LogRecord) -> None:  # noqa: D401
        try:
            LOG_QUEUE.put_nowait(record)
        except Exception:  # pragma: no cover - extremely rare
            pass


class QueueListener:
    """Simple queue listener that pulls records and dispatches to target logger."""

    def __init__(self, stop_event: Event, target: logging.Logger) -> None:
        self._stop = stop_event
        self._target = target

    def pump(self) -> None:
        while not self._stop.is_set():
            try:
                rec = LOG_QUEUE.get(timeout=0.2)
            except Empty:
                continue
            self._target.handle(rec)


def configure_logging(level: int = logging.INFO) -> Event:
    """Configure root logger with queue handling and timestamps.

    Returns a stop event the caller should set on shutdown to end listener.
    """
    root = logging.getLogger()
    root.setLevel(level)
    # Avoid duplicate handlers if called twice.
    if not any(isinstance(h, QueueHandler) for h in root.handlers):
        qh = QueueHandler()
        # Set format with timestamp for better monitoring
        formatter = logging.Formatter(
            fmt='[%(asctime)s] %(levelname)s: %(message)s',
            datefmt='%H:%M:%S'
        )
        qh.setFormatter(formatter)
        root.addHandler(qh)
    stop_event = Event()
    listener = QueueListener(stop_event, root)

    from threading import Thread

    t = Thread(target=listener.pump, name="LogPump", daemon=True)
    t.start()
    return stop_event


@dataclass
class Timer:
    """Context manager stopwatch for measuring code blocks."""

    label: str
    start: float | None = None

    def __enter__(self) -> "Timer":
        self.start = perf_counter()
        logging.debug("Timer %s start", self.label)
        return self

    def __exit__(self, exc_type, exc, tb) -> None:  # type: ignore[override]
        if self.start is None:
            return
        elapsed = perf_counter() - self.start
        logging.debug("Timer %s end: %.3fs", self.label, elapsed)


def format_bytes(num: int, precision: int = 2) -> str:
    """Human readable size formatting.

    Args:
        num: Number of bytes.
        precision: Decimal places.
    """
    step = 1024.0
    units = ["B", "KB", "MB", "GB", "TB"]
    size = float(num)
    for unit in units:
        if size < step:
            return f"{size:.{precision}f} {unit}"
        size /= step
    return f"{size:.{precision}f} PB"


def chunk_sequence(items: Iterable[int], max_sum: int) -> list[list[int]]:
    """Group integers into chunks with a maximum summed size (greedy).

    Used for tests as a deterministic analogue of item size grouping.
    """
    chunk: list[int] = []
    out: list[list[int]] = []
    total = 0
    for val in items:
        if val > max_sum:
            # Single oversize item becomes its own chunk.
            if chunk:
                out.append(chunk)
                chunk = []
                total = 0
            out.append([val])
            continue
        if total + val > max_sum:
            out.append(chunk)
            chunk = [val]
            total = val
        else:
            chunk.append(val)
            total += val
    if chunk:
        out.append(chunk)
    return out

# --- Preferences -----------------------------------------------------------------
PREF_PATH = Path.home() / ".pstsplitter_config.json"


def load_prefs() -> dict:
    """Load user preferences from disk (best-effort)."""
    try:
        if PREF_PATH.exists():
            with PREF_PATH.open("r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
    except Exception:  # pragma: no cover
        pass
    return {}


def save_prefs(prefs: dict) -> None:
    """Persist user preferences best-effort."""
    try:
        with PREF_PATH.open("w", encoding="utf-8") as f:
            json.dump(prefs, f, indent=2)
    except Exception:  # pragma: no cover
        pass


__all__ += [
    "configure_logging",
    "format_bytes",
    "chunk_sequence",
    "load_prefs",
    "save_prefs",
]
