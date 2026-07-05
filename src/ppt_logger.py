"""
ppt_logger.py — Central structured logging for the PPT generation pipeline.
================================================================================

Every log entry is written to TWO files inside the logs/ directory:

  ① <session>.txt     Human-readable pretty format — easy to read and trace.
  ② <session>.jsonl   One JSON object per line   — load directly into any DB.

Both files share the same session_id so you can always match them.

Log fields (present in every JSONL record):
  timestamp    ISO-8601 datetime with microseconds
  level        DEBUG / INFO / WARNING / ERROR / CRITICAL
  session_id   Unique per run  (YYYYMMDD_HHMMSS_<4-char suffix>)
  project      Active PPT project name
  slide        Active slide number   (int, 0 = not yet in a slide)
  chart        Active chart label    (str, e.g. "A", "B", "" = not in a chart)
  file         Python source filename  (e.g. "chart_builder.py")
  module       Python module name
  function     Enclosing function name at the call site
  line         Line number at the call site
  error_type   Exception class name   (ERROR/CRITICAL only, else "")
  message      Human-readable description
  traceback    Full formatted traceback (ERROR/CRITICAL only, else "")
  extra        Dict of any caller-supplied key-value pairs

QUICK START
───────────
  # 1. In main.py — initialise once at the top of run_pipeline():
  from src.ppt_logger import get_logger
  log = get_logger()
  log.init(log_dir="logs", project="NIIF_GRIX_PPT")

  # 2. Set / update context at any time:
  log.set_context(project="NIIF_GRIX_PPT", slide=2, chart="B")

  # 3. Block-level context (auto-reverts when the with-block exits):
  with log.context(slide=3, chart="A"):
      build_chart(...)

  # 4. Decorator — catches, logs, and re-raises exceptions automatically:
  from src.ppt_logger import log_exceptions

  @log_exceptions
  def _build_dual(cfg, x, y1, n, fw, fh):
      ...

  # 5. Manual log calls:
  log.info("Fetching widget", widget_id=1234)
  log.warning("y_max == y_min — chart may look flat", y_min=5, y_max=5)
  log.error("Unexpected shape", expected="(n,)", got=str(arr.shape))
"""

from __future__ import annotations

import inspect
import json
import os
import sys
import traceback as tb
import uuid
from contextlib import contextmanager
from datetime import datetime
from functools import wraps
from pathlib import Path
from threading import Lock


# ── ANSI colours for console output (auto-disabled on non-TTY) ───────────────
_COLOUR = sys.stderr.isatty() if hasattr(sys.stderr, "isatty") else False

_CLR = {
    "DEBUG":    "\033[36m",   # cyan
    "INFO":     "\033[32m",   # green
    "WARNING":  "\033[33m",   # yellow
    "ERROR":    "\033[31m",   # red
    "CRITICAL": "\033[35m",   # magenta
    "RESET":    "\033[0m",
}

def _clr(level: str, text: str) -> str:
    if not _COLOUR:
        return text
    return f"{_CLR.get(level, '')}{text}{_CLR['RESET']}"


# ── Separator widths ──────────────────────────────────────────────────────────
_SEP_WIDE  = "=" * 80
_SEP_SLIM  = "-" * 80


# ══════════════════════════════════════════════════════════════════════════════
class PPTLogger:
    """
    Singleton structured logger.

    Do NOT instantiate directly — always call ``get_logger()``.
    """

    _instance: "PPTLogger | None" = None
    _lock: Lock = Lock()

    # ── singleton access ──────────────────────────────────────────────────────
    @classmethod
    def _get(cls) -> "PPTLogger":
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = cls.__new__(cls)
                    cls._instance._setup()
        return cls._instance

    # ── internal init ─────────────────────────────────────────────────────────
    def _setup(self) -> None:
        self._write_lock  = Lock()
        self._context: dict = {
            "project": "",
            "slide":   0,
            "chart":   "",
        }
        self._ctx_stack: list[dict] = []   # for nested log_context() calls
        self._session_id: str = (
            datetime.now().strftime("%Y%m%d_%H%M%S") + "_" +
            uuid.uuid4().hex[:4].upper()
        )
        self._txt_path:   Path | None = None
        self._jsonl_path: Path | None = None
        self._txt_fh  = None   # open file handle (.txt)
        self._jsonl_fh = None  # open file handle (.jsonl)
        self._initialized = False
        self._min_console_level = "WARNING"   # only WARNING+ goes to stderr

    # ── public init ───────────────────────────────────────────────────────────
    def init(self,
             log_dir: str | Path = "logs",
             project: str        = "",
             console_level: str  = "WARNING") -> None:
        """
        Open log files.  Call ONCE at the start of main.py before any work.

        Parameters
        ----------
        log_dir        : Directory to write log files into (created if missing).
        project        : PPT project name, embedded in the filename.
        console_level  : Minimum level echoed to stderr (DEBUG/INFO/WARNING/...).
        """
        log_dir = Path(log_dir)
        log_dir.mkdir(parents=True, exist_ok=True)

        slug = project.replace(" ", "_") if project else "ppt"
        stem = f"{self._session_id}_{slug}"

        self._txt_path   = log_dir / f"{stem}.txt"
        self._jsonl_path = log_dir / f"{stem}.jsonl"

        self._txt_fh   = open(self._txt_path,   "a", encoding="utf-8")
        self._jsonl_fh = open(self._jsonl_path,  "a", encoding="utf-8")
        self._min_console_level = console_level.upper()
        self._initialized = True

        if project:
            self._context["project"] = project

        # Write session header to the .txt file
        header = (
            f"\n{_SEP_WIDE}\n"
            f"  PPT LOGGER SESSION\n"
            f"  Session ID : {self._session_id}\n"
            f"  Project    : {project or '(not set)'}\n"
            f"  Started    : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"  Log file   : {self._txt_path}\n"
            f"  JSONL file : {self._jsonl_path}\n"
            f"{_SEP_WIDE}\n"
        )
        self._txt_fh.write(header)
        self._txt_fh.flush()

        self._write("INFO", f"Logger initialised — session {self._session_id}",
                    _stacklevel=3)

    # ── context management ────────────────────────────────────────────────────
    def set_context(self, **kwargs) -> None:
        """Set persistent context fields (project, slide, chart, …)."""
        with self._write_lock:
            self._context.update(kwargs)

    @contextmanager
    def context(self, **kwargs):
        """
        Temporary context for a ``with`` block.

        Example::

            with log.context(slide=3, chart="B"):
                build_chart(...)
        """
        with self._write_lock:
            saved = dict(self._context)
            self._context.update(kwargs)
        try:
            yield
        finally:
            with self._write_lock:
                self._context.clear()
                self._context.update(saved)

    # ── caller frame resolution ───────────────────────────────────────────────
    @staticmethod
    def _caller(stacklevel: int = 2) -> tuple[str, str, str, int]:
        """Return (filename, module, function, lineno) of the real call site."""
        try:
            stack = inspect.stack()
            # stacklevel=1 → _write, 2 → debug/info/…, 3 → actual caller
            frame_info = stack[stacklevel]
            filepath = frame_info.filename or ""
            return (
                os.path.basename(filepath),
                os.path.splitext(os.path.basename(filepath))[0],
                frame_info.function,
                frame_info.lineno,
            )
        except Exception:
            return ("unknown", "unknown", "unknown", 0)

    # ── core writer ───────────────────────────────────────────────────────────
    def _write(self,
               level:      str,
               message:    str,
               exc_info:   bool        = False,
               _stacklevel: int        = 3,
               **extra) -> None:
        """Internal — write one log entry to both files and optionally stderr."""

        now = datetime.now()
        filename, module, function, lineno = self._caller(_stacklevel)

        # capture exception details if present
        error_type = ""
        traceback_str = ""
        if exc_info:
            exc = sys.exc_info()
            if exc[0] is not None:
                error_type    = exc[0].__name__
                traceback_str = tb.format_exc().strip()

        with self._write_lock:
            ctx = dict(self._context)

        # ── JSON record ───────────────────────────────────────────────────────
        record = {
            "timestamp":   now.isoformat(),
            "level":       level,
            "session_id":  self._session_id,
            "project":     ctx.get("project", ""),
            "slide":       ctx.get("slide",   0),
            "chart":       ctx.get("chart",   ""),
            "file":        filename,
            "module":      module,
            "function":    function,
            "line":        lineno,
            "error_type":  error_type,
            "message":     message,
            "traceback":   traceback_str,
            "extra":       extra if extra else {},
        }

        # ── human-readable TXT entry ──────────────────────────────────────────
        is_error = level in ("ERROR", "CRITICAL")
        ts_str   = now.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]   # ms precision

        if is_error:
            txt_lines = [
                _SEP_WIDE,
                f"[{ts_str}]  {level:<8s} | Session: {self._session_id}",
                f"  Project  : {ctx.get('project', '(none)')}",
                f"  Slide    : {ctx.get('slide', 0)}  |  Chart : {ctx.get('chart', '') or '—'}",
                f"  Location : {filename} → {function}() @ line {lineno}",
            ]
            if error_type:
                txt_lines.append(f"  Error    : {error_type}")
            txt_lines.append(f"  Message  : {message}")
            if extra:
                for k, v in extra.items():
                    txt_lines.append(f"  {k:<9s}: {v}")
            if traceback_str:
                txt_lines.append("  Traceback:")
                for tb_line in traceback_str.splitlines():
                    txt_lines.append(f"    {tb_line}")
            txt_lines.append(_SEP_WIDE)
            txt_entry = "\n".join(txt_lines) + "\n"
        else:
            # single compact line for INFO / DEBUG / WARNING
            ctx_tag  = (f"slide={ctx.get('slide',0)} "
                        f"chart={ctx.get('chart','') or '—'}")
            loc_tag  = f"{filename}:{function}:{lineno}"
            extra_s  = "  " + "  ".join(f"{k}={v}" for k, v in extra.items()) if extra else ""
            txt_entry = (f"[{ts_str}]  {level:<8s} | {ctx_tag} | "
                         f"{loc_tag} | {message}{extra_s}\n")

        # ── write to files ────────────────────────────────────────────────────
        if self._initialized:
            with self._write_lock:
                if self._jsonl_fh:
                    self._jsonl_fh.write(json.dumps(record, ensure_ascii=False) + "\n")
                    self._jsonl_fh.flush()
                if self._txt_fh:
                    self._txt_fh.write(txt_entry)
                    self._txt_fh.flush()

        # ── console echo (WARNING+) ───────────────────────────────────────────
        _levels = ["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]
        min_idx = _levels.index(self._min_console_level) \
                  if self._min_console_level in _levels else 2
        if _levels.index(level) >= min_idx:
            prefix = _clr(level, f"[{level}]")
            loc    = f"{filename}:{function}:{lineno}"
            print(f"  {prefix} {loc} — {message}", file=sys.stderr)
            if traceback_str and level in ("ERROR", "CRITICAL"):
                for line in traceback_str.splitlines()[-6:]:   # last 6 lines
                    print(f"    {line}", file=sys.stderr)

    # ── public log methods ────────────────────────────────────────────────────
    def debug(self, message: str, **extra) -> None:
        self._write("DEBUG", message, _stacklevel=3, **extra)

    def info(self, message: str, **extra) -> None:
        self._write("INFO", message, _stacklevel=3, **extra)

    def warning(self, message: str, **extra) -> None:
        self._write("WARNING", message, _stacklevel=3, **extra)

    def error(self, message: str, **extra) -> None:
        self._write("ERROR", message, _stacklevel=3, **extra)

    def critical(self, message: str, **extra) -> None:
        self._write("CRITICAL", message, _stacklevel=3, **extra)

    def exception(self, message: str, **extra) -> None:
        """Like error() but automatically captures the current exception."""
        self._write("ERROR", message, exc_info=True, _stacklevel=3, **extra)

    # ── shutdown ──────────────────────────────────────────────────────────────
    def close(self) -> None:
        """Write session footer and close file handles.  Call at end of main."""
        if not self._initialized:
            return
        footer = (
            f"\n{_SEP_WIDE}\n"
            f"  SESSION ENDED : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"{_SEP_WIDE}\n"
        )
        with self._write_lock:
            if self._txt_fh:
                self._txt_fh.write(footer)
                self._txt_fh.flush()
                self._txt_fh.close()
                self._txt_fh = None
            if self._jsonl_fh:
                self._jsonl_fh.flush()
                self._jsonl_fh.close()
                self._jsonl_fh = None
        self._initialized = False

    # ── log file paths (for reporting back to the user) ───────────────────────
    @property
    def txt_path(self) -> Path | None:
        return self._txt_path

    @property
    def jsonl_path(self) -> Path | None:
        return self._jsonl_path

    @property
    def session_id(self) -> str:
        return self._session_id


# ══════════════════════════════════════════════════════════════════════════════
# Module-level helpers
# ══════════════════════════════════════════════════════════════════════════════

def get_logger() -> PPTLogger:
    """Return the singleton PPTLogger.  Safe to call from any module."""
    return PPTLogger._get()


def log_exceptions(func):
    """
    Decorator — wraps any function so that unhandled exceptions are
    automatically logged (file / function / line / full traceback) before
    being re-raised.

    Usage::

        from src.ppt_logger import log_exceptions

        @log_exceptions
        def _build_dual(cfg, x, y1, n, fw, fh):
            ...
    """
    @wraps(func)
    def _wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as exc:
            _log = get_logger()
            # We call _write directly so we can set the exact stacklevel to
            # point at the line INSIDE func that raised — not at this wrapper.
            # traceback.format_exc() already contains the full chain.
            _log._write(
                "ERROR",
                f"Unhandled exception in {func.__qualname__}: {exc}",
                exc_info=True,
                _stacklevel=2,   # points at the wrapper (closest to func site)
            )
            raise   # always re-raise — callers decide whether to continue
    return _wrapper


# ── convenience re-export so callers can do: ─────────────────────────────────
#   from src.ppt_logger import get_logger, log_exceptions
__all__ = ["get_logger", "log_exceptions", "PPTLogger"]
