"""EPUB validation and Java runtime detection.

EpubCheck is an optional feature: it needs both the ``epubcheck`` package and a
Java runtime. Neither can be assumed in a Store-installed build, so the import
is lazy and every entry point degrades to "not validated" instead of raising.

The bundled ``epubcheck`` Python wrapper is deliberately bypassed. It launches
Java without a stack-size flag and sends stderr to os.devnull, so on a default
JVM the RelaxNG schema compilation dies with StackOverflowError, the wrapper
silently reports ``valid=False`` with an empty message list, and a perfectly
good EPUB looks broken. Invoking the jar directly lets us raise the stack and,
crucially, tell "the validator could not run" apart from "the book has errors".
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile

# The RelaxNG pattern compiler recurses deeply; the default JVM stack is not
# enough on Windows. 16 MB clears it with room to spare.
_JAVA_STACK = "-Xss16m"
_TIMEOUT_SECONDS = 180


def _java_executable() -> str | None:
    found = shutil.which("java")
    if found:
        return found
    java_home = os.environ.get("JAVA_HOME")
    if java_home:
        candidate = os.path.join(java_home, "bin", "java.exe" if os.name == "nt" else "java")
        if os.path.isfile(candidate):
            return candidate
    return None


def java_available() -> bool:
    return _java_executable() is not None


def _epubcheck_jar() -> str | None:
    try:
        import epubcheck
    except Exception:
        return None
    jar = os.path.join(os.path.dirname(epubcheck.__file__), "epubcheck.jar")
    return jar if os.path.isfile(jar) else None


def epubcheck_available() -> bool:
    """True when validation can actually run (jar present AND Java present)."""
    return _epubcheck_jar() is not None and java_available()


def _parse_messages(data: dict) -> list[str]:
    """Turn epubcheck JSON into display-ready strings.

    Deliberately not epubcheck.models.Message: its ``__str__`` joins the field
    values with str.join, which raises TypeError whenever EpubCheck omits a
    suggestion (a very common case). The UI formats messages with str(), so a
    real validation error would crash the result dialog instead of showing it.
    """
    messages: list[str] = []
    for m in data.get("messages", []):
        severity = str(m.get("severity") or "").upper()
        code = str(m.get("ID") or "")
        text = str(m.get("message") or "").strip()
        suggestion = str(m.get("suggestion") or "").strip()

        places = []
        for loc in m.get("locations") or []:
            path = str(loc.get("path") or "").strip()
            line, column = loc.get("line", -1), loc.get("column", -1)
            if line not in (None, -1):
                path += f":{line}"
                if column not in (None, -1):
                    path += f":{column}"
            if path:
                places.append(path)

        head = " ".join(part for part in (severity, code) if part)
        body = " — ".join(part for part in (text, suggestion) if part)
        where = f" [{'; '.join(places)}]" if places else ""
        messages.append(f"{head}{where}: {body}" if head else f"{body}{where}")
    return messages


def validate_epub(epub_bytes: bytes) -> tuple[bool, bool, list]:
    """Validate EPUB bytes with EpubCheck.

    Returns ``(ran, valid, messages)``:
      ran     - False when the validator is unavailable or failed to produce a
                verdict. Callers must not present this as a defect in the book.
      valid   - Only meaningful when ``ran`` is True.
      messages - EpubCheck findings, empty when ``ran`` is False.
    """
    java = _java_executable()
    jar = _epubcheck_jar()
    if not java or not jar:
        return False, True, []

    with tempfile.NamedTemporaryFile(suffix=".epub", delete=False) as tmp:
        tmp.write(epub_bytes)
        tmp_path = tmp.name

    creationflags = subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
    try:
        completed = subprocess.run(
            [
                java,
                _JAVA_STACK,
                "-Duser.language=en",
                "-jar",
                jar,
                tmp_path,
                "-q",
                "--profile",
                "default",
                "--json",
                "-",
            ],
            capture_output=True,
            timeout=_TIMEOUT_SECONDS,
            creationflags=creationflags,
        )
    except (OSError, subprocess.SubprocessError):
        return False, True, []
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass

    try:
        data = json.loads(completed.stdout.decode("utf-8", "replace"))
        counts = data["checker"]
    except (ValueError, KeyError, TypeError):
        # No parseable verdict - the validator crashed (e.g. StackOverflowError
        # on a small JVM stack). Report "did not run", never "invalid".
        return False, True, []

    fatal = int(counts.get("nFatal", 0))
    errors = int(counts.get("nError", 0))
    return True, (fatal == 0 and errors == 0), _parse_messages(data)
