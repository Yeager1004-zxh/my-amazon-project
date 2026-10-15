from __future__ import annotations

import json
import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .mcp_client import redact_sensitive


def write_validation_log(path: Path, records: list[dict[str, Any]]) -> Path:
    """Write a new, secret-redacted JSONL validation report atomically."""
    path = path.expanduser().resolve()
    if path.exists():
        raise FileExistsError(f"校验日志已存在，为保护已有记录不覆盖：{path.name}")
    path.parent.mkdir(parents=True, exist_ok=True)
    generated_at = datetime.now(timezone.utc).isoformat(timespec="seconds")
    safe_records = []
    for record in records:
        safe = redact_sensitive(record)
        safe.setdefault("generated_at_utc", generated_at)
        safe_records.append(json.dumps(safe, ensure_ascii=False, sort_keys=True, default=str))
    fd, temp_name = tempfile.mkstemp(prefix="validation-", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as stream:
            stream.write("\n".join(safe_records) + "\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temp_name, path)
    finally:
        try:
            os.unlink(temp_name)
        except FileNotFoundError:
            pass
    return path
