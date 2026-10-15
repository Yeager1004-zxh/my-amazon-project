from __future__ import annotations

from pathlib import Path
from typing import Protocol


class FileDataProvider(Protocol):
    """离线报表提供方的接口约定。"""

    def list_files(self, folder: Path) -> list[Path]:
        """返回提供方支持的输入文件。"""
        ...
