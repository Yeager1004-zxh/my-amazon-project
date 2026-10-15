from __future__ import annotations

import os
from typing import Protocol


class CredentialStoreError(RuntimeError):
    """Windows credential storage is unavailable or failed."""


class CredentialStore(Protocol):
    def get(self, account: str) -> str | None: ...

    def set(self, account: str, secret: str) -> None: ...

    def delete(self, account: str) -> None: ...


class WindowsCredentialStore:
    """Store MCP credentials in Windows Credential Manager via keyring."""

    service_name = "my-amazon-project.lingxing-mcp"

    def __init__(self) -> None:
        if os.name != "nt":
            raise CredentialStoreError(
                "领星凭证只在 Windows Credential Manager 中保存；请在 Windows 10/11 上配置。"
            )
        try:
            import keyring
            from keyring.backends.Windows import WinVaultKeyring

            # Pin the backend instead of trusting a user-configured plaintext backend.
            keyring.set_keyring(WinVaultKeyring())
            self._keyring = keyring
        except Exception as exc:
            raise CredentialStoreError("无法初始化 Windows Credential Manager。") from exc

    def get(self, account: str) -> str | None:
        try:
            return self._keyring.get_password(self.service_name, account)
        except Exception as exc:
            raise CredentialStoreError("从 Windows Credential Manager 读取凭证失败。") from exc

    def set(self, account: str, secret: str) -> None:
        if not secret or "\r" in secret or "\n" in secret:
            raise CredentialStoreError("MCP 凭证为空或包含非法换行。")
        try:
            self._keyring.set_password(self.service_name, account, secret)
        except Exception as exc:
            raise CredentialStoreError("写入 Windows Credential Manager 失败。") from exc

    def delete(self, account: str) -> None:
        try:
            self._keyring.delete_password(self.service_name, account)
        except self._keyring.errors.PasswordDeleteError:
            return
        except Exception as exc:
            raise CredentialStoreError("从 Windows Credential Manager 删除凭证失败。") from exc
