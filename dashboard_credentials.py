"""Save connection settings once per Windows account using Windows DPAPI."""
from __future__ import annotations

import argparse
import ctypes
from ctypes import wintypes
import getpass
import json
import os
from pathlib import Path
import sys


KEYS = (
    "C4C_USERNAME", "C4C_PASSWORD", "SAP_HANA_DSN",
    "C4C_HISTORY_USERNAME", "C4C_HISTORY_PASSWORD",
    "C4C_BASE_URL", "C4C_HISTORY_POSTMAN_COLLECTION", "SAP_SCHEMA", "SAP_CLIENT",
)
REQUIRED = ("C4C_USERNAME", "C4C_PASSWORD", "SAP_HANA_DSN")


def settings_path() -> Path:
    local = Path(os.environ["LOCALAPPDATA"]) if os.environ.get("LOCALAPPDATA") else Path.home() / "AppData/Local"
    return local / "ServiceCenterDashboard" / "connections.dpapi"


class DataBlob(ctypes.Structure):
    _fields_ = [("size", wintypes.DWORD), ("data", ctypes.POINTER(ctypes.c_ubyte))]


def protect(data: bytes, *, decrypt: bool = False) -> bytes:
    if os.name != "nt":
        raise RuntimeError("Saved connections require Windows.")
    crypt = ctypes.WinDLL("crypt32", use_last_error=True)
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    operation = crypt.CryptUnprotectData if decrypt else crypt.CryptProtectData
    operation.argtypes = [ctypes.POINTER(DataBlob), ctypes.c_void_p, ctypes.c_void_p,
                          ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(DataBlob)]
    operation.restype = wintypes.BOOL
    kernel.LocalFree.argtypes = [ctypes.c_void_p]
    kernel.LocalFree.restype = ctypes.c_void_p
    buffer = ctypes.create_string_buffer(data)
    source = DataBlob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)))
    result = DataBlob()
    # User scope (no LOCAL_MACHINE flag), with all DPAPI prompts disabled.
    if not operation(ctypes.byref(source), None, None, None, None, 1, ctypes.byref(result)):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return ctypes.string_at(result.data, result.size)
    finally:
        ctypes.memset(result.data, 0, result.size)
        kernel.LocalFree(result.data)


def load_settings() -> dict[str, str]:
    path = settings_path()
    if not path.exists():
        return {}
    try:
        values = json.loads(protect(path.read_bytes(), decrypt=True).decode("utf-8"))
        if not isinstance(values, dict) or any(not isinstance(v, str) for v in values.values()):
            raise ValueError("Invalid settings")
        return {key: value for key, value in values.items() if key in KEYS}
    except (OSError, ValueError, UnicodeError) as error:
        raise RuntimeError("Saved connections cannot be read by this Windows account. Run SETUP_CONNECTIONS.bat once to reconfigure.") from error


def save_settings(values: dict[str, str]) -> None:
    path = settings_path()
    encrypted = protect(json.dumps({k: v for k, v in values.items() if k in KEYS}).encode("utf-8"))
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".tmp")
    temporary.write_bytes(encrypted)
    temporary.replace(path)


def apply_saved_settings() -> None:
    for key, value in load_settings().items():
        if not os.environ.get(key) and value:
            os.environ[key] = value


def configured_values() -> dict[str, str]:
    values = load_settings()
    values.update({key: os.environ[key] for key in KEYS if os.environ.get(key)})
    return values


def prompt_value(label: str, existing: str = "", *, secret: bool = False) -> str:
    suffix = " [Enter to keep saved value]" if existing else ""
    while True:
        value = (getpass.getpass if secret else input)(label + suffix + ": ")
        if value:
            return value
        if existing:
            return existing
        print("This value is required.")


def setup(*, reconfigure: bool = False) -> None:
    try:
        values = configured_values()
    except RuntimeError:
        if not reconfigure:
            raise
        values = {key: os.environ[key] for key in KEYS if os.environ.get(key)}
    if not reconfigure and all(values.get(key) for key in REQUIRED):
        if not settings_path().exists():
            save_settings(values)
        print("[OK] Saved connection settings are ready; no input needed.")
        return
    print("[SETUP] Configure this Windows account once. Passwords are hidden and stored encrypted locally.")
    fields = (("C4C_USERNAME", "C4C username", False), ("C4C_PASSWORD", "C4C password", True),
              ("SAP_HANA_DSN", "SAP HANA connection string (DRIVER;SERVERNODE;UID;PWD)", True))
    for key, label, secret in fields:
        if reconfigure or not values.get(key):
            values[key] = prompt_value(label, values.get(key, ""), secret=secret)
    separate = bool(values.get("C4C_HISTORY_USERNAME") and values["C4C_HISTORY_USERNAME"] != values["C4C_USERNAME"])
    answer = input("Use the same C4C account for status history? " + ("[y/N]: " if separate else "[Y/n]: ")).strip().lower()
    use_same = answer == "y" or (not answer and not separate)
    if use_same:
        values["C4C_HISTORY_USERNAME"] = values["C4C_USERNAME"]
        values["C4C_HISTORY_PASSWORD"] = values["C4C_PASSWORD"]
    else:
        values["C4C_HISTORY_USERNAME"] = prompt_value("History username", values.get("C4C_HISTORY_USERNAME", ""))
        values["C4C_HISTORY_PASSWORD"] = prompt_value("History password", values.get("C4C_HISTORY_PASSWORD", ""), secret=True)
    save_settings(values)
    print("[SUCCESS] Connections saved. Future updates load them automatically.")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--setup", action="store_true", help="Change saved connections")
    parser.add_argument("--check", action="store_true", help="Check saved settings without prompting")
    args = parser.parse_args()
    try:
        if args.check:
            values = configured_values()
            if not all(values.get(key) for key in REQUIRED):
                raise RuntimeError("Run SETUP_CONNECTIONS.bat once to save connection settings.")
            print("[OK] Saved connection settings are ready.")
        else:
            setup(reconfigure=args.setup)
        return 0
    except (RuntimeError, OSError, EOFError, KeyboardInterrupt) as error:
        print(f"[ERROR] {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
