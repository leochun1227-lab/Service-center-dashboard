"""Verify encrypted persistence with synthetic test credentials only."""
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import dashboard_credentials as store


@unittest.skipUnless(os.name == "nt", "Windows DPAPI is required")
class CredentialTests(unittest.TestCase):
    def test_first_setup_then_no_prompts_across_processes(self):
        with tempfile.TemporaryDirectory(prefix="credentials-test-", dir=ROOT / "outputs") as directory:
            with patch.dict(os.environ, {"LOCALAPPDATA": directory}, clear=True):
                with patch("builtins.input", side_effect=["test-user", ""]), patch("getpass.getpass", side_effect=["fake-password", "fake-dsn"]):
                    store.setup()
                encrypted = store.settings_path().read_bytes()
                self.assertNotIn(b"fake-password", encrypted)
                self.assertNotIn(b"fake-dsn", encrypted)
                self.assertEqual(store.load_settings()["C4C_HISTORY_PASSWORD"], "fake-password")
                with patch("builtins.input", side_effect=AssertionError("Unexpected input")), patch("getpass.getpass", side_effect=AssertionError("Unexpected password prompt")):
                    store.setup()
                    store.setup()
                store.apply_saved_settings()
                self.assertEqual(os.environ["C4C_PASSWORD"], "fake-password")
                os.environ["C4C_PASSWORD"] = "temporary-override"
                store.apply_saved_settings()
                self.assertEqual(os.environ["C4C_PASSWORD"], "temporary-override")
                self.assertEqual(store.load_settings()["C4C_PASSWORD"], "fake-password")
            child_env = {key: value for key, value in os.environ.items() if key not in store.KEYS}
            child_env["LOCALAPPDATA"] = directory
            result = subprocess.run([sys.executable, str(ROOT / "dashboard_credentials.py"), "--check"], cwd=directory, env=child_env, stdin=subprocess.DEVNULL, capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertNotIn("fake-password", result.stdout + result.stderr)

    def test_separate_history_and_explicit_reconfiguration(self):
        with tempfile.TemporaryDirectory(prefix="credentials-test-", dir=ROOT / "outputs") as directory:
            with patch.dict(os.environ, {"LOCALAPPDATA": directory}, clear=True):
                with patch("builtins.input", side_effect=["test-user", "n", "history-user"]), patch("getpass.getpass", side_effect=["fake-password", "fake-dsn", "history-password"]):
                    store.setup()
                before = store.load_settings()
                self.assertEqual(before["C4C_HISTORY_USERNAME"], "history-user")
                with patch("builtins.input", side_effect=["", "", ""]), patch("getpass.getpass", side_effect=["new-password", "", ""]):
                    store.setup(reconfigure=True)
                after = store.load_settings()
                self.assertEqual(after["C4C_PASSWORD"], "new-password")
                self.assertEqual(after["C4C_HISTORY_PASSWORD"], "history-password")
                self.assertEqual(after["SAP_HANA_DSN"], before["SAP_HANA_DSN"])

    def test_unreadable_saved_settings_fail_without_prompt(self):
        with tempfile.TemporaryDirectory(prefix="credentials-test-", dir=ROOT / "outputs") as directory:
            with patch.dict(os.environ, {"LOCALAPPDATA": directory}, clear=True):
                store.settings_path().parent.mkdir()
                store.settings_path().write_bytes(b"invalid-test-data")
                with patch("builtins.input", side_effect=AssertionError("Unexpected input")), self.assertRaisesRegex(RuntimeError, "SETUP_CONNECTIONS"):
                    store.setup()


if __name__ == "__main__":
    unittest.main()
