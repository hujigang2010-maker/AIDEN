#!/usr/bin/env python3
"""Gemini 桌面端登录修复的单元测试。"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import gemini_desktop as gd  # noqa: E402


class LoginUrlTests(unittest.TestCase):
    def test_continue_points_to_gemini_dot_com(self) -> None:
        url = gd.login_url()
        self.assertTrue(url.startswith("https://accounts.google.com/ServiceLogin?"))
        self.assertIn("continue=", url)
        self.assertIn("gemini.google.com", url)
        self.assertNotIn("gemini.google/app", url.replace("gemini.google.com", ""))

    def test_rejects_broken_host(self) -> None:
        with self.assertRaises(ValueError):
            gd.login_url("https://gemini.google/app")

    def test_rejects_non_gemini_continue(self) -> None:
        with self.assertRaises(ValueError):
            gd.login_url("https://example.com/")


class ChromeArgsTests(unittest.TestCase):
    def test_app_mode_and_profile(self) -> None:
        profile = Path("/tmp/gemini-profile-test")
        start = gd.login_url()
        args = gd.build_chrome_args(
            chrome="/opt/google/chrome/google-chrome",
            profile_dir=profile,
            start_url=start,
            extension_dir=ROOT / "share" / "login-fix-extension",
        )
        self.assertEqual(args[0], "/opt/google/chrome/google-chrome")
        self.assertIn(f"--user-data-dir={profile}", args)
        self.assertIn("--class=GeminiDesktop", args)
        self.assertTrue(any(item.startswith("--app=") for item in args))
        app_url = next(item.split("=", 1)[1] for item in args if item.startswith("--app="))
        self.assertIn("accounts.google.com", app_url)
        self.assertIn("gemini.google.com", app_url)
        self.assertTrue(any(item.startswith("--load-extension=") for item in args))
        if Path("/.dockerenv").exists():
            self.assertIn("--no-sandbox", args)
            self.assertIn("--disable-dev-shm-usage", args)

    def test_find_chrome_skips_cloud_wrapper(self) -> None:
        chrome = gd.find_chrome()
        self.assertNotEqual(chrome, "/usr/local/bin/google-chrome")
        if Path("/opt/google/chrome/google-chrome").is_file():
            self.assertEqual(chrome, "/opt/google/chrome/google-chrome")

    def test_existing_cookies_open_gemini_directly(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            profile = Path(tmp)
            cookies = profile / "Default" / "Network" / "Cookies"
            cookies.parent.mkdir(parents=True)
            cookies.write_bytes(b"cookie")
            self.assertEqual(gd.resolve_start_url(False, profile), gd.GEMINI_APP_URL)
            self.assertTrue(gd.resolve_start_url(True, profile).startswith(gd.ACCOUNTS_LOGIN_PATH))


class ExtensionAndInstallTests(unittest.TestCase):
    def test_extension_manifest_and_rules(self) -> None:
        manifest = json.loads(
            (ROOT / "share" / "login-fix-extension" / "manifest.json").read_text(encoding="utf-8")
        )
        rules = json.loads(
            (ROOT / "share" / "login-fix-extension" / "rules.json").read_text(encoding="utf-8")
        )
        self.assertEqual(manifest["manifest_version"], 3)
        self.assertEqual(rules[0]["action"]["redirect"]["transform"]["host"], "gemini.google.com")
        self.assertEqual(rules[0]["condition"]["requestDomains"], ["gemini.google"])

    def test_install_to_temp_prefix(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            prefix = Path(tmp)
            paths = gd.install_user(prefix)
            self.assertTrue(paths["launcher"].is_file())
            self.assertTrue(paths["desktop"].is_file())
            desktop = paths["desktop"].read_text(encoding="utf-8")
            self.assertIn("Gemini 桌面端", desktop)
            self.assertIn(str(paths["launcher"]), desktop)
            wrapper = paths["launcher"].read_text(encoding="utf-8")
            self.assertIn("gemini_desktop.py", wrapper)


class PrintUrlCliTests(unittest.TestCase):
    def test_print_url_does_not_exec_chrome(self) -> None:
        with mock.patch.object(gd, "launch", side_effect=AssertionError("不应启动浏览器")):
            code = gd.main(["--print-url", "--fix-login"])
        self.assertEqual(code, 0)


if __name__ == "__main__":
    unittest.main()
