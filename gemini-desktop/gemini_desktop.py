#!/usr/bin/env python3
"""Gemini 桌面端启动与登录修复。

官方原生 Gemini 只提供 macOS（Apple Silicon + macOS 15+）和 Windows。
Linux 上常见的 PWA / Electron 套壳会把 Google 登录弹到另一窗口，
或把回跳写成错误的 gemini.google/app（缺少 .com），于是出现：
授权已经成功，应用窗口却仍未进入登录态。

本启动器使用本机 Google Chrome 的 --app 模式，并强制走
accounts.google.com 的 continue 回跳到 https://gemini.google.com/app。
"""

from __future__ import annotations

import argparse
import os
import shutil
import stat
import sys
import urllib.parse
from pathlib import Path

GEMINI_APP_URL = "https://gemini.google.com/app?hl=zh-CN"
ACCOUNTS_LOGIN_PATH = "https://accounts.google.com/ServiceLogin"
BROKEN_HOST_HINT = "gemini.google/app"

CHROME_CANDIDATES = (
    "google-chrome",
    "google-chrome-stable",
    "chromium-browser",
    "chromium",
    "/opt/google/chrome/google-chrome",
    "/usr/bin/google-chrome",
)


def repo_root() -> Path:
    return Path(__file__).resolve().parent


def default_profile_dir() -> Path:
    override = os.environ.get("GEMINI_DESKTOP_PROFILE")
    if override:
        return Path(override).expanduser()
    return Path.home() / ".local" / "share" / "gemini-desktop" / "chrome-profile"


def default_share_dir() -> Path:
    return repo_root() / "share"


def login_url(continue_url: str = GEMINI_APP_URL) -> str:
    """生成带 continue 的 Google 登录地址，登录成功后回到 Gemini。"""
    if "gemini.google.com" not in continue_url:
        raise ValueError(f"回跳地址必须指向 gemini.google.com，收到: {continue_url}")
    if BROKEN_HOST_HINT in continue_url.replace("gemini.google.com", ""):
        raise ValueError(f"禁止使用缺少 .com 的回跳: {continue_url}")
    query = urllib.parse.urlencode(
        {
            "hl": "zh-CN",
            "continue": continue_url,
            "ec": "Gawebapp",
        }
    )
    return f"{ACCOUNTS_LOGIN_PATH}?{query}"


def find_chrome() -> str:
    override = os.environ.get("GEMINI_DESKTOP_CHROME")
    if override:
        path = Path(override).expanduser()
        if path.is_file() and os.access(path, os.X_OK):
            return str(path)
        raise FileNotFoundError(f"GEMINI_DESKTOP_CHROME 不可执行: {override}")
    for name in CHROME_CANDIDATES:
        found = shutil.which(name) if "/" not in name else (name if Path(name).is_file() else None)
        if found:
            return found
    raise FileNotFoundError("未找到 Google Chrome / Chromium，无法启动 Gemini 桌面端。")


def extra_chrome_flags() -> list[str]:
    """容器环境里 Chrome 沙箱通常不可用，需要补上与系统 Chrome 一致的启动参数。"""
    flags = [item for item in os.environ.get("GEMINI_DESKTOP_CHROME_FLAGS", "").split() if item]
    in_container = Path("/.dockerenv").exists() or Path("/run/.containerenv").exists()
    if in_container:
        for flag in ("--no-sandbox", "--disable-dev-shm-usage", "--password-store=basic"):
            if flag not in flags:
                flags.append(flag)
    return flags


def build_chrome_args(
    *,
    chrome: str,
    profile_dir: Path,
    start_url: str,
    extension_dir: Path | None,
) -> list[str]:
    args = [
        chrome,
        f"--user-data-dir={profile_dir}",
        "--profile-directory=Default",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-features=TranslateUI,MediaRouter",
        "--disable-session-crashed-bubble",
        "--hide-crash-restore-bubble",
        *extra_chrome_flags(),
        f"--app={start_url}",
    ]
    if extension_dir and extension_dir.is_dir():
        args.insert(-1, f"--load-extension={extension_dir}")
    return args


def resolve_start_url(force_login: bool, profile_dir: Path) -> str:
    """无会话或强制修复时走 Google 登录回跳；已有配置目录则直接打开 Gemini。"""
    if force_login:
        return login_url()
    marker = profile_dir / "Default" / "Network" / "Cookies"
    if marker.is_file() and marker.stat().st_size > 0:
        return GEMINI_APP_URL
    return login_url()


def prepare_profile(profile_dir: Path) -> None:
    profile_dir.mkdir(parents=True, exist_ok=True)
    for stale in ("SingletonLock", "SingletonSocket", "SingletonCookie"):
        lock = profile_dir / stale
        if lock.exists() or lock.is_symlink():
            lock.unlink()


def reset_profile(profile_dir: Path) -> None:
    if profile_dir.exists():
        shutil.rmtree(profile_dir)
    profile_dir.mkdir(parents=True, exist_ok=True)


def launch(force_login: bool = False, reset: bool = False) -> int:
    chrome = find_chrome()
    profile_dir = default_profile_dir()
    if reset:
        reset_profile(profile_dir)
        force_login = True
    else:
        prepare_profile(profile_dir)
    start_url = resolve_start_url(force_login, profile_dir)
    extension_dir = default_share_dir() / "login-fix-extension"
    args = build_chrome_args(
        chrome=chrome,
        profile_dir=profile_dir,
        start_url=start_url,
        extension_dir=extension_dir,
    )
    os.execv(chrome, args)


def write_wrapper(dest: Path, target_py: Path, extra_args: str = "") -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    extra = f" {extra_args}" if extra_args else ""
    dest.write_text(
        "#!/usr/bin/env bash\n"
        "set -euo pipefail\n"
        f'exec python3 "{target_py}"{extra} "$@"\n',
        encoding="utf-8",
    )
    dest.chmod(dest.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)


def write_desktop_file(dest: Path, exec_path: Path, icon_path: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    content = f"""[Desktop Entry]
Version=1.0
Type=Application
Name=Gemini 桌面端
Name[zh_CN]=Gemini 桌面端
Comment=使用正确回跳地址登录的 Google Gemini 桌面窗口
Comment[zh_CN]=使用正确回跳地址登录的 Google Gemini 桌面窗口
Exec={exec_path}
Icon={icon_path}
Terminal=false
Categories=Network;WebBrowser;Utility;
StartupNotify=true
StartupWMClass=chrome-gemini.google.com__app
"""
    dest.write_text(content, encoding="utf-8")
    dest.chmod(dest.stat().st_mode | stat.S_IXUSR)


def write_plank_item(dest: Path, desktop_file: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(
        "[PlankDockItemPreferences]\n"
        f"Launcher=file://{desktop_file}\n",
        encoding="utf-8",
    )


def install_user(prefix: Path | None = None) -> dict[str, Path]:
    """把启动器安装到用户目录，并替换坏掉的 PWA 快捷方式。"""
    prefix = prefix or Path.home() / ".local"
    app_dir = prefix / "share" / "gemini-desktop"
    bin_dir = prefix / "bin"
    applications = prefix / "share" / "applications"
    icon_src = default_share_dir() / "icons" / "gemini.svg"
    if not icon_src.is_file():
        fallback = Path("/usr/share/icons/WhiteSur/apps/scalable/gemini.svg")
        icon_src = fallback if fallback.is_file() else icon_src

    share_src = default_share_dir()
    if app_dir.resolve() != repo_root().resolve():
        app_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(Path(__file__), app_dir / "gemini_desktop.py")
        share_dst = app_dir / "share"
        if share_dst.exists():
            shutil.rmtree(share_dst)
        shutil.copytree(share_src, share_dst)

    installed_py = app_dir / "gemini_desktop.py"
    if not installed_py.is_file():
        installed_py = Path(__file__).resolve()
    real_user_install = prefix.resolve() == (Path.home() / ".local").resolve()

    launcher = bin_dir / "gemini-desktop"
    fixer = bin_dir / "gemini-desktop-fix-login"
    write_wrapper(launcher, installed_py)
    write_wrapper(fixer, installed_py, "--fix-login")

    icon_path = app_dir / "share" / "icons" / "gemini.svg"
    if not icon_path.is_file() and icon_src.is_file():
        icon_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(icon_src, icon_path)

    desktop = applications / "gemini-desktop.desktop"
    write_desktop_file(desktop, launcher, icon_path if icon_path.is_file() else icon_src)

    if real_user_install:
        pwa = Path.home() / ".local" / "share" / "applications" / (
            "chrome-gdfaincndogidkdcdkhapmbffkckdkhn-Default.desktop"
        )
        if pwa.is_file():
            write_desktop_file(pwa, launcher, icon_path if icon_path.is_file() else icon_src)

        plank = Path.home() / ".config" / "plank" / "dock1"
        if plank.is_dir():
            write_plank_item(plank / "launchers" / "gemini-desktop.dockitem", desktop)
            settings = plank / "settings"
            if settings.is_file():
                text = settings.read_text(encoding="utf-8")
                if "gemini-desktop.dockitem" not in text:
                    settings.write_text(
                        text.replace(
                            "DockItems=",
                            "DockItems=gemini-desktop.dockitem;;",
                            1,
                        ),
                        encoding="utf-8",
                    )

    return {
        "app_dir": app_dir,
        "launcher": launcher,
        "desktop": desktop,
    }


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="启动或修复 Gemini 桌面端登录")
    parser.add_argument(
        "--fix-login",
        action="store_true",
        help="强制走 Google 登录回跳（授权成功但应用仍未登录时使用）",
    )
    parser.add_argument(
        "--reset-profile",
        action="store_true",
        help="清空本启动器的 Chrome 配置后重新登录",
    )
    parser.add_argument(
        "--print-url",
        action="store_true",
        help="只打印将要打开的地址，不启动浏览器",
    )
    parser.add_argument(
        "--install",
        action="store_true",
        help="安装到 ~/.local 并写入桌面快捷方式",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if args.install:
        paths = install_user()
        print(f"已安装 Gemini 桌面端: {paths['launcher']}")
        print(f"快捷方式: {paths['desktop']}")
        return 0
    profile_dir = default_profile_dir()
    url = resolve_start_url(args.fix_login or args.reset_profile, profile_dir)
    if args.print_url:
        print(url)
        return 0
    return launch(force_login=args.fix_login, reset=args.reset_profile)


if __name__ == "__main__":
    sys.exit(main())
