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

# 优先走真实二进制，避开 /usr/local/bin/google-chrome 这类会强行
# --user-data-dir 并 --class=google-chrome 的包装脚本，否则独立配置和 Dock 图标都会失效。
LINUX_CHROME_CANDIDATES = (
    "/opt/google/chrome/google-chrome",
    "/usr/bin/google-chrome-stable",
    "/usr/bin/google-chrome",
    "google-chrome-stable",
    "chromium-browser",
    "chromium",
)
MAC_CHROME_CANDIDATES = (
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Google Chrome Beta.app/Contents/MacOS/Google Chrome Beta",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
)

APP_WM_CLASS = "GeminiDesktop"
MISSING_APP_HINT = (
    "找不到 gemini_desktop.py。不要只复制 bin/ 启动脚本；"
    "请把仓库里的整个 gemini-desktop 目录放到 ~/gemini-desktop，"
    "或在 AIDEN 仓库执行 python3 gemini-desktop/gemini_desktop.py --install"
)


def repo_root() -> Path:
    return Path(__file__).resolve().parent


def default_profile_dir() -> Path:
    override = os.environ.get("GEMINI_DESKTOP_PROFILE")
    if override:
        return Path(override).expanduser()
    if sys.platform == "darwin":
        return (
            Path.home()
            / "Library"
            / "Application Support"
            / "gemini-desktop"
            / "chrome-profile"
        )
    return Path.home() / ".local" / "share" / "gemini-desktop" / "chrome-profile"


def is_real_app_script(path: Path) -> bool:
    try:
        text = path.read_text(encoding="utf-8")
    except OSError:
        return False
    return "GEMINI_APP_URL" in text and "def find_chrome" in text


def app_script_candidates(start: Path | None = None) -> list[Path]:
    here = (start or Path.cwd()).resolve()
    home = Path.home()
    seen: set[Path] = set()
    ordered: list[Path] = []
    for item in (
        here / "gemini_desktop.py",
        here.parent / "gemini_desktop.py",
        here / "gemini-desktop" / "gemini_desktop.py",
        here.parent / "gemini-desktop" / "gemini_desktop.py",
        home / "gemini-desktop" / "gemini_desktop.py",
        home / "gemini-desktop" / "gemini-desktop" / "gemini_desktop.py",
        home / ".local" / "share" / "gemini-desktop" / "gemini_desktop.py",
        home / "Library" / "Application Support" / "gemini-desktop" / "gemini_desktop.py",
    ):
        resolved = item.resolve() if item.exists() else item
        if resolved not in seen:
            seen.add(resolved)
            ordered.append(item)
    return ordered


def find_app_script(start: Path | None = None) -> Path:
    files = [path for path in app_script_candidates(start) if path.is_file()]
    for path in files:
        if is_real_app_script(path):
            return path
    if files:
        return files[0]
    raise FileNotFoundError(MISSING_APP_HINT)


def chrome_candidates() -> tuple[str, ...]:
    if sys.platform == "darwin":
        return MAC_CHROME_CANDIDATES + LINUX_CHROME_CANDIDATES
    return LINUX_CHROME_CANDIDATES + MAC_CHROME_CANDIDATES


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
    for name in chrome_candidates():
        found = shutil.which(name) if "/" not in name else (name if Path(name).is_file() else None)
        if not found:
            continue
        resolved = str(Path(found).resolve())
        if resolved == "/usr/local/bin/google-chrome":
            continue
        return found
    raise FileNotFoundError(
        "未找到 Google Chrome。macOS 请安装 /Applications/Google Chrome.app，"
        "Linux 请安装 google-chrome。"
    )


def extra_chrome_flags() -> list[str]:
    """容器环境里 Chrome 沙箱通常不可用，需要补上与系统 Chrome 一致的启动参数。"""
    flags = [item for item in os.environ.get("GEMINI_DESKTOP_CHROME_FLAGS", "").split() if item]
    in_container = Path("/.dockerenv").exists() or Path("/run/.containerenv").exists()
    if in_container:
        for flag in (
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--password-store=basic",
            "--use-gl=angle",
            "--use-angle=swiftshader-webgl",
            "--test-type",
        ):
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
        f"--class={APP_WM_CLASS}",
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


def discovery_wrapper_text(extra_args: str = "") -> str:
    extra_line = f'EXTRA="{extra_args}"\n' if extra_args else 'EXTRA=""\n'
    return (
        "#!/usr/bin/env bash\n"
        "set -euo pipefail\n"
        'HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"\n'
        + extra_line
        + """CANDIDATES=(
  "$HERE/gemini_desktop.py"
  "$HERE/../gemini_desktop.py"
  "$HERE/../gemini-desktop/gemini_desktop.py"
  "$HOME/gemini-desktop/gemini_desktop.py"
  "$HOME/gemini-desktop/gemini-desktop/gemini_desktop.py"
  "$HOME/.local/share/gemini-desktop/gemini_desktop.py"
  "$HOME/Library/Application Support/gemini-desktop/gemini_desktop.py"
)
PY=""
for c in "${CANDIDATES[@]}"; do
  if [ -f "$c" ] && grep -q "GEMINI_APP_URL" "$c" 2>/dev/null; then
    PY="$c"
    break
  fi
done
if [ -z "$PY" ]; then
  for c in "${CANDIDATES[@]}"; do
    if [ -f "$c" ]; then
      PY="$c"
      break
    fi
  done
fi
if [ -z "$PY" ]; then
  echo "找不到 gemini_desktop.py。不要只复制 bin/ 启动脚本。" >&2
  echo "请把仓库里的整个 gemini-desktop 目录放到 ~/gemini-desktop，" >&2
  echo "或执行：python3 gemini-desktop/gemini_desktop.py --install" >&2
  exit 1
fi
command -v python3 >/dev/null 2>&1 || { echo "需要 python3。" >&2; exit 1; }
exec python3 "$PY" $EXTRA "$@"
"""
    )


def write_wrapper(dest: Path, target_py: Path | None = None, extra_args: str = "") -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(discovery_wrapper_text(extra_args), encoding="utf-8")
    dest.chmod(dest.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)


def write_command_file(dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(
        "#!/bin/bash\n"
        "set -euo pipefail\n"
        'HERE="$(cd "$(dirname "$0")" && pwd)"\n'
        'if [ -x "$HERE/bin/gemini-desktop" ]; then exec "$HERE/bin/gemini-desktop" "$@"; fi\n'
        'if [ -f "$HERE/gemini_desktop.py" ]; then exec python3 "$HERE/gemini_desktop.py" "$@"; fi\n'
        'if [ -f "$HERE/gemini-desktop/gemini_desktop.py" ]; then\n'
        '  exec python3 "$HERE/gemini-desktop/gemini_desktop.py" "$@"\n'
        "fi\n"
        'echo "找不到 Gemini 桌面端主程序。请先运行：python3 gemini-desktop/gemini_desktop.py --install"\n'
        "read -r _\n"
        "exit 1\n",
        encoding="utf-8",
    )
    dest.chmod(dest.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)


def copy_package(dest: Path) -> Path:
    dest.mkdir(parents=True, exist_ok=True)
    shutil.copy2(Path(__file__), dest / "gemini_desktop.py")
    share_src = default_share_dir()
    share_dst = dest / "share"
    if share_dst.exists():
        shutil.rmtree(share_dst)
    shutil.copytree(share_src, share_dst)
    command_src = repo_root() / "启动 Gemini 桌面端.command"
    if command_src.is_file():
        shutil.copy2(command_src, dest / "启动 Gemini 桌面端.command")
        (dest / "启动 Gemini 桌面端.command").chmod(
            (dest / "启动 Gemini 桌面端.command").stat().st_mode | stat.S_IXUSR
        )
    write_wrapper(dest / "bin" / "gemini-desktop")
    write_wrapper(dest / "bin" / "gemini-desktop-fix-login", extra_args="--fix-login")
    return dest / "gemini_desktop.py"


def write_root_shim(dest: Path) -> None:
    dest.write_text(
        "#!/usr/bin/env python3\n"
        "from pathlib import Path\n"
        "import runpy, sys\n"
        "target = Path(__file__).resolve().parent / 'gemini-desktop' / 'gemini_desktop.py'\n"
        "if not target.is_file():\n"
        "    sys.stderr.write('找不到 gemini-desktop/gemini_desktop.py\\n')\n"
        "    raise SystemExit(1)\n"
        "sys.argv[0] = str(target)\n"
        "runpy.run_path(str(target), run_name='__main__')\n",
        encoding="utf-8",
    )
    dest.chmod(dest.stat().st_mode | stat.S_IXUSR)


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
StartupWMClass=GeminiDesktop
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


def looks_like_aiden_repo(path: Path) -> bool:
    return (path / ".git").exists() and (path / "AGENTS.md").is_file()


def install_macos() -> dict[str, Path]:
    """把完整程序装到 ~/gemini-desktop，避免只剩启动脚本却找不到主文件。"""
    support = Path.home() / "Library" / "Application Support" / "gemini-desktop"
    home_dir = Path.home() / "gemini-desktop"
    installed_py = copy_package(support)
    if looks_like_aiden_repo(home_dir):
        write_root_shim(home_dir / "gemini_desktop.py")
        write_command_file(home_dir / "启动 Gemini 桌面端.command")
        launcher = home_dir / "gemini-desktop" / "bin" / "gemini-desktop"
    else:
        installed_py = copy_package(home_dir)
        launcher = home_dir / "bin" / "gemini-desktop"
        write_command_file(home_dir / "启动 Gemini 桌面端.command")
    desktop_cmd = Path.home() / "Desktop" / "启动 Gemini 桌面端.command"
    write_command_file(desktop_cmd)
    local_bin = Path.home() / ".local" / "bin"
    write_wrapper(local_bin / "gemini-desktop")
    write_wrapper(local_bin / "gemini-desktop-fix-login", extra_args="--fix-login")
    return {
        "app_dir": home_dir if home_dir.exists() else support,
        "launcher": launcher if launcher.is_file() else local_bin / "gemini-desktop",
        "desktop": desktop_cmd,
        "script": installed_py,
    }


def install_user(prefix: Path | None = None) -> dict[str, Path]:
    """把启动器安装到用户目录，并替换坏掉的 PWA 快捷方式。"""
    if prefix is None and sys.platform == "darwin":
        return install_macos()
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
        help="安装到本机：Linux 写入 ~/.local，macOS 写入 ~/gemini-desktop",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if args.install:
        paths = install_user()
        print(f"已安装 Gemini 桌面端: {paths['launcher']}")
        desktop = paths.get("desktop")
        if desktop:
            print(f"快捷方式: {desktop}")
        script = paths.get("script")
        if script:
            print(f"主程序: {script}")
        return 0
    profile_dir = default_profile_dir()
    url = resolve_start_url(args.fix_login or args.reset_profile, profile_dir)
    if args.print_url:
        print(url)
        return 0
    return launch(force_login=args.fix_login, reset=args.reset_profile)


if __name__ == "__main__":
    sys.exit(main())
