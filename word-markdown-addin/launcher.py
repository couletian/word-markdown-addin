#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown 工具箱 · 单文件启动器
================================================================
打包为 exe 后双击即可：
  1. 自动生成自签名 HTTPS 证书（纯 Python，无需 openssl）
  2. 启动本地 HTTPS 静态服务器
  3. 自动把加载项注册到 Word（写注册表 + 复制清单）
  4. 打开浏览器预览面板
  5. 托盘式常驻，Ctrl+C 或关闭窗口退出

用法：
  Markdown工具箱.exe                # 默认端口 3000，自动注册并打开浏览器
  Markdown工具箱.exe --port 3001
  Markdown工具箱.exe --no-open      # 不自动打开浏览器
  Markdown工具箱.exe --no-install   # 跳过 Word 注册
  Markdown工具箱.exe --uninstall    # 从 Word 卸载加载项后退出
  Markdown工具箱.exe --http         # 仅 HTTP（浏览器预览用，Word 不支持）
"""
from __future__ import annotations

import argparse
import atexit
import datetime
import http.server
import io
import json
import os
import re
import socket
import socketserver
import sys
import threading
import time
import webbrowser
import zipfile
from pathlib import Path

# ------------------------------------------------------------------ #
# 路径解析：兼容「PyInstaller 单文件」与「源码运行」两种模式
# ------------------------------------------------------------------ #

APP_NAME = "Markdown工具箱"
FROZEN = getattr(sys, "frozen", False)

if FROZEN:
    # 单文件模式下资源解压在 sys._MEIPASS
    BUNDLE = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
else:
    BUNDLE = Path(__file__).resolve().parent

SRC_DIR = BUNDLE / "src"
ASSETS_DIR = BUNDLE / "assets"
MANIFEST = BUNDLE / "manifest.xml"

# 证书持久化到用户目录，避免每次重启重新生成
CERT_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "MarkdownToolbox" / "certs"
KEY_PATH = CERT_DIR / "key.pem"
CERT_PATH = CERT_DIR / "cert.pem"

# Word 加载项目录：存放 manifest.xml 供 Office 读取
CATALOG_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "MarkdownToolbox" / "catalog"

# 同时兼容的 Office 版本分支
WEF_VERSIONS = ("16.0", "15.0")

# 「共享文件夹」目录条目的固定 GUID：重复安装不会堆积垃圾键
CATALOG_GUID = "{7C4E9B21-3F6A-4D8E-9A15-2B7C0D5E8F31}"

# 清单解析失败时的兜底 GUID（正常情况下从 manifest.xml 读取）
MANIFEST_GUID_FALLBACK = "c572bc3f-65ba-4b89-9360-f68811b63a3b"


# ------------------------------------------------------------------ #
# 控制台输出
# ------------------------------------------------------------------ #

class C:
    OK = "\033[92m"
    WARN = "\033[93m"
    ERR = "\033[91m"
    DIM = "\033[90m"
    BOLD = "\033[1m"
    END = "\033[0m"


def _enable_ansi() -> None:
    """在 Windows 控制台启用 ANSI 颜色，并把输出编码切到 UTF-8。"""
    # 1) 强制 stdout/stderr 使用 UTF-8，避免 GBK 控制台无法输出 ⇄ 等字符
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if stream is not None and hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass

    if os.name != "nt":
        return

    # 2) 设置控制台代码页为 UTF-8
    try:
        import ctypes

        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        ctypes.windll.kernel32.SetConsoleCP(65001)
    except Exception:
        pass

    # 3) 启用 ANSI 转义序列
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        for handle in (-11, -12):  # stdout, stderr
            mode = ctypes.c_uint32()
            if kernel32.GetConsoleMode(kernel32.GetStdHandle(handle), ctypes.byref(mode)):
                kernel32.SetConsoleMode(kernel32.GetStdHandle(handle), mode.value | 0x0004)
    except Exception:
        pass


def say(msg: str = "", color: str = "") -> None:
    line = f"{color}{msg}{C.END}" if color else msg
    try:
        print(line, flush=True)
    except UnicodeEncodeError:
        # 极端情况下的兜底：降级为 ASCII 安全输出
        enc = getattr(sys.stdout, "encoding", None) or "ascii"
        print(line.encode(enc, errors="replace").decode(enc, errors="replace"), flush=True)


# ------------------------------------------------------------------ #
# 证书生成（纯 Python，不依赖 openssl）
# ------------------------------------------------------------------ #

def ensure_certificate() -> tuple[bytes, bytes]:
    """生成或复用自签名证书，返回 (cert_pem, key_pem)。"""
    if CERT_PATH.exists() and KEY_PATH.exists():
        return CERT_PATH.read_bytes(), KEY_PATH.read_bytes()

    try:
        from cryptography import x509
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import rsa
        from cryptography.x509.oid import NameOID
    except ImportError:
        say("  [!] 缺少 cryptography 库，无法生成证书", C.WARN)
        say("      请改用 node server.js，或运行：pip install cryptography", C.DIM)
        raise SystemExit(1)

    say("  正在生成自签名证书…", C.DIM)
    CERT_DIR.mkdir(parents=True, exist_ok=True)

    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COMMON_NAME, "localhost"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "Markdown Toolbox"),
    ])
    now = datetime.datetime.now(datetime.timezone.utc)
    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(issuer)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - datetime.timedelta(days=1))
        .not_valid_after(now + datetime.timedelta(days=825))
        .add_extension(
            x509.SubjectAlternativeName([
                x509.DNSName("localhost"),
                x509.IPAddress(__import__("ipaddress").ip_address("127.0.0.1")),
            ]),
            critical=False,
        )
        .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
        .sign(key, hashes.SHA256())
    )

    key_pem = key.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.TraditionalOpenSSL,
        encryption_algorithm=serialization.NoEncryption(),
    )
    cert_pem = cert.public_bytes(serialization.Encoding.PEM)
    KEY_PATH.write_bytes(key_pem)
    CERT_PATH.write_bytes(cert_pem)

    # 尝试把证书加入当前用户受信任根，避免浏览器告警
    # 注意：必须传 Windows 原生路径（含空格时尤其重要），并设超时防止挂起
    if os.name == "nt":
        try:
            import subprocess

            subprocess.run(
                ["certutil", "-user", "-addstore", "-f", "Root", str(CERT_PATH)],
                capture_output=True,
                timeout=15,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception:
            pass

    say("  [OK] 证书已生成并信任", C.OK)
    return cert_pem, key_pem


# ------------------------------------------------------------------ #
# 静态文件服务
# ------------------------------------------------------------------ #

MIME = {
    ".html": "text/html; charset=utf-8",
    ".htm": "text/html; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".svg": "image/svg+xml",
    ".ico": "image/x-icon",
    ".md": "text/markdown; charset=utf-8",
    ".xml": "application/xml; charset=utf-8",
    ".woff2": "font/woff2",
}


class AddinHandler(http.server.BaseHTTPRequestHandler):
    """静态文件处理器，限定在 src/ 与 assets/ 目录内。"""

    server_version = "MarkdownToolbox/1.0"
    roots: list[Path] = []

    def log_message(self, fmt, *args):  # 静音访问日志
        pass

    def _resolve(self, pathname: str) -> Path | None:
        clean = pathname.split("?")[0]
        try:
            from urllib.parse import unquote

            clean = unquote(clean)
        except Exception:
            pass
        if clean in ("", "/"):
            clean = "/taskpane.html"

        rel = clean.lstrip("/")
        for root in self.roots:
            candidate = (root / rel).resolve()
            try:
                candidate.relative_to(root.resolve())
            except ValueError:
                continue  # 防目录穿越
            if candidate.is_file():
                return candidate
        return None

    def do_GET(self):  # noqa: N802
        file = self._resolve(self.path)
        if file is None:
            self.send_response(404)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.end_headers()
            self.wfile.write(f"404 Not Found: {self.path}".encode("utf-8"))
            return
        try:
            data = file.read_bytes()
        except OSError as exc:
            self.send_response(500)
            self.end_headers()
            self.wfile.write(f"500 {exc}".encode("utf-8"))
            return
        self.send_response(200)
        self.send_header("Content-Type", MIME.get(file.suffix.lower(), "application/octet-stream"))
        self.send_header("Cache-Control", "no-store")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_HEAD(self):  # noqa: N802
        file = self._resolve(self.path)
        self.send_response(200 if file else 404)
        if file:
            self.send_header("Content-Type", MIME.get(file.suffix.lower(), "application/octet-stream"))
        self.end_headers()


class ReusableTCPServer(socketserver.ThreadingTCPServer):
    allow_reuse_address = True
    daemon_threads = True


def start_server(port: int, use_https: bool) -> ReusableTCPServer:
    AddinHandler.roots = [SRC_DIR, ASSETS_DIR, BUNDLE]
    if use_https:
        import ssl

        cert_pem, key_pem = ensure_certificate()
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        # 用内存中的证书，避免临时文件权限问题
        import tempfile

        tmp = Path(tempfile.mkdtemp(prefix="mdtoolbox_"))
        (tmp / "c.pem").write_bytes(cert_pem)
        (tmp / "k.pem").write_bytes(key_pem)
        ctx.load_cert_chain(certfile=str(tmp / "c.pem"), keyfile=str(tmp / "k.pem"))
        httpd = ReusableTCPServer(("127.0.0.1", port), AddinHandler)
        httpd.socket = ctx.wrap_socket(httpd.socket, server_side=True)
    else:
        httpd = ReusableTCPServer(("127.0.0.1", port), AddinHandler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    return httpd


# ------------------------------------------------------------------ #
# Word 加载项注册
# ------------------------------------------------------------------ #

def _read_manifest_id() -> str:
    """从 manifest.xml 里取 <Id>，取不到则用兜底值。"""
    try:
        text = MANIFEST.read_text(encoding="utf-8")
        m = re.search(r"<Id>\s*([^<\s]+)\s*</Id>", text)
        if m:
            return m.group(1).strip()
    except OSError:
        pass
    return MANIFEST_GUID_FALLBACK


def _find_unc_share_root() -> tuple[str, str] | None:
    """找一个能覆盖 CATALOG_DIR 的既有共享，返回 (UNC 根, 本地路径)。

    优先读注册表里的共享定义（最可靠），失败再回退到解析 `net share`。
    """
    try:
        target = CATALOG_DIR.resolve()
    except OSError:
        return None

    candidates: list[tuple[str, str]] = []

    # 方式一：读 LanmanServer 的共享定义
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SYSTEM\CurrentControlSet\Services\LanmanServer\Shares",
            0,
            winreg.KEY_READ,
        ) as key:
            _, n_val, _ = winreg.QueryInfoKey(key)
            for i in range(n_val):
                _, data, _ = winreg.EnumValue(key, i)
                fields = {}
                for item in data if isinstance(data, (list, tuple)) else [data]:
                    for part in str(item).split("\0"):
                        if "=" in part:
                            k, _, v = part.partition("=")
                            fields[k.strip()] = v.strip()
                share_name = fields.get("ShareName") or ""
                local_path = fields.get("Path") or ""
                if share_name and local_path and not share_name.endswith("$"):
                    candidates.append((share_name, local_path))
    except Exception:
        pass

    # 方式二：解析 `net share`
    if not candidates and os.name == "nt":
        try:
            import subprocess

            out = subprocess.run(
                ["net", "share"],
                capture_output=True,
                timeout=15,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            ).stdout
            for enc in ("gbk", "utf-8", "latin-1"):
                try:
                    text = out.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                text = out.decode("latin-1", errors="replace")

            for line in text.splitlines():
                parts = line.split()
                # 形如：Users  C:\Users
                if len(parts) >= 2 and re.match(r"^[A-Za-z]:\\", parts[1]):
                    if not parts[0].endswith("$"):
                        candidates.append((parts[0], parts[1]))
        except Exception:
            pass

    # 挑路径最长（最具体）且能覆盖目标目录的共享
    host = socket.gethostname()
    best: tuple[str, str] | None = None
    for share_name, local_path in candidates:
        try:
            base = Path(local_path).resolve()
        except OSError:
            continue
        try:
            rel = target.relative_to(base)
        except ValueError:
            continue
        if best is None or len(str(base)) > len(str(Path(best[1]))):
            best = (f"\\\\{host}\\{share_name}\\{rel}", str(base))
    return best


def _write_reg(path: str, name: str, value, reg_type) -> bool:
    """写一个注册表值，返回是否成功。"""
    try:
        import winreg

        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_WRITE) as key:
            winreg.SetValueEx(key, name, 0, reg_type, value)
        return True
    except OSError:
        return False


def _delete_reg_value(path: str, name: str) -> None:
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_WRITE) as key:
            winreg.DeleteValue(key, name)
    except OSError:
        pass


def _delete_reg_tree(path: str) -> None:
    try:
        import winreg

        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
    except OSError:
        pass


def install_addin() -> bool:
    """把加载项注册到 Word，返回是否至少成功一条通道。

    采用两条通道，任一成功即可让加载项出现：
      通道 A「共享文件夹」：WEF\\TrustedCatalogs\\{GUID}
              Url 必须是 UNC 网络路径，因此复用系统已有的共享。
      通道 B「开发者直挂」：WEF\\Developer 下以加载项 GUID 为值名、清单路径为数据。
              微软官方工具 office-addin-dev-settings 用的就是这条，允许本地路径。
    """
    if os.name != "nt":
        say("  [!] 非 Windows 系统，跳过 Word 注册", C.WARN)
        return False
    if not MANIFEST.exists():
        say(f"  [X] 找不到清单文件：{MANIFEST}", C.ERR)
        return False

    try:
        import winreg
    except ImportError:
        return False

    CATALOG_DIR.mkdir(parents=True, exist_ok=True)
    target = CATALOG_DIR / "manifest.xml"
    target.write_bytes(MANIFEST.read_bytes())

    manifest_id = _read_manifest_id()
    manifest_path = str(target.resolve())
    ok_a = ok_b = False

    for version in WEF_VERSIONS:
        wef = rf"Software\Microsoft\Office\{version}\WEF"

        # ---------- 通道 B：Developer 直挂清单 ----------
        dev_key = rf"{wef}\Developer"
        if _write_reg(dev_key, manifest_id, manifest_path, winreg.REG_SZ):
            ok_b = True
        # 通知 Office 启动时刷新加载项
        _write_reg(dev_key, "RefreshAddins", 1, winreg.REG_DWORD)
        # 清理旧版本脚本写错的残留值
        _delete_reg_value(dev_key, "TrustedCatalogs")

        # ---------- 通道 A：共享文件夹目录 ----------
        found = _find_unc_share_root()
        if found:
            unc_url = found[0]
            cat_key = rf"{wef}\TrustedCatalogs\{CATALOG_GUID}"
            wrote = True
            wrote &= _write_reg(cat_key, "Id", CATALOG_GUID, winreg.REG_SZ)
            wrote &= _write_reg(cat_key, "Url", unc_url, winreg.REG_SZ)
            wrote &= _write_reg(cat_key, "Flags", 1, winreg.REG_DWORD)
            if wrote:
                ok_a = True

    # ---------- 结果汇报 ----------
    if ok_b:
        say(f"  [OK] 已挂载到 Word（开发者通道 · 清单：{manifest_path}）", C.OK)
    if ok_a:
        found = _find_unc_share_root()
        say(f"  [OK] 已登记共享文件夹目录（{found[0] if found else 'UNC'}）", C.OK)
    if not ok_a and not ok_b:
        say("  [!] 注册表写入失败，请在 Word 中手动加载清单", C.WARN)
    return ok_a or ok_b


def uninstall_addin() -> None:
    import shutil

    if CATALOG_DIR.exists():
        shutil.rmtree(CATALOG_DIR, ignore_errors=True)
        say(f"  [OK] 已移除 {CATALOG_DIR}", C.OK)

    manifest_id = _read_manifest_id()
    for version in WEF_VERSIONS:
        wef = rf"Software\Microsoft\Office\{version}\WEF"
        _delete_reg_value(rf"{wef}\Developer", manifest_id)
        _delete_reg_value(rf"{wef}\Developer", "TrustedCatalogs")
        _delete_reg_value(rf"{wef}\Developer", "RefreshAddins")
        _delete_reg_tree(rf"{wef}\TrustedCatalogs\{CATALOG_GUID}")
    say("  [OK] 已清理注册表项", C.OK)

    wef_dir = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "Microsoft" / "Office" / "16.0" / "Wef"
    if wef_dir.exists():
        for child in wef_dir.iterdir():
            dev = child / "DeveloperSettings"
            if dev.exists():
                shutil.rmtree(dev, ignore_errors=True)
    say("  [OK] 卸载完成，请重启 Word", C.OK)


# ------------------------------------------------------------------ #
# 自检
# ------------------------------------------------------------------ #

def preflight() -> bool:
    """启动前检查必需资源。"""
    ok = True
    for name, path in [("src/", SRC_DIR), ("assets/", ASSETS_DIR), ("manifest.xml", MANIFEST)]:
        if not path.exists():
            say(f"  [X] 缺少 {name}：{path}", C.ERR)
            ok = False
    if ok and not (SRC_DIR / "taskpane.html").exists():
        say(f"  [X] 缺少任务面板：{SRC_DIR / 'taskpane.html'}", C.ERR)
        ok = False
    return ok


# ------------------------------------------------------------------ #
# 主流程
# ------------------------------------------------------------------ #

BANNER = r"""
  __  __            _          _            _      _____           _ _
 |  \/  |          | |        | |          | |    |_   _|         | | |
 | \  / | __ _ _ __| | ___   _| |__   ___  | |__    | | ___   ___ | | |__   _____  __
 | |\/| |/ _` | '__| |/ / | | | '_ \ / _ \ | '_ \   | |/ _ \ / _ \| | '_ \ / _ \ \/ /
 | |  | | (_| | |  |   <| |_| | |_) | (_) || |_) |  | | (_) | (_) | | |_) | (_) >  <
 |_|  |_|\__,_|_|  |_|\_\\__,_|_.__/ \___/ |_.__/   |_|\___/ \___/|_|_.__/ \___/_/\_\

 Word ⇄ Markdown 双向转换 · 本地加载项
"""


def main() -> int:
    _enable_ansi()
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--port", type=int, default=3000)
    parser.add_argument("--http", action="store_true", help="仅 HTTP 模式（浏览器预览）")
    parser.add_argument("--no-open", action="store_true", help="不自动打开浏览器")
    parser.add_argument("--no-install", action="store_true", help="跳过 Word 注册")
    parser.add_argument("--uninstall", action="store_true", help="从 Word 卸载加载项")
    parser.add_argument("-h", "--help", action="help")
    args = parser.parse_args()

    if args.uninstall:
        say("\n正在卸载 Markdown 工具箱…\n", C.BOLD)
        uninstall_addin()
        return 0

    say(BANNER, C.DIM)
    say("=" * 78, C.DIM)

    if not preflight():
        say("\n资源缺失，无法启动。", C.ERR)
        if FROZEN:
            say("exe 可能不完整，请重新下载。", C.DIM)
        input("\n按回车键退出…")
        return 1

    # 1) 注册到 Word
    if not args.no_install:
        say("\n[1/4] 注册加载项到 Word", C.BOLD)
        install_addin()
    else:
        say("\n[1/4] 跳过 Word 注册（--no-install）", C.DIM)

    # 2) 启动服务器
    say("\n[2/4] 启动本地服务器", C.BOLD)
    scheme = "http" if args.http else "https"
    try:
        httpd = start_server(args.port, not args.http)
    except OSError as exc:
        say(f"  [X] 端口 {args.port} 被占用或无法绑定：{exc}", C.ERR)
        say(f"      换一个端口：{APP_NAME}.exe --port 3001", C.DIM)
        input("\n按回车键退出…")
        return 1
    base = f"{scheme}://localhost:{args.port}"
    say(f"  [OK] 服务已启动：{base}", C.OK)
    if args.http:
        say("  [!] HTTP 模式仅供浏览器预览，Word 加载项必须用 HTTPS", C.WARN)

    atexit.register(httpd.shutdown)

    # 3) 打开浏览器
    say("\n[3/4] 打开预览", C.BOLD)
    if args.no_open:
        say(f"  已跳过，手动访问：{base}/taskpane.html", C.DIM)
    else:
        webbrowser.open(f"{base}/taskpane.html")
        say(f"  [OK] 已在浏览器打开 {base}/taskpane.html", C.OK)

    # 4) 使用指引
    say("\n[4/4] 在 Word 中使用", C.BOLD)
    say("  1. 完全关闭 Word（所有窗口），然后重新打开一个文档", "")
    say("  2. 功能区「开始」选项卡最右侧应出现「Markdown 工具箱」组", C.OK)
    say("  3. 若没出现：点击「插入」→「我的加载项」→「共享文件夹」", "")
    say("     选择「Markdown 工具箱」→「添加」", "")
    say("  功能区按钮：导出 Markdown / 粘贴 Markdown / 打开面板", C.DIM)

    say("\n" + "=" * 78, C.DIM)
    say(f"  服务运行中 · {base}", C.OK)
    say("  关闭此窗口或按 Ctrl+C 即可停止", C.DIM)
    say("=" * 78 + "\n", C.DIM)

    try:
        while True:
            time.sleep(3600)
    except KeyboardInterrupt:
        say("\n正在停止服务…", C.DIM)
        httpd.shutdown()
        say("已退出。", C.OK)
    return 0


if __name__ == "__main__":
    sys.exit(main())
