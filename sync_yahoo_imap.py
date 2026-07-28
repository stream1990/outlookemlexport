#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通过 IMAP 同步 Yahoo 邮箱（默认 yahoo.fr / imap.mail.yahoo.com）全部邮件为 .eml

目录结构、文件名规则与 export_outlook_eml.py 一致：
  <输出根目录>/<账号邮箱>/<文件夹层级>/YYYY-mm-dd_HHMMSS_主题_短id.eml

特性：
- 账号、密码参数登录
- 遍历全部可选择的邮件文件夹
- 默认从最近邮件开始同步（UID 从大到小）
- 断点续传（按 文件夹 + UIDVALIDITY + UID）
- 使用 BODY.PEEK 下载，尽量不改变服务器已读状态
- 中断/断线可重跑继续

用法示例：
  python sync_yahoo_imap.py -u you@yahoo.fr -p "应用专用密码"
  python sync_yahoo_imap.py -u you@yahoo.fr -p "xxx" --limit 200
  python sync_yahoo_imap.py -u you@yahoo.fr -p "xxx" --since 2024-01-01
  python sync_yahoo_imap.py -u you@yahoo.fr -p "xxx" --oldest-first
  python sync_yahoo_imap.py -u you@yahoo.fr -p "xxx" --list-folders
"""

from __future__ import annotations

import argparse
import email
import email.header
import hashlib
import imaplib
import json
import re
import socket
import ssl
import sys
import time
import traceback
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple


# 与 export_outlook_eml.py 保持一致
INVALID_FS_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
WHITESPACE_RUN = re.compile(r"\s+")

DEFAULT_HOST = "imap.mail.yahoo.com"
DEFAULT_PORT = 993

# 全局调试开关（由 -v / 异常路径打开）
VERBOSE = False


def log(msg: str = "", *, err: bool = False) -> None:
    stream = sys.stderr if err else sys.stdout
    print(msg, file=stream, flush=True)


def vlog(msg: str) -> None:
    if VERBOSE:
        log(f"[DEBUG] {msg}")


def explain_imap_error(exc: BaseException) -> str:
    """把常见 IMAP/网络异常翻译成人话。"""
    text = str(exc) or repr(exc)
    low = text.lower()
    hints: List[str] = [f"原始错误: {type(exc).__name__}: {text}"]

    if isinstance(exc, imaplib.IMAP4.error) or "authentication" in low or "login" in low:
        if any(k in low for k in ("invalid", "fail", "auth", "login", "password", "credentials")):
            hints.append(
                "判断: 认证失败（账号/密码不对，或未使用「应用专用密码」）。"
            )
            hints.append(
                "Yahoo 若开了两步验证，网页密码不能直接 IMAP，需在账号安全里生成 App Password。"
            )
        if "application-specific" in low or "app password" in low:
            hints.append("判断: 服务器明确要求应用专用密码。")

    if any(k in low for k in ("limit", "throttle", "rate", "too many", "slow down", "try again")):
        hints.append("判断: 可能触发限流/频率限制，稍等 10–30 分钟再试，或加大间隔。")

    if any(k in low for k in ("unavailable", "temporarily", "timeout", "timed out", "connection reset")):
        hints.append("判断: 网络超时或服务器临时不可用，可重试。")

    if any(k in low for k in ("ssl", "certificate", "handshake")):
        hints.append("判断: SSL/证书问题。")

    if "readonly" in low or "selected" in low:
        hints.append("判断: 文件夹状态异常。")

    if isinstance(exc, (socket.timeout, TimeoutError)):
        hints.append("判断: 连接超时，检查网络/代理/防火墙是否拦了 993 端口。")

    if isinstance(exc, ConnectionError) or "10054" in text or "10060" in text:
        hints.append("判断: TCP 连接被重置或连不上主机。")

    return "\n".join(f"  · {h}" for h in hints)


def safe_name(name: str, max_len: int = 80) -> str:
    name = (name or "").strip() or "untitled"
    name = INVALID_FS_CHARS.sub("_", name)
    name = WHITESPACE_RUN.sub(" ", name).strip(" .")
    if not name:
        name = "untitled"
    if len(name) > max_len:
        name = name[:max_len].rstrip(" .")
    return name


def short_id(text: str, n: int = 10) -> str:
    return hashlib.sha1(text.encode("utf-8", errors="replace")).hexdigest()[:n]


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def decode_mime_header(raw: Optional[str]) -> str:
    if not raw:
        return ""
    parts = email.header.decode_header(raw)
    out: List[str] = []
    for chunk, charset in parts:
        if isinstance(chunk, bytes):
            for cs in (charset, "utf-8", "latin-1", "gb18030"):
                if not cs:
                    continue
                try:
                    out.append(chunk.decode(cs, errors="replace"))
                    break
                except Exception:
                    continue
            else:
                out.append(chunk.decode("utf-8", errors="replace"))
        else:
            out.append(str(chunk))
    return "".join(out).strip()


def parse_mail_datetime(msg: email.message.Message, internaldate: Optional[str] = None) -> datetime:
    date_hdr = msg.get("Date")
    if date_hdr:
        try:
            dt = parsedate_to_datetime(date_hdr)
            if dt is not None:
                if dt.tzinfo is not None:
                    dt = dt.astimezone().replace(tzinfo=None)
                return dt
        except Exception:
            pass

    if internaldate:
        # IMAP INTERNALDATE 形如: 24-Jul-2026 12:34:56 +0000
        try:
            # imaplib.Internaldate2tuple
            tt = imaplib.Internaldate2tuple(
                f'INTERNALDATE "{internaldate}"'.encode("ascii", errors="ignore")
                if isinstance(internaldate, str)
                else internaldate
            )
            if tt:
                return datetime(*tt[:6])
        except Exception:
            pass
        try:
            # 宽松解析
            cleaned = internaldate.strip().strip('"')
            dt = datetime.strptime(cleaned[:20], "%d-%b-%Y %H:%M:%S")
            return dt
        except Exception:
            pass

    return datetime.now()


def unique_eml_path(folder_dir: Path, base_name: str) -> Path:
    candidate = folder_dir / f"{base_name}.eml"
    if not candidate.exists():
        return candidate
    n = 2
    while True:
        candidate = folder_dir / f"{base_name}_{n}.eml"
        if not candidate.exists():
            return candidate
        n += 1


def mail_filename_from_msg(
    msg: email.message.Message,
    stable_key: str,
    internaldate: Optional[str] = None,
) -> str:
    """与 export_outlook_eml.mail_filename 同格式。"""
    sid = short_id(stable_key, 10)
    dt = parse_mail_datetime(msg, internaldate)
    date_part = dt.strftime("%Y-%m-%d_%H%M%S")
    subject = safe_name(decode_mime_header(msg.get("Subject")) or "(no subject)", max_len=60)
    return f"{date_part}_{subject}_{sid}"


def imap_utf7_decode(name: str) -> str:
    """解码 IMAP modified UTF-7 文件夹名。"""
    try:
        # 常见已是可读 ASCII/UTF-8
        if "&" not in name:
            return name
        # 手工 modified UTF-7
        out = bytearray()
        i = 0
        while i < len(name):
            if name[i] != "&":
                out.extend(name[i].encode("ascii"))
                i += 1
                continue
            j = name.find("-", i)
            if j < 0:
                out.extend(name[i:].encode("ascii", errors="replace"))
                break
            token = name[i + 1 : j]
            i = j + 1
            if token == "":
                out.append(ord("&"))
                continue
            # modified base64: , -> /
            b64 = token.replace(",", "/")
            pad = (-len(b64)) % 4
            b64 += "=" * pad
            import base64

            raw = base64.b64decode(b64.encode("ascii"))
            out.extend(raw.decode("utf-16-be", errors="replace").encode("utf-8", errors="replace"))
        return out.decode("utf-8", errors="replace")
    except Exception:
        return name


def _line_to_text(line: Any) -> str:
    if line is None:
        return ""
    if isinstance(line, tuple):
        # 偶发 (bytes, ...) 
        line = line[0]
    if isinstance(line, bytes):
        return line.decode("utf-8", errors="replace")
    return str(line)


def parse_list_line(line: Any) -> Optional[Tuple[List[str], str, str]]:
    """
    解析 LIST/LSUB 一行。

    注意：imaplib 返回的 data 项通常已经去掉 "* LIST " 前缀，只有：
      (\\HasNoChildren) "/" "INBOX"
    完整协议行则是：
      * LIST (\\HasNoChildren) "/" "INBOX"
    两种都要支持。
    返回 (flags, delimiter, name)
    """
    text = _line_to_text(line).strip()
    if not text:
        return None

    # 去掉 * LIST / * LSUB 前缀（若有）
    text2 = re.sub(r"^\*\s+(?:LIST|LSUB)\s+", "", text, flags=re.IGNORECASE)

    m = re.match(
        r'^\((?P<flags>.*)\)\s+(?P<delim>"[^"]*"|NIL)\s+(?P<name>.*)$',
        text2,
        re.IGNORECASE,
    )
    if not m:
        # 再试：无 flags 括号的宽松形式
        m2 = re.match(
            r'^(?P<delim>"[^"]*"|NIL)\s+(?P<name>.*)$',
            text2,
            re.IGNORECASE,
        )
        if not m2:
            vlog(f"LIST 行无法解析: {text!r}")
            return None
        flags_norm: List[str] = []
        delim_raw = m2.group("delim")
        name_raw = m2.group("name").strip()
    else:
        flags_raw = m.group("flags") or ""
        flags_norm = []
        for f in flags_raw.replace("(", " ").replace(")", " ").split():
            flags_norm.append(f.lstrip("\\"))
        delim_raw = m.group("delim")
        name_raw = m.group("name").strip()

    if delim_raw.upper() == "NIL":
        delim = "/"
    else:
        delim = delim_raw.strip('"') or "/"

    if name_raw.startswith('"') and name_raw.endswith('"') and len(name_raw) >= 2:
        name = name_raw[1:-1]
    else:
        # 字面量 {n} 形式极少见，imaplib 通常已展开
        name = name_raw
    name = name.replace('\\"', '"').replace("\\\\", "\\")
    if not name:
        return None
    return flags_norm, delim, name


class SyncState:
    """按账号 + 文件夹 UIDVALIDITY + UID 断点。"""

    def __init__(self, path: Path, account: str):
        self.path = path
        self.account = account
        # folder_key -> {"uidvalidity": int, "done": [uid_str, ...]}
        self.folders: Dict[str, Dict[str, Any]] = {}
        self.stats: Dict[str, int] = {
            "exported": 0,
            "skipped": 0,
            "failed": 0,
            "folders": 0,
        }
        if path.exists():
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
                if data.get("account", "").lower() == account.lower():
                    self.folders = data.get("folders", {})
                    self.stats.update(data.get("stats", {}))
            except Exception:
                pass

    def get_done(self, folder: str, uidvalidity: int) -> Set[str]:
        info = self.folders.get(folder)
        if not info:
            return set()
        if int(info.get("uidvalidity") or 0) != int(uidvalidity):
            # UIDVALIDITY 变了，旧 UID 作废
            return set()
        return set(str(x) for x in info.get("done", []))

    def mark_done(self, folder: str, uidvalidity: int, uid: str) -> None:
        info = self.folders.get(folder)
        if not info or int(info.get("uidvalidity") or 0) != int(uidvalidity):
            info = {"uidvalidity": int(uidvalidity), "done": []}
            self.folders[folder] = info
        done_list = info.setdefault("done", [])
        uid = str(uid)
        if uid not in done_list:
            # 用 list 便于 JSON；查找时转 set 由 get_done 做
            # 这里为性能在内存再维护 set 更优，但简单起见追加并去重写回
            done_list.append(uid)

    def save(self) -> None:
        # done 列表去重保序
        folders_out: Dict[str, Any] = {}
        for folder, info in self.folders.items():
            seen = set()
            done_unique = []
            for u in info.get("done", []):
                s = str(u)
                if s not in seen:
                    seen.add(s)
                    done_unique.append(s)
            folders_out[folder] = {
                "uidvalidity": int(info.get("uidvalidity") or 0),
                "done": done_unique,
            }
        payload = {
            "account": self.account,
            "folders": folders_out,
            "stats": self.stats,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
        self.path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )


class YahooImapClient:
    def __init__(
        self,
        host: str,
        port: int,
        user: str,
        password: str,
        timeout: int = 120,
    ):
        self.host = host
        self.port = port
        self.user = user
        self.password = password
        self.timeout = timeout
        self.imap: Optional[imaplib.IMAP4_SSL] = None
        self.welcome: str = ""
        self.capabilities: List[str] = []

    def connect(self) -> None:
        self.close()
        socket.setdefaulttimeout(self.timeout)
        log(f"[连接] {self.host}:{self.port}  timeout={self.timeout}s  账号={self.user}")
        t0 = time.time()
        try:
            ctx = ssl.create_default_context()
            try:
                self.imap = imaplib.IMAP4_SSL(
                    self.host, self.port, timeout=self.timeout, ssl_context=ctx
                )
            except TypeError:
                self.imap = imaplib.IMAP4_SSL(self.host, self.port, timeout=self.timeout)
        except Exception as e:
            log(f"[连接失败] 无法建立 SSL 到 {self.host}:{self.port}", err=True)
            log(explain_imap_error(e), err=True)
            raise

        welcome = getattr(self.imap, "welcome", b"") or b""
        if isinstance(welcome, bytes):
            self.welcome = welcome.decode("utf-8", errors="replace")
        else:
            self.welcome = str(welcome)
        log(f"[连接] SSL 成功 ({time.time() - t0:.1f}s)")
        log(f"[欢迎语] {self.welcome.strip()}")

        try:
            typ, data = self.imap.capability()
            caps_raw = b" ".join(x for x in (data or []) if isinstance(x, bytes))
            self.capabilities = caps_raw.decode("ascii", errors="ignore").split()
            log(f"[能力] typ={typ}  count={len(self.capabilities)}")
            vlog("CAPABILITY: " + " ".join(self.capabilities[:80]))
        except Exception as e:
            log(f"[能力] 查询失败: {e}")

        try:
            caps_u = {c.upper() for c in self.capabilities}
            if "ID" in caps_u and hasattr(self.imap, "id"):
                typ, data = self.imap.id_("name", "outlookemlexport", "version", "1.0")
                vlog(f"ID => {typ} {_preview(data)}")
        except Exception as e:
            vlog(f"ID 命令忽略: {e}")

        log("[登录] 正在 LOGIN …")
        try:
            typ, data = self.imap.login(self.user, self.password)
        except imaplib.IMAP4.error as e:
            log("[登录失败] IMAP 拒绝认证", err=True)
            log(explain_imap_error(e), err=True)
            log(
                "建议:\n"
                "  1) Yahoo 安全中心开启 IMAP，并生成「应用专用密码」\n"
                "  2) -u 用完整邮箱，如 name@yahoo.fr\n"
                "  3) 密码含特殊字符时用引号包起来\n"
                "  4) Outlook 能用 ≠ IMAP 密码相同（Outlook 可能是 OAuth）",
                err=True,
            )
            raise
        except Exception as e:
            log("[登录失败] 非预期异常", err=True)
            log(explain_imap_error(e), err=True)
            raise

        log(f"[登录] 成功  typ={typ}  data={_preview(data)}")
        try:
            if getattr(self.imap, "capabilities", None):
                self.capabilities = list(self.imap.capabilities)
                vlog(
                    "login 后 capabilities: "
                    + " ".join(str(c) for c in self.capabilities[:80])
                )
        except Exception:
            pass

    def close(self) -> None:
        if self.imap is not None:
            try:
                self.imap.logout()
                vlog("已 LOGOUT")
            except Exception:
                try:
                    self.imap.shutdown()
                except Exception:
                    pass
            self.imap = None

    def ensure(self) -> imaplib.IMAP4_SSL:
        if self.imap is None:
            self.connect()
        assert self.imap is not None
        return self.imap

    def reconnect(self) -> None:
        log("  [重连] IMAP 连接…")
        time.sleep(2)
        self.connect()

    def list_folders(self) -> List[Tuple[str, str, List[str]]]:
        """返回 [(imap_name, display_name, flags), ...]"""
        imap = self.ensure()
        attempts: List[Tuple[str, str, str]] = [
            ("LIST", '""', "*"),
            ("LIST", '""', "%"),
            ("LSUB", '""', "*"),
        ]

        last_typ = ""
        last_data: Any = None

        for cmd, directory, pattern in attempts:
            log(f"[文件夹] 执行 {cmd} {directory} {pattern} …")
            try:
                if cmd == "LIST":
                    typ, data = imap.list(directory, pattern)
                else:
                    typ, data = imap.lsub(directory, pattern)
            except Exception as e:
                log(f"[文件夹] {cmd} 异常: {e}", err=True)
                log(explain_imap_error(e), err=True)
                continue

            last_typ, last_data = typ, data
            n = 0 if not data else len(data)
            log(f"[文件夹] {cmd} 返回 typ={typ}  原始条目数={n}")

            if data:
                show_n = len(data) if VERBOSE else min(8, len(data))
                for i, line in enumerate(data[:show_n]):
                    log(
                        f"[文件夹] 原始[{i}] type={type(line).__name__}  "
                        f"value={_preview(line, 240)}"
                    )
                if not VERBOSE and len(data) > 8:
                    log(f"[文件夹] … 另有 {len(data) - 8} 条（加 -v 打印全部）")

            if typ == "OK" and data:
                parsed_list = self._parse_folder_lines(data)
                if parsed_list:
                    log(f"[文件夹] {cmd} 解析成功 {len(parsed_list)} 个")
                    return parsed_list
                log(
                    f"[文件夹] {cmd} 有 {n} 条原始数据但解析为 0 —— 格式不兼容，试下一种",
                    err=True,
                )
            elif typ != "OK":
                log(f"[文件夹] {cmd} 非 OK: {_preview(data)}", err=True)
                low = _preview(data).lower()
                if any(k in low for k in ("limit", "rate", "too many", "throttle")):
                    log("  · 判断: 可能被限流", err=True)
                if any(k in low for k in ("auth", "login", "invalid")):
                    log("  · 判断: 会话认证状态异常", err=True)

        log(
            f"[文件夹] LIST/LSUB 均未解析出文件夹 (last typ={last_typ})",
            err=True,
        )
        if last_data is not None:
            log(f"[文件夹] 最后一次 data: {_preview(last_data, 500)}", err=True)
        log("[文件夹] 兜底使用 INBOX（多数账号至少有收件箱）", err=True)
        return [("INBOX", "INBOX", [])]

    def _parse_folder_lines(self, data: List[Any]) -> List[Tuple[str, str, List[str]]]:
        result: List[Tuple[str, str, List[str]]] = []
        seen: Set[str] = set()
        parse_fail = 0
        for line in data:
            if not line:
                continue
            parsed = parse_list_line(line)
            if not parsed:
                parse_fail += 1
                continue
            flags, _delim, name = parsed
            display = imap_utf7_decode(name)
            if name in seen:
                continue
            seen.add(name)
            result.append((name, display, flags))
        if parse_fail:
            log(f"[文件夹] 有 {parse_fail} 行解析失败（加 -v 看 [DEBUG]）")
        return result

    def select_folder(self, imap_name: str) -> int:
        """SELECT 文件夹，返回 UIDVALIDITY。"""
        imap = self.ensure()
        typ, data = imap.select(f'"{imap_name}"', readonly=True)
        vlog(f"SELECT \"{imap_name}\" => {typ} {_preview(data)}")
        if typ != "OK":
            typ, data = imap.select(imap_name, readonly=True)
            vlog(f"SELECT {imap_name} => {typ} {_preview(data)}")
        if typ != "OK":
            raise RuntimeError(f"无法打开文件夹: {imap_name} -> {_preview(data)}")

        typ, data = imap.response("UIDVALIDITY")
        if typ == "OK" and data and data[0] is not None:
            try:
                return int(data[0])
            except Exception:
                pass
        typ, data = imap.status(f'"{imap_name}"', "(UIDVALIDITY)")
        if typ == "OK" and data and data[0]:
            raw = data[0] if isinstance(data[0], bytes) else str(data[0]).encode()
            m = re.search(rb"UIDVALIDITY\s+(\d+)", raw)
            if m:
                return int(m.group(1))
        return 0

    def search_uids(self, since: Optional[str] = None) -> List[str]:
        imap = self.ensure()
        if since:
            try:
                d = datetime.strptime(since, "%Y-%m-%d")
            except ValueError as e:
                raise ValueError(f"--since 需为 YYYY-MM-DD: {since}") from e
            imap_date = d.strftime("%d-%b-%Y")
            vlog(f"UID SEARCH SINCE {imap_date}")
            typ, data = imap.uid("SEARCH", None, "SINCE", imap_date)
        else:
            vlog("UID SEARCH ALL")
            typ, data = imap.uid("SEARCH", None, "ALL")
        if typ != "OK":
            log(f"[SEARCH] 非 OK: typ={typ} data={_preview(data)}", err=True)
            return []
        if not data or data[0] is None:
            return []
        raw = data[0]
        if isinstance(raw, bytes):
            text = raw.decode("ascii", errors="ignore").strip()
        else:
            text = str(raw).strip()
        if not text:
            return []
        return text.split()

    def fetch_eml(self, uid: str) -> Tuple[bytes, Optional[str]]:
        """返回 (rfc822_bytes, internaldate_str_or_None)，BODY.PEEK 尽量不标已读。"""
        imap = self.ensure()
        typ, data = imap.uid("FETCH", uid, "(INTERNALDATE BODY.PEEK[])")
        if typ != "OK" or not data:
            raise RuntimeError(f"FETCH 失败 uid={uid}: {_preview(data)}")

        raw_msg: Optional[bytes] = None
        internaldate: Optional[str] = None

        for part in data:
            if not isinstance(part, tuple) or len(part) < 2:
                if isinstance(part, bytes):
                    m = re.search(rb'INTERNALDATE "([^"]+)"', part)
                    if m:
                        internaldate = m.group(1).decode("ascii", errors="ignore")
                continue
            meta, body = part[0], part[1]
            if isinstance(body, bytes) and len(body) > 0:
                raw_msg = body
            if isinstance(meta, bytes):
                m = re.search(rb'INTERNALDATE "([^"]+)"', meta)
                if m:
                    internaldate = m.group(1).decode("ascii", errors="ignore")

        if raw_msg is None:
            typ, data = imap.uid("FETCH", uid, "(RFC822)")
            if typ != "OK" or not data:
                raise RuntimeError(f"FETCH RFC822 失败 uid={uid}")
            for part in data:
                if isinstance(part, tuple) and len(part) >= 2 and isinstance(part[1], bytes):
                    raw_msg = part[1]
                    break
        if raw_msg is None:
            raise RuntimeError(f"空邮件 uid={uid}")

        return raw_msg, internaldate


def _preview(obj: Any, limit: int = 160) -> str:
    """安全预览调试信息（避免把密码打出来；密码不会出现在 IMAP data 里）。"""
    try:
        if obj is None:
            return "None"
        if isinstance(obj, (list, tuple)):
            parts = []
            for i, x in enumerate(obj[:5]):
                parts.append(_preview(x, limit // 2))
            more = f" …(+{len(obj)-5})" if len(obj) > 5 else ""
            return "[" + "; ".join(parts) + more + "]"
        if isinstance(obj, bytes):
            s = obj.decode("utf-8", errors="replace")
        else:
            s = str(obj)
        s = s.replace("\r", "\\r").replace("\n", "\\n")
        if len(s) > limit:
            return s[:limit] + f"…({len(s)} chars)"
        return s
    except Exception as e:
        return f"<preview-error {e}>"


def folder_local_path(account_dir: Path, display_name: str, delim: str = "/") -> Path:
    """
    把 IMAP 文件夹显示名映射到本地子目录。
    例: INBOX -> INBOX
        [Gmail]/已发邮件 -> [Gmail]/已发邮件
    Yahoo 常见: Inbox, Sent, Bulk, Trash, &xxx- 已解码
    """
    # 统一分隔符
    name = display_name.replace("\\", "/")
    if delim and delim != "/":
        name = name.replace(delim, "/")
    parts = [safe_name(p, max_len=80) for p in name.split("/") if p and p != ""]
    if not parts:
        parts = ["INBOX"]
    path = account_dir
    for p in parts:
        path = path / p
    return path


def sort_uids(uids: List[str], newest_first: bool) -> List[str]:
    """IMAP UID 通常随时间递增；按数值排序实现新/旧优先。"""
    def key(u: str) -> int:
        try:
            return int(u)
        except ValueError:
            return 0

    return sorted(uids, key=key, reverse=newest_first)


def sync_account(
    user: str,
    password: str,
    output_root: Path,
    host: str,
    port: int,
    save_every: int,
    verbose: bool,
    reset_state: bool,
    only_folders: Optional[List[str]],
    skip_folders: Set[str],
    timeout: int,
    newest_first: bool = True,
    limit: Optional[int] = None,
    since: Optional[str] = None,
) -> int:
    account_dir = output_root / safe_name(user, max_len=100)
    ensure_dir(account_dir)

    state_path = output_root / f".imap_state_{safe_name(user, max_len=60)}.json"
    if reset_state and state_path.exists():
        state_path.unlink()
        print(f"已清除断点: {state_path}")

    state = SyncState(state_path, user)
    client = YahooImapClient(host, port, user, password, timeout=timeout)

    order_desc = "最近优先" if newest_first else "最旧优先"
    extra = []
    if since:
        extra.append(f"since={since}")
    if limit:
        extra.append(f"每文件夹最多 {limit} 封")
    log(
        f"[开始] 账号={user}  主机={host}:{port}  顺序={order_desc}"
        + (f"  ({', '.join(extra)})" if extra else "")
        + f"  verbose={VERBOSE}"
    )
    try:
        client.connect()
    except imaplib.IMAP4.error as e:
        log(f"[失败] 登录阶段: {e}", err=True)
        log(explain_imap_error(e), err=True)
        return 1
    except Exception as e:
        log(f"[失败] 连接阶段: {e}", err=True)
        log(explain_imap_error(e), err=True)
        if VERBOSE:
            traceback.print_exc()
        return 1

    try:
        folders = client.list_folders()
        if not folders:
            log("[失败] 未列出任何文件夹（LIST 空且无 INBOX 兜底）。", err=True)
            log(
                "可能原因:\n"
                "  · 登录其实未完全成功 / 账号无邮箱权限\n"
                "  · 被限流，LIST 被拒\n"
                "  · 服务器返回格式异常\n"
                "请加 -v 重跑，把完整日志发出来。",
                err=True,
            )
            return 1

        log(f"[文件夹] 共 {len(folders)} 个将处理：")
        for imap_name, display, flags in folders:
            mark = " [\\Noselect]" if any(f.lower() == "noselect" for f in flags) else ""
            log(f"  - {display}  (imap={imap_name!r}){mark}  flags={flags}")

        only_set = {x.lower() for x in only_folders} if only_folders else None

        for imap_name, display, flags in folders:
            flags_l = {f.lower() for f in flags}
            if "noselect" in flags_l:
                log(f"  [跳过不可选] {display}")
                continue

            if display in skip_folders or imap_name in skip_folders:
                log(f"  [跳过] {display}")
                continue

            if only_set is not None:
                if display.lower() not in only_set and imap_name.lower() not in only_set:
                    continue

            dest_dir = folder_local_path(account_dir, display)
            ensure_dir(dest_dir)
            state.stats["folders"] = state.stats.get("folders", 0) + 1

            # 打开文件夹 + 拉 UID 列表（失败则重连重试）
            uidvalidity = 0
            uids: List[str] = []
            for attempt in range(1, 4):
                try:
                    log(f"  [SELECT] {display} ({imap_name!r}) 尝试 {attempt}/3")
                    uidvalidity = client.select_folder(imap_name)
                    log(f"  [SELECT] OK  UIDVALIDITY={uidvalidity}")
                    uids = client.search_uids(since=since)
                    log(f"  [SEARCH] 命中 {len(uids)} 封" + (f" (since={since})" if since else ""))
                    break
                except Exception as e:
                    log(f"  [打开失败 {attempt}/3] {display}: {e}", err=True)
                    log(explain_imap_error(e), err=True)
                    if VERBOSE:
                        traceback.print_exc()
                    try:
                        client.reconnect()
                    except Exception as e2:
                        log(f"  [重连失败] {e2}", err=True)
                        log(explain_imap_error(e2), err=True)
                        time.sleep(3)
            else:
                log(f"  [放弃文件夹] {display}", err=True)
                continue

            done = state.get_done(imap_name, uidvalidity)
            # 若 UIDVALIDITY 变化，写入新 validity 并清空 done（get_done 已返回空）
            if imap_name not in state.folders or int(
                state.folders.get(imap_name, {}).get("uidvalidity") or 0
            ) != int(uidvalidity):
                state.folders[imap_name] = {
                    "uidvalidity": int(uidvalidity),
                    "done": [],
                }
                done = set()

            # 最近优先：UID 从大到小
            uids = sort_uids(uids, newest_first=newest_first)
            pending = [u for u in uids if u not in done]
            total_pending = len(pending)
            if limit is not None and limit > 0:
                pending = pending[:limit]

            log(
                f"\n=== {display} ===  匹配 {len(uids)} 封, "
                f"已完成 {len(uids) - total_pending}, 待同步 {total_pending}"
                + (f", 本轮处理 {len(pending)}" if limit else "")
                + f"  (UIDVALIDITY={uidvalidity}, {order_desc})"
            )
            log(f"  目录: {dest_dir}")

            exported_here = 0
            # 仅统计断点已完成的跳过数，不含因 --limit 本轮未处理的
            skipped_here = len(uids) - total_pending
            failed_here = 0
            state.stats["skipped"] = state.stats.get("skipped", 0) + skipped_here

            for idx, uid in enumerate(pending, 1):
                try:
                    raw, internaldate = client.fetch_eml(uid)
                    # 规范化换行：部分服务器给的是裸 LF，eml 用原始 bytes 即可
                    if raw.startswith(b"\xef\xbb\xbf"):
                        raw = raw[3:]

                    try:
                        msg = email.message_from_bytes(raw)
                    except Exception:
                        msg = email.message_from_bytes(b"Subject: (parse-failed)\r\n\r\n" + raw)

                    stable_key = f"{user}|{imap_name}|{uidvalidity}|{uid}"
                    # 优先 Message-ID 做 short_id 源更稳，但仍纳入 stable_key 保证唯一
                    mid = decode_mime_header(msg.get("Message-ID") or msg.get("Message-Id") or "")
                    if mid:
                        stable_key = f"{stable_key}|{mid}"

                    base = mail_filename_from_msg(msg, stable_key, internaldate)
                    out_path = unique_eml_path(dest_dir, base)
                    out_path.write_bytes(raw)

                    state.mark_done(imap_name, uidvalidity, uid)
                    state.stats["exported"] = state.stats.get("exported", 0) + 1
                    exported_here += 1

                    if state.stats["exported"] % save_every == 0:
                        state.save()
                        log(
                            f"  …进度 {idx}/{len(pending)}  "
                            f"总导出 {state.stats['exported']}  "
                            f"失败 {state.stats.get('failed', 0)}"
                        )
                except (imaplib.IMAP4.abort, imaplib.IMAP4.error, socket.error, OSError) as e:
                    failed_here += 1
                    state.stats["failed"] = state.stats.get("failed", 0) + 1
                    log(f"  [网络/IMAP 失败] uid={uid}: {e}", err=True)
                    log(explain_imap_error(e), err=True)
                    state.save()
                    try:
                        client.reconnect()
                        client.select_folder(imap_name)
                    except Exception as e2:
                        log(f"  [重连失败，稍后继续] {e2}", err=True)
                        log(explain_imap_error(e2), err=True)
                        time.sleep(5)
                except Exception as e:
                    failed_here += 1
                    state.stats["failed"] = state.stats.get("failed", 0) + 1
                    log(f"  [失败] uid={uid}: {type(e).__name__}: {e}", err=True)
                    if verbose:
                        traceback.print_exc()

            state.save()
            log(
                f"  {display} 完成 → 新同步 {exported_here}, "
                f"跳过 {skipped_here}, 失败 {failed_here}"
            )

        state.save()
        log("\n—— 完成 ——")
        log(
            f"新同步: {state.stats.get('exported', 0)}  |  "
            f"跳过: {state.stats.get('skipped', 0)}  |  "
            f"失败: {state.stats.get('failed', 0)}  |  "
            f"文件夹: {state.stats.get('folders', 0)}"
        )
        log(f"断点文件: {state_path}")
        log(f"邮件目录: {account_dir}")
        return 0
    finally:
        client.close()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="IMAP 同步 Yahoo 邮箱全部邮件为 EML（目录/文件名与 Outlook 导出脚本一致）"
    )
    parser.add_argument("-u", "--user", required=True, help="邮箱账号，如 name@yahoo.fr")
    parser.add_argument(
        "-p",
        "--password",
        default=None,
        help="密码或应用专用密码；也可用环境变量 YAHOO_IMAP_PASSWORD",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="输出根目录，默认=脚本所在目录",
    )
    parser.add_argument(
        "--host",
        default=DEFAULT_HOST,
        help=f"IMAP 主机，默认 {DEFAULT_HOST}",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=DEFAULT_PORT,
        help=f"IMAP 端口，默认 {DEFAULT_PORT}",
    )
    parser.add_argument(
        "--list-folders",
        action="store_true",
        help="只登录并列出文件夹后退出",
    )
    parser.add_argument(
        "--folder",
        action="append",
        default=None,
        help="只同步指定文件夹显示名，可多次",
    )
    parser.add_argument(
        "--skip-folder",
        action="append",
        default=[],
        help="跳过的文件夹名，可多次",
    )
    parser.add_argument(
        "--save-every",
        type=int,
        default=20,
        help="每同步 N 封保存断点，默认 20",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=120,
        help="网络超时秒数，默认 120",
    )
    parser.add_argument(
        "--reset-state",
        action="store_true",
        help="清除该账号断点后重新同步（已存在 eml 不会自动删）",
    )
    parser.add_argument(
        "--oldest-first",
        action="store_true",
        help="从最旧邮件开始同步（默认从最近开始）",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="每个文件夹本轮最多同步 N 封（按当前顺序，默认最近 N 封）",
    )
    parser.add_argument(
        "--since",
        default=None,
        metavar="YYYY-MM-DD",
        help="只同步该日期及之后的邮件（IMAP SINCE）",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="详细调试日志（原始 LIST 行、堆栈等）",
    )
    args = parser.parse_args()

    global VERBOSE
    # 默认就输出关键诊断；-v 更啰嗦
    VERBOSE = bool(args.verbose)

    password = args.password
    if not password:
        import os

        password = os.environ.get("YAHOO_IMAP_PASSWORD")
    if not password:
        log("请通过 -p/--password 或环境变量 YAHOO_IMAP_PASSWORD 提供密码。", err=True)
        return 1

    script_dir = Path(__file__).resolve().parent
    output_root = Path(args.output).resolve() if args.output else script_dir
    ensure_dir(output_root)

    if args.list_folders:
        client = YahooImapClient(args.host, args.port, args.user, password, timeout=args.timeout)
        try:
            client.connect()
            folders = client.list_folders()
            log(f"账号 {args.user} 文件夹列表（{len(folders)}）：")
            for imap_name, display, flags in folders:
                log(f"  {display}  (raw={imap_name!r}, flags={flags})")
            return 0
        except Exception as e:
            log(f"失败: {e}", err=True)
            log(explain_imap_error(e), err=True)
            if VERBOSE:
                traceback.print_exc()
            return 1
        finally:
            client.close()

    if args.since:
        try:
            datetime.strptime(args.since, "%Y-%m-%d")
        except ValueError:
            log("--since 格式须为 YYYY-MM-DD，例如 2024-01-01", err=True)
            return 1

    return sync_account(
        user=args.user.strip(),
        password=password,
        output_root=output_root,
        host=args.host,
        port=args.port,
        save_every=max(1, args.save_every),
        verbose=args.verbose,
        reset_state=args.reset_state,
        only_folders=args.folder,
        skip_folders=set(args.skip_folder or []),
        timeout=args.timeout,
        newest_first=not args.oldest_first,
        limit=args.limit,
        since=args.since,
    )


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        log("\n已中断。下次运行会按断点跳过已同步 UID。")
        raise SystemExit(130)
    except Exception as e:
        log(f"\n[未捕获异常] {type(e).__name__}: {e}", err=True)
        log(explain_imap_error(e), err=True)
        traceback.print_exc()
        raise SystemExit(1)
