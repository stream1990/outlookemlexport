#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从本机 Outlook 导出已缓存/已下载的邮件为 .eml
- 不触发发送/接收，不主动拉新邮件
- 一个账号(Store)一个目录，文件夹层级映射为子目录
- 导出到脚本所在目录（或当前工作目录，见 OUTPUT_ROOT）
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import tempfile
import traceback
from datetime import datetime
from email import policy
from email.message import EmailMessage
from email.utils import format_datetime, formataddr, make_msgid
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
except ImportError:
    print("缺少 pywin32，请先执行: pip install -r requirements.txt")
    sys.exit(1)


# Outlook 对象 Class 常量
OL_MAIL = 43
OL_FOLDER = 2

# MAPI 属性：传输头（Internet headers）
PR_TRANSPORT_MESSAGE_HEADERS_W = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
PR_TRANSPORT_MESSAGE_HEADERS_A = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"

# 默认跳过的文件夹名（可按需改）
DEFAULT_SKIP_FOLDERS = {
    "日历",
    "Calendar",
    "联系人",
    "Contacts",
    "任务",
    "Tasks",
    "日记",
    "Journal",
    "便笺",
    "Notes",
    "RSS 订阅",
    "RSS Feeds",
    "同步问题",
    "Sync Issues",
    "快速步骤设置",
    "Quick Step Settings",
    "会话操作设置",
    "Conversation Action Settings",
}

INVALID_FS_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
WHITESPACE_RUN = re.compile(r"\s+")


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


def com_str(value: Any) -> str:
    if value is None:
        return ""
    try:
        return str(value).strip()
    except Exception:
        return ""


def get_prop(mail: Any, schema: str) -> Optional[str]:
    try:
        pa = mail.PropertyAccessor
        val = pa.GetProperty(schema)
        if val is None:
            return None
        return str(val)
    except Exception:
        return None


def recipients_to_header(recipients: Any, want_type: int) -> str:
    """
    Recipient.Type: 1=To, 2=CC, 3=BCC
    """
    parts: List[str] = []
    try:
        count = int(recipients.Count)
    except Exception:
        return ""
    for i in range(1, count + 1):
        try:
            r = recipients.Item(i)
            if int(r.Type) != want_type:
                continue
            name = com_str(getattr(r, "Name", ""))
            addr = com_str(getattr(r, "Address", ""))
            # Exchange 内部地址时 Address 可能是 DN，尽量用 Name
            if addr and "/" in addr and "=" in addr:
                addr = ""
            if name and addr and name.lower() != addr.lower():
                parts.append(formataddr((name, addr)))
            elif addr:
                parts.append(addr)
            elif name:
                parts.append(name)
        except Exception:
            continue
    return ", ".join(parts)


def format_outlook_time(dt: Any) -> Optional[str]:
    if not dt:
        return None
    try:
        # win32com 常返回 pywintypes.datetime
        if hasattr(dt, "tzinfo") and dt.tzinfo is not None:
            return format_datetime(dt)
        # 无时区：按本地naive处理
        if isinstance(dt, datetime):
            return dt.strftime("%a, %d %b %Y %H:%M:%S %z").strip() or dt.strftime(
                "%a, %d %b %Y %H:%M:%S"
            )
        return str(dt)
    except Exception:
        try:
            return str(dt)
        except Exception:
            return None


def build_eml_bytes(mail: Any) -> bytes:
    """
    从 MailItem 组装标准 MIME/EML。
    优先复用 PR_TRANSPORT_MESSAGE_HEADERS，再补正文与附件。
    """
    msg = EmailMessage()

    headers_raw = get_prop(mail, PR_TRANSPORT_MESSAGE_HEADERS_W) or get_prop(
        mail, PR_TRANSPORT_MESSAGE_HEADERS_A
    )

    # 基础头
    subject = com_str(getattr(mail, "Subject", ""))
    sender_name = com_str(getattr(mail, "SenderName", ""))
    sender_email = com_str(getattr(mail, "SenderEmailAddress", ""))
    if sender_email and "/" in sender_email and "=" in sender_email:
        # Exchange DN，尽量改用 SMTP
        sender_email = ""
    if not sender_email:
        try:
            eu = mail.Sender.GetExchangeUser()
            if eu is not None:
                sender_email = com_str(getattr(eu, "PrimarySmtpAddress", ""))
        except Exception:
            pass
        if not sender_email:
            try:
                ea = mail.Sender.GetExchangeDistributionList()
                if ea is not None:
                    sender_email = com_str(getattr(ea, "PrimarySmtpAddress", ""))
            except Exception:
                pass

    to_h = recipients_to_header(mail.Recipients, 1) or com_str(getattr(mail, "To", ""))
    cc_h = recipients_to_header(mail.Recipients, 2) or com_str(getattr(mail, "CC", ""))
    bcc_h = recipients_to_header(mail.Recipients, 3) or com_str(getattr(mail, "BCC", ""))

    if subject:
        msg["Subject"] = subject
    if sender_name or sender_email:
        if sender_name and sender_email:
            msg["From"] = formataddr((sender_name, sender_email))
        elif sender_email:
            msg["From"] = sender_email
        else:
            msg["From"] = sender_name
    if to_h:
        msg["To"] = to_h
    if cc_h:
        msg["Cc"] = cc_h
    if bcc_h:
        msg["Bcc"] = bcc_h

    sent = format_outlook_time(getattr(mail, "SentOn", None))
    received = format_outlook_time(getattr(mail, "ReceivedTime", None))
    if sent:
        msg["Date"] = sent
    elif received:
        msg["Date"] = received

    # 尝试保留原始 Internet 头里的 Message-ID 等（不覆盖已写关键字段时追加缺失项）
    if headers_raw:
        for line in headers_raw.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
            if not line or line[0] in " \t":
                continue
            if ":" not in line:
                continue
            key, val = line.split(":", 1)
            key = key.strip()
            val = val.strip()
            if not key or not val:
                continue
            lk = key.lower()
            # 已由结构化字段写入的跳过；其余有用的头尽量保留
            if lk in {
                "subject",
                "from",
                "to",
                "cc",
                "bcc",
                "date",
                "content-type",
                "content-transfer-encoding",
                "mime-version",
            }:
                continue
            if lk in msg:
                continue
            try:
                msg[key] = val
            except Exception:
                continue

    if "Message-ID" not in msg and "Message-Id" not in msg:
        entry = com_str(getattr(mail, "EntryID", "")) or short_id(subject + (sent or ""))
        msg["Message-ID"] = make_msgid(domain="outlook.local", idstring=short_id(entry, 12))

    # 正文
    html = com_str(getattr(mail, "HTMLBody", ""))
    body = com_str(getattr(mail, "Body", ""))
    if html and html.lower().strip() not in ("", "<html><body></body></html>"):
        if body:
            msg.set_content(body, subtype="plain", charset="utf-8")
            msg.add_alternative(html, subtype="html", charset="utf-8")
        else:
            msg.set_content(html, subtype="html", charset="utf-8")
    else:
        msg.set_content(body or "", subtype="plain", charset="utf-8")

    # 附件
    try:
        att_count = int(mail.Attachments.Count)
    except Exception:
        att_count = 0

    if att_count > 0:
        with tempfile.TemporaryDirectory(prefix="ol_eml_") as tmp:
            tmp_path = Path(tmp)
            for i in range(1, att_count + 1):
                att = None
                try:
                    att = mail.Attachments.Item(i)
                    fname = safe_name(com_str(getattr(att, "FileName", "")) or f"attachment_{i}")
                    # Type 1 = olByValue 普通附件；6 = 嵌入式 OLE 等，尽量都导出
                    save_as = tmp_path / f"{i}_{fname}"
                    att.SaveAsFile(str(save_as))
                    data = save_as.read_bytes()
                    ctype = "application/octet-stream"
                    maintype, subtype = ctype.split("/", 1)
                    msg.add_attachment(
                        data,
                        maintype=maintype,
                        subtype=subtype,
                        filename=fname,
                    )
                except Exception:
                    # 个别嵌入图片/受保护附件可能失败，跳过该附件
                    continue

    return msg.as_bytes(policy=policy.SMTP)


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


def mail_filename(mail: Any) -> str:
    entry = com_str(getattr(mail, "EntryID", "")) or os.urandom(8).hex()
    sid = short_id(entry, 10)
    dt = getattr(mail, "ReceivedTime", None) or getattr(mail, "SentOn", None)
    if dt:
        try:
            date_part = dt.strftime("%Y-%m-%d_%H%M%S")
        except Exception:
            date_part = "unknown-date"
    else:
        date_part = "unknown-date"
    subject = safe_name(com_str(getattr(mail, "Subject", "")) or "(no subject)", max_len=60)
    return f"{date_part}_{subject}_{sid}"


class ExportState:
    def __init__(self, path: Path):
        self.path = path
        self.done: Set[str] = set()
        self.stats: Dict[str, int] = {
            "exported": 0,
            "skipped": 0,
            "failed": 0,
            "folders": 0,
        }
        if path.exists():
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
                self.done = set(data.get("done", []))
                self.stats.update(data.get("stats", {}))
            except Exception:
                pass

    def is_done(self, entry_id: str) -> bool:
        return entry_id in self.done

    def mark_done(self, entry_id: str) -> None:
        self.done.add(entry_id)

    def save(self) -> None:
        payload = {
            "done": sorted(self.done),
            "stats": self.stats,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
        self.path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )


def iter_mail_items(folder: Any) -> Iterable[Any]:
    """
    只枚举当前文件夹内的邮件项（不含子文件夹）。
    使用限制条件尽量只取邮件；失败则全量扫并按 Class 过滤。
    """
    items = folder.Items
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass

    restricted = None
    try:
        # 43 = olMail
        restricted = items.Restrict("@SQL=\"http://schemas.microsoft.com/mapi/proptag/0x001A001F\" LIKE 'IPM.Note%'")
    except Exception:
        restricted = None

    source = restricted if restricted is not None else items
    try:
        count = int(source.Count)
    except Exception:
        count = 0

    # Outlook COM 集合 1-based；大文件夹用 GetFirst/GetNext 更稳
    try:
        item = source.GetFirst()
        while item is not None:
            yield item
            try:
                item = source.GetNext()
            except Exception:
                break
        return
    except Exception:
        pass

    for i in range(1, count + 1):
        try:
            yield source.Item(i)
        except Exception:
            continue


def should_skip_folder(name: str, skip_names: Set[str]) -> bool:
    return name.strip() in skip_names


def export_folder(
    folder: Any,
    dest_dir: Path,
    state: ExportState,
    skip_names: Set[str],
    save_every: int,
    verbose: bool,
) -> None:
    name = com_str(getattr(folder, "Name", "")) or "folder"
    if should_skip_folder(name, skip_names):
        if verbose:
            print(f"  [跳过文件夹] {name}")
        return

    ensure_dir(dest_dir)
    state.stats["folders"] = state.stats.get("folders", 0) + 1

    # 导出本文件夹邮件
    exported_here = 0
    failed_here = 0
    skipped_here = 0

    try:
        for item in iter_mail_items(folder):
            try:
                # 有的 Restrict 仍可能混入非邮件
                cls = int(getattr(item, "Class", 0) or 0)
                if cls != OL_MAIL:
                    # 有时会议请求等也想保留：仅当 MessageClass 以 IPM.Note 开头
                    mc = com_str(getattr(item, "MessageClass", ""))
                    if not mc.startswith("IPM.Note"):
                        continue

                entry_id = com_str(getattr(item, "EntryID", ""))
                if not entry_id:
                    # 无 EntryID 极少见，用内容哈希顶一下
                    entry_id = short_id(
                        com_str(getattr(item, "Subject", ""))
                        + com_str(getattr(item, "Body", ""))[:200],
                        40,
                    )

                if state.is_done(entry_id):
                    skipped_here += 1
                    state.stats["skipped"] = state.stats.get("skipped", 0) + 1
                    continue

                # 未完整下载的邮件 Body 可能空且 Size 异常，仍尝试导出（用户要求只处理本地已有）
                raw = build_eml_bytes(item)
                out_path = unique_eml_path(dest_dir, mail_filename(item))
                out_path.write_bytes(raw)

                state.mark_done(entry_id)
                state.stats["exported"] = state.stats.get("exported", 0) + 1
                exported_here += 1

                if state.stats["exported"] % save_every == 0:
                    state.save()
                    print(
                        f"  …已导出 {state.stats['exported']} 封 "
                        f"(跳过 {state.stats.get('skipped', 0)}, 失败 {state.stats.get('failed', 0)})"
                    )
            except Exception as e:
                failed_here += 1
                state.stats["failed"] = state.stats.get("failed", 0) + 1
                if verbose:
                    subj = ""
                    try:
                        subj = com_str(getattr(item, "Subject", ""))
                    except Exception:
                        pass
                    print(f"  [失败] {subj or '(unknown)'}: {e}")
    except Exception as e:
        print(f"  [读文件夹失败] {name}: {e}")

    print(
        f"  {dest_dir.as_posix()}  →  新导出 {exported_here}, "
        f"跳过 {skipped_here}, 失败 {failed_here}"
    )

    # 子文件夹
    try:
        sub_count = int(folder.Folders.Count)
    except Exception:
        sub_count = 0

    for i in range(1, sub_count + 1):
        try:
            sub = folder.Folders.Item(i)
            sub_name = safe_name(com_str(getattr(sub, "Name", "")) or f"sub_{i}")
            export_folder(
                sub,
                dest_dir / sub_name,
                state,
                skip_names,
                save_every,
                verbose,
            )
        except Exception as e:
            print(f"  [子文件夹失败] {e}")


def list_stores(namespace: Any) -> List[Any]:
    stores = []
    try:
        count = int(namespace.Stores.Count)
    except Exception:
        count = 0
    for i in range(1, count + 1):
        try:
            stores.append(namespace.Stores.Item(i))
        except Exception:
            continue
    return stores


def get_store_root_folder(store: Any) -> Optional[Any]:
    try:
        return store.GetRootFolder()
    except Exception:
        return None


def main() -> int:
    parser = argparse.ArgumentParser(
        description="导出本机 Outlook 已下载邮件为 EML（不拉新邮件）"
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="输出根目录，默认=脚本所在目录",
    )
    parser.add_argument(
        "--store",
        action="append",
        default=None,
        help="只导出指定账号/数据文件显示名，可多次指定；默认全部",
    )
    parser.add_argument(
        "--list-stores",
        action="store_true",
        help="仅列出本机 Outlook 账号/数据文件后退出",
    )
    parser.add_argument(
        "--skip-folder",
        action="append",
        default=[],
        help="额外跳过的文件夹名，可多次指定（默认已跳过日历/联系人/任务等）",
    )
    parser.add_argument(
        "--save-every",
        type=int,
        default=50,
        help="每导出 N 封保存一次断点状态，默认 50",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="打印更细的失败信息",
    )
    parser.add_argument(
        "--reset-state",
        action="store_true",
        help="忽略已有断点，重新导出（会覆盖同名策略下的新文件名去重）",
    )
    args = parser.parse_args()

    script_dir = Path(__file__).resolve().parent
    output_root = Path(args.output).resolve() if args.output else script_dir
    ensure_dir(output_root)

    print("连接本机 Outlook（只读本地存储，不触发发送/接收）…")
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        # 不调用 ns.SendAndReceive()，避免拉新

        stores = list_stores(ns)
        if not stores:
            print("未找到任何 Outlook 数据文件/账号。请先打开 Outlook 并确认已登录。")
            return 1

        print(f"发现 {len(stores)} 个账号/数据文件：")
        for i, st in enumerate(stores, 1):
            print(f"  {i}. {com_str(getattr(st, 'DisplayName', ''))}")

        if args.list_stores:
            return 0

        selected = stores
        if args.store:
            want = {s.strip().lower() for s in args.store}
            selected = [
                st
                for st in stores
                if com_str(getattr(st, "DisplayName", "")).lower() in want
            ]
            if not selected:
                print("未匹配到 --store 指定的账号，请用 --list-stores 查看名称。")
                return 1

        skip_names = set(DEFAULT_SKIP_FOLDERS)
        for extra in args.skip_folder:
            skip_names.add(extra.strip())

        state_path = output_root / ".export_state.json"
        if args.reset_state and state_path.exists():
            state_path.unlink()
        state = ExportState(state_path)

        print(f"\n输出目录: {output_root}")
        print("开始导出…\n")

        for st in selected:
            store_name = com_str(getattr(st, "DisplayName", "")) or "Account"
            store_dir_name = safe_name(store_name, max_len=100)
            store_dir = output_root / store_dir_name
            root = get_store_root_folder(st)
            if root is None:
                print(f"[跳过] 无法打开: {store_name}")
                continue

            print(f"=== 账号: {store_name} ===")
            ensure_dir(store_dir)

            # 根下各顶级文件夹分别导出（收件箱、已发送等）
            try:
                top_count = int(root.Folders.Count)
            except Exception:
                top_count = 0

            if top_count == 0:
                # 少数 pst 结构特殊，直接导出 root
                export_folder(
                    root,
                    store_dir / safe_name(com_str(root.Name) or "Root"),
                    state,
                    skip_names,
                    args.save_every,
                    args.verbose,
                )
            else:
                for i in range(1, top_count + 1):
                    try:
                        top = root.Folders.Item(i)
                        top_name = safe_name(
                            com_str(getattr(top, "Name", "")) or f"Folder_{i}"
                        )
                        export_folder(
                            top,
                            store_dir / top_name,
                            state,
                            skip_names,
                            args.save_every,
                            args.verbose,
                        )
                    except Exception as e:
                        print(f"  [顶级文件夹失败] {e}")
                        if args.verbose:
                            traceback.print_exc()

            state.save()
            print()

        state.save()
        print("—— 完成 ——")
        print(
            f"新导出: {state.stats.get('exported', 0)}  |  "
            f"跳过(已导出过): {state.stats.get('skipped', 0)}  |  "
            f"失败: {state.stats.get('failed', 0)}  |  "
            f"遍历文件夹: {state.stats.get('folders', 0)}"
        )
        print(f"断点文件: {state_path}")
        print(f"邮件目录: {output_root}")
        return 0
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\n已中断。下次运行会按断点跳过已导出的邮件。")
        raise SystemExit(130)
