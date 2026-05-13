# -*- coding: utf-8 -*-
"""
MikroTik Ops - Unlimited Edition
SSH Automation, WinBox Launcher, Router Backups
No license, no trial limits.
"""

import os
import re
import csv
import json
import math
import time
import threading
import queue
import subprocess
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import customtkinter as ctk
import pandas as pd
import paramiko

from scp import SCPClient
from apscheduler.schedulers.background import BackgroundScheduler


ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

APP_TITLE = "MikroTik Ops • SSH / WinBox / BackUp"
OUT_BASE = os.path.join(os.getcwd(), "outputs")
os.makedirs(OUT_BASE, exist_ok=True)


def safe_filename(s: str, max_len: int = 80) -> str:
    s = (s or "").strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^0-9A-Za-z_+=\-\.]", "", s)
    s = s.strip("._-")
    if not s:
        s = "export"
    return s[:max_len]


def now_ts() -> str:
    return time.strftime("%Y%m%d-%H%M%S")


def _is_context_line(s: str) -> bool:
    if not s:
        return False
    if s.startswith(":"):
        return False
    tokens = s.lstrip("/").split()
    if not tokens:
        return False
    ops = {
        "add", "set", "get", "print", "remove", "enable", "disable",
        "reset", "export", "import", "edit", "find", "monitor", "reboot"
    }
    return all(t.lower() not in ops for t in tokens)


def _is_operation_line(s: str) -> bool:
    if not s or s.startswith(":"):
        return False
    head = s.split()[0].lower()
    return head in {
        "add", "set", "get", "print", "remove", "enable", "disable",
        "reset", "export", "import", "edit", "find", "monitor", "reboot"
    }


def normalize_ros_commands(lines):
    normalized = []
    current_ctx = ""
    for raw in lines:
        s = str(raw).strip()
        if not s or s.startswith("#"):
            continue
        if s.strip().lower() in {"yes", "y"}:
            continue
        if s.lstrip("/").lower().startswith("system reboot"):
            normalized.append("/system reboot")
            continue
        if s.startswith(":") or s.startswith("{") or s.endswith("\\"):
            normalized.append(s)
            continue
        if s.startswith("/"):
            if _is_context_line(s):
                current_ctx = s
                continue
            else:
                parts = s.split()
                if len(parts) >= 2:
                    ctx_guess = " ".join(parts[:-1])
                    if _is_context_line(ctx_guess):
                        current_ctx = ctx_guess
                normalized.append(s)
                continue
        if _is_context_line(s):
            current_ctx = "/" + s
            continue
        if _is_operation_line(s) and current_ctx:
            normalized.append(f"{current_ctx} {s}")
            continue
        normalized.append(s)
    return normalized


def load_commands_list(cmd_path: str):
    with open(cmd_path, "r", encoding="utf-8") as f:
        raw = f.readlines()
    cmds = [ln.strip() for ln in raw if ln.strip() and not ln.strip().startswith("#")]
    return normalize_ros_commands(cmds)


def _strip_bidi(s: str) -> str:
    if s is None:
        return ""
    for bad in ["\u200f", "\u200e", "\u202a", "\u202b", "\u202c", "\u202d", "\u202e", "\ufeff"]:
        s = s.replace(bad, "")
    return s


def read_ips_from_excel(xlsx_path, sheet_name, ip_col, start_row, end_row):
    df0 = pd.read_excel(xlsx_path, sheet_name=sheet_name or 0, nrows=1, engine="openpyxl")
    cmap = {_strip_bidi(c).strip().lower(): c for c in df0.columns}
    key = _strip_bidi(ip_col).strip().lower()
    if key not in cmap:
        raise ValueError(f"IP column '{ip_col}' not found. Available: {list(df0.columns)}")
    real_ip_col = cmap[key]

    df = pd.read_excel(xlsx_path, sheet_name=sheet_name or 0, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.iloc[start_row:end_row] if end_row is not None else df.iloc[start_row:]
    df = df.reset_index(drop=True)

    ips = []
    for _, r in df.iterrows():
        v = r.get(real_ip_col, None)
        if pd.notna(v):
            s = str(v).strip()
            s = re.sub(r'\.0$', '', s)
            if s:
                ips.append(s)
    return df, real_ip_col, ips


def get_cell(row, colname, default=None):
    if not colname:
        return default
    if colname in row and pd.notna(row[colname]):
        val = str(row[colname]).strip()
        val = re.sub(r'\.0$', '', val)
        return val
    return default


def ssh_connect(ip, port, username, password, timeout=30):
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        ssh.connect(
            ip, port=int(port), username=username, password=password,
            timeout=timeout, auth_timeout=timeout, banner_timeout=timeout
        )
        return ssh
    except paramiko.AuthenticationException:
        return f"Auth failed for {ip}"
    except paramiko.SSHException as e:
        return f"SSH error on {ip}: {e}"
    except Exception as e:
        return f"Conn error on {ip}: {e}"


def exec_commands_fast(ssh, commands, per_cmd_timeout=25, cmd_delay=0.0, *,
                       on_cmd=None, on_line=None, cancel_event=None):
    outputs = []

    def _run(cmd):
        try:
            stdin, stdout, stderr = ssh.exec_command(cmd, timeout=per_cmd_timeout)
            out = stdout.read().decode(errors="ignore")
            err = stderr.read().decode(errors="ignore")
            return out, err, None
        except Exception as e:
            return "", "", e

    def _reboot_via_script():
        name = "__rb_auto__"
        add_cmd = f'/system script add name={name} source="/system reboot"'
        if on_cmd: on_cmd(add_cmd)
        out1, err1, ex1 = _run(add_cmd)

        if any(x in (out1 + err1).lower() for x in ["already have such item", "failure: already exists"]):
            if on_cmd: on_cmd(f"/system script remove {name}")
            _run(f"/system script remove {name}")
            if on_cmd: on_cmd(add_cmd)
            out1, err1, ex1 = _run(add_cmd)

        if ex1 or any(x in (out1 + err1).lower() for x in ["failure", "syntax error", "expected end of command"]):
            return False, [(add_cmd, out1, err1 if err1 else ("" if not ex1 else f"EXC: {ex1}"))]

        logs = [(add_cmd, out1, err1)]

        run_cmd = f"/system script run {name}"
        if on_cmd: on_cmd(run_cmd)
        out2, err2, ex2 = _run(run_cmd)
        logs.append((run_cmd, out2, err2 if err2 else ("" if not ex2 else f"EXC: {ex2}")))

        rem_cmd = f"/system script remove {name}"
        if on_cmd: on_cmd(rem_cmd)
        out3, err3, ex3 = _run(rem_cmd)
        logs.append((rem_cmd, out3, err3 if err3 else ("" if not ex3 else f"EXC: {ex3}")))

        bad = (out2 + err2).lower()
        ok = not any(x in bad for x in ["syntax error", "expected end of command", "failure", "invalid"])
        return ok, logs

    def _reboot_interactive():
        out_buf = ""
        try:
            chan = ssh.invoke_shell()
            chan.settimeout(per_cmd_timeout)
            if on_cmd: on_cmd("/system reboot (interactive)")
            chan.send("/system reboot\r\n")
            time.sleep(0.5)
            chan.send("y\r\n")
            time.sleep(1.0)
            try:
                while chan.recv_ready():
                    out_buf += chan.recv(4096).decode(errors="ignore")
                    time.sleep(0.1)
            except Exception:
                pass
            try:
                chan.close()
            except Exception:
                pass
            return [("/system reboot (interactive)", out_buf, "")]
        except Exception as e:
            return [("/system reboot (interactive)", out_buf, f"note: {e}")]

    reboot_regex = re.compile(r"(^|[\s;])/?\s*system\s+reboot\b", re.IGNORECASE)

    for cmd in commands:
        if cancel_event and cancel_event.is_set():
            if on_line:
                on_line("[CANCEL] stopped before next command.")
            break

        c = (cmd or "").strip()
        if not c:
            continue

        if reboot_regex.search(c):
            before, _, _ = c.partition(";")
            if before and not reboot_regex.search(before):
                if on_cmd: on_cmd(before.strip())
                out, err, ex = _run(before.strip())
                outputs.append((before.strip(), out, err if err else ("" if not ex else f"EXC: {ex}")))

            ok, logs = _reboot_via_script()
            outputs.extend(logs)
            if not ok:
                outputs.extend(_reboot_interactive())
            break
        else:
            if on_cmd: on_cmd(c)
            out, err, ex = _run(c)
            outputs.append((c, out, err if err else ("" if not ex else f"EXC: {ex}")))
            if cmd_delay:
                time.sleep(cmd_delay)

    return outputs


def run_stage1_exec_concurrent(*, excel_path, sheet_name, ip_col, cmd_file,
                               fallback_user, fallback_pass, fallback_port,
                               user_col, pass_col, port_col, start_row, end_row,
                               timeout_conn, timeout_cmd, cmd_delay,
                               concurrency: int, log_q: queue.Queue,
                               prog_q: queue.Queue, cancel_event: threading.Event):
    log_q.put(("info", f"[INFO] Stage 1 starting • concurrency={concurrency}"))
    try:
        df, ip_col_resolved, ips = read_ips_from_excel(excel_path, sheet_name, ip_col, start_row, end_row)
    except Exception as e:
        log_q.put(("err", f"[FATAL] Excel read failed: {e}"))
        prog_q.put(("done", 0, 0, None, True))
        return

    try:
        commands = load_commands_list(cmd_file)
        if not commands:
            log_q.put(("err", "[FATAL] Command file has no effective lines."))
            prog_q.put(("done", 0, 0, None, True))
            return
    except Exception as e:
        log_q.put(("err", f"[FATAL] Cannot read command file: {e}"))
        prog_q.put(("done", 0, 0, None, True))
        return

    out_dir = os.path.join(OUT_BASE, now_ts() + "_stage1")
    os.makedirs(out_dir, exist_ok=True)
    summary_csv = os.path.join(out_dir, "summary.csv")
    with open(summary_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["timestamp", "ip", "port", "username", "status", "error", "log_file"])

    devices = []
    for _, row in df.iterrows():
        ip = str(row.get(ip_col_resolved, "")).strip()
        if not ip:
            continue
        username = get_cell(row, user_col, fallback_user) if user_col else fallback_user
        password = get_cell(row, pass_col, fallback_pass) if pass_col else fallback_pass
        port = get_cell(row, port_col, fallback_port) if port_col else fallback_port
        devices.append({"ip": ip, "username": username, "password": password, "port": port})

    total = len(devices)

    log_q.put(("info", f"[INFO] Devices to process: {total} • outputs → {out_dir}"))
    if total == 0:
        prog_q.put(("done", 0, 0, out_dir, False))
        return

    done = 0
    lock = threading.Lock()

    def write_summary(ip, port, username, status, error, log_path):
        with lock:
            with open(summary_csv, "a", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow([
                    datetime.now().isoformat(timespec="seconds"),
                    ip, port, username, status, error, log_path
                ])

    def process_device(dev):
        ip = dev["ip"]
        username = (dev.get("username") or "").strip()
        password = dev.get("password") or ""
        port = dev.get("port") or "22"
        port = re.sub(r'\.0$', '', port)
        log_path = os.path.join(out_dir, f"{ip.replace(':', '_')}.log")

        if cancel_event.is_set():
            return (ip, port, username, "cancelled", "cancelled-before-start", log_path)

        if not ip or not username:
            msg = "missing-username"
            log_q.put(("warn", f"[{ip}] [SKIP] {msg}"))
            try:
                with open(log_path, "w", encoding="utf-8") as fo:
                    fo.write(msg + "\n")
            except Exception:
                pass
            return (ip, port, username, "skipped", msg, log_path)

        log_q.put(("run", f"[{ip}] connect {ip}:{port} as {username} …"))
        ssh = ssh_connect(ip, port, username, password, timeout=timeout_conn)
        if isinstance(ssh, str):
            err_msg = ssh
            log_q.put(("err", f"[{ip}] [FAIL] {err_msg}"))
            try:
                with open(log_path, "w", encoding="utf-8") as fo:
                    fo.write(err_msg + "\n")
            except Exception:
                pass
            return (ip, port, username, "failed", err_msg, log_path)

        any_err = False

        def on_cmd(cmd_str: str):
            log_q.put(("cmd", f"[{ip}] $ {cmd_str}"))

        def on_line(s: str):
            log_q.put(("info", f"[{ip}] {s}"))

        results = exec_commands_fast(
            ssh, commands, per_cmd_timeout=timeout_cmd, cmd_delay=cmd_delay,
            on_cmd=on_cmd, on_line=on_line, cancel_event=cancel_event)

        try:
            ssh.close()
        except Exception:
            pass

        try:
            with open(log_path, "w", encoding="utf-8") as fo:
                for cmd, out, err in results:
                    fo.write(f"$ {cmd}\n")
                    if out:
                        fo.write(out)
                    if err:
                        any_err = True
                        fo.write(f"[ERR] {err}\n")
                    fo.write("\n")
        except Exception as e:
            any_err = True
            log_q.put(("err", f"[{ip}] failed to write log: {e}"))

        status = "ok" if not any_err else "failed"
        if cancel_event.is_set():
            status = "cancelled"

        log_q.put(("ok" if status == "ok" else "warn", f"[{ip}] [{status.upper()}]"))
        return (ip, port, username, status, "" if status == "ok" else "one/more cmds err", log_path)

    concurrency = max(1, int(concurrency or 1))
    log_q.put(("info", f"[INFO] ThreadPoolExecutor(max_workers={concurrency})"))

    cancelled = False
    with ThreadPoolExecutor(max_workers=concurrency) as ex:
        futures = []
        for dev in devices:
            if cancel_event.is_set():
                cancelled = True
                break
            futures.append(ex.submit(process_device, dev))

        for fut in as_completed(futures):
            if cancel_event.is_set():
                cancelled = True
            try:
                ip, port, username, status, err, log_path = fut.result()
            except Exception as e:
                ip, port, username, status, err, log_path = ("?", "?", "?", "failed", f"EXC: {e}", "")
            write_summary(ip, port, username, status, err, log_path)
            done += 1
            prog_q.put(("tick", done, total, out_dir))
            if cancelled and done >= total:
                break

    prog_q.put(("done", done, total, out_dir, cancelled))


try:
    import keyring
    KEYRING_OK = True
except Exception:
    keyring = None
    KEYRING_OK = False

SERVICE_NAME = "WinboxLauncher"
ACCOUNT_USER = "username_default"
ACCOUNT_PASS = "password_default"

DEFAULT_CONFIG = {
    "excel_path": "",
    "port": 2121,
    "winbox_path": "",
    "auto_reload_seconds": 10,
    "pass_credentials_on_cli": True
}

COMMON_WINBOX_PATHS = [
    r"C:\Program Files\MikroTik\Winbox\winbox64.exe",
    r"C:\Program Files (x86)\MikroTik\Winbox\winbox.exe",
    os.path.join(os.environ.get("USERPROFILE", ""), "Downloads", "winbox64.exe"),
    os.path.join(os.environ.get("USERPROFILE", ""), "Downloads", "winbox.exe"),
    "winbox64.exe",
    "winbox.exe",
]


def get_config_path() -> str:
    base = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "WinboxLauncher")
    os.makedirs(base, exist_ok=True)
    return os.path.join(base, "config.json")


def get_ui_config_path() -> str:
    base = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "MikroTikOps")
    os.makedirs(base, exist_ok=True)
    return os.path.join(base, "ui.json")


def load_ui_settings() -> dict:
    cfg_path = get_ui_config_path()
    try:
        with open(cfg_path, "r", encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}


def save_ui_settings(d: dict) -> None:
    cfg_path = get_ui_config_path()
    try:
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump(d or {}, f, indent=2)
    except Exception:
        pass


DEFAULT_STYLE = "Ocean"

STYLE_PRESETS = {
    "Ocean": {
        "accent": ("#2563EB", "#3B82F6"),
        "accent_hover": ("#1D4ED8", "#2563EB"),
        "ok": ("#16A34A", "#22C55E"),
        "warn": ("#D97706", "#F59E0B"),
        "bad": ("#DC2626", "#EF4444"),
        "card": ("#F7F8FA", "#111827"),
        "card2": ("#FFFFFF", "#0B1220"),
        "border": ("#E5E7EB", "#1F2937"),
        "text": ("#111827", "#E5E7EB"),
        "muted": ("#6B7280", "#9CA3AF"),
        "table_head": ("#EEF2FF", "#0F172A"),
        "table_head_text": ("#1F2937", "#E5E7EB"),
        "select": ("#DBEAFE", "#1E3A8A"),
    },
    "Emerald": {
        "accent": ("#059669", "#10B981"),
        "accent_hover": ("#047857", "#059669"),
        "ok": ("#16A34A", "#22C55E"),
        "warn": ("#D97706", "#F59E0B"),
        "bad": ("#DC2626", "#EF4444"),
        "card": ("#F7F8FA", "#0B1B17"),
        "card2": ("#FFFFFF", "#071612"),
        "border": ("#E5E7EB", "#0F2E24"),
        "text": ("#0F172A", "#E5E7EB"),
        "muted": ("#6B7280", "#9CA3AF"),
        "table_head": ("#ECFDF5", "#052e2a"),
        "table_head_text": ("#064E3B", "#E5E7EB"),
        "select": ("#D1FAE5", "#064E3B"),
    },
    "Teal": {
        "accent": ("#0F766E", "#14B8A6"),
        "accent_hover": ("#115E59", "#0F766E"),
        "ok": ("#16A34A", "#22C55E"),
        "warn": ("#D97706", "#F59E0B"),
        "bad": ("#DC2626", "#EF4444"),
        "card": ("#F7F8FA", "#0B1B1A"),
        "card2": ("#FFFFFF", "#061514"),
        "border": ("#E5E7EB", "#0F2F2D"),
        "text": ("#0F172A", "#E5E7EB"),
        "muted": ("#6B7280", "#9CA3AF"),
        "table_head": ("#E6FFFB", "#042f2e"),
        "table_head_text": ("#134E4A", "#E5E7EB"),
        "select": ("#CCFBF1", "#134E4A"),
    },
    "Violet": {
        "accent": ("#7C3AED", "#8B5CF6"),
        "accent_hover": ("#6D28D9", "#7C3AED"),
        "ok": ("#16A34A", "#22C55E"),
        "warn": ("#D97706", "#F59E0B"),
        "bad": ("#DC2626", "#EF4444"),
        "card": ("#F7F8FA", "#120B1A"),
        "card2": ("#FFFFFF", "#0B0712"),
        "border": ("#E5E7EB", "#2A1447"),
        "text": ("#111827", "#E5E7EB"),
        "muted": ("#6B7280", "#9CA3AF"),
        "table_head": ("#F5F3FF", "#1b1033"),
        "table_head_text": ("#4C1D95", "#E5E7EB"),
        "select": ("#EDE9FE", "#4C1D95"),
    },
    "Sunset": {
        "accent": ("#EA580C", "#F97316"),
        "accent_hover": ("#C2410C", "#EA580C"),
        "ok": ("#16A34A", "#22C55E"),
        "warn": ("#CA8A04", "#FBBF24"),
        "bad": ("#DC2626", "#EF4444"),
        "card": ("#FFF7ED", "#1A0F0B"),
        "card2": ("#FFFFFF", "#120A06"),
        "border": ("#FED7AA", "#3B1C12"),
        "text": ("#111827", "#F3F4F6"),
        "muted": ("#6B7280", "#9CA3AF"),
        "table_head": ("#FFEDD5", "#2a130b"),
        "table_head_text": ("#7C2D12", "#F3F4F6"),
        "select": ("#FED7AA", "#7C2D12"),
    },
    "Ruby": {
        "accent": ("#BE123C", "#FB7185"),
        "accent_hover": ("#9F1239", "#BE123C"),
        "ok": ("#16A34A", "#22C55E"),
        "warn": ("#D97706", "#F59E0B"),
        "bad": ("#DC2626", "#EF4444"),
        "card": ("#FFF1F2", "#1A0B10"),
        "card2": ("#FFFFFF", "#12070B"),
        "border": ("#FECDD3", "#3B1321"),
        "text": ("#111827", "#F3F4F6"),
        "muted": ("#6B7280", "#9CA3AF"),
        "table_head": ("#FFE4E6", "#2b0b14"),
        "table_head_text": ("#9F1239", "#F3F4F6"),
        "select": ("#FECDD3", "#9F1239"),
    },
    "Amber": {
        "accent": ("#B45309", "#F59E0B"),
        "accent_hover": ("#92400E", "#B45309"),
        "ok": ("#16A34A", "#22C55E"),
        "warn": ("#B45309", "#F59E0B"),
        "bad": ("#DC2626", "#EF4444"),
        "card": ("#FFFBEB", "#16120B"),
        "card2": ("#FFFFFF", "#100C06"),
        "border": ("#FDE68A", "#3A2B12"),
        "text": ("#111827", "#F3F4F6"),
        "muted": ("#6B7280", "#9CA3AF"),
        "table_head": ("#FEF3C7", "#2a1f0c"),
        "table_head_text": ("#92400E", "#F3F4F6"),
        "select": ("#FDE68A", "#92400E"),
    },
    "Slate": {
        "accent": ("#334155", "#64748B"),
        "accent_hover": ("#1F2937", "#334155"),
        "ok": ("#16A34A", "#22C55E"),
        "warn": ("#D97706", "#F59E0B"),
        "bad": ("#DC2626", "#EF4444"),
        "card": ("#F8FAFC", "#0B1220"),
        "card2": ("#FFFFFF", "#070D18"),
        "border": ("#E2E8F0", "#1F2937"),
        "text": ("#0F172A", "#E5E7EB"),
        "muted": ("#64748B", "#94A3B8"),
        "table_head": ("#E2E8F0", "#0f172a"),
        "table_head_text": ("#0F172A", "#E5E7EB"),
        "select": ("#CBD5E1", "#1F2937"),
    },
    "Mono": {
        "accent": ("#111827", "#E5E7EB"),
        "accent_hover": ("#0B1220", "#CBD5E1"),
        "ok": ("#111827", "#E5E7EB"),
        "warn": ("#111827", "#E5E7EB"),
        "bad": ("#111827", "#E5E7EB"),
        "card": ("#F9FAFB", "#0B0F19"),
        "card2": ("#FFFFFF", "#070A12"),
        "border": ("#E5E7EB", "#1F2937"),
        "text": ("#111827", "#E5E7EB"),
        "muted": ("#6B7280", "#9CA3AF"),
        "table_head": ("#F3F4F6", "#0f172a"),
        "table_head_text": ("#111827", "#E5E7EB"),
        "select": ("#E5E7EB", "#1F2937"),
    },
}


def iter_children_recursive(root_widget):
    stack = [root_widget]
    while stack:
        w = stack.pop()
        yield w
        try:
            stack.extend(w.winfo_children())
        except Exception:
            pass


def backup_single_router(ip, username, password, port, output_dir, backups, log_fn):
    try:
        log_fn(f"Connecting to {ip} ...")
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(ip, port=int(port), username=username, password=password, timeout=15)

        timestamp = datetime.now().strftime("%Y-%m-%d")
        safe_ip = str(ip).replace(".", "_")

        if backups.get("system"):
            backup_name = f"{safe_ip}_{timestamp}"
            cmd = f'/system backup save name={backup_name}'
            stdin, stdout, stderr = client.exec_command(cmd)
            stdout.channel.recv_exit_status()
            remote_path = f"/{backup_name}.backup"
            local_path = os.path.join(output_dir, f"{backup_name}.backup")
            with SCPClient(client.get_transport()) as scp:
                scp.get(remote_path, local_path)
            log_fn(f"System backup saved for {ip}.")

        if backups.get("config"):
            config_name = f"{safe_ip}_{timestamp}"
            cmd = f'/export file={config_name}'
            stdin, stdout, stderr = client.exec_command(cmd)
            stdout.channel.recv_exit_status()
            remote_path = f"/{config_name}.rsc"
            local_path = os.path.join(output_dir, f"{config_name}.rsc")
            with SCPClient(client.get_transport()) as scp:
                scp.get(remote_path, local_path)
            log_fn(f"Config export saved for {ip}.")

        client.close()
    except Exception as e:
        log_fn(f"ERROR on {ip}: {e}")
class SSHFrame(ctk.CTkFrame):
    """Stage 1 tab: concurrent SSH exec - NO LICENSE"""

    def __init__(self, master):
        super().__init__(master)
        self._running = False
        self.cancel_event = threading.Event()
        self.log_q = queue.Queue()
        self.prog_q = queue.Queue()
        self._build()

    def _build(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)

        form = ctk.CTkScrollableFrame(self, corner_radius=12, width=520)
        form.grid(row=0, column=0, sticky="nsew", padx=(12, 8), pady=12)
        form.grid_columnconfigure(0, weight=0)
        form.grid_columnconfigure(1, weight=1)
        form.grid_columnconfigure(2, weight=0)

        log = ctk.CTkFrame(self, corner_radius=12)
        log.grid(row=0, column=1, sticky="nsew", padx=(8, 12), pady=12)
        log.grid_rowconfigure(2, weight=1)
        log.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(form, text="SSH • Run commands on multiple routers", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(12, 6))

        self.var_excel = tk.StringVar()
        self.var_sheet = tk.StringVar()
        self.var_ipcol = tk.StringVar(value="RouterIP")
        self.var_cmd = tk.StringVar()
        self.var_user = tk.StringVar()
        self.var_pass = tk.StringVar()
        self.var_port = tk.StringVar(value="22")
        self.var_usercol = tk.StringVar()
        self.var_passcol = tk.StringVar()
        self.var_portcol = tk.StringVar()
        self.var_startrow = tk.StringVar(value="0")
        self.var_endrow = tk.StringVar(value="")
        self.var_cto = tk.StringVar(value="30")
        self.var_cmdto = tk.StringVar(value="25")
        self.var_cmddelay = tk.StringVar(value="0")
        self.var_conc = tk.StringVar(value="10")

        r = 1

        def row_label(text):
            nonlocal r
            ctk.CTkLabel(form, text=text).grid(row=r, column=0, sticky="w", padx=14, pady=(8, 2))
            r += 1

        def row_entry(var, width=280, placeholder=""):
            nonlocal r
            e = ctk.CTkEntry(form, textvariable=var, width=width, placeholder_text=placeholder)
            e.grid(row=r, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 6))
            r += 1
            return e

        def row_browse(btn_text, cmd):
            nonlocal r
            b = ctk.CTkButton(form, text=btn_text, command=cmd, width=90)
            b.grid(row=r-1, column=2, sticky="e", padx=14, pady=(0, 6))
            return b

        row_label("Routers Excel file")
        row_entry(self.var_excel, width=320)
        row_browse("Browse", self._choose_excel)

        row_label("Excel options (optional)")
        pair = ctk.CTkFrame(form, fg_color="transparent")
        pair.grid(row=r, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 6))
        pair.grid_columnconfigure(0, weight=1)
        pair.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(pair, text="Sheet").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ctk.CTkLabel(pair, text="IP Column").grid(row=0, column=1, sticky="w")

        ctk.CTkEntry(pair, textvariable=self.var_sheet, placeholder_text="e.g. Sheet1").grid(
            row=1, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkEntry(pair, textvariable=self.var_ipcol, placeholder_text="RouterIP").grid(
            row=1, column=1, sticky="ew")
        r += 1

        row_label("Commands file (.txt / .rsc)")
        row_entry(self.var_cmd, width=320)
        row_browse("Browse", self._choose_cmd)

        ctk.CTkLabel(form, text="Username").grid(row=r, column=0, sticky="w", padx=14, pady=(10, 2))
        ctk.CTkLabel(form, text="Password (optional)").grid(row=r, column=1, sticky="w", padx=14, pady=(10, 2))
        ctk.CTkLabel(form, text="Port").grid(row=r, column=2, sticky="w", padx=14, pady=(10, 2))
        r += 1
        ctk.CTkEntry(form, textvariable=self.var_user, width=120, placeholder_text="admin").grid(row=r, column=0, sticky="w", padx=14, pady=(0, 6))
        ctk.CTkEntry(form, textvariable=self.var_pass, width=120, show="•", placeholder_text="optional").grid(row=r, column=1, sticky="w", padx=14, pady=(0, 6))
        ctk.CTkEntry(form, textvariable=self.var_port, width=80, placeholder_text="22").grid(row=r, column=2, sticky="w", padx=14, pady=(0, 6))
        r += 1

        row_label("Credentials from Excel (optional)")
        ctk.CTkLabel(form, text="User Col").grid(row=r, column=0, sticky="w", padx=14)
        ctk.CTkLabel(form, text="Pass Col").grid(row=r, column=1, sticky="w", padx=14)
        ctk.CTkLabel(form, text="Port Col").grid(row=r, column=2, sticky="w", padx=14)
        r += 1
        ctk.CTkEntry(form, textvariable=self.var_usercol, width=120, placeholder_text="e.g. Username").grid(row=r, column=0, sticky="w", padx=14, pady=(0, 6))
        ctk.CTkEntry(form, textvariable=self.var_passcol, width=120, placeholder_text="e.g. Password").grid(row=r, column=1, sticky="w", padx=14, pady=(0, 6))
        ctk.CTkEntry(form, textvariable=self.var_portcol, width=80, placeholder_text="Port").grid(row=r, column=2, sticky="w", padx=14, pady=(0, 6))
        r += 1

        # ========== SINGLE IP SECTION ==========
        ctk.CTkLabel(form, text="━━━━━━━━━━━━ Single IP Mode ━━━━━━━━━━━━", font=ctk.CTkFont(size=12, weight="bold")).grid(
            row=r, column=0, columnspan=3, sticky="ew", padx=14, pady=(10, 5))
        r += 1
        
        self.var_single_ip_enabled = tk.BooleanVar(value=False)
        self.single_ip_check = ctk.CTkCheckBox(form, text="Enable Single IP Mode", variable=self.var_single_ip_enabled)
        self.single_ip_check.grid(row=r, column=0, columnspan=3, sticky="w", padx=14, pady=(5, 5))
        r += 1
        
        single_frame = ctk.CTkFrame(form, fg_color="transparent")
        single_frame.grid(row=r, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 10))
        single_frame.grid_columnconfigure(0, weight=0)
        single_frame.grid_columnconfigure(1, weight=0)
        single_frame.grid_columnconfigure(2, weight=0)
        single_frame.grid_columnconfigure(3, weight=0)
        single_frame.grid_columnconfigure(4, weight=0)
        single_frame.grid_columnconfigure(5, weight=1)
        r += 1
        
        self.var_single_ip = tk.StringVar()
        self.var_single_ip_user = tk.StringVar()
        self.var_single_ip_pass = tk.StringVar()
        self.var_single_ip_port = tk.StringVar(value="22")
        
        ctk.CTkLabel(single_frame, text="IP:", width=30).grid(row=0, column=0, sticky="w", padx=(0, 5))
        ctk.CTkEntry(single_frame, textvariable=self.var_single_ip, width=130, placeholder_text="192.168.1.1").grid(row=0, column=1, padx=(0, 15))
        
        ctk.CTkLabel(single_frame, text="User:", width=35).grid(row=0, column=2, sticky="w", padx=(0, 5))
        ctk.CTkEntry(single_frame, textvariable=self.var_single_ip_user, width=110, placeholder_text="admin").grid(row=0, column=3, padx=(0, 15))
        
        ctk.CTkLabel(single_frame, text="Pass:", width=35).grid(row=0, column=4, sticky="w", padx=(0, 5))
        ctk.CTkEntry(single_frame, textvariable=self.var_single_ip_pass, width=110, show="•", placeholder_text="optional").grid(row=0, column=5, padx=(0, 15))
        
        ctk.CTkLabel(single_frame, text="Port:", width=30).grid(row=1, column=0, sticky="w", padx=(0, 5), pady=(5, 0))
        ctk.CTkEntry(single_frame, textvariable=self.var_single_ip_port, width=70, placeholder_text="22").grid(row=1, column=1, sticky="w", padx=(0, 15), pady=(5, 0))
        
        ctk.CTkLabel(form, text="─────────────────────────────────────────────────").grid(
            row=r, column=0, columnspan=3, sticky="ew", padx=14, pady=(5, 10))
        r += 1
        # ========== END SINGLE IP SECTION ==========

        row_label("Row Range / Timeouts")
        grid = ctk.CTkFrame(form, fg_color="transparent")
        grid.grid(row=r, column=0, columnspan=3, sticky="w", padx=14, pady=(0, 6))
        r += 1
        ctk.CTkLabel(grid, text="Start").grid(row=0, column=0, sticky="w")
        ctk.CTkEntry(grid, textvariable=self.var_startrow, width=60).grid(row=1, column=0, padx=(0, 8))
        ctk.CTkLabel(grid, text="End").grid(row=0, column=1, sticky="w")
        ctk.CTkEntry(grid, textvariable=self.var_endrow, width=60, placeholder_text="empty").grid(row=1, column=1, padx=(0, 8))
        ctk.CTkLabel(grid, text="Conn(s)").grid(row=0, column=2, sticky="w")
        ctk.CTkEntry(grid, textvariable=self.var_cto, width=60).grid(row=1, column=2, padx=(0, 8))
        ctk.CTkLabel(grid, text="Cmd(s)").grid(row=0, column=3, sticky="w")
        ctk.CTkEntry(grid, textvariable=self.var_cmdto, width=60).grid(row=1, column=3, padx=(0, 8))
        ctk.CTkLabel(grid, text="Delay").grid(row=0, column=4, sticky="w")
        ctk.CTkEntry(grid, textvariable=self.var_cmddelay, width=60).grid(row=1, column=4, padx=(0, 8))
        ctk.CTkLabel(grid, text="Concurrency").grid(row=0, column=5, sticky="w")
        ctk.CTkEntry(grid, textvariable=self.var_conc, width=80).grid(row=1, column=5, padx=(0, 0))

        btnrow = ctk.CTkFrame(form, fg_color="transparent")
        btnrow.grid(row=r, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 12))
        r += 1
        self.btn_run = ctk.CTkButton(btnrow, text="▶ Run", command=self._run, width=110)
        self.btn_run.pack(side="left")
        self.btn_cancel = ctk.CTkButton(btnrow, text="⛔ Cancel", command=self._cancel, width=110, fg_color="#9b1c1c")
        self.btn_cancel.pack(side="left", padx=8)
        self.btn_reset = ctk.CTkButton(btnrow, text="↺ Reset", command=self._reset, width=110, fg_color="#374151")
        self.btn_reset.pack(side="left", padx=8)
        self._set_running(False)

        ctk.CTkLabel(log, text="Logs", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, sticky="w", padx=14, pady=(12, 6))
        self.lbl_prog = ctk.CTkLabel(log, text="0/0")
        self.lbl_prog.grid(row=0, column=0, sticky="e", padx=14, pady=(12, 6))
        self.pbar = ctk.CTkProgressBar(log)
        self.pbar.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 8))
        self.pbar.set(0)

        self.txt_log = ctk.CTkTextbox(log, wrap="none")
        self.txt_log.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 14))
        self.txt_log.configure(state="disabled")

        self.after(200, self._pump_queues)

    def _set_running(self, running: bool):
        self._running = running
        self.btn_run.configure(state="disabled" if running else "normal")
        self.btn_cancel.configure(state="normal" if running else "disabled")

    def _append_log(self, s: str):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", s + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _choose_excel(self):
        p = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel", "*.xlsx;*.xls")])
        if p:
            self.var_excel.set(p)

    def _choose_cmd(self):
        p = filedialog.askopenfilename(
            title="Select commands text file",
            filetypes=[("Text", "*.txt;*.rsc;*.cfg;*.conf;*.scr;*.ros"), ("All", "*.*")]
        )
        if p:
            self.var_cmd.set(p)

    def _reset(self):
        if self._running:
            messagebox.showwarning("Running", "Already running. Click Cancel first.")
            return
        self.var_excel.set("")
        self.var_sheet.set("")
        self.var_ipcol.set("RouterIP")
        self.var_cmd.set("")
        self.var_user.set("")
        self.var_pass.set("")
        self.var_port.set("22")
        self.var_usercol.set("")
        self.var_passcol.set("")
        self.var_portcol.set("")
        self.var_startrow.set("0")
        self.var_endrow.set("")
        self.var_cto.set("30")
        self.var_cmdto.set("25")
        self.var_cmddelay.set("0")
        self.var_conc.set("10")
        self.var_single_ip_enabled.set(False)
        self.var_single_ip.set("")
        self.var_single_ip_user.set("")
        self.var_single_ip_pass.set("")
        self.var_single_ip_port.set("22")
        self.pbar.set(0)
        self.lbl_prog.configure(text="0/0")
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

    def _cancel(self):
        if not self._running:
            return
        self.cancel_event.set()
        self._append_log("[CANCEL] Cancel requested...")

    def _run(self):
        if self._running:
            return

        excel = self.var_excel.get().strip()
        sheet = self.var_sheet.get().strip() or None
        ipcol = self.var_ipcol.get().strip() or "RouterIP"
        cmdf = self.var_cmd.get().strip()

        user = self.var_user.get().strip()
        pw = self.var_pass.get()
        port = self.var_port.get().strip() or "22"
        port = re.sub(r'\.0$', '', port)

        usercol = self.var_usercol.get().strip() or None
        passcol = self.var_passcol.get().strip() or None
        portcol = self.var_portcol.get().strip() or None

        try:
            start_row = int(self.var_startrow.get() or "0")
        except Exception:
            start_row = 0
        end_raw = self.var_endrow.get().strip()
        end_row = int(end_raw) if end_raw else None

        try:
            t_conn = int(self.var_cto.get() or "30")
        except Exception:
            t_conn = 30
        try:
            t_cmd = int(self.var_cmdto.get() or "25")
        except Exception:
            t_cmd = 25
        try:
            delay = float(self.var_cmddelay.get() or "0")
        except Exception:
            delay = 0.0
        try:
            conc = int(self.var_conc.get() or "1")
        except Exception:
            conc = 1

        if conc < 1:
            messagebox.showerror("Concurrency", "Concurrency must be 1 or higher.")
            return

        # ========== SINGLE IP MODE CHECK ==========
        if self.var_single_ip_enabled.get():
            single_ip = self.var_single_ip.get().strip()
            if not single_ip:
                messagebox.showerror("Missing", "Single IP Mode is enabled but no IP address entered.")
                return
            
            single_user = self.var_single_ip_user.get().strip()
            if not single_user:
                single_user = user if user else "admin"
            
            single_pass = self.var_single_ip_pass.get()
            single_port = self.var_single_ip_port.get().strip() or "22"
            single_port = re.sub(r'\.0$', '', single_port)
            
            if not cmdf or not os.path.exists(cmdf):
                messagebox.showerror("Missing", "Commands file path is invalid.")
                return
            
            self.cancel_event.clear()
            self._set_running(True)
            self._append_log(f"[RUN] Single IP Mode - Processing: {single_ip}")
            
            def worker_single():
                try:
                    out_dir = os.path.join(OUT_BASE, now_ts() + "_single")
                    os.makedirs(out_dir, exist_ok=True)
                    
                    commands = load_commands_list(cmdf)
                    
                    log_path = os.path.join(out_dir, f"{single_ip.replace(':', '_')}.log")
                    
                    self.log_q.put(("run", f"[{single_ip}] Connecting..."))
                    ssh = ssh_connect(single_ip, single_port, single_user, single_pass, timeout=t_conn)
                    
                    if isinstance(ssh, str):
                        self.log_q.put(("err", f"[{single_ip}] {ssh}"))
                        with open(log_path, "w", encoding="utf-8") as fo:
                            fo.write(f"Connection failed: {ssh}\n")
                    else:
                        any_err = False
                        results = exec_commands_fast(
                            ssh, commands, per_cmd_timeout=t_cmd, cmd_delay=delay,
                            on_cmd=lambda c: self.log_q.put(("cmd", f"[{single_ip}] $ {c}")),
                            cancel_event=self.cancel_event
                        )
                        ssh.close()
                        
                        with open(log_path, "w", encoding="utf-8") as fo:
                            for cmd, out, err in results:
                                fo.write(f"$ {cmd}\n")
                                if out:
                                    fo.write(out)
                                if err:
                                    any_err = True
                                    fo.write(f"[ERR] {err}\n")
                                fo.write("\n")
                        
                        status = "ok" if not any_err else "failed"
                        self.log_q.put(("ok" if status == "ok" else "warn", f"[{single_ip}] Done."))
                    
                    self.prog_q.put(("done", 1, 1, out_dir, self.cancel_event.is_set()))
                except Exception as e:
                    self.log_q.put(("err", f"[FATAL] {e}"))
                    self.prog_q.put(("done", 0, 0, None, True))
                finally:
                    self._set_running(False)
            
            threading.Thread(target=worker_single, daemon=True).start()
            return
        # ========== END SINGLE IP MODE CHECK ==========

        # حالت عادی با فایل اکسل
        if not (excel and os.path.exists(excel) and cmdf and os.path.exists(cmdf)):
            messagebox.showerror("Missing", "Excel / Commands path is invalid.")
            return

        self.cancel_event.clear()
        self._set_running(True)
        self._append_log("[RUN] Stage 1 starting...")

        def worker():
            try:
                run_stage1_exec_concurrent(
                    excel_path=excel, sheet_name=sheet, ip_col=ipcol,
                    cmd_file=cmdf,
                    fallback_user=user, fallback_pass=pw, fallback_port=port,
                    user_col=usercol, pass_col=passcol, port_col=portcol,
                    start_row=start_row, end_row=end_row,
                    timeout_conn=t_conn, timeout_cmd=t_cmd, cmd_delay=delay,
                    concurrency=conc,
                    log_q=self.log_q, prog_q=self.prog_q,
                    cancel_event=self.cancel_event,
                )
            except Exception as e:
                self.log_q.put(("err", f"[FATAL] {e}"))
                self.prog_q.put(("done", 0, 0, None, True))

        threading.Thread(target=worker, daemon=True).start()

    def _pump_queues(self):
        try:
            while True:
                typ, msg = self.log_q.get_nowait()
                prefix = ""
                if typ == "cmd":
                    prefix = ""
                elif typ == "info":
                    prefix = ""
                elif typ == "warn":
                    prefix = "⚠ "
                elif typ == "err":
                    prefix = "❌ "
                elif typ == "ok":
                    prefix = "✅ "
                elif typ == "run":
                    prefix = "▶ "
                self._append_log(prefix + msg)
        except queue.Empty:
            pass

        try:
            while True:
                item = self.prog_q.get_nowait()
                if not item:
                    continue
                typ = item[0]
                if typ == "tick":
                    _, a, b, _outdir = item
                    self.lbl_prog.configure(text=f"{a}/{b}")
                    self.pbar.set(0 if b == 0 else (a / b))
                elif typ == "done":
                    _, a, b, outdir, cancelled = item
                    self.lbl_prog.configure(text=f"{a}/{b}")
                    self.pbar.set(0 if b == 0 else (a / b))
                    self._set_running(False)
                    if cancelled:
                        messagebox.showinfo("SSH", f"Cancelled.\nProgress: {a}/{b}\nOutputs: {outdir or ''}")
                    else:
                        messagebox.showinfo("SSH", f"Done.\nProcessed: {a}/{b}\nOutputs: {outdir or ''}")
        except queue.Empty:
            pass

        self.after(200, self._pump_queues)
class WinBoxFrame(ctk.CTkFrame):
    """Stage 2 tab: excel viewer + filters + context menu + export - NO LICENSE"""

    def __init__(self, master):
        super().__init__(master)

        self.config_path = get_config_path()
        self.config = self._load_config()

        self.df = pd.DataFrame()
        self.view_df = pd.DataFrame()
        self.sort_col = None
        self.sort_asc = True

        self._ip_col = None
        self._port_col = None

        self.filters = []
        self._stop_event = threading.Event()
        self._auto_thread = None
        
        # متغیرهای مربوط به انتخاب شیت
        self.sheet_var = tk.StringVar(value="")
        self.sheet_menu = None
        self.current_sheet = None

        # Credential management
        self.credentials = self._load_creds()
        self.current_cred = self.credentials[0] if self.credentials else None
        self.current_cred_name = self.current_cred["name"] if self.current_cred else ""
        self.cred_var = tk.StringVar(value=self.current_cred_name)
        self.cred_menu = None

        self._build()
        self.load_excel(force=True)
        self.start_auto_reload()

    def _load_config(self):
        if not os.path.exists(self.config_path):
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_CONFIG, f, indent=2)
            return DEFAULT_CONFIG.copy()
        with open(self.config_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        for k, v in DEFAULT_CONFIG.items():
            cfg.setdefault(k, v)
        return cfg

    def _save_config(self):
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=2)

    def _get_creds_path(self):
        """مسیر فایل ذخیره Credential ها"""
        base = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "MikroTikOps")
        os.makedirs(base, exist_ok=True)
        return os.path.join(base, "winbox_creds.json")

    def _load_creds(self):
        """بارگذاری Credential ها از فایل"""
        default_creds = [
            {"name": "Default", "username": "admin", "password": ""},
            {"name": "ISP1", "username": "admin", "password": ""},
            {"name": "ISP2", "username": "admin", "password": ""},
            {"name": "Backup", "username": "backup", "password": ""}
        ]
        path = self._get_creds_path()
        try:
            with open(path, "r", encoding="utf-8") as f:
                creds = json.load(f)
                if creds and len(creds) > 0:
                    return creds
        except Exception:
            pass
        return default_creds

    def _save_creds(self):
        """ذخیره Credential ها در فایل"""
        path = self._get_creds_path()
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.credentials, f, indent=2)
        except Exception as e:
            print(f"Error saving creds: {e}")

    def _refresh_cred_menu(self):
        """به روز رسانی منوی Credential ها"""
        cred_names = [c["name"] for c in self.credentials]
        if self.cred_menu:
            self.cred_menu.configure(values=cred_names)
            if self.current_cred_name not in cred_names:
                self.current_cred_name = cred_names[0] if cred_names else ""
                self.cred_var.set(self.current_cred_name)

    def _on_cred_change(self, choice):
        """وقتی Credential عوض میشه"""
        self.current_cred_name = choice
        for cred in self.credentials:
            if cred["name"] == choice:
                self.current_cred = cred
                break
        if hasattr(self, 'cred_status_update'):
            self.cred_status_update()

    def _add_cred(self):
        """افزودن Credential جدید"""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Add Credential")
        dialog.geometry("400x300")
        dialog.grab_set()
        dialog.transient(self)

        ctk.CTkLabel(dialog, text="Profile Name:", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=20, pady=(20, 5))
        name_entry = ctk.CTkEntry(dialog, width=300, placeholder_text="e.g. Office, DataCenter")
        name_entry.pack(anchor="w", padx=20, pady=(0, 15))

        ctk.CTkLabel(dialog, text="Username:", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=20, pady=(0, 5))
        user_entry = ctk.CTkEntry(dialog, width=300, placeholder_text="admin")
        user_entry.pack(anchor="w", padx=20, pady=(0, 15))

        ctk.CTkLabel(dialog, text="Password:", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=20, pady=(0, 5))
        pass_entry = ctk.CTkEntry(dialog, width=300, show="*", placeholder_text="optional")
        pass_entry.pack(anchor="w", padx=20, pady=(0, 15))

        def save():
            name = name_entry.get().strip()
            username = user_entry.get().strip()
            password = pass_entry.get()
            if not name:
                messagebox.showwarning("Warning", "Profile name is required.", parent=dialog)
                return
            if not username:
                username = "admin"
            for cred in self.credentials:
                if cred["name"] == name:
                    messagebox.showwarning("Warning", "Profile name already exists.", parent=dialog)
                    return
            self.credentials.append({"name": name, "username": username, "password": password})
            self._save_creds()
            self._refresh_cred_menu()
            self.cred_var.set(name)
            self._on_cred_change(name)
            dialog.destroy()
            self.set_status(f"Credential '{name}' added.")

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(10, 20))
        ctk.CTkButton(btn_frame, text="Save", width=100, command=save).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_frame, text="Cancel", width=100, command=dialog.destroy).pack(side="left")

    def _edit_cred(self):
        """ویرایش Credential فعلی"""
        if not self.current_cred:
            return
        
        dialog = ctk.CTkToplevel(self)
        dialog.title(f"Edit Credential: {self.current_cred['name']}")
        dialog.geometry("400x300")
        dialog.grab_set()
        dialog.transient(self)

        ctk.CTkLabel(dialog, text="Profile Name:", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=20, pady=(20, 5))
        name_entry = ctk.CTkEntry(dialog, width=300)
        name_entry.insert(0, self.current_cred["name"])
        name_entry.pack(anchor="w", padx=20, pady=(0, 15))

        ctk.CTkLabel(dialog, text="Username:", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=20, pady=(0, 5))
        user_entry = ctk.CTkEntry(dialog, width=300)
        user_entry.insert(0, self.current_cred["username"])
        user_entry.pack(anchor="w", padx=20, pady=(0, 15))

        ctk.CTkLabel(dialog, text="Password:", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=20, pady=(0, 5))
        pass_entry = ctk.CTkEntry(dialog, width=300, show="*")
        pass_entry.insert(0, self.current_cred["password"])
        pass_entry.pack(anchor="w", padx=20, pady=(0, 15))

        def save():
            new_name = name_entry.get().strip()
            username = user_entry.get().strip()
            password = pass_entry.get()
            if not new_name:
                messagebox.showwarning("Warning", "Profile name is required.", parent=dialog)
                return
            if not username:
                username = "admin"
            
            if new_name != self.current_cred["name"]:
                for cred in self.credentials:
                    if cred["name"] == new_name:
                        messagebox.showwarning("Warning", "Profile name already exists.", parent=dialog)
                        return
            
            self.current_cred["name"] = new_name
            self.current_cred["username"] = username
            self.current_cred["password"] = password
            self._save_creds()
            self._refresh_cred_menu()
            self.cred_var.set(new_name)
            self.current_cred_name = new_name
            dialog.destroy()
            self.set_status(f"Credential '{new_name}' updated.")

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(10, 20))
        ctk.CTkButton(btn_frame, text="Save", width=100, command=save).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_frame, text="Cancel", width=100, command=dialog.destroy).pack(side="left")

    def _delete_cred(self):
        """حذف Credential فعلی"""
        if not self.current_cred:
            return
        if len(self.credentials) <= 1:
            messagebox.showwarning("Warning", "You must keep at least one credential profile.")
            return
        
        result = messagebox.askyesno("Delete", f"Delete credential '{self.current_cred['name']}'?")
        if result:
            self.credentials.remove(self.current_cred)
            self._save_creds()
            self._refresh_cred_menu()
            new_name = self.credentials[0]["name"]
            self.cred_var.set(new_name)
            self._on_cred_change(new_name)
            self.set_status("Credential deleted.")

    def _refresh(self):
        """Refresh the table view based on current data and filters"""
        if self.df is None or self.df.empty:
            self.view_df = pd.DataFrame()
            self._build_tree_columns([])
            return

        view = self.df.copy()

        q = self.search_var.get().strip().lower()
        if q:
            mask = view.apply(lambda r: any(q in str(v).lower() for v in r.values), axis=1)
            view = view[mask].copy()

        for f in self.filters:
            col = f.get("col")
            val = (f.get("val") or "").strip()
            mode = f.get("mode", "contains")
            if not col or col not in view.columns or val == "":
                continue
            series = view[col].astype(str)
            if mode == "equals":
                view = view[series.str.lower() == val.lower()].copy()
            elif mode == "startswith":
                view = view[series.str.lower().str.startswith(val.lower())].copy()
            else:
                view = view[series.str.lower().str.contains(val.lower(), na=False)].copy()

        if self.sort_col and self.sort_col in view.columns:
            try:
                view["_sort_key_"] = pd.to_numeric(view[self.sort_col], errors="coerce")
                if view["_sort_key_"].notna().any():
                    view = view.sort_values(by="_sort_key_", ascending=self.sort_asc, na_position="last")
                else:
                    view = view.sort_values(by=self.sort_col, ascending=self.sort_asc, na_position="last",
                                            key=lambda s: s.str.lower())
            except Exception:
                view = view.sort_values(by=self.sort_col, ascending=self.sort_asc, na_position="last")
            if "_sort_key_" in view.columns:
                view = view.drop(columns=["_sort_key_"])

        self.view_df = view.reset_index(drop=True)
        self._build_tree_columns(self.view_df.columns.tolist())
        self._populate_table(self.view_df)

    def _refresh_sheets(self):
        """دریافت لیست شیت‌های فایل اکسل"""
        path = self.config.get("excel_path")
        if not path or not os.path.exists(path):
            if self.sheet_menu:
                self.sheet_menu.configure(values=[])
                self.sheet_var.set("")
            return
        try:
            xl = pd.ExcelFile(path)
            sheets = xl.sheet_names
            if self.sheet_menu:
                self.sheet_menu.configure(values=sheets)
                if sheets:
                    current = self.sheet_var.get()
                    if current not in sheets:
                        self.sheet_var.set(sheets[0])
                        self._on_sheet_change(sheets[0])
        except Exception as e:
            print(f"Error loading sheets: {e}")

    def _on_sheet_change(self, sheet_name):
        """وقتی شیت عوض میشه، دیتا رو دوباره بارگذاری کن"""
        if not sheet_name:
            return
        path = self.config.get("excel_path")
        if not path or not os.path.exists(path):
            return
        try:
            df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
            df.columns = [str(c).strip() for c in df.columns]
            df = df.fillna("")
            for c in df.columns:
                df[c] = df[c].astype(str).str.strip()
                df[c] = df[c].replace("nan", "")
            self.df = df
            self.current_sheet = sheet_name
            
            cols_lower = [c.lower() for c in df.columns]
            ip_idx = self.detect_ip_column(cols_lower)
            if ip_idx is not None and ip_idx < len(df.columns):
                self._ip_col = df.columns[ip_idx]
                self.df = self.df[self.df[self._ip_col] != ""].reset_index(drop=True)
            
            port_idx = self.detect_port_column(cols_lower)
            self._port_col = df.columns[port_idx] if port_idx is not None else None
            
            self._refresh()
            self.set_status(f"Loaded sheet '{sheet_name}' with {len(self.df)} rows")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheet '{sheet_name}': {e}")

    def _render_filters(self):
        for w in list(self.filters_row.winfo_children()):
            w.destroy()

        if not self.filters:
            ctk.CTkLabel(self.filters_row, text="(No column filters)", text_color="#9CA3AF").pack(side="left")
            return

        ctk.CTkLabel(self.filters_row, text="Filters:", text_color="#9CA3AF").pack(side="left", padx=(0, 8))
        for i, f in enumerate(self.filters):
            chip = ctk.CTkFrame(self.filters_row, corner_radius=999, fg_color="#1f2937")
            chip.pack(side="left", padx=4, pady=2)
            txt = f'{f["col"]} {f["mode"]} "{f["val"]}"'
            ctk.CTkLabel(chip, text=txt).pack(side="left", padx=(10, 6), pady=4)
            ctk.CTkButton(chip, text="X", width=26, height=26, corner_radius=999,
                          fg_color="#374151", hover_color="#4b5563",
                          command=lambda idx=i: self.remove_filter(idx)).pack(side="left", padx=(0, 6), pady=4)

        ctk.CTkButton(self.filters_row, text="Clear", width=70, fg_color="#374151", hover_color="#4b5563",
                      command=self.clear_filters).pack(side="left", padx=8)

    def _on_right_click(self, event):
        iid = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if iid:
            self.tree.selection_set(iid)
        self._rc_row_iid = iid
        self._rc_col_id = col
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    def _clipboard_set(self, text: str):
        try:
            self.clipboard_clear()
            self.clipboard_append(text)
            self.set_status("Copied to clipboard")
        except Exception:
            pass
    def copy_row(self):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        vals = self.tree.item(iid, "values")
        self._clipboard_set("\t".join(map(str, vals)))

    def copy_cell(self):
        iid = self._rc_row_iid
        colid = self._rc_col_id
        if not iid or not colid:
            return
        try:
            col_index = int(colid.replace("#", "")) - 1
        except Exception:
            return
        vals = self.tree.item(iid, "values")
        if 0 <= col_index < len(vals):
            self._clipboard_set(str(vals[col_index]))

    def copy_ip(self):
        if self._ip_col is None:
            return
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        try:
            row = self.view_df.iloc[int(iid)]
            self._clipboard_set(str(row[self._ip_col]).strip())
        except Exception:
            pass

    def copy_selected_text(self):
        w = self.focus_get()
        if isinstance(w, (tk.Entry, ttk.Entry)):
            try:
                txt = w.selection_get()
                self._clipboard_set(txt)
            except Exception:
                pass
        elif isinstance(w, tk.Text):
            try:
                txt = w.get("sel.first", "sel.last")
                self._clipboard_set(txt)
            except Exception:
                pass

    def open_cell_viewer(self):
        iid = self._rc_row_iid
        colid = self._rc_col_id
        if not iid or not colid:
            return
        try:
            col_index = int(colid.replace("#", "")) - 1
        except Exception:
            return
        vals = self.tree.item(iid, "values")
        if not (0 <= col_index < len(vals)):
            return
        cell_text = str(vals[col_index])

        win = ctk.CTkToplevel(self)
        win.title("Cell Viewer")
        win.geometry("520x220")
        win.grab_set()

        ctk.CTkLabel(win, text="Cell text (select with mouse to copy):").pack(anchor="w", padx=14, pady=(14, 6))
        entry = ctk.CTkEntry(win, width=480)
        entry.pack(padx=14, pady=(0, 10))
        entry.insert(0, cell_text)
        entry.select_range(0, "end")
        entry.focus_set()

        def copy_all():
            self._clipboard_set(entry.get())

        ctk.CTkButton(win, text="Copy (All)", width=140, command=copy_all).pack(pady=(0, 14))

    def open_excel_file(self):
        path = filedialog.askopenfilename(title="Select Excel", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        self.config["excel_path"] = path
        self._save_config()
        self._refresh_sheets()
        self.load_excel(force=True)

    def detect_ip_column(self, cols_lower):
        if "ip" in cols_lower:
            return cols_lower.index("ip")
        for i, c in enumerate(cols_lower):
            if "ip" in c:
                return i
        return 1 if len(cols_lower) > 1 else 0

    def detect_port_column(self, cols_lower):
        if "port" in cols_lower:
            return cols_lower.index("port")
        for i, c in enumerate(cols_lower):
            if "port" in c:
                return i
        return None

    def load_excel(self, force=False):
        path = self.config.get("excel_path")
        if not path or not os.path.exists(path):
            self.set_status(f"Excel not found: {path}")
            return

        try:
            mtime = os.path.getmtime(path)
            if not force and hasattr(self, "_last_mtime") and self._last_mtime == mtime:
                return

            if self.sheet_var.get():
                sheet_name = self.sheet_var.get()
            else:
                sheet_name = 0
            
            df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
            if df.empty:
                raise ValueError("Excel has no rows.")

            df.columns = [str(c).strip() for c in df.columns]
            cols_lower = [c.lower() for c in df.columns]

            ip_idx = self.detect_ip_column(cols_lower)
            if ip_idx is None or ip_idx >= len(df.columns):
                raise ValueError("No IP column found. Ensure an 'IP' column exists.")

            df = df.fillna("")
            for c in df.columns:
                df[c] = df[c].astype(str).str.strip()
                df[c] = df[c].replace("nan", "")

            df = df[df.iloc[:, ip_idx] != ""].reset_index(drop=True)

            self.df = df
            self._ip_col = df.columns[ip_idx]

            port_idx = self.detect_port_column(cols_lower)
            self._port_col = df.columns[port_idx] if port_idx is not None else None

            self._refresh()
            self._last_mtime = mtime
            
            self._refresh_sheets()
            
            self.set_status(f"Loaded {len(self.df)} rows")
        except Exception as e:
            messagebox.showerror("Excel Error", str(e))
            self.set_status("Excel load failed")

    def _build_tree_columns(self, columns):
        for c in self.tree["columns"]:
            try:
                self.tree.heading(c, command=None)
            except Exception:
                pass
        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.on_heading_click(c))
            self.tree.column(col, width=170, anchor=tk.W)

    def _populate_table(self, df: pd.DataFrame):
        self.tree.delete(*self.tree.get_children())
        for i, row in df.iterrows():
            vals = [row.get(c, "") for c in df.columns]
            self.tree.insert("", tk.END, iid=str(i), values=vals)

    def on_heading_click(self, col):
        if self.sort_col == col:
            self.sort_asc = not self.sort_asc
        else:
            self.sort_col = col
            self.sort_asc = True
        self._refresh()

    def open_filter_dialog(self):
        if self.df is None or self.df.empty:
            messagebox.showinfo("Filter", "Load an Excel file first.")
            return

        win = ctk.CTkToplevel(self)
        win.title("Add Filter")
        win.geometry("420x260")
        win.grab_set()

        cols = list(self.df.columns)

        var_col = tk.StringVar(value=cols[0] if cols else "")
        var_mode = tk.StringVar(value="contains")
        var_val = tk.StringVar()

        ctk.CTkLabel(win, text="Column").pack(anchor="w", padx=14, pady=(14, 2))
        ctk.CTkOptionMenu(win, values=cols, variable=var_col).pack(anchor="w", padx=14, pady=(0, 10))

        ctk.CTkLabel(win, text="Match").pack(anchor="w", padx=14, pady=(0, 2))
        ctk.CTkOptionMenu(win, values=["contains", "equals", "startswith"], variable=var_mode).pack(anchor="w", padx=14, pady=(0, 10))

        ctk.CTkLabel(win, text="Value").pack(anchor="w", padx=14, pady=(0, 2))
        ctk.CTkEntry(win, textvariable=var_val, width=360, placeholder_text="e.g. Tehran or 10.10").pack(anchor="w", padx=14, pady=(0, 14))

        def add():
            col = var_col.get().strip()
            val = var_val.get().strip()
            mode = var_mode.get().strip() or "contains"
            if not col or not val:
                messagebox.showwarning("Filter", "Please enter both Column and Value.")
                return
            self.filters.append({"col": col, "val": val, "mode": mode})
            self._render_filters()
            self._refresh()
            win.destroy()

        ctk.CTkButton(win, text="Add", command=add, width=120).pack(pady=(0, 14))

    def remove_filter(self, idx: int):
        try:
            self.filters.pop(idx)
        except Exception:
            return
        self._render_filters()
        self._refresh()

    def clear_filters(self):
        self.filters = []
        self._render_filters()
        self._refresh()

    def export_filtered(self):
        if self.view_df is None or self.view_df.empty:
            messagebox.showinfo("Export", "Nothing to export.")
            return

        parts = []
        q = self.search_var.get().strip()
        if q:
            parts.append(q)
        if self.filters:
            for f in self.filters:
                parts.append(f'{f["col"]}-{f["val"]}')
        name = safe_filename("_".join(parts) if parts else "all")

        out_dir = os.path.join(OUT_BASE, "exports")
        os.makedirs(out_dir, exist_ok=True)

        fpath = filedialog.asksaveasfilename(
            title="Save Export",
            initialdir=out_dir,
            initialfile=f"{name}.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")]
        )
        if not fpath:
            return

        try:
            if fpath.lower().endswith(".csv"):
                self.view_df.to_csv(fpath, index=False, encoding="utf-8-sig")
            else:
                self.view_df.to_excel(fpath, index=False, engine="openpyxl")
            self.set_status(f"Exported -> {fpath}")
            messagebox.showinfo("Export", f"Export done.\n{fpath}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))

    def _toggle_pass_cli(self):
        self.config["pass_credentials_on_cli"] = bool(self.pass_cli_var.get())
        self._save_config()

    def set_credentials(self):
        if not KEYRING_OK:
            messagebox.showwarning("Credentials", "keyring is not installed.\nEnable with: pip install keyring")
            return

        user = keyring.get_password(SERVICE_NAME, ACCOUNT_USER) or ""
        new_user = simpledialog.askstring("Username", "Enter Winbox username:", initialvalue=user, parent=self.winfo_toplevel())
        if new_user is None:
            return
        new_pass = simpledialog.askstring("Password", "Enter Winbox password:", show="*", parent=self.winfo_toplevel())
        if new_pass is None:
            return
        try:
            keyring.set_password(SERVICE_NAME, ACCOUNT_USER, new_user)
            keyring.set_password(SERVICE_NAME, ACCOUNT_PASS, new_pass)
            messagebox.showinfo("Saved", "Credentials saved to Windows Credential Manager.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save credentials: {e}")

    def get_credentials(self):
        if not KEYRING_OK:
            return None, None

        user = keyring.get_password(SERVICE_NAME, ACCOUNT_USER)
        pwd = keyring.get_password(SERVICE_NAME, ACCOUNT_PASS)
        if not user or not pwd:
            if not messagebox.askyesno("Credentials", "Credentials not set. Set now?"):
                return None, None
            self.set_credentials()
            user = keyring.get_password(SERVICE_NAME, ACCOUNT_USER)
            pwd = keyring.get_password(SERVICE_NAME, ACCOUNT_PASS)
        return user, pwd

    def open_settings(self):
        try:
            cur_port = str(self.config.get("port", 2121))
            new_port = simpledialog.askstring("Settings", "Default Port (used if no 'Port' column):",
                                              initialvalue=cur_port, parent=self.winfo_toplevel())
            if new_port is not None:
                try:
                    self.config["port"] = int(new_port)
                except ValueError:
                    messagebox.showerror("Settings", "Port must be an integer.")
                    return

            if messagebox.askyesno("Settings", "Change Excel file path?"):
                f = filedialog.askopenfilename(title="Select Excel", filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")])
                if f:
                    self.config["excel_path"] = f
                    self._refresh_sheets()

            if messagebox.askyesno("Settings", "Change winbox.exe path?"):
                f = filedialog.askopenfilename(title="Select winbox.exe", filetypes=[("Winbox", "winbox*.exe"), ("Exe", "*.exe")])
                if f:
                    self.config["winbox_path"] = f

            cur_reload = str(self.config.get("auto_reload_seconds", 10))
            new_reload = simpledialog.askstring("Settings", "Auto-reload seconds (0 to disable):",
                                                initialvalue=cur_reload, parent=self.winfo_toplevel())
            if new_reload is not None:
                try:
                    self.config["auto_reload_seconds"] = int(new_reload)
                except ValueError:
                    messagebox.showerror("Settings", "Auto-reload must be an integer.")
                    return

            self._save_config()
            self.load_excel(force=True)
            messagebox.showinfo("Settings", "Settings saved.")
        except Exception as e:
            messagebox.showerror("Settings", str(e))

    def resolve_winbox(self):
        cfg_path = (self.config.get("winbox_path") or "").strip()
        if cfg_path and os.path.exists(cfg_path):
            return cfg_path
        for p in COMMON_WINBOX_PATHS:
            if p and os.path.exists(p):
                return p
        messagebox.showwarning("Winbox", "Select your winbox.exe")
        f = filedialog.askopenfilename(title="Select winbox.exe", filetypes=[("Winbox", "winbox*.exe"), ("Exe", "*.exe")])
        if f:
            self.config["winbox_path"] = f
            self._save_config()
            return f
        return None

    def connect_selected(self, no_creds=False):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Connect", "Select a row first.")
            return

        iid = sel[0]
        row = self.view_df.iloc[int(iid)]
        ip = str(row[self._ip_col]).strip()

        port = int(self.config.get("port", 2121))
        if self._port_col and self._port_col in row.index:
            pval = str(row[self._port_col]).strip()
            if pval.isdigit():
                port = int(pval)

        if no_creds:
            self.run_winbox_with_creds(ip, port, "", "")
        else:
            if self.current_cred:
                username = self.current_cred.get("username", "admin")
                password = self.current_cred.get("password", "")
                self.run_winbox_with_creds(ip, port, username, password)
            else:
                self.run_winbox_with_creds(ip, port, "", "")

    def run_winbox_with_creds(self, ip: str, port: int, username: str, password: str):
        winbox = self.resolve_winbox()
        if not winbox:
            messagebox.showerror("Winbox", "winbox.exe not found.")
            return

        target = f"{ip}:{port}"
        
        if username and self.pass_cli_var.get():
            args = [winbox, target, username, password]
        else:
            args = [winbox, target]

        try:
            subprocess.Popen(args, creationflags=0x00000008, close_fds=True)
            if username:
                self.set_status(f"Launching Winbox -> {target} as {username}")
            else:
                self.set_status(f"Launching Winbox -> {target}")
        except Exception as e:
            messagebox.showerror("Launch Error", str(e))

    def start_auto_reload(self):
        sec = int(self.config.get("auto_reload_seconds", 10))
        if sec <= 0:
            return

        def loop():
            while not self._stop_event.is_set():
                time.sleep(sec)
                try:
                    self.load_excel(force=False)
                except Exception:
                    pass

        self._auto_thread = threading.Thread(target=loop, daemon=True)
        self._auto_thread.start()

    def stop(self):
        self._stop_event.set()

    def set_status(self, msg: str):
        self.status_var.set(msg)

    def _build(self):
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)
        self.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(self, corner_radius=12)
        top.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 8))
        top.grid_columnconfigure(2, weight=1)

        ctk.CTkLabel(top, text="WinBox • Router list", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 6))

        # ========== ROW 1: SHEET + CREDENTIAL (کنار هم) ==========
        row1_frame = ctk.CTkFrame(top, fg_color="transparent")
        row1_frame.grid(row=1, column=0, columnspan=3, sticky="ew", padx=14, pady=(5, 5))
        row1_frame.grid_columnconfigure(2, weight=1)
        
        # Sheet selection
        ctk.CTkLabel(row1_frame, text="Sheet:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.sheet_var = tk.StringVar(value="")
        self.sheet_menu = ctk.CTkOptionMenu(row1_frame, values=[], variable=self.sheet_var, width=200,
                                             command=self._on_sheet_change)
        self.sheet_menu.grid(row=0, column=1, sticky="w", padx=(0, 20))
        
        # Credential selection
        ctk.CTkLabel(row1_frame, text="Credential:").grid(row=0, column=2, sticky="w", padx=(0, 8))
        cred_names = [c["name"] for c in self.credentials]
        self.cred_menu = ctk.CTkOptionMenu(row1_frame, values=cred_names, variable=self.cred_var, width=150,
                                            command=self._on_cred_change)
        self.cred_menu.grid(row=0, column=3, sticky="w", padx=(0, 5))
        
        # Credential buttons
        btn_add = ctk.CTkButton(row1_frame, text="Add", width=45, command=self._add_cred)
        btn_add.grid(row=0, column=4, sticky="w", padx=(0, 2))
        
        btn_edit = ctk.CTkButton(row1_frame, text="Edit", width=45, command=self._edit_cred)
        btn_edit.grid(row=0, column=5, sticky="w", padx=(0, 2))
        
        btn_delete = ctk.CTkButton(row1_frame, text="Del", width=45, command=self._delete_cred)
        btn_delete.grid(row=0, column=6, sticky="w", padx=(0, 10))
        
        # نمایش Credential فعلی
        self.cred_status_label = ctk.CTkLabel(row1_frame, text="", text_color="gray70", font=ctk.CTkFont(size=11))
        self.cred_status_label.grid(row=0, column=7, sticky="w")
        
        def update_cred_display():
            if self.current_cred:
                pwd_display = '*' * min(6, len(self.current_cred['password'])) if self.current_cred['password'] else '(no pass)'
                self.cred_status_label.configure(text=f"{self.current_cred['username']} / {pwd_display}")
        self.cred_status_update = update_cred_display
        
        if self.current_cred:
            self._on_cred_change(self.current_cred_name)
            update_cred_display()
        # ========== END ROW 1 ==========

        # ========== ROW 2: Search and Buttons ==========
        row2_frame = ctk.CTkFrame(top, fg_color="transparent")
        row2_frame.grid(row=2, column=0, columnspan=3, sticky="ew", padx=14, pady=(5, 5))
        row2_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(row2_frame, text="Search:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._refresh())
        ctk.CTkEntry(row2_frame, textvariable=self.search_var, width=250, placeholder_text="Search...").grid(
            row=0, column=1, sticky="ew", padx=(0, 10))
        
        btns = ctk.CTkFrame(row2_frame, fg_color="transparent")
        btns.grid(row=0, column=2, sticky="e")
        
        ctk.CTkButton(btns, text="Open Excel", width=100, command=self.open_excel_file).pack(side="left", padx=2)
        ctk.CTkButton(btns, text="Add Filter", width=100, command=self.open_filter_dialog).pack(side="left", padx=2)
        ctk.CTkButton(btns, text="Export", width=80, command=self.export_filtered).pack(side="left", padx=2)
        ctk.CTkButton(btns, text="Settings", width=80, command=self.open_settings).pack(side="left", padx=2)
        # ========== END ROW 2 ==========

        # Filters row
        self.filters_row = ctk.CTkFrame(top, fg_color="transparent")
        self.filters_row.grid(row=3, column=0, columnspan=3, sticky="ew", padx=14, pady=(5, 10))
        self._render_filters()

        mid = ctk.CTkFrame(self, corner_radius=12)
        mid.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 8))
        mid.grid_rowconfigure(0, weight=1)
        mid.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        try:
            if "vista" in style.theme_names():
                style.theme_use("vista")
            elif "clam" in style.theme_names():
                style.theme_use("clam")
        except Exception:
            pass
        style.configure("Treeview", rowheight=26)

        wrap = ctk.CTkFrame(mid, corner_radius=10)
        wrap.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        wrap.grid_rowconfigure(0, weight=1)
        wrap.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(wrap, columns=(), show="headings")
        self.tree.grid(row=0, column=0, sticky="nsew")

        yscroll = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll = ttk.Scrollbar(wrap, orient="horizontal", command=self.tree.xview)
        xscroll.grid(row=1, column=0, sticky="ew")

        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.bind("<Double-1>", lambda e: self.connect_selected())
        self.tree.bind("<Button-3>", self._on_right_click)

        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Copy Row", command=self.copy_row)
        self.menu.add_command(label="Copy Cell", command=self.copy_cell)
        self.menu.add_command(label="Open Cell Viewer", command=self.open_cell_viewer)
        self.menu.add_separator()
        self.menu.add_command(label="Copy IP", command=self.copy_ip)
        self.menu.add_separator()
        self.menu.add_command(label="Copy Selected Text", command=self.copy_selected_text)

        bottom = ctk.CTkFrame(self, corner_radius=12)
        bottom.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 12))
        bottom.grid_columnconfigure(2, weight=1)

        ctk.CTkButton(bottom, text="Connect", width=110, command=self.connect_selected).grid(row=0, column=0, padx=14, pady=10, sticky="w")
        ctk.CTkButton(bottom, text="Connect (no creds)", width=160, command=lambda: self.connect_selected(no_creds=True)).grid(row=0, column=1, padx=(0, 8), pady=10, sticky="w")

        self.pass_cli_var = tk.BooleanVar(value=self.config.get("pass_credentials_on_cli", True))
        ctk.CTkCheckBox(bottom, text="Pass user/pass on CLI", variable=self.pass_cli_var, command=self._toggle_pass_cli).grid(
            row=0, column=2, padx=10, pady=10, sticky="w")

        self.status_var = tk.StringVar(value="Ready")
        ctk.CTkLabel(bottom, textvariable=self.status_var).grid(row=0, column=3, padx=14, pady=10, sticky="e")

        if not KEYRING_OK:
            self.set_status("keyring not installed -> Credentials via CLI may be limited (pip install keyring)")

        self._rc_row_iid = None
        self._rc_col_id = None            
            
class BackupFrame(ctk.CTkFrame):
    """BackUp tab: all options in ONE page + schedule - NO LICENSE"""

    def __init__(self, master):
        super().__init__(master)
        self.scheduler = BackgroundScheduler()
        self.scheduler.start()

        self.excel_path = ""
        self.output_dir = os.path.join(os.getcwd(), "backups")
        os.makedirs(self.output_dir, exist_ok=True)

        self._running = False
        self.log_q = queue.Queue()
        self._build()
        self.after(200, self._pump_logs)

    def _build(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        container = ctk.CTkScrollableFrame(self, corner_radius=12)
        container.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)
        container.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(container, text="BackUp • Router backups", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=(10, 8))

        r = 1

        sec1 = ctk.CTkFrame(container, corner_radius=12)
        sec1.grid(row=r, column=0, sticky="ew", padx=10, pady=(0, 10))
        sec1.grid_columnconfigure(1, weight=1)
        r += 1

        ctk.CTkLabel(sec1, text="Routers Excel file:").grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))
        self.excel_entry = ctk.CTkEntry(sec1, width=520)
        self.excel_entry.grid(row=0, column=1, sticky="ew", padx=8, pady=(12, 6))
        ctk.CTkButton(sec1, text="Browse", width=90, command=self.browse_excel).grid(row=0, column=2, padx=12, pady=(12, 6))

        grid = ctk.CTkFrame(sec1, fg_color="transparent")
        grid.grid(row=1, column=0, columnspan=3, sticky="ew", padx=12, pady=(0, 10))
        ctk.CTkLabel(grid, text="Sheet").grid(row=0, column=0, sticky="w")
        self.sheet_entry = ctk.CTkEntry(grid, width=140, placeholder_text="Sheet1")
        self.sheet_entry.insert(0, "Sheet1")
        self.sheet_entry.grid(row=1, column=0, padx=(0, 10))

        ctk.CTkLabel(grid, text="IP Column").grid(row=0, column=1, sticky="w")
        self.ip_col_entry = ctk.CTkEntry(grid, width=160, placeholder_text="IP_Address")
        self.ip_col_entry.insert(0, "IP_Address")
        self.ip_col_entry.grid(row=1, column=1, padx=(0, 10))

        ctk.CTkLabel(grid, text="Start row").grid(row=0, column=2, sticky="w")
        self.start_row_entry = ctk.CTkEntry(grid, width=90)
        self.start_row_entry.insert(0, "2")
        self.start_row_entry.grid(row=1, column=2, padx=(0, 10))

        ctk.CTkLabel(grid, text="End row").grid(row=0, column=3, sticky="w")
        self.end_row_entry = ctk.CTkEntry(grid, width=90, placeholder_text="empty")
        self.end_row_entry.grid(row=1, column=3, padx=(0, 10))

        ctk.CTkButton(grid, text="Preview", width=110, command=self.preview_data).grid(row=1, column=4, padx=(0, 0))

        self.preview_text = ctk.CTkTextbox(sec1, height=110, wrap="none")
        self.preview_text.grid(row=2, column=0, columnspan=3, sticky="ew", padx=12, pady=(0, 12))

        sec2 = ctk.CTkFrame(container, corner_radius=12)
        sec2.grid(row=r, column=0, sticky="ew", padx=10, pady=(0, 10))
        sec2.grid_columnconfigure(1, weight=1)
        r += 1

        ctk.CTkLabel(sec2, text="SSH Username:").grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))
        self.user_entry = ctk.CTkEntry(sec2, width=220)
        self.user_entry.insert(0, "admin")
        self.user_entry.grid(row=0, column=1, sticky="w", padx=8, pady=(12, 6))

        ctk.CTkLabel(sec2, text="Password (optional):").grid(row=1, column=0, sticky="w", padx=12, pady=(0, 6))
        self.pass_entry = ctk.CTkEntry(sec2, width=220, show="*", placeholder_text="optional")
        self.pass_entry.grid(row=1, column=1, sticky="w", padx=8, pady=(0, 6))

        ctk.CTkLabel(sec2, text="Port:").grid(row=0, column=2, sticky="w", padx=12, pady=(12, 6))
        self.port_entry = ctk.CTkEntry(sec2, width=90)
        self.port_entry.insert(0, "22")
        self.port_entry.grid(row=0, column=3, sticky="w", padx=8, pady=(12, 6))

        ctk.CTkLabel(sec2, text="Output folder:").grid(row=2, column=0, sticky="w", padx=12, pady=(0, 6))
        self.output_entry = ctk.CTkEntry(sec2, width=520)
        self.output_entry.insert(0, self.output_dir)
        self.output_entry.grid(row=2, column=1, columnspan=2, sticky="ew", padx=8, pady=(0, 6))
        ctk.CTkButton(sec2, text="Browse", width=90, command=self.browse_output).grid(row=2, column=3, padx=12, pady=(0, 6))

        ctk.CTkLabel(sec2, text="Backup type:").grid(row=3, column=0, sticky="w", padx=12, pady=(6, 6))
        self.backup_type_sys = ctk.CTkCheckBox(sec2, text="System Backup (.backup)")
        self.backup_type_cfg = ctk.CTkCheckBox(sec2, text="Config Export (.rsc)")
        self.backup_type_sys.grid(row=4, column=0, sticky="w", padx=12, pady=(0, 12))
        self.backup_type_cfg.grid(row=4, column=1, sticky="w", padx=12, pady=(0, 12))
        self.backup_type_sys.select()
        self.backup_type_cfg.select()

        sec3 = ctk.CTkFrame(container, corner_radius=12)
        sec3.grid(row=r, column=0, sticky="ew", padx=10, pady=(0, 10))
        sec3.grid_columnconfigure(0, weight=1)
        r += 1

        self.run_button = ctk.CTkButton(sec3, text="Run backup now", command=self.run_backup_thread, height=44,
                                        font=ctk.CTkFont(size=14, weight="bold"))
        self.run_button.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 10))

        sch = ctk.CTkFrame(sec3, fg_color="transparent")
        sch.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 12))

        self.schedule_var = tk.BooleanVar(value=False)
        self.schedule_check = ctk.CTkCheckBox(sch, text="Enable daily schedule", variable=self.schedule_var)
        self.schedule_check.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        ctk.CTkLabel(sch, text="Hour (0-23)").grid(row=1, column=0, sticky="w")
        self.schedule_hour = ctk.CTkEntry(sch, width=70)
        self.schedule_hour.insert(0, "02")
        self.schedule_hour.grid(row=2, column=0, sticky="w", padx=(0, 10))

        ctk.CTkLabel(sch, text="Minute (0-59)").grid(row=1, column=1, sticky="w")
        self.schedule_minute = ctk.CTkEntry(sch, width=70)
        self.schedule_minute.insert(0, "00")
        self.schedule_minute.grid(row=2, column=1, sticky="w", padx=(0, 10))

        ctk.CTkButton(sch, text="Save schedule", width=140, command=self.set_schedule).grid(row=2, column=2, sticky="w")

        self.log_text = ctk.CTkTextbox(container, height=170, wrap="none")
        self.log_text.grid(row=r, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.log_text.configure(state="disabled")

    def log(self, msg: str):
        self.log_q.put(f'{datetime.now().strftime("%H:%M:%S")} - {msg}')

    def _pump_logs(self):
        try:
            while True:
                line = self.log_q.get_nowait()
                self.log_text.configure(state="normal")
                self.log_text.insert("end", line + "\n")
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(200, self._pump_logs)

    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            self.excel_path = path
            self.excel_entry.delete(0, "end")
            self.excel_entry.insert(0, path)

    def browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir = path
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, path)

    def preview_data(self):
        try:
            if not self.excel_path:
                messagebox.showerror("Error", "Please select an Excel file.")
                return
            sheet = self.sheet_entry.get().strip() or "Sheet1"
            ip_col = self.ip_col_entry.get().strip()
            start = int(self.start_row_entry.get()) - 2
            end_input = self.end_row_entry.get().strip()
            end = int(end_input) - 1 if end_input else None

            df = pd.read_excel(self.excel_path, sheet_name=sheet, engine="openpyxl")
            ip_series = df[ip_col].iloc[start:end]
            self.preview_text.delete("1.0", "end")
            self.preview_text.insert("1.0", "Backup IP list:\n" + "\n".join(ip_series.dropna().astype(str)))
        except Exception as e:
            messagebox.showerror("Preview error", str(e))

    def set_schedule(self):
        if not self.schedule_var.get():
            self.scheduler.remove_all_jobs()
            self.log("Schedule disabled.")
            return

        try:
            hour = int(self.schedule_hour.get())
            minute = int(self.schedule_minute.get())
            if not (0 <= hour <= 23 and 0 <= minute <= 59):
                raise ValueError("Invalid hour or minute.")
            self.scheduler.remove_all_jobs()
            self.scheduler.add_job(self.run_backup_thread, 'cron', hour=hour, minute=minute)
            self.log(f"Schedule enabled: daily at {hour:02d}:{minute:02d}")
            messagebox.showinfo("Schedule", f"Saved: daily {hour:02d}:{minute:02d}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to set schedule:\n{e}")

    def run_backup_thread(self):
        if self._running:
            return

        excel_file = self.excel_entry.get().strip()
        sheet = self.sheet_entry.get().strip() or "Sheet1"
        ip_col = self.ip_col_entry.get().strip()

        try:
            start_row = int(self.start_row_entry.get() or "1") - 2
        except Exception:
            start_row = 0
        end_input = self.end_row_entry.get().strip()
        end_row = None
        if end_input:
            try:
                end_row = int(end_input) - 1
            except Exception:
                end_row = None

        if not excel_file or not os.path.exists(excel_file):
            messagebox.showerror("BackUp", "Excel path is invalid.")
            return
        if not ip_col:
            messagebox.showerror("BackUp", "IP column is required.")
            return

        # NO LICENSE CHECK - unlimited devices
        self._running = True
        self.run_button.configure(state="disabled")
        thread = threading.Thread(target=self.run_backup, daemon=True)
        thread.start()

    def run_backup(self):
        try:
            self.log("Starting backup...")

            excel_file = self.excel_entry.get().strip()
            sheet = self.sheet_entry.get().strip() or "Sheet1"
            ip_col = self.ip_col_entry.get().strip()
            start_row = int(self.start_row_entry.get()) - 2
            end_input = self.end_row_entry.get().strip()
            end_row = int(end_input) - 1 if end_input else None

            username = self.user_entry.get().strip()
            password = self.pass_entry.get()
            port = int(self.port_entry.get() or "22")
            output_dir = self.output_entry.get().strip() or self.output_dir

            if not excel_file or not os.path.exists(excel_file):
                self.log("ERROR: Excel path is invalid.")
                return
            if not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)

            df = pd.read_excel(excel_file, sheet_name=sheet, engine="openpyxl")

            ip_list = df[ip_col].iloc[start_row:end_row].dropna().tolist()
            ip_list = [str(x).strip() for x in ip_list if str(x).strip()]

            backups = {"system": bool(self.backup_type_sys.get()),
                       "config": bool(self.backup_type_cfg.get())}

            for ip in ip_list:
                ip = str(ip).strip()
                if not ip:
                    continue
                backup_single_router(ip, username, password, port, output_dir, backups, self.log)

            self.log("Backup completed successfully.")
        except Exception as e:
            self.log(f"ERROR: {e}")
        finally:
            self._running = False
            self.run_button.configure(state="normal")

class IOSFrame(ctk.CTkFrame):
    """Stage 4: Firmware update on MikroTik routers - NO LICENSE"""

    def __init__(self, master):
        super().__init__(master)
        self._running = False
        self.cancel_event = threading.Event()
        self.log_q = queue.Queue()
        self.prog_q = queue.Queue()
        self._build()
        self.after(200, self._pump_queues)

    def _build(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)

        # فرم سمت چپ
        form = ctk.CTkScrollableFrame(self, corner_radius=12, width=520)
        form.grid(row=0, column=0, sticky="nsew", padx=(12, 8), pady=12)
        form.grid_columnconfigure(0, weight=0)
        form.grid_columnconfigure(1, weight=1)
        form.grid_columnconfigure(2, weight=0)

        # لاگ سمت راست
        log = ctk.CTkFrame(self, corner_radius=12)
        log.grid(row=0, column=1, sticky="nsew", padx=(8, 12), pady=12)
        log.grid_rowconfigure(2, weight=1)
        log.grid_columnconfigure(0, weight=1)

        r = 0
        
        def row_label(text, bold=False):
            nonlocal r
            font = ctk.CTkFont(weight="bold") if bold else None
            ctk.CTkLabel(form, text=text, font=font).grid(row=r, column=0, sticky="w", padx=14, pady=(8, 2))
            r += 1

        def row_entry(var, width=280, placeholder=""):
            nonlocal r
            e = ctk.CTkEntry(form, textvariable=var, width=width, placeholder_text=placeholder)
            e.grid(row=r, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 6))
            r += 1
            return e

        def row_browse(btn_text, cmd):
            nonlocal r
            b = ctk.CTkButton(form, text=btn_text, command=cmd, width=90)
            b.grid(row=r-1, column=2, sticky="e", padx=14, pady=(0, 6))
            return b

        # عنوان
        ctk.CTkLabel(form, text="IOS • Firmware Update", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=r, column=0, columnspan=3, sticky="w", padx=14, pady=(12, 6))
        r += 1

        # انتخاب فایل IOS
        row_label("Firmware File (.npk)", bold=True)
        self.var_firmware = tk.StringVar()
        row_entry(self.var_firmware, width=320, placeholder="e.g. routeros-7.x.npk")
        row_browse("Browse", self._choose_firmware)

        # انتخاب فایل اکسل
        row_label("Routers Excel file", bold=True)
        self.var_excel = tk.StringVar()
        row_entry(self.var_excel, width=320)
        row_browse("Browse", self._choose_excel)

        # تنظیمات اکسل
        row_label("Excel options", bold=True)
        sheet_ip_frame = ctk.CTkFrame(form, fg_color="transparent")
        sheet_ip_frame.grid(row=r, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 6))
        sheet_ip_frame.grid_columnconfigure(0, weight=1)
        sheet_ip_frame.grid_columnconfigure(1, weight=1)
        r += 1
        
        ctk.CTkLabel(sheet_ip_frame, text="Sheet").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ctk.CTkLabel(sheet_ip_frame, text="IP Column").grid(row=0, column=1, sticky="w")
        
        self.var_sheet = tk.StringVar()
        ctk.CTkEntry(sheet_ip_frame, textvariable=self.var_sheet, placeholder_text="e.g. Sheet1", width=150).grid(
            row=1, column=0, sticky="ew", padx=(0, 8))
        
        self.var_ipcol = tk.StringVar(value="RouterIP")
        ctk.CTkEntry(sheet_ip_frame, textvariable=self.var_ipcol, placeholder_text="RouterIP", width=150).grid(
            row=1, column=1, sticky="ew")

        # Credentials (فیلدهای جداگانه)
        row_label("SSH Credentials", bold=True)
        cred_frame = ctk.CTkFrame(form, fg_color="transparent")
        cred_frame.grid(row=r, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 6))
        cred_frame.grid_columnconfigure(1, weight=1)
        r += 1
        
        ctk.CTkLabel(cred_frame, text="Username:").grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.var_username = tk.StringVar(value="admin")
        ctk.CTkEntry(cred_frame, textvariable=self.var_username, width=150, placeholder_text="admin").grid(row=0, column=1, sticky="w", padx=(0, 15))
        
        ctk.CTkLabel(cred_frame, text="Password:").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(5, 0))
        self.var_password = tk.StringVar()
        ctk.CTkEntry(cred_frame, textvariable=self.var_password, width=150, show="*", placeholder_text="optional").grid(row=1, column=1, sticky="w", padx=(0, 15), pady=(5, 0))
        
        ctk.CTkLabel(cred_frame, text="Port:").grid(row=0, column=2, sticky="w", padx=(15, 10))
        self.var_port = tk.StringVar(value="22")
        ctk.CTkEntry(cred_frame, textvariable=self.var_port, width=80, placeholder_text="22").grid(row=0, column=3, sticky="w")

        # Row Range
        row_label("Row Range", bold=True)
        range_frame = ctk.CTkFrame(form, fg_color="transparent")
        range_frame.grid(row=r, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 6))
        r += 1
        
        ctk.CTkLabel(range_frame, text="Start").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.var_startrow = tk.StringVar(value="0")
        ctk.CTkEntry(range_frame, textvariable=self.var_startrow, width=60).grid(row=1, column=0, padx=(0, 15))
        
        ctk.CTkLabel(range_frame, text="End").grid(row=0, column=1, sticky="w", padx=(0, 8))
        self.var_endrow = tk.StringVar(value="")
        ctk.CTkEntry(range_frame, textvariable=self.var_endrow, width=60, placeholder_text="empty").grid(row=1, column=1, padx=(0, 15))
        
        ctk.CTkLabel(range_frame, text="Concurrency").grid(row=0, column=2, sticky="w", padx=(0, 8))
        self.var_conc = tk.StringVar(value="3")
        ctk.CTkEntry(range_frame, textvariable=self.var_conc, width=70).grid(row=1, column=2)

        # Timeouts
        row_label("Timeouts", bold=True)
        timeout_frame = ctk.CTkFrame(form, fg_color="transparent")
        timeout_frame.grid(row=r, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 6))
        r += 1
        
        ctk.CTkLabel(timeout_frame, text="Conn(s)").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.var_cto = tk.StringVar(value="30")
        ctk.CTkEntry(timeout_frame, textvariable=self.var_cto, width=60).grid(row=1, column=0, padx=(0, 15))
        
        ctk.CTkLabel(timeout_frame, text="Upload(s)").grid(row=0, column=1, sticky="w", padx=(0, 8))
        self.var_upload_to = tk.StringVar(value="120")
        ctk.CTkEntry(timeout_frame, textvariable=self.var_upload_to, width=60).grid(row=1, column=1, padx=(0, 15))
        
        ctk.CTkLabel(timeout_frame, text="Reboot(s)").grid(row=0, column=2, sticky="w", padx=(0, 8))
        self.var_reboot_to = tk.StringVar(value="60")
        ctk.CTkEntry(timeout_frame, textvariable=self.var_reboot_to, width=60).grid(row=1, column=2)

        # گزینه‌ها
        row_label("Options", bold=True)
        opt_frame = ctk.CTkFrame(form, fg_color="transparent")
        opt_frame.grid(row=r, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 6))
        r += 1
        
        self.var_reboot = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(opt_frame, text="Reboot after update", variable=self.var_reboot).grid(row=0, column=0, sticky="w")
        
        self.var_verify = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(opt_frame, text="Skip if same version", variable=self.var_verify).grid(row=0, column=1, sticky="w", padx=(20, 0))

        # دکمه‌ها
        btnrow = ctk.CTkFrame(form, fg_color="transparent")
        btnrow.grid(row=r, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 12))
        r += 1
        
        self.btn_run = ctk.CTkButton(btnrow, text="Update Firmware", command=self._run, width=160, height=40,
                                      font=ctk.CTkFont(size=14, weight="bold"))
        self.btn_run.pack(side="left")
        
        self.btn_cancel = ctk.CTkButton(btnrow, text="Cancel", command=self._cancel, width=100, height=40,
                                         fg_color="#9b1c1c")
        self.btn_cancel.pack(side="left", padx=8)
        self.btn_cancel.configure(state="disabled")
        
        self.btn_reset = ctk.CTkButton(btnrow, text="Reset", command=self._reset, width=100, fg_color="#374151")
        self.btn_reset.pack(side="left", padx=8)

        # لاگ
        ctk.CTkLabel(log, text="Logs", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, sticky="w", padx=14, pady=(12, 6))
        self.lbl_prog = ctk.CTkLabel(log, text="0/0")
        self.lbl_prog.grid(row=0, column=0, sticky="e", padx=14, pady=(12, 6))
        self.pbar = ctk.CTkProgressBar(log)
        self.pbar.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 8))
        self.pbar.set(0)

        self.txt_log = ctk.CTkTextbox(log, wrap="none")
        self.txt_log.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 14))
        self.txt_log.configure(state="disabled")

    def _choose_firmware(self):
        p = filedialog.askopenfilename(title="Select firmware file", filetypes=[("NPK", "*.npk"), ("All", "*.*")])
        if p:
            self.var_firmware.set(p)

    def _choose_excel(self):
        p = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel", "*.xlsx;*.xls")])
        if p:
            self.var_excel.set(p)

    def _append_log(self, s: str):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", s + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _set_running(self, running: bool):
        self._running = running
        self.btn_run.configure(state="disabled" if running else "normal")
        self.btn_cancel.configure(state="normal" if running else "disabled")

    def _cancel(self):
        if not self._running:
            return
        self.cancel_event.set()
        self._append_log("[CANCEL] Cancel requested...")

    def _reset(self):
        if self._running:
            messagebox.showwarning("Running", "Already running. Click Cancel first.")
            return
        self.var_firmware.set("")
        self.var_excel.set("")
        self.var_sheet.set("")
        self.var_ipcol.set("RouterIP")
        self.var_username.set("admin")
        self.var_password.set("")
        self.var_port.set("22")
        self.var_startrow.set("0")
        self.var_endrow.set("")
        self.var_cto.set("30")
        self.var_upload_to.set("120")
        self.var_reboot_to.set("60")
        self.var_conc.set("3")
        self.var_reboot.set(True)
        self.var_verify.set(False)
        self.pbar.set(0)
        self.lbl_prog.configure(text="0/0")
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

    def _pump_queues(self):
        try:
            while True:
                typ, msg = self.log_q.get_nowait()
                prefix = ""
                if typ == "cmd":
                    prefix = ""
                elif typ == "info":
                    prefix = ""
                elif typ == "warn":
                    prefix = "⚠ "
                elif typ == "err":
                    prefix = "❌ "
                elif typ == "ok":
                    prefix = "✅ "
                elif typ == "run":
                    prefix = "▶ "
                self._append_log(prefix + msg)
        except queue.Empty:
            pass

        try:
            while True:
                item = self.prog_q.get_nowait()
                if not item:
                    continue
                typ = item[0]
                if typ == "tick":
                    _, a, b, _outdir = item
                    self.lbl_prog.configure(text=f"{a}/{b}")
                    if b > 0:
                        self.pbar.set(a / b)
                elif typ == "done":
                    _, a, b, outdir, cancelled = item
                    self.lbl_prog.configure(text=f"{a}/{b}")
                    if b > 0:
                        self.pbar.set(a / b)
                    self._set_running(False)
                    if cancelled:
                        messagebox.showinfo("Update", f"Cancelled.\nProgress: {a}/{b}")
                    else:
                        msg = f"Done.\nProcessed: {a}/{b}"
                        if outdir:
                            msg += f"\nLogs: {outdir}"
                        messagebox.showinfo("Update", msg)
        except queue.Empty:
            pass

        self.after(200, self._pump_queues)

    def _run(self):
        if self._running:
            return

        firmware = self.var_firmware.get().strip()
        excel = self.var_excel.get().strip()
        sheet = self.var_sheet.get().strip() or None
        ipcol = self.var_ipcol.get().strip() or "RouterIP"
        username = self.var_username.get().strip() or "admin"
        password = self.var_password.get()
        port = self.var_port.get().strip() or "22"
        port = re.sub(r'\.0$', '', port)

        if not firmware or not os.path.exists(firmware):
            messagebox.showerror("Missing", "Firmware file is invalid.")
            return
        if not excel or not os.path.exists(excel):
            messagebox.showerror("Missing", "Excel file is invalid.")
            return

        try:
            start_row = int(self.var_startrow.get() or "0")
        except Exception:
            start_row = 0
        end_raw = self.var_endrow.get().strip()
        end_row = int(end_raw) if end_raw else None

        try:
            t_conn = int(self.var_cto.get() or "30")
        except Exception:
            t_conn = 30
        try:
            t_upload = int(self.var_upload_to.get() or "120")
        except Exception:
            t_upload = 120
        try:
            t_reboot = int(self.var_reboot_to.get() or "60")
        except Exception:
            t_reboot = 60
        try:
            conc = int(self.var_conc.get() or "3")
        except Exception:
            conc = 3

        if conc < 1:
            conc = 1

        self.cancel_event.clear()
        self._set_running(True)
        self._append_log("[RUN] Starting firmware update...")

        def worker():
            try:
                self._run_update(
                    excel_path=excel, sheet_name=sheet, ip_col=ipcol,
                    firmware_path=firmware,
                    username=username, password=password, port=port,
                    start_row=start_row, end_row=end_row,
                    timeout_conn=t_conn, timeout_upload=t_upload, timeout_reboot=t_reboot,
                    concurrency=conc,
                    reboot=self.var_reboot.get(),
                    verify=self.var_verify.get()
                )
            except Exception as e:
                self.log_q.put(("err", f"[FATAL] {e}"))
                self.prog_q.put(("done", 0, 0, None, True))

        threading.Thread(target=worker, daemon=True).start()

    def _run_update(self, excel_path, sheet_name, ip_col, firmware_path,
                    username, password, port, start_row, end_row,
                    timeout_conn, timeout_upload, timeout_reboot, concurrency,
                    reboot, verify):

        self.log_q.put(("info", f"[INFO] Loading devices from Excel..."))
        
        try:
            df, ip_col_resolved, ips = read_ips_from_excel(excel_path, sheet_name, ip_col, start_row, end_row)
        except Exception as e:
            self.log_q.put(("err", f"[FATAL] Excel read failed: {e}"))
            self.prog_q.put(("done", 0, 0, None, True))
            return

        devices = []
        for _, row in df.iterrows():
            ip = str(row.get(ip_col_resolved, "")).strip()
            if not ip:
                continue
            devices.append({"ip": ip})

        total = len(devices)
        self.log_q.put(("info", f"[INFO] Devices to update: {total}"))

        if total == 0:
            self.prog_q.put(("done", 0, 0, None, False))
            return

        out_dir = os.path.join(OUT_BASE, now_ts() + "_firmware_update")
        os.makedirs(out_dir, exist_ok=True)

        done = 0
        lock = threading.Lock()

        def write_log(ip, status, msg, log_path):
            with lock:
                try:
                    with open(log_path, "a", encoding="utf-8") as f:
                        f.write(f"{datetime.now().isoformat()} | {status} | {msg}\n")
                except Exception:
                    pass

        def update_device(dev):
            ip = dev["ip"]
            log_path = os.path.join(out_dir, f"{ip.replace(':', '_')}.log")
            
            if self.cancel_event.is_set():
                write_log(ip, "CANCELLED", "Update cancelled", log_path)
                return (ip, "cancelled")

            self.log_q.put(("run", f"[{ip}] Connecting..."))
            write_log(ip, "START", "Starting firmware update", log_path)
            
            ssh = ssh_connect(ip, port, username, password, timeout=timeout_conn)
            if isinstance(ssh, str):
                self.log_q.put(("err", f"[{ip}] Connection failed: {ssh}"))
                write_log(ip, "FAILED", f"Connection failed: {ssh}", log_path)
                return (ip, "failed")

            try:
                firmware_filename = os.path.basename(firmware_path)
                
                # آپلود فایل
                self.log_q.put(("cmd", f"[{ip}] Uploading {firmware_filename}..."))
                write_log(ip, "UPLOAD", f"Uploading {firmware_filename}", log_path)
                
                with SCPClient(ssh.get_transport()) as scp:
                    scp.put(firmware_path, f"/{firmware_filename}")
                
                self.log_q.put(("ok", f"[{ip}] File uploaded successfully."))
                write_log(ip, "UPLOAD_OK", "File uploaded", log_path)
                
                # ریبوت (فایل NPK خودکار در هنگام ریبوت نصب میشه)
                if reboot:
                    self.log_q.put(("cmd", f"[{ip}] Rebooting (firmware will be installed automatically)..."))
                    write_log(ip, "REBOOT", "Sending reboot command - firmware will install", log_path)
                    
                    time.sleep(2)
                    try:
                        ssh.exec_command("/system reboot", timeout=timeout_reboot)
                    except Exception:
                        pass
                    
                    self.log_q.put(("ok", f"[{ip}] Reboot command sent. Router will update on next boot."))
                else:
                    self.log_q.put(("info", f"[{ip}] File uploaded. Manual reboot required for update."))
                
                ssh.close()
                self.log_q.put(("ok", f"[{ip}] Update process completed."))
                write_log(ip, "SUCCESS", "Update process completed", log_path)
                return (ip, "success")
                
            except Exception as e:
                self.log_q.put(("err", f"[{ip}] Error: {e}"))
                write_log(ip, "FAILED", f"Error: {e}", log_path)
                try:
                    ssh.close()
                except Exception:
                    pass
                return (ip, "failed")
        concurrency = max(1, concurrency)
        self.log_q.put(("info", f"[INFO] Running with concurrency={concurrency}"))

        with ThreadPoolExecutor(max_workers=concurrency) as ex:
            futures = [ex.submit(update_device, dev) for dev in devices if not self.cancel_event.is_set()]
            for fut in as_completed(futures):
                if self.cancel_event.is_set():
                    break
                try:
                    ip, status = fut.result()
                except Exception as e:
                    self.log_q.put(("err", f"Exception: {e}"))
                done += 1
                self.prog_q.put(("tick", done, total, out_dir))

        self.prog_q.put(("done", done, total, out_dir, self.cancel_event.is_set()))
class HelpDialog(ctk.CTkToplevel):
    """پنجره راهنما برای نمایش توضیحات تب‌ها"""
    
    def __init__(self, parent, tab_name):
        super().__init__(parent)
        self.title(f"Help - {tab_name}")
        self.geometry("650x500")
        self.minsize(550, 400)
        self.transient(parent)
        self.grab_set()
        
        # مرکز کردن پنجره
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - 325
        y = parent.winfo_y() + (parent.winfo_height() // 2) - 250
        self.geometry(f"+{x}+{y}")
        
        # محتوای راهنما
        help_texts = {
            "SSH": """=== SSH Tab - Run Commands on Multiple Routers ===

OVERVIEW:
This tab allows you to execute RouterOS commands on multiple MikroTik routers simultaneously via SSH.

FEATURES:
• Read router IPs from Excel file
• Load command list from .txt or .rsc file
• Run commands concurrently on multiple devices
• Save detailed logs for each router

INPUT FIELDS:

1. Routers Excel file:
   - Select an Excel file containing router IP addresses
   - Required columns: IP address column (default: RouterIP)

2. Excel options:
   - Sheet: Name of the sheet in Excel file (e.g., Sheet1)
   - IP Column: Name of the column containing router IPs

3. Commands file (.txt / .rsc):
   - Text file with RouterOS commands (one per line)
   - Supports context-aware command parsing

4. Credential Profile:
   - Select saved credentials from WinBox tab
   - Contains username and password for SSH access

5. Row Range:
   - Start: First row to read from Excel (0-indexed)
   - End: Last row to read (leave empty for all rows)

6. Timeouts:
   - Conn(s): SSH connection timeout in seconds
   - Cmd(s): Command execution timeout per command
   - Delay: Delay between commands (seconds)

7. Concurrency:
   - Number of routers to process simultaneously

OUTPUT:
• Log files saved in 'outputs/YYYYMMDD-HHMMSS_stage1/'
• Summary.csv with execution status for each router
• Individual .log file per router with command outputs

BUTTONS:
• Run: Start the SSH command execution
• Cancel: Stop the ongoing operation
• Reset: Clear all input fields

NOTE: Empty password is allowed for routers without authentication.""",

            "WinBox": """=== WinBox Tab - Router List Viewer & Launcher ===

OVERVIEW:
This tab displays router information from Excel and launches WinBox connections with saved credentials.

FEATURES:
• View Excel data in a sortable/filterable table
• Multiple column filters
• Save and manage multiple credential profiles
• Launch WinBox directly from the table

INPUT FIELDS & CONTROLS:

1. Sheet selector:
   - Choose which sheet from Excel file to display

2. Credential Profile:
   - Select saved username/password for WinBox
   - Add/Edit/Delete profiles as needed
   - Profiles saved to 'winbox_creds.json'

3. Search:
   - Search across all columns in the table
   - Real-time filtering

4. Table columns:
   - Click on column headers to sort
   - Right-click on any cell for copy options

FILTERS:
• Add Filter: Create column-specific filters (contains/equals/startswith)
• Filter chips appear above the table
• Clear button removes all filters

CONTEXT MENU (Right-click on table):
• Copy Row: Copy entire row as tab-separated
• Copy Cell: Copy selected cell content
• Open Cell Viewer: View full cell text in popup window
• Copy IP: Copy IP address from selected row
• Copy Selected Text: Copy highlighted text from any field

EXPORT:
• Export filtered/sorted data to Excel or CSV
• Files saved in 'outputs/exports/'

SETTINGS:
• Default WinBox port
• WinBox executable path
• Auto-reload interval (seconds)

BUTTONS:
• Refresh: Reload Excel file
• Open Excel: Select a different Excel file
• Export: Save filtered data to file
• Settings: Configure WinBox path and port
• Connect: Launch WinBox with selected credential
• Connect (no creds): Launch WinBox without credentials

NOTE: Credential profiles are shared with SSH and Backup tabs.""",

            "BackUp": """=== Backup Tab - Router Backup Automation ===

OVERVIEW:
This tab creates system backups (.backup) and configuration exports (.rsc) from multiple routers via SSH.

FEATURES:
• Backup multiple routers from Excel list
• Both system backup and config export options
• Scheduled daily backups
• SCP file transfer from routers

INPUT FIELDS:

1. Routers Excel file:
   - Select Excel file with router IPs

2. Excel options:
   - Sheet: Sheet name in Excel file
   - IP Column: Column containing router IPs
   - Start/End rows: Row range to process

3. Credential Profile:
   - Shared with SSH and WinBox tabs
   - Contains SSH username/password

4. Output folder:
   - Where backup files will be saved
   - Default: './backups/'

5. Backup Type:
   - System Backup (.backup): Full system backup file
   - Config Export (.rsc): Plain-text configuration export

SCHEDULE:
• Enable daily schedule for automatic backups
• Set hour (0-23) and minute (0-59)
• Uses APScheduler for background execution

PREVIEW:
• Shows list of IPs that will be backed up
• Based on current Excel settings

OUTPUT:
• Backup files named: IP_YYYY-MM-DD.backup
• Config files named: IP_YYYY-MM-DD.rsc

BUTTONS:
• Run backup now: Execute backup immediately
• Save schedule: Save daily schedule settings

NOTE: Empty password is allowed for routers without authentication.""",

            "IOS": """=== IOS Tab - Firmware Update ===

OVERVIEW:
This tab uploads RouterOS firmware files (.npk) to multiple routers and performs upgrade via reboot.

FEATURES:
• Upload .npk firmware files via SCP
• Upgrade multiple routers from Excel list
• Automatic reboot after upload
• Concurrent uploads support

IMPORTANT NOTES:
• Only UPGRADE is supported (e.g., v6 → v7)
• For DOWNGRADE (e.g., v7 → v6), use Netinstall tool
• Firmware file must match router architecture (mipsbe, arm, x86, etc.)

INPUT FIELDS:

1. Firmware File (.npk):
   - Select RouterOS package file
   - File name format: routeros-version-architecture.npk

2. Routers Excel file:
   - Excel file containing router IPs

3. Excel options:
   - Sheet: Sheet name in Excel file
   - IP Column: Column containing router IPs

4. SSH Credentials:
   - Username: SSH login (default: admin)
   - Password: SSH password (optional)
   - Port: SSH port (default: 22)

5. Row Range:
   - Start: First row to read (0-indexed)
   - End: Last row (leave empty for all)

6. Concurrency:
   - Number of routers to update simultaneously (default: 3)

7. Timeouts:
   - Conn(s): SSH connection timeout
   - Upload(s): File upload timeout
   - Reboot(s): Time to wait after reboot command

OPTIONS:
• Auto Reboot: Reboot router after upload (recommended)

UPDATE MODE:
• Upgrade (Normal): Upload file → Reboot → Router auto-installs on boot
• Downgrade: DISABLED - Use Netinstall for downgrade

WORKFLOW:
1. Upload .npk file to router's root directory
2. Send reboot command (if enabled)
3. Router installs firmware during boot

OUTPUT:
• Logs saved in 'outputs/YYYYMMDD_HHMMSS_firmware_upgrade/'
• Individual .log file per router

BUTTONS:
• Update Firmware: Start the upgrade process
• Cancel: Stop ongoing operation
• Reset: Clear all input fields

WARNING: Router will reboot during upgrade. Ensure no critical traffic is passing through.""",

            "General": """=== GENERAL TIPS ===

LICENSE:
• No license required - Unlimited devices
• No trial limitations

FILE LOCATIONS:
• Outputs: ./outputs/
• Backups: ./backups/
• Config: %APPDATA%/MikroTikOps/
• WinBox Credentials: winbox_creds.json
• UI Settings: ui.json

COMMON ISSUES:

1. Connection failed:
   - Check IP, username, password
   - Ensure SSH is enabled on router
   - Check firewall rules

2. Excel file not loading:
   - Ensure file is not open in another program
   - Check column names match settings

3. WinBox not launching:
   - Verify winbox.exe path in Settings
   - Default paths: Program Files, Downloads

4. Import errors:
   - Run: pip install customtkinter pandas paramiko scp apscheduler openpyxl keyring

REQUIRED DEPENDENCIES:
• customtkinter - Modern UI
• pandas - Excel handling
• paramiko - SSH connections
• scp - File transfers
• apscheduler - Scheduling
• openpyxl - Excel support
• keyring - Credential storage (optional)

SUPPORT:
For issues and feature requests, check documentation or contact developer."""
        }
        
        # عنوان
        header = ctk.CTkFrame(self, corner_radius=0, height=50)
        header.pack(fill="x", padx=0, pady=0)
        ctk.CTkLabel(header, text=f"Help: {tab_name} Tab", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(10, 5))
        
        # متن راهنما
        text_frame = ctk.CTkFrame(self)
        text_frame.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        text_widget = ctk.CTkTextbox(text_frame, wrap="word", font=ctk.CTkFont(size=12))
        text_widget.pack(fill="both", expand=True)
        
        content = help_texts.get(tab_name, help_texts["General"])
        text_widget.insert("1.0", content)
        text_widget.configure(state="disabled")
        
        # دکمه بستن
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=(0, 15))
        ctk.CTkButton(btn_frame, text="Close", width=100, command=self.destroy).pack()        
class MikroTikOpsApp(ctk.CTk):
    def __init__(self):
        ui = load_ui_settings()
        appearance = ui.get("appearance_mode", "Dark")
        theme = ui.get("color_theme", "blue")
        style = ui.get("style", DEFAULT_STYLE)
        if style not in STYLE_PRESETS:
            style = DEFAULT_STYLE
        try:
            ctk.set_appearance_mode(appearance)
        except Exception:
            ctk.set_appearance_mode("Dark")
        try:
            ctk.set_default_color_theme(theme)
        except Exception:
            ctk.set_default_color_theme("blue")

        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1200x760")
        self.minsize(1080, 680)

        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)
        self.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(self, corner_radius=14)
        header.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 8))
        header.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(header, text="MikroTik Ops", font=ctk.CTkFont(size=20, weight="bold"))
        title.grid(row=0, column=0, sticky="w", padx=14, pady=(12, 0))

        subtitle = ctk.CTkLabel(header, text="SSH automation • WinBox launcher • Router backups • No License Limits",
                                font=ctk.CTkFont(size=12), text_color=("gray30", "gray70"))
        subtitle.grid(row=1, column=0, sticky="w", padx=14, pady=(2, 12))

        controls = ctk.CTkFrame(header, fg_color="transparent")
        controls.grid(row=0, column=1, rowspan=2, sticky="e", padx=14, pady=12)

        self._appearance_var = tk.StringVar(value=appearance)
        self._theme_var = tk.StringVar(value=theme)
        self._style_var = tk.StringVar(value=style)

        ctk.CTkLabel(controls, text="Appearance").grid(row=0, column=0, sticky="e", padx=(0, 8))
        self.appearance_menu = ctk.CTkOptionMenu(controls, values=["Light", "Dark", "System"],
                                                  variable=self._appearance_var, width=130,
                                                  command=self._on_change_appearance)
        self.appearance_menu.grid(row=0, column=1, sticky="e", padx=(0, 14))

        ctk.CTkLabel(controls, text="Color Style").grid(row=0, column=2, sticky="e", padx=(0, 8))
        self.style_menu = ctk.CTkOptionMenu(controls, values=list(STYLE_PRESETS.keys()),
                                             variable=self._style_var, width=150,
                                             command=self._on_change_style)
        self.style_menu.grid(row=0, column=3, sticky="e", padx=(0, 14))

        # NO LICENSE - removed license badge and button

        self.tabs = ctk.CTkTabview(self, corner_radius=14)
        self.tabs.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 8))

        status = ctk.CTkFrame(self, corner_radius=14)
        status.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 12))
        status.grid_columnconfigure(0, weight=1)
        self.status_var_app = tk.StringVar(value="Ready - Unlimited Devices")
        ctk.CTkLabel(status, textvariable=self.status_var_app).grid(row=0, column=0, sticky="w", padx=14, pady=10)

        self.tabs.add("SSH")
        self.tabs.add("WinBox")
        self.tabs.add("BackUp")
        self.tabs.add("IOS")
        
        self.ssh_frame = SSHFrame(self.tabs.tab("SSH"))
        self.ssh_frame.pack(fill="both", expand=True)

        self.winbox_frame = WinBoxFrame(self.tabs.tab("WinBox"))
        self.winbox_frame.pack(fill="both", expand=True)

        self.backup_frame = BackupFrame(self.tabs.tab("BackUp"))
        self.backup_frame.pack(fill="both", expand=True)
        
        self.ios_frame = IOSFrame(self.tabs.tab("IOS"))
        self.ios_frame.pack(fill="both", expand=True) 
        # ========== اضافه کردن دکمه‌های راهنما (?) به هر تب ==========
        def add_help_button(tab_frame, tab_name):
            help_btn = ctk.CTkButton(
                tab_frame,
                text="?",
                width=8,
                height=8,
                corner_radius=15,
                fg_color="#3b82f6",
                hover_color="#2563eb",
                font=ctk.CTkFont(size=11, weight="bold"),
                command=lambda: HelpDialog(self, tab_name)
            )
            help_btn.place(relx=0.01, rely=0.02, anchor="nw")
        
        add_help_button(self.tabs.tab("SSH"), "SSH")
        add_help_button(self.tabs.tab("WinBox"), "WinBox")
        add_help_button(self.tabs.tab("BackUp"), "BackUp")
        add_help_button(self.tabs.tab("IOS"), "IOS")
        # ========== پایان اضافه کردن دکمه‌های راهنما ==========                 
        self.apply_style()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_change_appearance(self, value: str):
        try:
            ctk.set_appearance_mode(value)
        except Exception:
            return
        ui = load_ui_settings()
        ui["appearance_mode"] = value
        save_ui_settings(ui)
        self.apply_style()

    def _on_change_style(self, value: str):
        ui = load_ui_settings()
        ui["style"] = value
        save_ui_settings(ui)
        self.apply_style()

    def apply_style(self):
        style_name = self._style_var.get() if hasattr(self, "_style_var") else DEFAULT_STYLE
        pal = STYLE_PRESETS.get(style_name, STYLE_PRESETS[DEFAULT_STYLE])

        try:
            self.configure(fg_color=pal["card"])
        except Exception:
            pass

        try:
            seg = getattr(self.tabs, "_segmented_button", None)
            if seg is not None:
                seg.configure(selected_color=pal["accent"], selected_hover_color=pal["accent_hover"],
                              fg_color=pal["card2"], unselected_color=pal["card2"],
                              unselected_hover_color=pal["table_head"], text_color=pal["text"],
                              selected_text_color=pal["text"])
        except Exception:
            pass

        for w in iter_children_recursive(self):
            try:
                if isinstance(w, ctk.CTkButton):
                    txt = str(w.cget("text") or "").strip().lower()
                    if any(k in txt for k in ("stop", "cancel", "delete", "remove")):
                        w.configure(fg_color=pal["bad"], hover_color=pal["bad"])
                    else:
                        w.configure(fg_color=pal["accent"], hover_color=pal["accent_hover"])
                elif isinstance(w, ctk.CTkOptionMenu):
                    w.configure(fg_color=pal["card2"], text_color=pal["text"],
                                dropdown_fg_color=pal["card2"], dropdown_hover_color=pal["table_head"],
                                dropdown_text_color=pal["text"], button_color=pal["accent"],
                                button_hover_color=pal["accent_hover"])
                elif isinstance(w, ctk.CTkProgressBar):
                    w.configure(progress_color=pal["accent"])
                elif isinstance(w, ctk.CTkCheckBox):
                    w.configure(fg_color=pal["accent"], hover_color=pal["accent_hover"])
            except Exception:
                pass

    def _on_close(self):
        try:
            self.winbox_frame.stop()
        except Exception:
            pass
        try:
            self.backup_frame.scheduler.shutdown(wait=False)
        except Exception:
            pass
        self.destroy()


if __name__ == "__main__":
    app = MikroTikOpsApp()
    app.mainloop()                        