# -*- coding: utf-8 -*-
"""依出貨清單調整採購單數量與金額。

本工具處理內容為 HTML 的 .xls，不依賴 Excel 或 VBA。
只有採購單號、項次與料號完全相同的項目才會更新。

用法:
  python 採購單大量修改數量與總價.py
      → 開啟操作視窗
  python 採購單大量修改數量與總價.py <採購單.xls> <出貨清單.xlsx|csv>
      → 直接指定路徑

規則:
  - 採購單號 + 項次 + 料號必須完全符合才修改
  - 出貨清單未列出的 PO / 項次，保留原始內容，不改 0
  - 採購單號、項次或料號任一不符，直接略過
  - 重算已修改項目的未稅金額，採購金額合計僅套用修改差額
  - 不覆寫原檔，另存 *_出貨調整_時間戳.xls
"""
from __future__ import annotations

import csv
import os
import re
import sys
import time
from datetime import datetime
from html import unescape
from pathlib import Path

try:
    import openpyxl
except ImportError:  # pragma: no cover
    openpyxl = None


# ---------------------------------------------------------------------------
# paths / normalize
# ---------------------------------------------------------------------------

def script_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def normalize_po(s: str) -> str:
    t = str(s or "").strip().replace("\u3000", " ")
    t = re.sub(r"\s+", " ", t)
    if "採購單號" in t:
        t = t.split("採購單號", 1)[-1].lstrip(" :：").strip()
    return t


def clean_number(v):
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip().replace(",", "").replace(" ", "").replace("\u3000", "")
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def fmt_qty(n: float) -> str:
    if abs(n - round(n)) < 1e-9:
        return f"{int(round(n)):,}"
    return f"{n:,.4f}".rstrip("0").rstrip(".")


def fmt_amt(n: float) -> str:
    return f"{n:,.2f}"


def cell_text(td_html: str) -> str:
    t = re.sub(r"<[^>]+>", " ", td_html)
    t = unescape(t)
    return re.sub(r"\s+", " ", t).strip()


def replace_first_span_text(td_html: str, new_text: str) -> str:
    """Replace the first visible text node inside a <td> (span or plain)."""
    m = re.search(r"(<span[^>]*>)(.*?)(</span>)", td_html, flags=re.I | re.S)
    if m:
        return td_html[: m.start()] + m.group(1) + new_text + m.group(3) + td_html[m.end() :]
    # fallback: replace text between tags loosely
    m2 = re.search(r"(>)([^<>]*?)(<)", td_html)
    if m2:
        return td_html[: m2.start()] + m2.group(1) + new_text + m2.group(3) + td_html[m2.end() :]
    return td_html


# ---------------------------------------------------------------------------
# ship list
# ---------------------------------------------------------------------------

def _header_map(headers: list[str]) -> dict[str, int]:
    def norm(h: str) -> str:
        return re.sub(r"\s+", "", str(h or "").upper().replace("\u3000", ""))

    idx = {}
    for i, h in enumerate(headers):
        nh = norm(h)
        if not nh:
            continue
        if "c_po" not in idx and ("採購單號" in str(h) or nh in {"PO", "PONO", "PONUMBER"}):
            idx["c_po"] = i
        if "c_line" not in idx and ("項次" in str(h) or nh in {"LINE", "LINENO"} or "行號" in str(h)):
            idx["c_line"] = i
        if "c_part" not in idx and (
            "料號" in str(h) or "料件" in str(h) or nh in {"PN", "PART", "PARTNO"}
        ):
            idx["c_part"] = i
        if "c_qty" not in idx and (
            "出貨數量" in str(h) or nh in {"SHIPQTY", "QTY"} or ("數量" in str(h) and "項次" not in str(h))
        ):
            idx["c_qty"] = i
    idx.setdefault("c_po", 0)
    idx.setdefault("c_line", 1)
    idx.setdefault("c_part", 2)
    idx.setdefault("c_qty", 3)
    return idx


def load_ship(ship_path: Path, emit=None):
    ship: dict[str, dict] = {}
    affected: set[str] = set()
    logs: list[tuple[str, str]] = []

    def emit_log(entry):
        logs.extend((entry,))
        if emit is not None:
            emit(entry[0], entry[1])

    ext = ship_path.suffix.lower()
    rows: list[list] = []

    if ext == ".csv":
        raw = ship_path.read_bytes()
        text = None
        for enc in ("utf-8-sig", "utf-8", "cp950", "big5"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            text = raw.decode("utf-8", errors="replace")
        reader = csv.reader(text.splitlines())
        rows = [list(r) for r in reader if any(str(c).strip() for c in r)]
    else:
        if openpyxl is None:
            raise RuntimeError("讀取 xlsx 需要 openpyxl，請先 pip install openpyxl")
        wb = openpyxl.load_workbook(ship_path, data_only=True, read_only=True)
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            if row is None or all(v is None or str(v).strip() == "" for v in row):
                continue
            rows.append(list(row))
        wb.close()

    if not rows:
        return ship, affected, logs

    cols = _header_map([str(h or "") for h in rows[0]])
    for rno, row in enumerate(rows[1:], start=2):
        def get(i):
            return row[i] if i < len(row) else None

        po = normalize_po(str(get(cols["c_po"]) or ""))
        if not po:
            continue
        line = clean_number(get(cols["c_line"]))
        qty = clean_number(get(cols["c_qty"]))
        part = str(get(cols["c_part"]) or "").strip().upper()
        if (
            line is None
            or line <= 0
            or abs(line - round(line)) > 1e-9
            or qty is None
        ):
            emit_log(("WARN", f"出貨列 {rno} 項次/數量無法解析，略過"))
            continue
        if qty < 0:
            emit_log(("WARN", f"出貨列 {rno} 數量為負，略過"))
            continue
        if not part:
            emit_log(("WARN", f"出貨列 {rno} 料號空白，無法完整比對，略過"))
            continue
        key = f"{po}|{int(line)}"
        if key in ship:
            # 同一 PO + 項次出現多筆時無法判斷哪一筆才正確，整組停用。
            ship[key]["valid"] = False
            emit_log(("WARN", f"出貨鍵重複，該鍵全部略過: {key}"))
            affected.add(po)
            continue
        ship[key] = {
            "qty": float(qty),
            "part": part,
            "matched": False,
            "seen": False,
            "valid": True,
        }
        affected.add(po)

    emit_log(("INFO", f"出貨載入 筆數={len(ship)} 影響PO數={len(affected)}"))
    return ship, affected, logs


# ---------------------------------------------------------------------------
# PO HTML process
# ---------------------------------------------------------------------------

TR_RE = re.compile(r"(<tr\b[^>]*>)(.*?)(</tr>)", re.I | re.S)
TD_RE = re.compile(r"<t[dh]\b[^>]*>.*?</t[dh]>", re.I | re.S)


def is_line_item_cells(texts: list[str]) -> bool:
    if len(texts) < 6:
        return False
    line = clean_number(texts[0])
    if line is None or line <= 0 or abs(line - int(line)) > 1e-9:
        return False
    part = texts[1].strip()
    if not part or "料件" in part or "品名" in part or part == "項次":
        return False
    # unit often PCS at index 2
    qty = clean_number(texts[3]) if len(texts) > 3 else None
    price = clean_number(texts[4]) if len(texts) > 4 else None
    # tolerate missing unit text: try scan
    if qty is None or price is None:
        # find PCS then next two numbers
        for i, t in enumerate(texts):
            if t.upper() in {"PCS", "EA", "SET", "PC"}:
                if i + 2 < len(texts):
                    qty = clean_number(texts[i + 1])
                    price = clean_number(texts[i + 2])
                break
    if qty is None or price is None:
        return False
    return True


def parse_line_fields(texts: list[str]):
    line_no = int(clean_number(texts[0]))
    part = texts[1].strip().upper()
    unit_i = None
    for i, t in enumerate(texts):
        if t.upper() in {"PCS", "EA", "SET", "PC"}:
            unit_i = i
            break
    if unit_i is not None and unit_i + 3 < len(texts):
        qty = clean_number(texts[unit_i + 1])
        price = clean_number(texts[unit_i + 2])
        amt = clean_number(texts[unit_i + 3])
        qty_text_i = unit_i + 1
        price_text_i = unit_i + 2
        amt_text_i = unit_i + 3
    else:
        qty = clean_number(texts[3])
        price = clean_number(texts[4])
        amt = clean_number(texts[5]) if len(texts) > 5 else None
        qty_text_i, price_text_i, amt_text_i = 3, 4, 5
    return line_no, part, qty, price, amt, qty_text_i, price_text_i, amt_text_i


def process_po_html(html: str, ship: dict, affected: set[str], emit=None):
    logs: list[tuple[str, str]] = []

    def emit_log(entry):
        logs.extend((entry,))
        if emit is not None:
            emit(entry[0], entry[1])

    stats = {
        "updated": 0,
        "unchanged": 0,
        "skipped_not_listed": 0,
        "skipped_part_mismatch": 0,
        "skipped_invalid": 0,
        "totals": 0,
    }

    current_po = ""
    in_aff = False
    po_delta = 0.0
    po_updated = 0
    out_parts: list[str] = []
    last = 0

    for m in TR_RE.finditer(html):
        out_parts.append(html[last : m.start()])
        open_tr, inner, close_tr = m.group(1), m.group(2), m.group(3)
        tds = list(TD_RE.finditer(inner))
        texts = [cell_text(td.group(0)) for td in tds]
        nonempty = [t for t in texts if t]

        # PO header
        if nonempty and "採購單號" in nonempty[0]:
            # value is next nonempty cell usually
            po_val = ""
            for t in nonempty[1:]:
                if t and "送貨" not in t and "地址" not in t:
                    po_val = t
                    break
            current_po = normalize_po(po_val)
            in_aff = current_po in affected
            po_delta = 0.0
            po_updated = 0
            out_parts.append(m.group(0))
            last = m.end()
            continue

        # total row
        if any("採購金額合計" in t for t in nonempty):
            if in_aff and po_updated > 0:
                # replace amount in last numeric-looking td among later cells
                new_inner = inner
                # find td that currently holds amount (often last with number)
                amt_td_idx = None
                for i, t in enumerate(texts):
                    if clean_number(t) is not None and "採購" not in t:
                        amt_td_idx = i
                if amt_td_idx is None and tds:
                    amt_td_idx = len(tds) - 1
                if amt_td_idx is not None:
                    old_total = clean_number(texts[amt_td_idx])
                    if old_total is None:
                        emit_log(
                            (
                                "WARN",
                                f"無法解析採購金額合計，保留原值: {current_po}",
                            )
                        )
                    else:
                        new_total = round(old_total + po_delta, 2)
                        old_td = tds[amt_td_idx].group(0)
                        new_td = replace_first_span_text(old_td, fmt_amt(new_total))
                        # splice in original inner by replacing exact old_td once
                        new_inner = inner.replace(old_td, new_td, 1)
                        stats["totals"] += 1
                        emit_log(
                            (
                                "INFO",
                                f"更新合計 {current_po} "
                                f"{fmt_amt(old_total)}->{fmt_amt(new_total)}",
                            )
                        )
                out_parts.append(open_tr + new_inner + close_tr)
            else:
                out_parts.append(m.group(0))
            last = m.end()
            continue

        # line item
        if in_aff and current_po and is_line_item_cells(nonempty):
            (
                line_no,
                part,
                old_qty,
                price,
                old_amt,
                qty_i,
                price_i,
                amt_i,
            ) = parse_line_fields(nonempty)
            if old_amt is None and old_qty is not None and price is not None:
                old_amt = round(old_qty * price, 2)

            key = f"{current_po}|{line_no}"
            ship_row = ship.get(key)

            # 出貨清單沒有這個 PO + 項次：整列保持原樣，絕不改成 0。
            if ship_row is None:
                stats["skipped_not_listed"] += 1
                emit_log(
                    (
                        "SKIP",
                        f"出貨清單未列，保留原值 {current_po} 項次{line_no} {part}",
                    )
                )
                out_parts.append(m.group(0))
                last = m.end()
                continue

            ship_row["seen"] = True

            # 重複鍵等無法唯一判定的資料不允許修改。
            if not ship_row.get("valid", True):
                stats["skipped_invalid"] += 1
                emit_log(("SKIP", f"出貨資料不唯一，保留原值: {key}"))
                out_parts.append(m.group(0))
                last = m.end()
                continue

            # 料號也必須完全一致；不一致時連數量、金額與合計都不動。
            if ship_row["part"] != part:
                stats["skipped_part_mismatch"] += 1
                emit_log(
                    (
                        "SKIP",
                        f"料號不符，保留原值 {key}: "
                        f"PO={part} 出貨={ship_row['part']}",
                    )
                )
                out_parts.append(m.group(0))
                last = m.end()
                continue

            new_qty = ship_row["qty"]
            ship_row["matched"] = True

            new_amt = round(new_qty * price, 2)

            # 數量與金額原本就一致時，不重寫 HTML，僅記錄已完全命中。
            if (
                abs(new_qty - old_qty) < 1e-9
                and old_amt is not None
                and abs(new_amt - old_amt) < 0.005
            ):
                stats["unchanged"] += 1
                emit_log(
                    (
                        "OK",
                        f"完全符合且無須修改 {current_po} 項次{line_no} {part}",
                    )
                )
                out_parts.append(m.group(0))
                last = m.end()
                continue

            stats["updated"] += 1
            po_updated += 1
            po_delta += new_amt - old_amt

            # map nonempty indices back to original td indices
            # texts includes empty cells; nonempty was filtered — re-parse from full texts
            full_texts = texts
            # recompute field indices on full_texts (with empties)
            qty_td = price_td = amt_td = None
            # find by scanning full tds texts
            ft = [cell_text(td.group(0)) for td in tds]
            # line is first numeric-ish
            unit_i = None
            for i, t in enumerate(ft):
                if t.upper() in {"PCS", "EA", "SET", "PC"}:
                    unit_i = i
                    break
            if unit_i is not None:
                qty_td, price_td, amt_td = unit_i + 1, unit_i + 2, unit_i + 3
            else:
                qty_td, price_td, amt_td = 3, 4, 5

            new_inner = inner
            # replace from end to start so offsets stay valid if using positions —
            # we use exact td string replace once each
            replacements = []
            if 0 <= qty_td < len(tds):
                replacements.append((tds[qty_td].group(0), fmt_qty(new_qty)))
            if 0 <= amt_td < len(tds):
                replacements.append((tds[amt_td].group(0), fmt_amt(new_amt)))
            for old_td, new_val in replacements:
                new_td = replace_first_span_text(old_td, new_val)
                new_inner = new_inner.replace(old_td, new_td, 1)

            out_parts.append(open_tr + new_inner + close_tr)
            emit_log(
                (
                    "OK",
                    f"更新數量 {current_po} 項次{line_no} {part} "
                    f"{fmt_qty(old_qty)}->{fmt_qty(new_qty)} "
                    f"金額 {fmt_amt(old_amt)}->{fmt_amt(new_amt)}",
                )
            )
            last = m.end()
            continue

        out_parts.append(m.group(0))
        last = m.end()

    out_parts.append(html[last:])
    new_html = "".join(out_parts)

    miss = [k for k, v in ship.items() if not v.get("seen", False)]
    for k in miss:
        emit_log(("WARN", f"找不到完全相同的 PO + 項次，已略過: {k}"))
    stats["miss"] = len(miss)
    return new_html, stats, logs


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

APP_NAME = "採購單數量與金額調整工具"


def run_conversion(po_path: Path, ship_path: Path, on_log=None):
    """Process two selected files and return paths, statistics, and logs.

    on_log, when provided, is called as on_log(level, message) for each log
    entry as it happens, so a GUI can stream progress live. It runs in the
    calling thread; keep it thread-safe and non-blocking.
    """
    po_path = Path(po_path)
    ship_path = Path(ship_path)
    t0 = time.perf_counter()

    if not po_path.exists():
        raise FileNotFoundError(f"找不到採購單檔案：{po_path}")
    if not ship_path.exists():
        raise FileNotFoundError(f"找不到出貨清單：{ship_path}")

    ship, affected, logs1 = load_ship(ship_path, emit=on_log)
    if not ship:
        raise ValueError("出貨清單沒有有效資料，請檢查欄位名稱與內容。")

    raw = po_path.read_bytes()
    html = raw.decode("utf-8")
    new_html, stats, logs2 = process_po_html(html, ship, affected, emit=on_log)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = po_path.with_name(f"{po_path.stem}_出貨調整_{stamp}.xls")
    out.write_bytes(new_html.encode("utf-8"))

    log_path = out.with_suffix(".log.txt")
    all_logs = logs1 + logs2
    elapsed = time.perf_counter() - t0
    with log_path.open("w", encoding="utf-8") as f:
        f.write(f"result={out}\nstats={stats}\nelapsed_sec={elapsed:.3f}\n")
        for lv, msg in all_logs:
            f.write(f"[{lv}] {msg}\n")

    return {
        "output": out,
        "log": log_path,
        "stats": stats,
        "logs": all_logs,
        "elapsed": elapsed,
        "html_warning": not html.lstrip().lower().startswith("<html"),
    }


def launch_gui():
    try:
        import queue
        import threading
        import tkinter as tk
        from tkinter import filedialog, font as tkfont, messagebox, ttk
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("此環境無法開啟操作視窗。") from exc

    # Light app shell (timeline) paired with a dark console (live log).
    ACCENT = "#0E6E78"
    ACCENT_WEAK = "#CDE6E8"
    LINE = "#C7CFD8"
    SURFACE = "#FFFFFF"
    OK_C = "#1F7A4D"
    UN_C = "#586472"
    SK_C = "#9A6712"
    CON_BG = "#0B1015"
    CON_FG = "#95A1AD"

    root = tk.Tk()
    root.title(APP_NAME)
    root.geometry("960x620")
    root.minsize(880, 580)

    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    frame_bg = style.lookup("TFrame", "background") or "#F0F0F0"
    root.configure(bg=frame_bg)
    style.configure("Title.TLabel", font=("Segoe UI", 17, "bold"))
    style.configure("Subtitle.TLabel", font=("Segoe UI", 9), foreground=UN_C)
    style.configure("Section.TLabel", font=("Segoe UI", 11, "bold"))
    style.configure("Hint.TLabel", font=("Segoe UI", 9), foreground="#7A8593")
    style.configure("Done.TLabel", font=("Segoe UI", 9), foreground=UN_C)
    style.configure("Con.TLabel", font=("Segoe UI", 10, "bold"))
    style.configure("Ok.TLabel", font=("Segoe UI", 13, "bold"), foreground=OK_C)
    style.configure("Un.TLabel", font=("Segoe UI", 13, "bold"), foreground=UN_C)
    style.configure("Sk.TLabel", font=("Segoe UI", 13, "bold"), foreground=SK_C)
    style.configure("TallyCap.TLabel", font=("Segoe UI", 8), foreground="#7A8593")
    style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=(16, 7))
    style.configure("Action.TButton", padding=(9, 5))

    mono = tkfont.Font(family="Consolas", size=9)

    po_var = tk.StringVar()
    ship_var = tk.StringVar()
    stage_status_var = tk.StringVar(value="選好兩個檔案後即可開始")
    done_status_var = tk.StringVar(value="尚未處理")
    result_path_var = tk.StringVar(value="尚無結果檔")
    updated_var = tk.StringVar(value="0")
    unchanged_var = tk.StringVar(value="0")
    skipped_var = tk.StringVar(value="0")
    q = queue.Queue()
    ui = {"output": None, "log": None, "processing": False}

    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)

    header = ttk.Frame(root, padding=(26, 20, 26, 12))
    header.grid(row=0, column=0, sticky="ew")
    ttk.Label(header, text=APP_NAME, style="Title.TLabel").pack(anchor="w")
    ttk.Label(
        header,
        text="依出貨清單更新指定項目的數量與金額 · 原始採購單不會被覆寫",
        style="Subtitle.TLabel",
    ).pack(anchor="w", pady=(4, 0))

    body = ttk.Frame(root, padding=(26, 0, 26, 20))
    body.grid(row=1, column=0, sticky="nsew")
    body.columnconfigure(0, weight=0, minsize=380)
    body.columnconfigure(1, weight=1)
    body.rowconfigure(0, weight=1)

    # ---------------- left: timeline spine ----------------
    timeline = ttk.Frame(body)
    timeline.grid(row=0, column=0, sticky="nsew", padx=(0, 24))
    timeline.columnconfigure(1, weight=1)

    nodes = {}
    NODE_W = 38

    def make_stage(name, row, is_first=False, is_last=False):
        canvas = tk.Canvas(
            timeline, width=NODE_W, highlightthickness=0, bd=0, bg=frame_bg
        )
        canvas.grid(row=row, column=0, sticky="ns")
        content = ttk.Frame(timeline, padding=(8, 12, 0, 14))
        content.grid(row=row, column=1, sticky="nwe")
        content.columnconfigure(0, weight=1)
        node = {"canvas": canvas, "state": "pending"}

        def redraw(_event=None):
            canvas.delete("all")
            w = int(canvas.winfo_width()) or NODE_W
            h = int(canvas.winfo_height()) or 1
            cx, cy, r = w // 2, 22, 8
            top = cy if is_first else 0
            bot = cy if is_last else h
            if bot > top:
                canvas.create_line(cx, top, cx, bot, fill=LINE, width=2)
            st = node["state"]
            if st == "done":
                canvas.create_oval(
                    cx - r, cy - r, cx + r, cy + r, fill=ACCENT, outline=ACCENT
                )
                canvas.create_line(
                    cx - 3, cy, cx - 1, cy + 3, cx + 4, cy - 4, fill="white", width=2
                )
            elif st == "active":
                canvas.create_oval(
                    cx - r - 3, cy - r - 3, cx + r + 3, cy + r + 3,
                    outline=ACCENT_WEAK, width=3,
                )
                canvas.create_oval(
                    cx - r, cy - r, cx + r, cy + r, fill=SURFACE, outline=ACCENT, width=2
                )
                canvas.create_oval(cx - 3, cy - 3, cx + 3, cy + 3, fill=ACCENT, outline=ACCENT)
            else:
                canvas.create_oval(
                    cx - r, cy - r, cx + r, cy + r, fill=SURFACE, outline=LINE, width=2
                )

        canvas.bind("<Configure>", redraw)
        node["redraw"] = redraw
        nodes[name] = node
        return content

    def set_state(name, st):
        nodes[name]["state"] = st
        nodes[name]["redraw"]()

    stage_po = make_stage("po", 0, is_first=True)
    ttk.Label(stage_po, text="採購單檔案", style="Section.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Label(stage_po, text="內容為 HTML 的 .xls", style="Hint.TLabel").grid(
        row=1, column=0, sticky="w", pady=(1, 6)
    )
    po_row = ttk.Frame(stage_po)
    po_row.grid(row=2, column=0, sticky="ew")
    po_row.columnconfigure(0, weight=1)
    ttk.Entry(po_row, textvariable=po_var, state="readonly").grid(row=0, column=0, sticky="ew", padx=(0, 8))
    ttk.Button(po_row, text="選擇", style="Action.TButton", command=lambda: choose_po()).grid(row=0, column=1)

    stage_ship = make_stage("ship", 1)
    ttk.Label(stage_ship, text="出貨清單", style="Section.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Label(stage_ship, text=".xlsx 或 .csv", style="Hint.TLabel").grid(
        row=1, column=0, sticky="w", pady=(1, 6)
    )
    ship_row = ttk.Frame(stage_ship)
    ship_row.grid(row=2, column=0, sticky="ew")
    ship_row.columnconfigure(0, weight=1)
    ttk.Entry(ship_row, textvariable=ship_var, state="readonly").grid(row=0, column=0, sticky="ew", padx=(0, 8))
    ttk.Button(ship_row, text="選擇", style="Action.TButton", command=lambda: choose_ship()).grid(row=0, column=1)

    stage_run = make_stage("run", 2)
    ttk.Label(stage_run, text="處理", style="Section.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Label(stage_run, textvariable=stage_status_var, style="Hint.TLabel").grid(
        row=1, column=0, sticky="w", pady=(1, 8)
    )
    run_button = ttk.Button(stage_run, text="開始處理", style="Primary.TButton", command=lambda: start_processing())
    run_button.grid(row=2, column=0, sticky="w")
    run_button.state(["disabled"])

    stage_done = make_stage("done", 3, is_last=True)
    ttk.Label(stage_done, text="完成", style="Section.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Label(stage_done, textvariable=done_status_var, style="Done.TLabel").grid(
        row=1, column=0, sticky="w", pady=(1, 8)
    )
    tallies = ttk.Frame(stage_done)
    tallies.grid(row=2, column=0, sticky="w")

    def add_tally(col, cap, var, val_style):
        cell = ttk.Frame(tallies)
        cell.grid(row=0, column=col, padx=(0, 18))
        ttk.Label(cell, textvariable=var, style=val_style).pack(anchor="w")
        ttk.Label(cell, text=cap, style="TallyCap.TLabel").pack(anchor="w")

    add_tally(0, "已更新", updated_var, "Ok.TLabel")
    add_tally(1, "原值相同", unchanged_var, "Un.TLabel")
    add_tally(2, "已略過", skipped_var, "Sk.TLabel")

    ttk.Label(
        timeline,
        text="比對規則：採購單號、項次與料號完全相同才會更新；未列入或不相符的項目保留原值。",
        wraplength=330,
        style="Hint.TLabel",
    ).grid(row=4, column=1, sticky="w", pady=(16, 0))

    # ---------------- right: live console ----------------
    console = ttk.Frame(body)
    console.grid(row=0, column=1, sticky="nsew")
    console.columnconfigure(0, weight=1)
    console.rowconfigure(1, weight=1)

    ttk.Label(console, text="處理紀錄", style="Con.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))

    text_wrap = tk.Frame(console, bg=LINE)
    text_wrap.grid(row=1, column=0, sticky="nsew")
    text_wrap.columnconfigure(0, weight=1)
    text_wrap.rowconfigure(0, weight=1)
    log_text = tk.Text(
        text_wrap, bg=CON_BG, fg=CON_FG, insertbackground=CON_FG, relief="flat",
        highlightthickness=0, bd=0, padx=12, pady=10, wrap="word", font=mono,
        state="disabled", height=10,
    )
    log_text.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
    scroll = ttk.Scrollbar(text_wrap, orient="vertical", command=log_text.yview)
    scroll.grid(row=0, column=1, sticky="ns", padx=(0, 1), pady=1)
    log_text.configure(yscrollcommand=scroll.set)
    for tag, color in (
        ("prompt", "#38B9C2"), ("cmd", "#E6ECF2"), ("msg", CON_FG),
        ("lvl_OK", "#52C088"), ("lvl_SKIP", "#D6A64A"), ("lvl_WARN", "#D6A64A"),
        ("lvl_INFO", "#8FE1E7"), ("lvl_ERR", "#E36C63"), ("sys", "#5E6975"),
    ):
        log_text.tag_configure(tag, foreground=color)

    ttk.Label(console, textvariable=result_path_var, wraplength=520, style="Hint.TLabel").grid(
        row=2, column=0, sticky="w", pady=(10, 6)
    )
    result_actions = ttk.Frame(console)
    result_actions.grid(row=3, column=0, sticky="w")

    def open_path(path):
        if not path:
            return
        try:
            os.startfile(str(path))
        except OSError as exc:
            messagebox.showerror("無法開啟", str(exc), parent=root)

    open_result_button = ttk.Button(
        result_actions, text="開啟結果檔", style="Action.TButton",
        command=lambda: open_path(ui["output"]),
    )
    open_result_button.grid(row=0, column=0, padx=(0, 8))
    open_folder_button = ttk.Button(
        result_actions, text="開啟所在資料夾", style="Action.TButton",
        command=lambda: open_path(Path(ui["output"]).parent if ui["output"] else None),
    )
    open_folder_button.grid(row=0, column=1, padx=(0, 8))
    open_log_button = ttk.Button(
        result_actions, text="查看詳細紀錄", style="Action.TButton",
        command=lambda: open_path(ui["log"]),
    )
    open_log_button.grid(row=0, column=2)
    result_buttons = (open_result_button, open_folder_button, open_log_button)
    for _button in result_buttons:
        _button.state(["disabled"])

    def con_write(segments):
        log_text.configure(state="normal")
        for txt, tag in segments:
            log_text.insert("end", txt, tag)
        log_text.see("end")
        log_text.configure(state="disabled")

    def con_clear():
        log_text.configure(state="normal")
        log_text.delete("1.0", "end")
        log_text.configure(state="disabled")

    def con_cmd(line):
        con_write([("› ", "prompt"), (line + "\n", "cmd")])

    def con_log(level, msg):
        tag = {"OK": "lvl_OK", "SKIP": "lvl_SKIP", "WARN": "lvl_WARN", "INFO": "lvl_INFO"}.get(
            level, "lvl_INFO"
        )
        con_write([(f"[{level}] ", tag), (msg + "\n", "msg")])

    def update_flow():
        if ui["processing"]:
            return
        has_po = bool(po_var.get())
        has_ship = bool(ship_var.get())
        set_state("po", "done" if has_po else "active")
        set_state("ship", ("done" if has_ship else "active") if has_po else "pending")
        if has_po and has_ship:
            set_state("run", "active")
            run_button.state(["!disabled"])
            stage_status_var.set("兩個檔案已就緒，可以開始")
        else:
            set_state("run", "pending")
            run_button.state(["disabled"])
            stage_status_var.set("選好兩個檔案後即可開始")
        set_state("done", "pending")

    def choose_po():
        initial = Path(po_var.get()).parent if po_var.get() else script_dir()
        selected = filedialog.askopenfilename(
            parent=root, title="選擇採購單檔案", initialdir=str(initial),
            filetypes=[("HTML 格式採購單", "*.xls;*.html;*.htm"), ("所有檔案", "*.*")],
        )
        if selected:
            po_var.set(selected)
            update_flow()

    def choose_ship():
        if ship_var.get():
            initial = Path(ship_var.get()).parent
        elif po_var.get():
            initial = Path(po_var.get()).parent
        else:
            initial = script_dir()
        selected = filedialog.askopenfilename(
            parent=root, title="選擇出貨清單", initialdir=str(initial),
            filetypes=[
                ("出貨清單", "*.xlsx;*.csv"),
                ("Excel 活頁簿", "*.xlsx"),
                ("CSV", "*.csv"),
                ("所有檔案", "*.*"),
            ],
        )
        if selected:
            ship_var.set(selected)
            update_flow()

    def start_processing():
        if ui["processing"] or not (po_var.get() and ship_var.get()):
            return
        ui["processing"] = True
        run_button.state(["disabled"])
        for _button in result_buttons:
            _button.state(["disabled"])
        set_state("po", "done")
        set_state("ship", "done")
        set_state("run", "active")
        set_state("done", "pending")
        stage_status_var.set("處理中…")
        done_status_var.set("處理中…")
        for var in (updated_var, unchanged_var, skipped_var):
            var.set("0")
        result_path_var.set("處理中…")
        con_clear()
        con_cmd(f"採購單 = {Path(po_var.get()).name}")
        con_cmd(f"出貨清單 = {Path(ship_var.get()).name}")
        con_cmd("開始處理…")
        root.configure(cursor="wait")

        po_path = Path(po_var.get())
        ship_path = Path(ship_var.get())

        def worker():
            try:
                result = run_conversion(
                    po_path, ship_path, on_log=lambda lv, m: q.put(("log", lv, m))
                )
            except Exception as exc:  # noqa: BLE001 — surfaced to the user below
                q.put(("error", exc))
            else:
                q.put(("done", result))

        threading.Thread(target=worker, daemon=True).start()
        root.after(40, drain_queue)

    def drain_queue():
        try:
            while True:
                item = q.get_nowait()
                if item[0] == "log":
                    con_log(item[1], item[2])
                elif item[0] == "done":
                    finish_ok(item[1])
                elif item[0] == "error":
                    finish_err(item[1])
        except queue.Empty:
            pass
        if ui["processing"]:
            root.after(40, drain_queue)

    def finish_ok(result):
        ui["processing"] = False
        root.configure(cursor="")
        stats = result["stats"]
        skipped = stats["skipped_part_mismatch"] + stats["skipped_invalid"] + stats["miss"]
        updated_var.set(str(stats["updated"]))
        unchanged_var.set(str(stats["unchanged"]))
        skipped_var.set(str(skipped))
        ui["output"] = result["output"]
        ui["log"] = result["log"]
        result_path_var.set(str(result["output"]))
        set_state("run", "done")
        set_state("done", "done")
        elapsed = result["elapsed"]
        done_status_var.set(f"處理完成（{elapsed:.2f} 秒）")
        stage_status_var.set("已完成")
        con_write([("── 完成 ", "sys"), (f"{elapsed:.2f}s\n", "sys")])
        for _button in result_buttons:
            _button.state(["!disabled"])
        run_button.state(["!disabled"])
        if result["html_warning"]:
            messagebox.showwarning(
                "格式提醒",
                "採購單檔案不是標準 HTML 開頭，請務必檢查輸出結果。",
                parent=root,
            )

    def finish_err(exc):
        ui["processing"] = False
        root.configure(cursor="")
        set_state("run", "active")
        set_state("done", "pending")
        con_write([("[錯誤] ", "lvl_ERR"), (f"{exc}\n", "msg")])
        done_status_var.set("處理失敗")
        stage_status_var.set("處理失敗，請檢查檔案")
        result_path_var.set("尚無結果檔")
        run_button.state(["!disabled"])
        messagebox.showerror(
            "處理失敗",
            f"請檢查選擇的檔案與內容。\n\n詳細錯誤：{exc}",
            parent=root,
        )

    con_write([("等待開始…\n", "sys")])
    root.after(60, update_flow)
    root.mainloop()


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)

    if len(argv) == 1 and argv[0] in {"-h", "--help", "/?"}:
        print(__doc__)
        return 0
    if len(argv) < 2:
        launch_gui()
        return 0

    po_path = Path(argv[0])
    ship_path = Path(argv[1])
    print("使用命令列指定的檔案:")
    print("  PO  =", po_path)
    print("  SHIP=", ship_path)
    try:
        result = run_conversion(po_path, ship_path)
    except FileNotFoundError as exc:
        print(exc)
        return 2
    except (ValueError, UnicodeError, RuntimeError) as exc:
        print(exc)
        return 1

    if result["html_warning"]:
        print("警告: 採購單不是 HTML 格式的 .xls，仍嘗試處理")
    out = result["output"]
    stats = result["stats"]
    elapsed = result["elapsed"]
    all_logs = result["logs"]
    log_path = result["log"]
    print("RESULT", out)
    print("STATS", stats)
    print(f"TIME   {elapsed:.3f}s")
    for lv, msg in all_logs[:80]:
        print(f"[{lv}] {msg}")
    if len(all_logs) > 80:
        print(f"... 另有 {len(all_logs) - 80} 行寫入 log")
    print("LOG", log_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
