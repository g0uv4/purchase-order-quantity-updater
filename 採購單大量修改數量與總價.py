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


def load_ship(ship_path: Path):
    ship: dict[str, dict] = {}
    affected: set[str] = set()
    logs: list[tuple[str, str]] = []

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
            logs.append(("WARN", f"出貨列 {rno} 項次/數量無法解析，略過"))
            continue
        if qty < 0:
            logs.append(("WARN", f"出貨列 {rno} 數量為負，略過"))
            continue
        if not part:
            logs.append(("WARN", f"出貨列 {rno} 料號空白，無法完整比對，略過"))
            continue
        key = f"{po}|{int(line)}"
        if key in ship:
            # 同一 PO + 項次出現多筆時無法判斷哪一筆才正確，整組停用。
            ship[key]["valid"] = False
            logs.append(("WARN", f"出貨鍵重複，該鍵全部略過: {key}"))
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

    logs.append(("INFO", f"出貨載入 筆數={len(ship)} 影響PO數={len(affected)}"))
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


def process_po_html(html: str, ship: dict, affected: set[str]):
    logs: list[tuple[str, str]] = []
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
                        logs.append(
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
                        logs.append(
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
                logs.append(
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
                logs.append(("SKIP", f"出貨資料不唯一，保留原值: {key}"))
                out_parts.append(m.group(0))
                last = m.end()
                continue

            # 料號也必須完全一致；不一致時連數量、金額與合計都不動。
            if ship_row["part"] != part:
                stats["skipped_part_mismatch"] += 1
                logs.append(
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
                logs.append(
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
            logs.append(
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
        logs.append(("WARN", f"找不到完全相同的 PO + 項次，已略過: {k}"))
    stats["miss"] = len(miss)
    return new_html, stats, logs


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

APP_NAME = "採購單數量與金額調整工具"


def run_conversion(po_path: Path, ship_path: Path):
    """Process two selected files and return paths, statistics, and logs."""
    po_path = Path(po_path)
    ship_path = Path(ship_path)
    t0 = time.perf_counter()

    if not po_path.exists():
        raise FileNotFoundError(f"找不到採購單檔案：{po_path}")
    if not ship_path.exists():
        raise FileNotFoundError(f"找不到出貨清單：{ship_path}")

    ship, affected, logs1 = load_ship(ship_path)
    if not ship:
        raise ValueError("出貨清單沒有有效資料，請檢查欄位名稱與內容。")

    raw = po_path.read_bytes()
    html = raw.decode("utf-8")
    new_html, stats, logs2 = process_po_html(html, ship, affected)

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
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("此環境無法開啟操作視窗。") from exc

    root = tk.Tk()
    root.title(APP_NAME)
    root.geometry("760x590")
    root.minsize(640, 540)

    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"))
    style.configure("Subtitle.TLabel", font=("Segoe UI", 10))
    style.configure("Section.TLabel", font=("Segoe UI", 10, "bold"))
    style.configure("Hint.TLabel", font=("Segoe UI", 9))
    style.configure("ResultValue.TLabel", font=("Segoe UI", 15, "bold"))
    style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=(18, 8))
    style.configure("Action.TButton", padding=(10, 6))

    po_var = tk.StringVar()
    ship_var = tk.StringVar()
    status_var = tk.StringVar(value="請先選擇兩個檔案")
    result_path_var = tk.StringVar()
    updated_var = tk.StringVar(value="0")
    unchanged_var = tk.StringVar(value="0")
    skipped_var = tk.StringVar(value="0")
    last_result = {"output": None, "log": None}

    main_frame = ttk.Frame(root, padding=(28, 24, 28, 22))
    main_frame.pack(fill="both", expand=True)
    main_frame.columnconfigure(0, weight=1)

    ttk.Label(main_frame, text=APP_NAME, style="Title.TLabel").grid(
        row=0, column=0, sticky="w"
    )
    ttk.Label(
        main_frame,
        text="依出貨清單更新指定項目的數量與金額。原始採購單不會被覆寫。",
        style="Subtitle.TLabel",
    ).grid(row=1, column=0, sticky="w", pady=(6, 18))

    ttk.Separator(main_frame).grid(row=2, column=0, sticky="ew", pady=(0, 18))

    files_frame = ttk.Frame(main_frame)
    files_frame.grid(row=3, column=0, sticky="ew")
    files_frame.columnconfigure(0, weight=1)

    def update_ready_state():
        ready = bool(po_var.get() and ship_var.get())
        if ready:
            run_button.state(["!disabled"])
            status_var.set("檔案已就緒，可以開始處理")
        else:
            run_button.state(["disabled"])
            status_var.set("請先選擇兩個檔案")

    def choose_po():
        initial = Path(po_var.get()).parent if po_var.get() else script_dir()
        selected = filedialog.askopenfilename(
            parent=root,
            title="選擇採購單檔案",
            initialdir=str(initial),
            filetypes=[
                ("HTML 格式採購單", "*.xls;*.html;*.htm"),
                ("所有檔案", "*.*"),
            ],
        )
        if selected:
            po_var.set(selected)
            update_ready_state()

    def choose_ship():
        if ship_var.get():
            initial = Path(ship_var.get()).parent
        elif po_var.get():
            initial = Path(po_var.get()).parent
        else:
            initial = script_dir()
        selected = filedialog.askopenfilename(
            parent=root,
            title="選擇出貨清單",
            initialdir=str(initial),
            filetypes=[
                ("出貨清單", "*.xlsx;*.csv"),
                ("Excel 活頁簿", "*.xlsx"),
                ("CSV", "*.csv"),
                ("所有檔案", "*.*"),
            ],
        )
        if selected:
            ship_var.set(selected)
            update_ready_state()

    def add_file_row(row, title, hint, variable, command):
        frame = ttk.Frame(files_frame)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 16))
        frame.columnconfigure(0, weight=1)
        ttk.Label(frame, text=title, style="Section.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(frame, text=hint, style="Hint.TLabel").grid(
            row=1, column=0, sticky="w", pady=(2, 6)
        )
        ttk.Entry(frame, textvariable=variable, state="readonly").grid(
            row=2, column=0, sticky="ew", padx=(0, 10)
        )
        ttk.Button(frame, text="選擇檔案", command=command, style="Action.TButton").grid(
            row=2, column=1, sticky="e"
        )

    add_file_row(0, "1. 採購單檔案", "支援內容為 HTML 的 .xls", po_var, choose_po)
    add_file_row(1, "2. 出貨清單", "支援 .xlsx 或 .csv", ship_var, choose_ship)

    rule_text = (
        "比對規則：採購單號、項次與料號完全相同才會更新；"
        "未列入或不相符的項目保留原值。"
    )
    ttk.Label(main_frame, text=rule_text, wraplength=680, style="Hint.TLabel").grid(
        row=4, column=0, sticky="w", pady=(0, 18)
    )

    action_frame = ttk.Frame(main_frame)
    action_frame.grid(row=5, column=0, sticky="ew")
    action_frame.columnconfigure(0, weight=1)
    ttk.Label(action_frame, textvariable=status_var, style="Hint.TLabel").grid(
        row=0, column=0, sticky="w"
    )

    result_frame = ttk.LabelFrame(main_frame, text="處理結果", padding=(16, 12))
    result_frame.columnconfigure(0, weight=1)

    stats_frame = ttk.Frame(result_frame)
    stats_frame.grid(row=0, column=0, sticky="ew")
    for column in range(3):
        stats_frame.columnconfigure(column, weight=1)

    def add_stat(column, label, variable):
        cell = ttk.Frame(stats_frame)
        cell.grid(row=0, column=column, sticky="ew")
        ttk.Label(cell, text=label, style="Hint.TLabel").pack(anchor="center")
        ttk.Label(cell, textvariable=variable, style="ResultValue.TLabel").pack(
            anchor="center", pady=(2, 0)
        )

    add_stat(0, "已更新", updated_var)
    add_stat(1, "原值相同", unchanged_var)
    add_stat(2, "已略過", skipped_var)

    ttk.Label(result_frame, text="結果檔", style="Section.TLabel").grid(
        row=1, column=0, sticky="w", pady=(14, 3)
    )
    ttk.Label(
        result_frame,
        textvariable=result_path_var,
        wraplength=650,
        justify="left",
        style="Hint.TLabel",
    ).grid(row=2, column=0, sticky="w")

    result_actions = ttk.Frame(result_frame)
    result_actions.grid(row=3, column=0, sticky="w", pady=(12, 0))

    def open_path(path):
        try:
            os.startfile(str(path))
        except OSError as exc:
            messagebox.showerror("無法開啟", str(exc), parent=root)

    open_result_button = ttk.Button(
        result_actions,
        text="開啟結果檔",
        command=lambda: open_path(last_result["output"]),
        style="Action.TButton",
    )
    open_result_button.grid(row=0, column=0, padx=(0, 8))
    open_folder_button = ttk.Button(
        result_actions,
        text="開啟所在資料夾",
        command=lambda: open_path(Path(last_result["output"]).parent),
        style="Action.TButton",
    )
    open_folder_button.grid(row=0, column=1, padx=(0, 8))
    open_log_button = ttk.Button(
        result_actions,
        text="查看詳細紀錄",
        command=lambda: open_path(last_result["log"]),
        style="Action.TButton",
    )
    open_log_button.grid(row=0, column=2)

    def run_selected_files():
        run_button.state(["disabled"])
        result_frame.grid_remove()
        status_var.set("處理中…")
        root.configure(cursor="wait")
        root.update_idletasks()
        try:
            result = run_conversion(Path(po_var.get()), Path(ship_var.get()))
        except Exception as exc:
            status_var.set("處理失敗，檔案未完成輸出")
            messagebox.showerror(
                "處理失敗",
                f"請檢查選擇的檔案與內容。\n\n詳細錯誤：{exc}",
                parent=root,
            )
        else:
            stats = result["stats"]
            skipped = (
                stats["skipped_part_mismatch"]
                + stats["skipped_invalid"]
                + stats["miss"]
            )
            updated_var.set(str(stats["updated"]))
            unchanged_var.set(str(stats["unchanged"]))
            skipped_var.set(str(skipped))
            result_path_var.set(str(result["output"]))
            last_result["output"] = result["output"]
            last_result["log"] = result["log"]
            status_var.set(f"處理完成（{result['elapsed']:.2f} 秒）")
            result_frame.grid(row=6, column=0, sticky="ew", pady=(18, 0))
            if result["html_warning"]:
                messagebox.showwarning(
                    "格式提醒",
                    "採購單檔案不是標準 HTML 開頭，請務必檢查輸出結果。",
                    parent=root,
                )
        finally:
            root.configure(cursor="")
            run_button.state(["!disabled"])

    run_button = ttk.Button(
        action_frame,
        text="開始處理",
        command=run_selected_files,
        style="Primary.TButton",
    )
    run_button.grid(row=0, column=1, sticky="e")
    run_button.state(["disabled"])

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
