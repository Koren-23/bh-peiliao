#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BH型鋼鋼板配料系統 GUI
版本: 24.5  日期: 2026/10
"""

import sys, subprocess, importlib

def _auto_install(pkg, import_name=None):
    """自動安裝缺少的套件"""
    import_name = import_name or pkg
    try:
        importlib.import_module(import_name)
    except ImportError:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk(); root.withdraw()
        messagebox.showinfo("套件安裝",
            f"正在自動安裝必要套件：{pkg}\n請稍候...")
        root.destroy()
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg,
                               "--quiet", "--break-system-packages"],
                              stderr=subprocess.DEVNULL)

_auto_install("openpyxl")
_auto_install("reportlab")

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import re
import math
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import (Font, Alignment, PatternFill, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter
import os

# ─── 常數 ─────────────────────────────────────────────────────────────
DENSITY = {"黑鐵板（碳鋼）": 7.85, "不鏽鋼板（304）": 7.93, "不鏽鋼板（316）": 7.98}
BW_MIN, BW_MAX = 1524, 2499   # 板寬範圍
BL_MIN       = 8000            # 板長最小值
BL_MAX       = 18499           # 板長最大值
W_MIN        = 2600            # 最低重量 kg
W_MAX        = 12500           # 最高重量 kg

# ─── 顏色主題 ─────────────────────────────────────────────────────────
CLR_BG       = "#F0F4F8"
CLR_HEADER   = "#1A3050"
CLR_BTN_MAIN = "#2B6CB0"
CLR_BTN_ADD  = "#2B6CB0"
CLR_BTN_DEL  = "#7A3535"
CLR_BTN_OUT  = "#276749"
CLR_SCRAP    = "#D4EDDA"       # 餘料綠
CLR_WHITE    = "#FFFFFF"
CLR_ROW_ODD  = "#F7FAFC"
CLR_ROW_EVEN = "#EBF4FF"

# ═══════════════════════════════════════════════════════════════════════
# 計算核心
# ═══════════════════════════════════════════════════════════════════════

def parse_bh(spec: str):
    """解析 BH高×寬×腹厚×翼厚，回傳 (H, B, tw, tf) 或 None"""
    spec = spec.strip().upper().replace("*", "×").replace("X", "×")
    m = re.fullmatch(r"BH\s*(\d+)[×x](\d+)[×x](\d+)[×x](\d+)", spec)
    if not m:
        return None
    H, B, tw, tf = tuple(int(v) for v in m.groups())
    if H <= 0 or B <= 0 or tw <= 0 or tf <= 0:
        return None
    if H - 2*tf <= 0:
        return None
    return H, B, tw, tf


def calc_weight(width_mm, thick_mm, length_mm, density):
    """計算鋼板重量 kg"""
    return width_mm / 1000 * thick_mm / 1000 * length_mm / 1000 * density * 1000


def decompose_bh(row, density, new_kerf, new_trim, scrap_kerf, scrap_trim, existing_scraps):
    """
    將單一 BH 構件拆為翼板 / 腹板零件，並進行配料（優先餘料）。
    row = {"comp": str, "part": str, "spec": str, "length": int, "qty": int, "mat": str}
    回傳 dict 含 flange / web 資訊及配料結果
    """
    parsed = parse_bh(row["spec"])
    if not parsed:
        return None
    H, B, tw, tf = parsed
    hw = H - 2 * tf       # 腹板高

    flange = {"name": f"{row['comp']}-F", "width": B, "thick": tf,
               "length": row["length"], "qty": 2 * row["qty"],
               "unit": 2, "total": 2 * row["qty"], "mat": row["mat"],
               "spec_str": f"PL{tf}×{B}×{row['length']}"}
    web    = {"name": f"{row['comp']}-W", "width": hw, "thick": tw,
               "length": row["length"], "qty": row["qty"],
               "unit": 1, "total": row["qty"], "mat": row["mat"],
               "spec_str": f"PL{tw}×{hw}×{row['length']}"}

    for part in (flange, web):
        w = calc_weight(part["width"], part["thick"], part["length"], density)
        part["unit_wt"]  = round(w, 1)
        part["total_wt"] = round(w * part["total"], 1)

    return {"comp": row["comp"], "part": row.get("part",""), "spec": row["spec"],
            "length": row["length"], "qty": row["qty"], "mat": row["mat"],
            "H": H, "B": B, "tw": tw, "tf": tf, "hw": hw,
            "flange": flange, "web": web}


def scrap_runs(rems):
    """
    多段切割各排的長度方向餘料：相鄰且等長（對齊）的排合併為一塊。
    rems：每排餘料長度（0＝無餘料）→ [{"start", "end", "len"}]（start/end 為排索引）
    """
    runs = []
    for i, v in enumerate(rems):
        if v <= 0:
            continue
        if runs and runs[-1]["len"] == v and runs[-1]["end"] == i - 1:
            runs[-1]["end"] = i
        else:
            runs.append({"start": i, "end": i, "len": v})
    return runs


def run_width(run, part_w, kerf):
    """合併後餘料寬度：排寬 × 排數 + 排間鋸縫（餘料未切開，鋸縫寬度仍在）"""
    n = run["end"] - run["start"] + 1
    return n * part_w + (n - 1) * kerf


def board_leftovers(thick, bw, bl, part_w, trim, kerf, col_used, density, src, typ, mat, new_scraps):
    """
    一張板（新板或現有餘料）的餘料：新板 / 餘料、單段 / 多段、配料計算 / 採購修正共用。
    col_used：每排已用長度（含前端修邊）。寬度方向扣兩側修邊及一刀鋸縫；
    長度方向每排扣尾端修邊及最後一刀鋸縫，相鄰等長的排合併成一塊。
    餘料物件依序加入 new_scraps；回傳 {"left_spec", "width_obj", "row_objs"}
    """
    n = len(col_used)
    specs = []

    def mk(w, l):
        wt = round(calc_weight(w, thick, l, density), 0)
        specs.append(f"PL{thick}×{w}×{l}（{wt}kg）")
        obj = {"src": src, "type": typ, "spec": f"PL{thick}×{w}×{l}", "mat": mat, "wt": int(wt)}
        new_scraps.append(obj)
        return obj

    lw = max(0, bw - (n * part_w + max(0, n - 1) * kerf + trim * 2) - kerf)
    width_obj = mk(lw, bl) if lw > 0 else None
    row_objs = [None] * n
    for run in scrap_runs([max(0, bl - c - trim - kerf) for c in col_used]):
        obj = mk(run_width(run, part_w, kerf), run["len"])
        for ri in range(run["start"], run["end"] + 1):
            row_objs[ri] = obj
    return {"left_spec": "　".join(specs) if specs else "無餘料",
            "width_obj": width_obj, "row_objs": row_objs}


def make_layout(bw, bl, part_w, thick, trim, kerf, rows, lo):
    """
    排列圖用結構化資料。每排的 scrap_ref 直接存餘料物件參照，PreviewWindow 顯示餘料清單時，
    會把最終編號（餘01/餘02...）寫回該物件的 "_no" 欄位，排列圖繪製時就能讀到正確的餘NO。
    """
    return {
        "board_w": bw, "board_l": bl, "part_w": part_w, "thick": thick,
        "trim": trim, "kerf": kerf, "width_scrap_ref": lo["width_obj"],
        "rows": [{"parts": [{"name": p["name"], "length": p["length"]} for p in row],
                  "scrap_ref": lo["row_objs"][ri]} for ri, row in enumerate(rows)],
    }


def _plan_once(by_group, density, new_kerf, new_trim, scrap_kerf, scrap_trim,
               existing_scraps, seed=None,
               bw_min=BW_MIN, bw_max=BW_MAX, bl_min=BL_MIN, bl_max=BL_MAX,
               w_min=W_MIN, w_max=W_MAX,
               cut_mode="single"):
    """單次配料計算（可帶不同 seed 做隨機排列）"""
    import random
    rng = random.Random(seed)

    purchase_list  = []
    cut_details    = []
    new_scraps     = []
    used_scrap_ids = set()
    cut_idx        = 0
    circled = ["①","②","③","④","⑤","⑥","⑦","⑧","⑨","⑩",
               "⑪","⑫","⑬","⑭","⑮","⑯","⑰","⑱","⑲","⑳"]

    def next_idx():
        nonlocal cut_idx
        s = circled[cut_idx] if cut_idx < len(circled) else f"({cut_idx+1})"
        cut_idx += 1
        return s

    # 餘料依 qty 展開成單片，每片獨立追蹤
    remaining_scraps = []
    for s in existing_scraps:
        qty = int(s.get("qty", 1))
        for i in range(qty):
            piece = dict(s)
            piece["_piece_idx"] = i
            remaining_scraps.append(piece)

    for (thick, width, mat), group in sorted(by_group.items()):
        # 攤平成單片需求
        demand = []
        for p in group:
            for _ in range(p["total"]):
                demand.append({
                    "name": p["name"], "width": width, "thick": thick,
                    "length": p["length"], "unit_wt": p["unit_wt"],
                    "mat": p["mat"], "type": "F" if p["name"].endswith("-F") else "W"
                })

        # 隨機打亂（保持長度降冪為主，加入微擾）
        if seed is not None:
            rng.shuffle(demand)
        remaining = list(demand)

        # ── 餘料優先 ──────────────────────────────────────────────
        for sc in remaining_scraps:
            if not remaining: break
            if sc.get("used", 0) or sc["thick"] != thick: continue
            if sc.get("mat", mat) != mat: continue
            usable_w = sc["width"]  - scrap_trim * 2
            usable_l = sc["length"] - scrap_trim * 2
            if usable_w < width: continue   # 餘料寬度不足一排
            max_cols = (usable_w + scrap_kerf) // (width + scrap_kerf)
            eligible = sorted([d for d in remaining if d["length"] <= usable_l],
                              key=lambda d: -d["length"])
            if not eligible: continue
            fits  = min(max_cols, len(eligible))
            taken = eligible[:fits]
            for t in taken: remaining.remove(t)
            sc["used"] = True
            used_scrap_ids.add(sc["id"])
            idx      = next_idx() + "♻"
            board_wt = round(calc_weight(sc["width"], sc["thick"], sc["length"], density), 0)
            # 每排一片；寬度方向與長度方向餘料皆與新板相同算法（扣鋸縫及修邊）
            lo = board_leftovers(sc["thick"], sc["width"], sc["length"], width, scrap_trim, scrap_kerf,
                                 [scrap_trim + t["length"] for t in taken], density,
                                 idx, taken[0]["type"], taken[0]["mat"], new_scraps)
            comp_str = "、".join(dict.fromkeys(t["name"] for t in taken))
            cut_details.append({
                "idx": idx, "type": taken[0]["type"],
                "board_spec": f"PL{sc["thick"]}×{sc["width"]}×{sc["length"]}（餘料）",
                "mat": taken[0]["mat"], "board_wt": int(board_wt),
                "comp": comp_str,
                "part_spec": f"PL{thick}×{width}×{max(t["length"] for t in taken)}",
                "qty": fits, "unit_wt": taken[0]["unit_wt"],
                "leftover": lo["left_spec"], "is_scrap": True,
                "layout": make_layout(sc["width"], sc["length"], width, thick, scrap_trim, scrap_kerf,
                                      [[t] for t in taken], lo)
            })

        # ── 新板裝箱 ──────────────────────────────────────────────
        # max_cols：最多可以幾排並排，嚴格限制在 bw_max 以內
        max_cols = max(1, int((bw_max + new_kerf) // (width + new_kerf)))

        def _add_board(cols, bl, taken, flag="", seg_desc=None, col_used_list=None,
                       col_parts_layout=None):
            """
            col_used_list:     每一排實際用掉的長度（含前端修邊）
            col_parts_layout:  每一排的零件清單（list of list），
                               用來在排列圖視窗繪製實際排列位置（單段切割為每排一片）
            """
            nonlocal bw_min, bw_max, w_min
            bw  = cols * width + (cols-1)*new_kerf + new_trim*2
            bw  = max(bw, bw_min)
            bw  = min(bw, bw_max)   # 嚴格限制在 bw_max 以內
            # 板長限制
            bl  = max(bl, bl_min)
            bl  = min(bl, bl_max)   # 嚴格限制在 bl_max 以內
            bwt = calc_weight(bw, thick, bl, density)
            f   = flag
            if bwt < w_min:
                f = flag + "*"
                need_w = w_min / (thick/1000 * bl/1000 * density * 1000)
                bw = min(int(math.ceil(need_w*1000/10)*10), bw_max)
                bwt = calc_weight(bw, thick, bl, density)
            # 板重上限檢查（保留可能已加上的「*」）
            if bwt > w_max:
                f = f + "H"   # H = Heavy，超重標記
            bwt_r      = round(bwt, 0)
            idx        = next_idx() + f
            board_spec = f"PL{thick}×{bw}×{bl}"
            comp_str   = "、".join(dict.fromkeys(t["name"] for t in taken))
            ps = (f"PL{thick}×{width}×多段({seg_desc})" if seg_desc
                  else f"PL{thick}×{width}×{max(t['length'] for t in taken)}")

            # 餘料：寬度方向一塊 + 每排長度方向（相鄰等長的排合併），單段 / 多段相同算法
            lo = board_leftovers(thick, bw, bl, width, new_trim, new_kerf, col_used_list,
                                 density, idx, taken[0]["type"], taken[0]["mat"], new_scraps)
            layout = make_layout(bw, bl, width, thick, new_trim, new_kerf, col_parts_layout, lo)
            left_spec = lo["left_spec"]

            cut_details.append({
                "idx": idx, "type": taken[0]["type"],
                "board_spec": board_spec, "mat": taken[0]["mat"],
                "board_wt": int(bwt_r), "comp": comp_str,
                "part_spec": ps, "qty": len(taken),
                "unit_wt": taken[0]["unit_wt"],
                "leftover": left_spec, "is_scrap": False,
                "layout": layout
            })
            purchase_list.append({
                "spec": board_spec, "qty": 1, "mat": taken[0]["mat"],
                "unit_wt": int(bwt_r), "total_wt": int(bwt_r)
            })

        if cut_mode == "single":
            while remaining:
                remaining.sort(key=lambda d: -d["length"])
                max_len  = remaining[0]["length"]
                eligible = [d for d in remaining if d["length"] <= max_len]
                cols     = min(max_cols, len(eligible))
                taken    = eligible[:cols]
                for t in taken: remaining.remove(t)
                bl = max_len + new_trim * 2
                # 單段：每排一片（各排長度可能不同，短的排尾端留下長度方向餘料）
                _add_board(cols, bl, taken, "", None,
                           [new_trim + t["length"] for t in taken], [[t] for t in taken])

        else:
            # 多段切割：長度方向分段排入，充分利用板長
            BL_USE = bl_max   # 使用使用者設定的板長最大值

            def _try_arrange(pieces, cols, BL_USE, new_trim, new_kerf):
                """
                嘗試把 pieces 以最省餘料的方式分配到 cols 排。
                策略：先用貪心法建立初始排列，再做「排間互換」優化——
                把不同排的零件互換位置，如果互換後整體最大 col_used 縮短則接受。
                允許長度穿插（不要求同排同長），確保找到更省料的組合。
                回傳 (col_parts, col_used) 或 None（若塞不下任何東西）。
                """
                # ── 初始貪心排列 ──
                col_used  = [new_trim] * cols
                col_parts = [[] for _ in range(cols)]

                # 先依長度由長到短排序，但允許短件穿插到有空間的排
                for piece in sorted(pieces, key=lambda d: -d["length"]):
                    # 找「塞入後該排 col_used 最小」的排（best fit decreasing）
                    best_c, best_used = None, float('inf')
                    for c in range(cols):
                        gap = new_kerf if col_parts[c] else 0
                        new_used = col_used[c] + gap + piece["length"] + new_trim
                        if new_used <= BL_USE and new_used < best_used:
                            best_c, best_used = c, new_used
                    if best_c is not None:
                        gap = new_kerf if col_parts[best_c] else 0
                        col_used[best_c] += gap + piece["length"]
                        col_parts[best_c].append(piece)

                if not any(col_parts):
                    return None, None

                # ── 互換優化：嘗試交換不同排的零件，若最大col_used縮短則接受 ──
                improved = True
                max_iter = 50   # 避免無限循環
                itr = 0
                while improved and itr < max_iter:
                    improved = False
                    itr += 1
                    for ca in range(cols):
                        for cb in range(ca+1, cols):
                            for ia in range(len(col_parts[ca])):
                                for ib in range(len(col_parts[cb])):
                                    # 交換後兩排內容已改變，每次都取最新的零件（避免同一零件重複、另一零件遺失）
                                    pa, pb = col_parts[ca][ia], col_parts[cb][ib]
                                    # 試算交換 pa（在排ca第ia個）與 pb（在排cb第ib個）後的新 col_used
                                    # 重算排ca：把pa換成pb
                                    new_used_ca = new_trim
                                    for k, p in enumerate(col_parts[ca]):
                                        pp = pb if k == ia else p
                                        new_used_ca += (new_kerf if k > 0 else 0) + pp["length"]
                                    new_used_ca += new_trim

                                    # 重算排cb：把pb換成pa
                                    new_used_cb = new_trim
                                    for k, p in enumerate(col_parts[cb]):
                                        pp = pa if k == ib else p
                                        new_used_cb += (new_kerf if k > 0 else 0) + pp["length"]
                                    new_used_cb += new_trim

                                    # 若兩排都合法（不超BL_USE）且整體最大col_used有改善則接受
                                    if new_used_ca <= BL_USE and new_used_cb <= BL_USE:
                                        old_max = max(col_used[ca], col_used[cb])
                                        new_max = max(new_used_ca - new_trim,
                                                      new_used_cb - new_trim)
                                        if new_max < old_max:
                                            # 接受互換
                                            col_parts[ca][ia], col_parts[cb][ib] = pb, pa
                                            col_used[ca] = new_used_ca - new_trim
                                            col_used[cb] = new_used_cb - new_trim
                                            improved = True

                return col_parts, col_used

            while remaining:
                remaining.sort(key=lambda d: -d["length"])
                cols = min(max_cols, len(remaining))

                # 先試一次排列，若完全塞不下任何東西才單獨處理最長的
                col_parts, col_used = _try_arrange(remaining, cols, BL_USE,
                                                    new_trim, new_kerf)
                placed = [p for cp in col_parts for p in cp] if col_parts else []

                if not placed:
                    piece = remaining.pop(0)
                    bl    = piece["length"] + new_trim*2
                    _add_board(1, bl, [piece], "", None, [new_trim + piece["length"]], [[piece]])
                    continue

                for p in placed: remaining.remove(p)
                used_cols = sum(1 for cp in col_parts if cp)
                bl = max((col_used[c] for c in range(cols) if col_parts[c]),
                          default=new_trim) + new_trim
                seg_desc = "　".join(
                    f"排{c+1}:"+"+".join(str(p["length"]) for p in col_parts[c])
                    for c in range(cols) if col_parts[c])
                all_taken = [p for cp in col_parts for p in cp]
                used_lengths = [col_used[c] for c in range(cols) if col_parts[c]]
                active_rows = [cp for cp in col_parts if cp]
                _add_board(used_cols, bl, all_taken, seg_desc=seg_desc,
                          col_used_list=used_lengths,
                          col_parts_layout=active_rows)

    return purchase_list, cut_details, new_scraps, list(used_scrap_ids)


def plan_purchase(parts_list, density, new_kerf, new_trim, scrap_kerf, scrap_trim,
                  existing_scraps, iterations=600,
                  bw_min=BW_MIN, bw_max=BW_MAX, bl_min=BL_MIN, bl_max=BL_MAX,
                  w_min=W_MIN, w_max=W_MAX, cut_mode="single",
                  progress_callback=None, cancel_event=None):
    """
    跑 iterations 次隨機配料，取「採購總重最輕 + 餘料最少」的最佳方案。
    cut_mode: "single"=單段切割  "multi"=多段切割
    """
    by_group = {}
    for p in parts_list:
        key = (p["thick"], p["width"], p["mat"])
        by_group.setdefault(key, []).append(p)

    best = None
    best_score = None

    for i in range(iterations):
        if cancel_event is not None and cancel_event.is_set():
            return None
        result = _plan_once(by_group, density, new_kerf, new_trim,
                            scrap_kerf, scrap_trim, existing_scraps, seed=i,
                            bw_min=bw_min, bw_max=bw_max,
                            bl_min=bl_min, bl_max=bl_max,
                            w_min=w_min, w_max=w_max,
                            cut_mode=cut_mode)
        purchase_list, cut_details, new_scraps, used_ids = result
        buy_wt   = sum(p["total_wt"] for p in purchase_list)
        scrap_wt = sum(s["wt"] for s in new_scraps)
        score    = (buy_wt, scrap_wt)
        if best_score is None or score < best_score:
            best_score = score
            best = result
        if progress_callback and (i % 5 == 0 or i + 1 == iterations):
            progress_callback(i + 1, iterations)

    return best





# ═══════════════════════════════════════════════════════════════════════
# Excel 輸出
# ═══════════════════════════════════════════════════════════════════════

def make_border(thin=True):
    s = Side(style="thin" if thin else "medium")
    return Border(left=s, right=s, top=s, bottom=s)

def hdr_fill():
    return PatternFill("solid", fgColor="1E3A5F")

def scrap_fill():
    return PatternFill("solid", fgColor="C6EFCE")

def write_xlsx(path, proj_no, proj_name, date_str, mat_name, density,
               new_kerf, new_trim, scrap_kerf, scrap_trim,
               bh_rows, decomposed, purchase_list, cut_details,
               new_scraps, existing_scraps, used_scrap_ids):
    wb = Workbook()
    wb.remove(wb.active)
    param_line = (f"工程：{proj_no}　{proj_name}　材質：{mat_name}　比重：{density}"
                  f"　新板每刀：{new_kerf}mm 頭尾：{new_trim}mm"
                  f"　餘料每刀：{scrap_kerf}mm 頭尾：{scrap_trim}mm")

    def add_ws(title):
        ws = wb.create_sheet(title)
        return ws

    def set_col_widths(ws, widths):
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    def write_title(ws, title, ncols, prod_date):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols-1)
        ws.merge_cells(start_row=1, start_column=ncols, end_row=1, end_column=ncols)
        c = ws.cell(1, 1, title)
        c.font = Font(bold=True, size=13, color="FFFFFF")
        c.fill = hdr_fill()
        c.alignment = Alignment(horizontal="left", vertical="center")
        c2 = ws.cell(1, ncols, f"產出日期：{prod_date}")
        c2.font = Font(size=10, color="FFFFFF")
        c2.fill = hdr_fill()
        c2.alignment = Alignment(horizontal="right", vertical="center")
        ws.row_dimensions[1].height = 22

    def write_param(ws, ncols):
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
        c = ws.cell(2, 1, param_line)
        c.font = Font(size=9, color="C53030")
        c.fill = PatternFill("solid", fgColor="FFF5F5")
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[2].height = 16

    def write_header(ws, row, headers, start_col=1):
        for j, h in enumerate(headers, start_col):
            c = ws.cell(row, j, h)
            c.font = Font(bold=True, color="FFFFFF", size=10)
            c.fill = hdr_fill()
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = make_border()
        ws.row_dimensions[row].height = 20

    def data_cell(ws, r, c, v, bold=False, align="center", bg=None):
        cell = ws.cell(r, c, v)
        cell.font = Font(bold=bold, size=10)
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.border = make_border()
        if bg:
            cell.fill = PatternFill("solid", fgColor=bg)
        ws.row_dimensions[r].height = 18
        return cell

    prod = datetime.now().strftime("%Y/%m/%d")

    # ── 1. 配料清單 ──────────────────────────────────────────────
    # 建立 comp → (serial_flange, serial_web) 對應
    serial_lookup = {}
    for d in decomposed:
        serial_lookup[d["comp"]] = {
            "serial":   d.get("serial", ""),   # 構件流水號（若有）
            "fl_serial": d["flange"].get("serial", ""),
            "wb_serial": d["web"].get("serial", ""),
        }

    ws1 = add_ws("配料清單")
    cols = ["構件編號", "零件編號", "斷面規格", "長度(mm)", "數量", "材質"]
    write_title(ws1, "配料清單　BH配料報表", 6, prod)
    write_param(ws1, 6)
    write_header(ws1, 3, cols)
    set_col_widths(ws1, [12, 12, 22, 12, 8, 8])
    for i, row in enumerate(bh_rows):
        r = i + 4
        bg = "F7FAFC" if i % 2 == 0 else "EBF4FF"
        data_cell(ws1, r, 1, row["comp"], align="center", bg=bg)
        data_cell(ws1, r, 2, row["part"], align="center", bg=bg)
        data_cell(ws1, r, 3, row["spec"], align="left",   bg=bg)
        data_cell(ws1, r, 4, row["length"], bg=bg)
        data_cell(ws1, r, 5, row["qty"],    bg=bg)
        data_cell(ws1, r, 6, row["mat"],    bg=bg)

    # ── 2. BH拆板明細 ────────────────────────────────────────────
    ws2 = add_ws("BH拆板明細")
    cols2 = ["構件編號", "零件編號", "拆板零件", "流水號", "流水號(W/F)", "斷面規格", "長度(mm)", "數量", "材質", "單重量(kg)", "總重量(kg)"]
    write_title(ws2, "BH拆板明細", 11, prod)
    write_param(ws2, 11)
    write_header(ws2, 3, cols2)
    set_col_widths(ws2, [12, 12, 12, 12, 16, 22, 12, 8, 8, 12, 12])
    r = 4
    for d in decomposed:
        comp_wt  = round(d["flange"]["unit_wt"] * 2 + d["web"]["unit_wt"], 1)
        comp_tot = round(comp_wt * d["qty"], 1)
        fl_sn = d["flange"].get("serial","")
        wb_sn = d["web"].get("serial","")
        wf_str = f"{wb_sn}W / {fl_sn}F" if (wb_sn or fl_sn) else ""
        # 構件主行
        for col in range(1, 12):
            data_cell(ws2, r, col, "", bg="D6EAF8")
        data_cell(ws2, r, 1, d["comp"],    bold=True, bg="D6EAF8", align="center")
        data_cell(ws2, r, 6, d["spec"],    bold=True, bg="D6EAF8", align="left")
        data_cell(ws2, r, 7, d["length"],  bold=True, bg="D6EAF8")
        data_cell(ws2, r, 8, d["qty"],     bold=True, bg="D6EAF8")
        data_cell(ws2, r, 9, d["mat"],     bold=True, bg="D6EAF8")
        data_cell(ws2, r, 10, comp_wt,     bold=True, bg="D6EAF8")
        data_cell(ws2, r, 11, comp_tot,    bold=True, bg="D6EAF8")
        r += 1
        # 翼板
        fl = d["flange"]
        fl_spec = f"PL{fl['thick']}×{fl['width']}×{fl['length']}"
        fl_sn_wf = f"{fl_sn}F" if fl_sn else ""
        data_cell(ws2, r, 1, d["comp"],    align="center")
        data_cell(ws2, r, 2, d.get("part",""), align="center")
        data_cell(ws2, r, 3, fl["name"],   align="left")
        data_cell(ws2, r, 4, fl_sn,        align="center")
        data_cell(ws2, r, 5, fl_sn_wf,     align="center")
        data_cell(ws2, r, 6, fl_spec,      align="left")
        data_cell(ws2, r, 7, fl["length"])
        data_cell(ws2, r, 8, fl["total"])
        data_cell(ws2, r, 9, fl["mat"])
        data_cell(ws2, r, 10, fl["unit_wt"])
        data_cell(ws2, r, 11, fl["total_wt"])
        r += 1
        # 腹板
        wb_ = d["web"]
        wb_spec = f"PL{wb_['thick']}×{wb_['width']}×{wb_['length']}"
        wb_sn_wf = f"{wb_sn}W" if wb_sn else ""
        data_cell(ws2, r, 1, d["comp"],    align="center")
        data_cell(ws2, r, 2, d.get("part",""), align="center")
        data_cell(ws2, r, 3, wb_["name"],  align="left")
        data_cell(ws2, r, 4, wb_sn,        align="center")
        data_cell(ws2, r, 5, wb_sn_wf,     align="center")
        data_cell(ws2, r, 6, wb_spec,      align="left")
        data_cell(ws2, r, 7, wb_["length"])
        data_cell(ws2, r, 8, wb_["total"])
        data_cell(ws2, r, 9, wb_["mat"])
        data_cell(ws2, r, 10, wb_["unit_wt"])
        data_cell(ws2, r, 11, wb_["total_wt"])
        r += 1

    # ── 3. BH合板清單 ────────────────────────────────────────────
    ws_combine = add_ws("BH合板清單")
    cols_c = ["構件編號", "流水號(W/F)", "斷面規格", "長度(mm)", "數量"]
    write_title(ws_combine, "BH合板清單", 5, prod)
    write_param(ws_combine, 5)
    write_header(ws_combine, 3, cols_c)
    set_col_widths(ws_combine, [12, 16, 22, 12, 8])
    r_c = 4
    for d in decomposed:
        fl    = d["flange"]
        wb_   = d["web"]
        fl_sn = fl.get("serial","")
        wb_sn = wb_.get("serial","")
        fl_wf = f"{fl_sn}F" if fl_sn else ""
        wb_wf = f"{wb_sn}W" if wb_sn else ""
        bg_f  = "EBF4FF"
        bg_w  = "F7FAFC"
        data_cell(ws_combine, r_c, 1, d["comp"], align="center", bg=bg_f)
        data_cell(ws_combine, r_c, 2, fl_wf,     align="center", bg=bg_f)
        data_cell(ws_combine, r_c, 3, f'PL{fl["thick"]}×{fl["width"]}×{fl["length"]}', align="left", bg=bg_f)
        data_cell(ws_combine, r_c, 4, fl["length"], bg=bg_f)
        data_cell(ws_combine, r_c, 5, fl["total"],  bg=bg_f)
        r_c += 1
        data_cell(ws_combine, r_c, 1, d["comp"], align="center", bg=bg_w)
        data_cell(ws_combine, r_c, 2, wb_wf,     align="center", bg=bg_w)
        data_cell(ws_combine, r_c, 3, f'PL{wb_["thick"]}×{wb_["width"]}×{wb_["length"]}', align="left", bg=bg_w)
        data_cell(ws_combine, r_c, 4, wb_["length"], bg=bg_w)
        data_cell(ws_combine, r_c, 5, wb_["total"],  bg=bg_w)
        r_c += 1

    # ── 4. 板厚清單 ──────────────────────────────────────────────
    ws_thick = add_ws("板厚清單")
    cols_t = ["流水號(W/F)", "斷面規格", "數量", "材質", "單重量(kg)", "總重量(kg)"]
    write_title(ws_thick, "板厚清單", 6, prod)
    write_param(ws_thick, 6)
    write_header(ws_thick, 3, cols_t)
    set_col_widths(ws_thick, [16, 22, 8, 8, 14, 14])
    from collections import OrderedDict
    thick_groups = OrderedDict()
    for d in decomposed:
        for part, suffix in [(d["flange"], "F"), (d["web"], "W")]:
            sn    = part.get("serial","")
            sn_wf = f"{sn}{suffix}" if sn else ""
            spec  = f'PL{part["thick"]}×{part["width"]}×{part["length"]}'
            key   = (sn_wf, spec, part["mat"], part["unit_wt"])
            if key not in thick_groups:
                thick_groups[key] = {"qty": 0, "total_wt": 0.0}
            thick_groups[key]["qty"]      += part["total"]
            thick_groups[key]["total_wt"] += part["total_wt"]
    r_t = 4
    for i, ((sn_wf, spec, mat, unit_wt), val) in enumerate(thick_groups.items()):
        bg = "F7FAFC" if i % 2 == 0 else "EBF4FF"
        data_cell(ws_thick, r_t, 1, sn_wf,                  align="center", bg=bg)
        data_cell(ws_thick, r_t, 2, spec,                    align="left",   bg=bg)
        data_cell(ws_thick, r_t, 3, val["qty"],              bg=bg)
        data_cell(ws_thick, r_t, 4, mat,                     bg=bg)
        data_cell(ws_thick, r_t, 5, round(unit_wt, 1),       bg=bg)
        data_cell(ws_thick, r_t, 6, round(val["total_wt"],1), bg=bg)
        r_t += 1

    # ── 5. 採購清單 ──────────────────────────────────────────────
    ws3 = add_ws("採購清單")
    cols3 = ["NO", "採購規格(mm)", "採購數量", "材質", "單片重量(kg)", "總重量(kg)", "備註"]
    write_title(ws3, "採購清單　BH型鋼鋼板配料", 7, prod)
    write_param(ws3, 7)
    write_header(ws3, 3, cols3)
    set_col_widths(ws3, [6, 24, 10, 8, 14, 14, 16])

    # 相同規格合計
    from collections import OrderedDict
    merged = OrderedDict()
    for p in purchase_list:
        key = (p["spec"], p["mat"], p["unit_wt"])
        if key not in merged:
            merged[key] = {"qty": 0, "total_wt": 0}
        merged[key]["qty"]      += p["qty"]
        merged[key]["total_wt"] += p["total_wt"]

    total_wt  = 0
    total_qty = 0
    for i, ((spec, mat, unit_wt), val) in enumerate(merged.items()):
        r = i + 4
        bg = "F7FAFC" if i % 2 == 0 else "EBF4FF"
        data_cell(ws3, r, 1, i+1,             bg=bg)
        data_cell(ws3, r, 2, spec,            align="left", bg=bg)
        data_cell(ws3, r, 3, val["qty"],      bg=bg)
        data_cell(ws3, r, 4, mat,             bg=bg)
        data_cell(ws3, r, 5, unit_wt,         bg=bg)
        data_cell(ws3, r, 6, round(val["total_wt"], 0), bg=bg)
        total_wt  += val["total_wt"]
        total_qty += val["qty"]

    # 餘料使用備註
    scrap_used = [s for s in existing_scraps if s["id"] in used_scrap_ids]
    if scrap_used:
        r_note = len(merged) + 4
        note_txt = "♻ 餘料使用：" + "　".join(
            f"既{i+1:02d} PL{s['thick']}×{s['width']}×{s['length']}×{s['qty']}片"
            for i, s in enumerate(scrap_used))
        ws3.merge_cells(start_row=r_note, start_column=1, end_row=r_note, end_column=7)
        c = ws3.cell(r_note, 1, note_txt)
        c.font = Font(size=9, color="276749")
        c.fill = scrap_fill()
        r_note += 1
    else:
        r_note = len(merged) + 4
    ws3.merge_cells(start_row=r_note, start_column=1, end_row=r_note, end_column=5)
    c = ws3.cell(r_note, 1, f"合計：共 {total_qty} 片（不含餘料）")
    c.font = Font(bold=True, size=10)
    ws3.cell(r_note, 6, round(total_wt, 0)).font = Font(bold=True, size=10)

    # ── 4. 切割明細 ──────────────────────────────────────────────
    ws4 = add_ws("切割明細")
    cols4 = ["片次","板別","採購規格(mm)","材質","板重(kg)",
             "構件編號（裁切尺寸mm）","數量","單板重(kg)","餘料規格(mm)/重量(kg)"]
    write_title(ws4, "切割明細　BH型鋼鋼板配料", 9, prod)
    write_param(ws4, 9)
    write_header(ws4, 3, cols4)
    set_col_widths(ws4, [6,10,22,8,10,28,6,10,24])
    for i, ct in enumerate(cut_details):
        r = i + 4
        bg = "C6EFCE" if ct["is_scrap"] else ("F7FAFC" if i%2==0 else "EBF4FF")
        for c_, v in enumerate([
            ct["idx"],
            ("F=翼鈑" if ct["type"]=="F" else "W=腹鈑"),
            ct["board_spec"], ct["mat"], ct["board_wt"],
            f"{ct['comp']}（{ct['part_spec']}）",
            ct["qty"], ct["unit_wt"], ct["leftover"]
        ], 1):
            data_cell(ws4, r, c_, v, align="left" if c_ in (3,6,9) else "center", bg=bg)
    # 圖例
    legend_r = len(cut_details) + 4
    ws4.merge_cells(start_row=legend_r, start_column=1, end_row=legend_r, end_column=9)
    c = ws4.cell(legend_r, 1, "F=翼鈑　W=腹鈑　*=因板重限制加大　♻=使用現有餘料（綠色底）")
    c.font = Font(size=9, italic=True, color="666666")

    # ── 5. 餘料清單 ──────────────────────────────────────────────
    ws5 = add_ws("餘料清單")
    cols5 = ["餘NO","來源片次","板別","餘料規格(mm)","材質","重量(kg)"]
    write_title(ws5, "餘料清單　BH型鋼鋼板配料", 6, prod)
    write_param(ws5, 6)
    write_header(ws5, 3, cols5)
    set_col_widths(ws5, [8,10,10,24,8,12])
    r = 4
    for i, sc in enumerate(new_scraps):
        bg = "F7FAFC" if i%2==0 else "EBF4FF"
        data_cell(ws5, r, 1, f"餘{i+1:02d}", bg=bg)
        data_cell(ws5, r, 2, sc["src"],  bg=bg)
        data_cell(ws5, r, 3, f"{'F=翼鈑' if sc['type']=='F' else 'W=腹鈑'}", bg=bg)
        data_cell(ws5, r, 4, sc["spec"], align="left", bg=bg)
        data_cell(ws5, r, 5, sc["mat"],  bg=bg)
        data_cell(ws5, r, 6, sc["wt"],   bg=bg)
        r += 1
    # 未使用既有餘料
    unused = [s for s in existing_scraps if s["id"] not in used_scrap_ids]
    for s in unused:
        bg = "FFF3CD"
        data_cell(ws5, r, 1, f"*既{s['id']}", bg=bg)
        data_cell(ws5, r, 2, "—", bg=bg)
        data_cell(ws5, r, 3, "—", bg=bg)
        data_cell(ws5, r, 4, f"PL{s['thick']}×{s['width']}×{s['length']}", align="left", bg=bg)
        data_cell(ws5, r, 5, s["mat"], bg=bg)
        wt = round(calc_weight(s["width"], s["thick"], s["length"], density))
        data_cell(ws5, r, 6, wt, bg=bg)
        r += 1
    total_sc = sum(sc["wt"] for sc in new_scraps)
    ws5.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    c = ws5.cell(r, 1, f"餘料合計：共 {len(new_scraps)} 件")
    c.font = Font(bold=True)
    ws5.cell(r, 6, total_sc).font = Font(bold=True)
    r += 1
    ws5.cell(r, 1, "* = 現有餘料未使用").font = Font(size=9, italic=True, color="666666")

    wb.save(path)


# ═══════════════════════════════════════════════════════════════════════
# HTML / PDF 輸出
# ═══════════════════════════════════════════════════════════════════════

def write_html(path, proj_no, proj_name, date_str, mat_name, density,
               new_kerf, new_trim, scrap_kerf, scrap_trim,
               bh_rows, decomposed, purchase_list, cut_details,
               new_scraps, existing_scraps, used_scrap_ids):

    prod = datetime.now().strftime("%Y/%m/%d")
    param_line = (f"工程：{proj_no}　{proj_name}　材質：{mat_name}　比重：{density}"
                  f"　新板每刀：{new_kerf}mm 頭尾：{new_trim}mm"
                  f"　餘料每刀：{scrap_kerf}mm 頭尾：{scrap_trim}mm")

    def tbl_start(title, headers):
        ths = "".join(f"<th>{h}</th>" for h in headers)
        return (f'<div class="section">'
                f'<div class="tbl-title"><span>{title}</span>'
                f'<span class="date">產出日期：{prod}</span></div>'
                f'<div class="param">{param_line}</div>'
                f'<table><thead><tr>{ths}</tr></thead><tbody>')

    def tbl_end(legend=""):
        leg = f'<tr><td colspan="99" class="legend">{legend}</td></tr>' if legend else ""
        return f'{leg}</tbody></table></div>'

    def tr(cells, cls=""):
        tds = "".join(f"<td>{c}</td>" for c in cells)
        return f'<tr class="{cls}">{tds}</tr>'

    # ── 1. 配料清單 ──
    s1 = tbl_start("配料清單　BH配料報表",
                    ["構件編號","零件編號","斷面規格","長度(mm)","數量","材質"])
    for i, row in enumerate(bh_rows):
        s1 += tr([row["comp"], row["part"], row["spec"],
                  row["length"], row["qty"], row["mat"]],
                 "odd" if i%2==0 else "even")
    s1 += tbl_end()

    # ── 2. BH拆板明細 ──
    s2 = tbl_start("BH拆板明細",
                    ["構件","流水號","拆板零件","斷面規格","長度(mm)","數量","材質","單重(kg)","總重(kg)"])
    for d in decomposed:
        comp_wt  = round(d["flange"]["unit_wt"]*2 + d["web"]["unit_wt"], 1)
        comp_tot = round(comp_wt * d["qty"], 1)
        s2 += tr([d["comp"], "", "", d["spec"], d["length"], d["qty"],
                  d["mat"], comp_wt, comp_tot], "comp-row")
        fl = d["flange"]
        s2 += tr(["", fl.get("serial",""), fl["name"],
                  f'PL{fl["thick"]}×{fl["width"]}×{fl["length"]}',
                  fl["length"], fl["total"], fl["mat"], fl["unit_wt"], fl["total_wt"]])
        wb_ = d["web"]
        s2 += tr(["", wb_.get("serial",""), wb_["name"],
                  f'PL{wb_["thick"]}×{wb_["width"]}×{wb_["length"]}',
                  wb_["length"], wb_["total"], wb_["mat"], wb_["unit_wt"], wb_["total_wt"]])
    s2 += tbl_end()

    # ── 3. 採購清單 ──
    s3 = tbl_start("採購清單　BH型鋼鋼板配料",
                    ["NO","採購規格(mm)","採購數量","材質","單片重量(kg)","總重量(kg)","備註"])
    total_wt = 0
    for i, p in enumerate(purchase_list):
        s3 += tr([i+1, p["spec"], p["qty"], p["mat"],
                  p["unit_wt"], p["total_wt"], ""], "odd" if i%2==0 else "even")
        total_wt += p["total_wt"]
    scrap_used = [s for s in existing_scraps if s["id"] in used_scrap_ids]
    if scrap_used:
        note = "♻ 餘料使用：" + "　".join(
            f'既{i+1:02d} PL{s["thick"]}×{s["width"]}×{s["length"]}×{s["qty"]}片'
            for i, s in enumerate(scrap_used))
        s3 += f'<tr class="scrap-note"><td colspan="7">{note}</td></tr>'
    s3 += f'<tr class="total-row"><td colspan="5">合計：共 {len(purchase_list)} 片（不含餘料）</td><td>{total_wt}</td><td></td></tr>'
    s3 += tbl_end()

    # ── 4. 切割明細 ──
    s4 = tbl_start("切割明細　BH型鋼鋼板配料",
                    ["片次","板別","採購規格(mm)","材質","板重(kg)",
                     "構件編號（裁切尺寸mm）","數量","單板重(kg)","餘料規格/重量"])
    for ct in cut_details:
        cls = "scrap-row" if ct["is_scrap"] else ""
        s4 += tr([ct["idx"],
                  "F=翼鈑" if ct["type"]=="F" else "W=腹鈑",
                  ct["board_spec"], ct["mat"], ct["board_wt"],
                  f'{ct["comp"]}（{ct["part_spec"]}）',
                  ct["qty"], ct["unit_wt"], ct["leftover"]], cls)
    s4 += tbl_end("F=翼鈑　W=腹鈑　*=因板重限制加大　♻=使用現有餘料（綠色底）")

    # ── 5. 餘料清單 ──
    s5 = tbl_start("餘料清單　BH型鋼鋼板配料",
                    ["餘NO","來源片次","板別","餘料規格(mm)","材質","重量(kg)"])
    for i, sc in enumerate(new_scraps):
        s5 += tr([f"餘{i+1:02d}", sc["src"],
                  "F=翼鈑" if sc["type"]=="F" else "W=腹鈑",
                  sc["spec"], sc["mat"], sc["wt"]],
                 "odd" if i%2==0 else "even")
    unused = [s for s in existing_scraps if s["id"] not in used_scrap_ids]
    for s in unused:
        wt = round(calc_weight(s["width"], s["thick"], s["length"], density))
        s5 += tr([f'*{s["id"]}', "—", "—",
                  f'PL{s["thick"]}×{s["width"]}×{s["length"]}',
                  s["mat"], wt], "unused-scrap")
    total_sc = sum(sc["wt"] for sc in new_scraps)
    s5 += f'<tr class="total-row"><td colspan="5">餘料合計：共 {len(new_scraps)} 件</td><td>{total_sc}</td></tr>'
    s5 += tbl_end("* = 現有餘料未使用（黃色底）")

    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<title>BH配料報表　{proj_no} {proj_name}</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: "Microsoft JhengHei", "微軟正黑體", sans-serif;
          font-size: 11px; color: #1a1a1a; background: #fff; padding: 16px; }}
  .section {{ margin-bottom: 28px; page-break-inside: avoid; }}
  .tbl-title {{ background: #1E3A5F; color: #fff; font-size: 13px; font-weight: bold;
                padding: 6px 10px; display: flex; justify-content: space-between; }}
  .date {{ font-size: 10px; font-weight: normal; }}
  .param {{ background: #FFF5F5; color: #C53030; font-size: 9px;
            padding: 4px 10px; border: 1px solid #FED7D7; }}
  table {{ width: 100%; border-collapse: collapse; }}
  th {{ background: #D0D8E4; color: #1E3A5F; font-weight: bold;
        border: 1px solid #999; padding: 5px 6px; text-align: center; }}
  td {{ border: 1px solid #bbb; padding: 4px 6px; text-align: center; }}
  .odd  {{ background: #F7FAFC; }}
  .even {{ background: #EBF4FF; }}
  .comp-row {{ background: #D6EAF8; font-weight: bold; }}
  .scrap-row {{ background: #C6EFCE; }}
  .scrap-note {{ background: #C6EFCE; color: #276749; }}
  .unused-scrap {{ background: #FFF3CD; }}
  .total-row {{ background: #EDF2F7; font-weight: bold; }}
  .legend {{ font-size: 9px; color: #666; font-style: italic;
             text-align: left; padding: 4px 8px; }}
  @media print {{
    body {{ padding: 0; }}
    .section {{ page-break-after: always; }}
    .section:last-child {{ page-break-after: avoid; }}
  }}
</style>
</head>
<body>
{s1}
{s2}
{s3}
{s4}
{s5}
<p style="font-size:9px;color:#999;margin-top:16px;text-align:right;">
  BH型鋼鋼板配料系統　產出日期：{prod}
</p>
</body>
</html>"""

    with open(path, "w", encoding="utf-8") as f:
        f.write(html)


# ═══════════════════════════════════════════════════════════════════════
# 排列圖 PDF 輸出（仿照片格式：兩欄方框）
# ═══════════════════════════════════════════════════════════════════════

def leftover_specs(left):
    """切割明細餘料文字 → 規格清單（去掉重量）"""
    if not left or left == "無餘料":
        return []
    return [p.split("（")[0].strip() for p in left.split("　") if p.split("（")[0].strip()]


def write_layout_pdf(path, proj_no, proj_name, mat_name, cut_details, new_scraps,
                     serial_map=None, modified_specs=None):
    """
    產生鐵板裁切排列圖 PDF。
    每張板用方框表示，兩欄排列，框內顯示裁切內容。
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont

    try:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        FONT  = "STSong-Light"
        FONTB = "STSong-Light"
    except Exception:
        FONT  = "Helvetica"
        FONTB = "Helvetica-Bold"

    FS      = 12         # 內文字型大小
    FS_HDR  = 13         # 標題字型大小
    FS_SUB  = 10         # 副標題（材質、頁碼）

    prod    = datetime.now().strftime("%Y/%m/%d")
    W, H    = A4
    ML, MR  = 14*mm, 14*mm
    MT      = 24*mm      # 頂部留給兩行標題
    MB      = 12*mm
    GAP     = 5*mm       # 兩欄間距
    COL_W   = (W - ML - MR - GAP) / 2
    LINE_H  = (FS + 3) * 0.352778 * mm   # 行高（pt→mm）
    BOX_H   = LINE_H * 7 + 10*mm         # 框高：約7行 + 上下留白

    c = rl_canvas.Canvas(path, pagesize=A4)

    def draw_header(page_no, mat=""):
        """每頁頂部兩行"""
        # 第一行：材質（左，紅色加粗）、標題（中）、日期（右）
        y1 = H - 10*mm
        # 材質標示（左側，紅色加粗）
        c.setFont(FONTB, FS_SUB + 1)
        c.setFillColor(colors.HexColor("#C53030"))
        c.drawString(ML, y1, f"材質:{mat}" if mat else "")
        c.setFillColor(colors.black)
        # 標題（中央）
        title = f"{proj_no}-BH拆版"
        c.setFont(FONTB, FS_HDR)
        tw = c.stringWidth(title, FONTB, FS_HDR)
        c.drawString((W - tw) / 2, y1, title)
        # 日期（右，移除頁碼）
        c.setFont(FONT, FS_SUB)
        c.drawRightString(W - MR, y1, prod)
        # 第二行：工程名稱（左）
        y2 = H - 17*mm
        c.setFont(FONT, FS_SUB)
        c.drawString(ML, y2, f"{proj_no} {proj_name}")
        c.setStrokeColor(colors.black)

    def draw_footer(page_no):
        """每頁底部中央：頁次"""
        c.setFont(FONT, FS_SUB)
        label = f"第 {page_no} 頁"
        lw = c.stringWidth(label, FONT, FS_SUB)
        c.drawString((W - lw) / 2, MB - 4*mm, label)

    def draw_box(x, y, ct):
        bw = COL_W
        bh = BOX_H

        # 判斷此板是否有被修改（板寬或板長有變更）
        import re as _re2
        mb = _re2.match(r"PL(\d+)×(\d+)×(\d+)", ct["board_spec"])
        is_modified = False
        if mb and modified_specs:
            key = (int(mb.group(1)), int(mb.group(2)), int(mb.group(3)), ct.get("mat"))
            is_modified = key in modified_specs

        TEXT_COLOR  = colors.HexColor("#1A6B1A") if is_modified else colors.black
        TITLE_COLOR = colors.HexColor("#1A6B1A") if is_modified else colors.black

        # 框外標題：① PL 厚×寬×長 × N PC
        qty_label = ct.get("_qty", 1)
        c.setFont(FONTB, FS)
        c.setFillColor(TITLE_COLOR)
        c.drawString(x, y + 1.5*mm, f"{ct['idx']} {ct['board_spec']} × {qty_label}  PC")
        c.setFillColor(colors.black)

        # 方框（修改過用綠色框線）
        c.setStrokeColor(colors.HexColor("#1A6B1A") if is_modified else colors.black)
        c.setLineWidth(1.2 if is_modified else 0.7)
        c.rect(x, y - bh, bw, bh)
        c.setStrokeColor(colors.black)
        c.setLineWidth(0.7)

        # 框內文字
        c.setFont(FONT, FS)
        ly     = y - LINE_H - 2*mm
        bottom = y - bh + 3*mm

        def write_line(text, red=False):
            nonlocal ly
            if ly < bottom:
                return
            if red:
                c.setFillColor(colors.HexColor("#C53030"))
            elif is_modified:
                c.setFillColor(colors.HexColor("#1A6B1A"))
            else:
                c.setFillColor(colors.black)
            c.drawString(x + 3*mm, ly, text)
            c.setFillColor(colors.black)
            ly -= LINE_H

        def write_leftovers(left):
            """列出全部餘料（每件一行）；框內放不下時最後一行改為「…另 N 件」"""
            lefts = leftover_specs(left)
            room  = max(0, int((ly - bottom) // LINE_H) + 1)
            shown = lefts[:max(0, room - 1)] if len(lefts) > room else lefts
            for i, sp in enumerate(shown):
                write_line(("　　　" if i else "餘料：") + sp, red=True)
            if len(shown) < len(lefts):
                write_line(("　　　" if shown else "餘料：")
                           + f"…另 {len(lefts) - len(shown)} 件（詳見餘料清單）", red=True)

        sm = serial_map or {}

        layout = ct.get("layout")
        if layout:
            thick = layout.get("thick", "")
            pw    = layout.get("part_w", "")
            # 收集所有零件，按流水號合計
            from collections import Counter
            all_parts = [(p["name"], p["length"]) for ri in layout.get("rows",[]) for p in ri.get("parts",[])]
            part_counter = Counter(all_parts)
            for (name, length), cnt in part_counter.items():
                sn = sm.get(name, name)
                write_line(f"{thick}*{pw}*{length}*{cnt}({sn})")
            write_leftovers(ct.get("leftover", ""))
        else:
            comps = ct.get("comp", "").split("、")
            ps    = ct.get("part_spec", "")
            import re as _re
            m = _re.match(r"PL(\d+)×(\d+)×(\d+)", ps)
            if m:
                thick, pw, pl = m.group(1), m.group(2), m.group(3)
                total_qty    = ct.get("qty", 1)
                unique_comps = list(dict.fromkeys(comps))
                qty_each     = max(1, total_qty // len(unique_comps)) if unique_comps else total_qty
                for comp in unique_comps:
                    sn = sm.get(comp, comp)
                    write_line(f"{thick}*{pw}*{pl}*{qty_each}({sn})")
            else:
                write_line(f"{ps} × {ct.get('qty','')}（{ct.get('comp','')}）")
            write_leftovers(ct.get("leftover", ""))

        sm = serial_map or {}

    # ── 排版（按材質分頁）─────────────────────────────────────────
    usable_h      = H - MT - MB
    rows_per_page = max(1, int(usable_h / (BOX_H + 6*mm)))

    # 按材質分組，保持原始順序
    from itertools import groupby
    groups = []
    for mat, items in groupby(cut_details, key=lambda ct: ct.get("mat", "")):
        groups.append((mat, list(items)))

    page_no     = 1
    cur_mat     = None

    for mat, cts in groups:
        # 每換一種材質 → 新頁
        if cur_mat is not None:
            x_end = ML + col * (COL_W + GAP)
            y_end = H - MT - row_i * (BOX_H + 6*mm) - BOX_H - 4*mm
            c.setFont(FONT, FS_SUB)
            c.drawString(x_end, y_end, "以下空白")
            draw_footer(page_no)
            c.showPage()
            page_no += 1

        cur_mat = mat
        col     = 0
        row_i   = 0
        draw_header(page_no, mat)

        # 合併完全相同的板次（board_spec + comp + leftover 相同則合併，數量加總）
        from collections import OrderedDict
        merged_cts = []
        seen = OrderedDict()
        for ct in cts:
            key = (ct["board_spec"], ct.get("comp",""), ct.get("leftover",""))
            if key in seen:
                seen[key]["_qty"] = seen[key].get("_qty", 1) + 1
            else:
                import copy as _copy
                ct2 = _copy.copy(ct)
                ct2["_qty"] = 1
                seen[key] = ct2
                merged_cts.append(ct2)

        for ct in merged_cts:
            x = ML + col * (COL_W + GAP)
            y = H - MT - row_i * (BOX_H + 6*mm)

            draw_box(x, y, ct)

            col += 1
            if col >= 2:
                col = 0
                row_i += 1
            if row_i >= rows_per_page:
                row_i = 0
                col   = 0
                draw_footer(page_no)
                c.showPage()
                page_no += 1
                draw_header(page_no, mat)

    # 最後一頁「以下空白」
    x_end = ML + col * (COL_W + GAP)
    y_end = H - MT - row_i * (BOX_H + 6*mm) - BOX_H - 4*mm
    c.setFont(FONT, FS_SUB)
    c.drawString(x_end, y_end, "以下空白")
    draw_footer(page_no)

    c.save()


# ═══════════════════════════════════════════════════════════════════════
# PDF 輸出（reportlab）
# ═══════════════════════════════════════════════════════════════════════

def write_pdf(path, proj_no, proj_name, date_str, mat_name, density,
              new_kerf, new_trim, scrap_kerf, scrap_trim,
              bh_rows, decomposed, purchase_list, cut_details,
              new_scraps, existing_scraps, used_scrap_ids):

    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer, PageBreak)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont

    # 註冊中文字型
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        FONT = "STSong-Light"
    except Exception:
        FONT = "Helvetica"

    prod     = datetime.now().strftime("%Y/%m/%d")
    W, H     = A4
    margin   = 12 * mm
    doc      = SimpleDocTemplate(path, pagesize=A4,
                                 leftMargin=margin, rightMargin=margin,
                                 topMargin=margin, bottomMargin=margin)
    styles   = getSampleStyleSheet()
    param_line = (f"工程：{proj_no}　{proj_name}　材質：{mat_name}　比重：{density}"
                  f"　新板每刀：{new_kerf}mm 頭尾：{new_trim}mm"
                  f"　餘料每刀：{scrap_kerf}mm 頭尾：{scrap_trim}mm")

    N_STYLE  = ParagraphStyle("n", fontName=FONT, fontSize=8, leading=10)
    T_STYLE  = ParagraphStyle("t", fontName=FONT, fontSize=11, leading=14, textColor=colors.white)
    P_STYLE  = ParagraphStyle("p", fontName=FONT, fontSize=7, leading=9,
                               textColor=colors.HexColor("#C53030"))

    COL_HDR  = colors.HexColor("#1A3050")
    COL_ODD  = colors.HexColor("#F7FAFC")
    COL_EVEN = colors.HexColor("#EBF4FF")
    COL_COMP = colors.HexColor("#D6EAF8")
    COL_SCRP = colors.HexColor("#C6EFCE")
    COL_UNUS = colors.HexColor("#FFF3CD")
    COL_TOT  = colors.HexColor("#EDF2F7")

    def cell(txt):
        return Paragraph(str(txt), N_STYLE)

    def tbl_style(data, col_widths, scrap_rows=None, comp_rows=None, unused_rows=None):
        n = len(data)
        style = [
            ("FONTNAME",    (0,0), (-1,-1), FONT),
            ("FONTSIZE",    (0,0), (-1,-1), 7),
            ("FONTNAME",    (0,0), (-1,0),  FONT),
            ("FONTSIZE",    (0,0), (-1,0),  8),
            ("FONTNAME",    (0,0), (-1,0),  FONT),
            ("BACKGROUND",  (0,0), (-1,0),  COL_HDR),
            ("TEXTCOLOR",   (0,0), (-1,0),  colors.white),
            ("ALIGN",       (0,0), (-1,-1), "CENTER"),
            ("VALIGN",      (0,0), (-1,-1), "MIDDLE"),
            ("GRID",        (0,0), (-1,-1), 0.4, colors.grey),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [COL_ODD, COL_EVEN]),
            ("TOPPADDING",  (0,0), (-1,-1), 2),
            ("BOTTOMPADDING",(0,0),(-1,-1), 2),
        ]
        if scrap_rows:
            for r in scrap_rows:
                style.append(("BACKGROUND", (0,r), (-1,r), COL_SCRP))
        if comp_rows:
            for r in comp_rows:
                style.append(("BACKGROUND", (0,r), (-1,r), COL_COMP))
                style.append(("FONTNAME",   (0,r), (-1,r), FONT))
        if unused_rows:
            for r in unused_rows:
                style.append(("BACKGROUND", (0,r), (-1,r), COL_UNUS))
        return TableStyle(style)

    def title_block(title):
        return Table(
            [[Paragraph(title, T_STYLE),
              Paragraph(f"產出日期：{prod}", P_STYLE)]],
            colWidths=[W - 2*margin - 60*mm, 60*mm],
            style=TableStyle([
                ("BACKGROUND", (0,0), (-1,-1), COL_HDR),
                ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
                ("TOPPADDING", (0,0), (-1,-1), 4),
                ("BOTTOMPADDING",(0,0),(-1,-1), 4),
            ])
        )

    def param_block():
        return Table(
            [[Paragraph(param_line, P_STYLE)]],
            colWidths=[W - 2*margin],
            style=TableStyle([
                ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#FFF5F5")),
                ("TOPPADDING", (0,0), (-1,-1), 3),
                ("BOTTOMPADDING",(0,0),(-1,-1), 3),
            ])
        )

    story = []
    avail_w = W - 2 * margin

    # ── 1. 配料清單 ──────────────────────────────────────────────
    story.append(title_block("配料清單　BH配料報表"))
    story.append(param_block())
    story.append(Spacer(1, 2*mm))
    cw1 = [avail_w*r for r in [0.12,0.12,0.32,0.14,0.10,0.10,0.10]]
    data1 = [["構件編號","零件編號","斷面規格","長度(mm)","數量","材質","備註"]]
    for row in bh_rows:
        data1.append([cell(row["comp"]), cell(row["part"]), cell(row["spec"]),
                      cell(row["length"]), cell(row["qty"]), cell(row["mat"]), cell("")])
    t1 = Table(data1, colWidths=cw1, repeatRows=1)
    t1.setStyle(tbl_style(data1, cw1))
    story.append(t1)
    story.append(PageBreak())

    # ── 2. BH拆板明細 ────────────────────────────────────────────
    story.append(title_block("BH拆板明細"))
    story.append(param_block())
    story.append(Spacer(1, 2*mm))
    cw2 = [avail_w*r for r in [0.09,0.10,0.12,0.26,0.11,0.08,0.08,0.08,0.08]]
    data2 = [["構件","流水號","拆板零件","斷面規格","長度(mm)","數量","材質","單重(kg)","總重(kg)"]]
    comp_rows2 = []
    for d in decomposed:
        comp_wt  = round(d["flange"]["unit_wt"]*2 + d["web"]["unit_wt"], 1)
        comp_tot = round(comp_wt * d["qty"], 1)
        comp_rows2.append(len(data2))
        data2.append([cell(d["comp"]), cell(""), cell(""), cell(d["spec"]),
                      cell(d["length"]), cell(d["qty"]), cell(d["mat"]),
                      cell(comp_wt), cell(comp_tot)])
        fl = d["flange"]
        data2.append([cell(""), cell(fl.get("serial","")), cell(fl["name"]),
                      cell(f'PL{fl["thick"]}×{fl["width"]}×{fl["length"]}'),
                      cell(fl["length"]), cell(fl["total"]), cell(fl["mat"]),
                      cell(fl["unit_wt"]), cell(fl["total_wt"])])
        wb_ = d["web"]
        data2.append([cell(""), cell(wb_.get("serial","")), cell(wb_["name"]),
                      cell(f'PL{wb_["thick"]}×{wb_["width"]}×{wb_["length"]}'),
                      cell(wb_["length"]), cell(wb_["total"]), cell(wb_["mat"]),
                      cell(wb_["unit_wt"]), cell(wb_["total_wt"])])
    t2 = Table(data2, colWidths=cw2, repeatRows=1)
    t2.setStyle(tbl_style(data2, cw2, comp_rows=comp_rows2))
    story.append(t2)
    story.append(PageBreak())

    # ── 3. 採購清單 ──────────────────────────────────────────────
    story.append(title_block("採購清單　BH型鋼鋼板配料"))
    story.append(param_block())
    story.append(Spacer(1, 2*mm))
    cw3 = [avail_w*r for r in [0.06,0.30,0.10,0.10,0.16,0.16,0.12]]
    data3 = [["NO","採購規格(mm)","數量","材質","單片重(kg)","總重(kg)","備註"]]
    total_wt = 0
    for i, p in enumerate(purchase_list):
        data3.append([cell(i+1), cell(p["spec"]), cell(p["qty"]),
                      cell(p["mat"]), cell(p["unit_wt"]), cell(p["total_wt"]), cell("")])
        total_wt += p["total_wt"]
    data3.append([cell(""), cell(f"合計：共 {len(purchase_list)} 片"), cell(""),
                  cell(""), cell(""), cell(total_wt), cell("")])
    t3 = Table(data3, colWidths=cw3, repeatRows=1)
    st3 = tbl_style(data3, cw3)
    st3.add("BACKGROUND", (0, len(data3)-1), (-1, len(data3)-1), COL_TOT)
    t3.setStyle(st3)
    story.append(t3)
    story.append(PageBreak())

    # ── 4. 切割明細 ──────────────────────────────────────────────
    story.append(title_block("切割明細　BH型鋼鋼板配料"))
    story.append(param_block())
    story.append(Spacer(1, 2*mm))
    cw4 = [avail_w*r for r in [0.06,0.08,0.20,0.07,0.09,0.22,0.07,0.09,0.12]]
    data4 = [["片次","板別","採購規格(mm)","材質","板重(kg)",
               "構件編號（裁切尺寸）","數量","單重(kg)","餘料規格/重量"]]
    scrap_rows4 = []
    for ct in cut_details:
        if ct["is_scrap"]:
            scrap_rows4.append(len(data4))
        data4.append([cell(ct["idx"]),
                      cell("F=翼" if ct["type"]=="F" else "W=腹"),
                      cell(ct["board_spec"]), cell(ct["mat"]), cell(ct["board_wt"]),
                      cell(f'{ct["comp"]}（{ct["part_spec"]}）'),
                      cell(ct["qty"]), cell(ct["unit_wt"]), cell(ct["leftover"])])
    data4.append([cell(""), cell("F=翼鈑　W=腹鈑　*=板重限制加大　♻=使用餘料（綠底）"),
                  cell(""),cell(""),cell(""),cell(""),cell(""),cell(""),cell("")])
    t4 = Table(data4, colWidths=cw4, repeatRows=1)
    t4.setStyle(tbl_style(data4, cw4, scrap_rows=scrap_rows4))
    story.append(t4)
    story.append(PageBreak())

    # ── 5. 餘料清單 ──────────────────────────────────────────────
    story.append(title_block("餘料清單　BH型鋼鋼板配料"))
    story.append(param_block())
    story.append(Spacer(1, 2*mm))
    cw5 = [avail_w*r for r in [0.10,0.12,0.12,0.32,0.12,0.12,0.10]]
    data5 = [["餘NO","來源片次","板別","餘料規格(mm)","材質","重量(kg)","備註"]]
    unused_rows5 = []
    for i, sc in enumerate(new_scraps):
        data5.append([cell(f"餘{i+1:02d}"), cell(sc["src"]),
                      cell("F=翼" if sc["type"]=="F" else "W=腹"),
                      cell(sc["spec"]), cell(sc["mat"]), cell(sc["wt"]), cell("")])
    unused = [s for s in existing_scraps if s["id"] not in used_scrap_ids]
    for s in unused:
        wt = round(calc_weight(s["width"], s["thick"], s["length"], density))
        unused_rows5.append(len(data5))
        data5.append([cell(f'*{s["id"]}'), cell("—"), cell("—"),
                      cell(f'PL{s["thick"]}×{s["width"]}×{s["length"]}'),
                      cell(s["mat"]), cell(wt), cell("未使用")])
    total_sc = sum(sc["wt"] for sc in new_scraps)
    data5.append([cell(""), cell(f"合計：共 {len(new_scraps)} 件"),
                  cell(""),cell(""),cell(""),cell(total_sc),cell("")])
    data5.append([cell(""), cell("* = 現有餘料未使用（黃底）"),
                  cell(""),cell(""),cell(""),cell(""),cell("")])
    t5 = Table(data5, colWidths=cw5, repeatRows=1)
    st5 = tbl_style(data5, cw5, unused_rows=unused_rows5)
    st5.add("BACKGROUND", (0, len(data5)-2), (-1, len(data5)-2), COL_TOT)
    t5.setStyle(st5)
    story.append(t5)

    doc.build(story)


# ═══════════════════════════════════════════════════════════════════════
# 程式內預覽視窗
# ═══════════════════════════════════════════════════════════════════════

def assign_scrap_numbers(result):
    """
    統一賦予「餘料NO」並寫回每個餘料物件的 "_no" 欄位。
    任何需要在排列圖（LayoutWindow）顯示餘NO 的進入路徑
    （無論是先開「預覽結果」還是直接點「排列圖」按鈕），
    都必須在開圖之前呼叫這個函式一次，確保編號已經寫入。
    重複呼叫安全（每次都用相同規則重新編號，結果一致）。
    """
    for i, sc in enumerate(result["new_scraps"]):
        sc["_no"] = f"餘{i+1:02d}"


class TkDrawer:
    """繪圖後端：把通用繪圖指令轉發到 tk.Canvas（畫面顯示用）"""
    def __init__(self, canvas):
        self.canvas = canvas

    def rect(self, x0, y0, x1, y1, fill="", outline="", width=1, dash=None):
        kwargs = {"fill": fill, "outline": outline, "width": width}
        if dash: kwargs["dash"] = dash
        self.canvas.create_rectangle(x0, y0, x1, y1, **kwargs)

    def line(self, x0, y0, x1, y1, fill="#000", width=1, dash=None):
        kwargs = {"fill": fill, "width": width}
        if dash: kwargs["dash"] = dash
        self.canvas.create_line(x0, y0, x1, y1, **kwargs)

    def text(self, x, y, text, font_size=10, bold=False, fill="#000",
             anchor="center", angle=0):
        weight = "bold" if bold else "normal"
        self.canvas.create_text(x, y, text=text,
                                font=("Microsoft JhengHei", font_size, weight),
                                fill=fill, anchor=anchor, angle=angle, justify="center")


class PILDrawer:
    """
    繪圖後端：把通用繪圖指令轉發到 PIL ImageDraw（下載圖片用）。
    座標、顏色、文字內容與 TkDrawer 完全相同邏輯，
    確保下載出來的圖片跟畫面上看到的排列圖一致。
    """
    def __init__(self, draw, font_finder):
        self.draw = draw
        self.font_finder = font_finder   # 函式：(size, bold) -> PIL ImageFont 物件

    def rect(self, x0, y0, x1, y1, fill="", outline="", width=1, dash=None):
        # PIL 不支援虛線矩形，dash 樣式一律改畫實線框（顏色/位置仍正確，僅線型差異）
        kwargs = {}
        if fill: kwargs["fill"] = fill
        if outline:
            kwargs["outline"] = outline
            kwargs["width"] = width
        self.draw.rectangle([x0, y0, x1, y1], **kwargs)

    def line(self, x0, y0, x1, y1, fill="#000", width=1, dash=None):
        if dash:
            # 簡易虛線：依固定間隔切成多段短線
            import math as _m
            total = _m.hypot(x1-x0, y1-y0)
            if total > 0:
                seg, gap = dash[0], dash[1] if len(dash) > 1 else dash[0]
                step = seg + gap
                n = int(total / step) + 1
                ux, uy = (x1-x0)/total, (y1-y0)/total
                d = 0
                while d < total:
                    sx, sy = x0+ux*d, y0+uy*d
                    ex, ey = x0+ux*min(d+seg, total), y0+uy*min(d+seg, total)
                    self.draw.line([sx, sy, ex, ey], fill=fill, width=width)
                    d += step
            return
        self.draw.line([x0, y0, x1, y1], fill=fill, width=width)

    def text(self, x, y, text, font_size=10, bold=False, fill="#000",
             anchor="center", angle=0):
        font = self.font_finder(font_size, bold)
        if angle == 90:
            # 直書文字：用獨立小畫布畫橫排文字後旋轉貼回，PIL 原生不支援旋轉文字
            bbox = font.getbbox(text)
            tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
            from PIL import Image as _Image, ImageDraw as _ImageDraw
            txt_img = _Image.new("RGBA", (max(tw,1)+4, max(th,1)+4), (255,255,255,0))
            _ImageDraw.Draw(txt_img).text((2,2), text, font=font, fill=fill)
            txt_img = txt_img.rotate(90, expand=True)
            self.draw._image.paste(txt_img, (int(x-txt_img.width/2), int(y-txt_img.height/2)), txt_img)
            return
        # anchor 對應：tk 的 "center" → PIL 的 "mm"（水平垂直皆置中）；"e" → "rm"
        anchor_map = {"center": "mm", "e": "rm", "w": "lm"}
        pil_anchor = anchor_map.get(anchor, "mm")
        self.draw.text((x, y), text, font=font, fill=fill, anchor=pil_anchor, align="center")


def find_cjk_font(size, bold=False):
    """
    跨平台尋找系統可用的中文字型檔，回傳 PIL ImageFont 物件。
    找不到任何中文字型時退回 PIL 預設字型（會無法顯示中文，但至少不會崩潰）。
    """
    import platform
    from PIL import ImageFont
    system = platform.system()
    if system == "Windows":
        candidates = (["C:/Windows/Fonts/msjhbd.ttc", "C:/Windows/Fonts/msyhbd.ttc"] if bold else []) + [
            "C:/Windows/Fonts/msjh.ttc", "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/mingliu.ttc", "C:/Windows/Fonts/simsun.ttc",
        ]
    elif system == "Darwin":
        candidates = [
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Light.ttc",
            "/System/Library/Fonts/STHeiti Medium.ttc",
        ]
    else:
        candidates = [
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc" if bold else
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
            "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue
    return ImageFont.load_default()


def draw_scrap_block(drawer, x0, y0, x1, y1):
    """
    繪製醒目的餘料色塊：橘黃底色 + 斜線網底紋路 + 橘紅虛線外框。
    Module-level 純函式（不依賴任何視窗實例），畫面顯示與批次下載都呼叫這個函式。
    """
    drawer.rect(x0, y0, x1, y1, fill="#FDE7C8")

    w, h = x1 - x0, y1 - y0
    if w > 6 and h > 6:
        step = 14
        diag = w + h
        for offset in range(0, int(diag) + step, step):
            raw_x_start, raw_y_start = x0 + offset, y1
            raw_x_end,   raw_y_end   = x0 + offset - h, y0
            cx0 = min(max(raw_x_start, x0), x1)
            cx1 = min(max(raw_x_end,   x0), x1)
            if raw_x_end != raw_x_start:
                t0 = (cx0 - raw_x_start) / (raw_x_end - raw_x_start)
                t1 = (cx1 - raw_x_start) / (raw_x_end - raw_x_start)
                cy0 = raw_y_start + t0 * (raw_y_end - raw_y_start)
                cy1 = raw_y_start + t1 * (raw_y_end - raw_y_start)
                if 0 <= t0 <= 1 or 0 <= t1 <= 1:
                    drawer.line(cx0, cy0, cx1, cy1, fill="#F0B868", width=1)

    drawer.rect(x0, y0, x1, y1, outline="#DC6803", width=2, dash=(4, 2))


LAYOUT_EDGE_CLR = "#E2E8F0"   # 修邊


def layout_row_used(L, row):
    """排的已用長度（含前端修邊），與配料計算的 col_used 相同"""
    used = L["trim"]
    for i, p in enumerate(row.get("parts", [])):
        used += (L["kerf"] if i else 0) + p["length"]
    return used


def layout_row_left(L, row):
    """該排長度方向餘料：扣最後一刀鋸縫及尾端修邊（與切割明細相同算法）"""
    return max(0, L["board_l"] - layout_row_used(L, row) - L["trim"] - L["kerf"])


def layout_scrap_runs(L):
    """各排長度方向餘料，相鄰等長的排合併"""
    return scrap_runs([layout_row_left(L, r) for r in L["rows"]])


def layout_width_left(L):
    """寬度方向餘料：扣一刀鋸縫及兩側修邊（與切割明細相同算法）"""
    nr = len([r for r in L["rows"] if r.get("parts")])
    used_w = nr * L["part_w"] + max(0, nr - 1) * L["kerf"] + L["trim"] * 2
    return max(0, L["board_w"] - used_w - L["kerf"])


def render_layout(drawer, layout, zoom=1.0):
    """
    共用繪圖核心邏輯：計算所有座標並透過 drawer 畫出排列圖。
    Module-level 純函式，只依賴傳入的 layout 字典與 zoom 倍率，
    不依賴任何 Tkinter 視窗實例，因此：
      - 畫面顯示（LayoutWindow._draw 呼叫，搭配 TkDrawer）
      - 單張下載（LayoutWindow._download_image 呼叫，搭配 PILDrawer）
      - 批次下載（不開視窗，直接對多個 layout 呼叫，搭配 PILDrawer）
    三種情境都呼叫同一份邏輯，確保結果永遠一致。
    回傳 (scale, margin, dims)，dims=(x0,y0,x1,y1) 供呼叫端計算畫布總尺寸。
    """
    L = layout
    board_w, board_l = L["board_w"], L["board_l"]
    part_w, thick    = L["part_w"], L["thick"]
    trim, kerf        = L["trim"], L["kerf"]
    rows              = L["rows"]

    # 每排最小高度 80px（加大避免標籤文字互相重疊）
    MIN_ROW_H = 80
    scale = (MIN_ROW_H / part_w) * zoom

    # 左側區塊配置（由外而內）：
    # [板寬直書文字 40px] [間距 30px] [排N+W數值 兩行標籤 52px] [間距 20px] [板框]
    LABEL_W_BOARD = 40
    GAP_1         = 30
    LABEL_W_ROW   = 52
    GAP_2         = 20
    margin = LABEL_W_BOARD + GAP_1 + LABEL_W_ROW + GAP_2

    # 頂部留白加大，避免「板長」文字與其他元素重疊
    TOP_MARGIN = 50
    BOT_MARGIN = 30

    x0 = margin
    y0 = TOP_MARGIN
    x1 = margin + board_l*scale
    y1 = TOP_MARGIN + board_w*scale

    # 底色＝修邊（四周修邊與無法成為餘料的邊角）；零件、切割縫、餘料畫在上面
    drawer.rect(x0, y0, x1, y1, outline="#333", width=2, fill=LAYOUT_EDGE_CLR)

    # 板長標示：放在板框上方正中央
    drawer.text((x0+x1)/2, y0-28, f"板長 {board_l} mm", font_size=14, bold=True)

    # 板寬標示：直書放在最左側獨立欄位，不與排次標籤重疊
    board_label_x = LABEL_W_BOARD // 2
    drawer.text(board_label_x, (y0+y1)/2, f"板寬 {board_w} mm",
                font_size=13, bold=True, angle=90)

    row_h = part_w * scale
    row_top = []
    cur_y = y0 + trim*scale
    for ri, row in enumerate(rows):
        if ri > 0:
            drawer.rect(x0, cur_y, x1, cur_y + kerf*scale, fill="#FFD966")
            cur_y += kerf*scale
        row_top.append(cur_y)
        cur_x = x0 + trim*scale
        parts = row["parts"]

        for pi, part in enumerate(parts):
            if pi > 0:
                kx0 = cur_x
                kx1 = cur_x + kerf*scale
                drawer.rect(kx0, cur_y, kx1, cur_y+row_h, fill="#FFD966")
                cur_x = kx1
            pw = part["length"] * scale
            drawer.rect(cur_x, cur_y, cur_x+pw, cur_y+row_h,
                       fill="#4A90D9", outline="white", width=2)

            if pw > 90 and row_h > 36:
                fs = max(9, min(14, int(row_h / 4)))
                drawer.text(cur_x+pw/2, cur_y+row_h/2 - fs*0.8, part['name'],
                            font_size=fs, bold=True, fill="white")
                drawer.text(cur_x+pw/2, cur_y+row_h/2 + fs*0.8, f"{part['length']}mm",
                            font_size=max(8, fs-2), fill="#EAF3FF")
            elif pw > 50:
                fs = max(8, min(12, int(row_h / 3)))
                drawer.text(cur_x+pw/2, cur_y+row_h/2, f"{part['length']}",
                            font_size=fs, fill="white")
            elif pw > 26:
                fs = max(7, min(9, int(row_h / 5)))
                drawer.text(cur_x+pw/2, cur_y+row_h/2, f"{part['length']}",
                            font_size=fs, fill="white")
            cur_x += pw

        # 最後一段零件後的鋸縫（該排有餘料時）
        if layout_row_left(L, row) > 0:
            drawer.rect(cur_x, cur_y, cur_x + kerf*scale, cur_y+row_h, fill="#FFD966")

        # 排次標籤：兩行（排N / W零件寬），置中對齊該排，確保不跟板寬文字重疊
        label_cx = x0 - GAP_2 - LABEL_W_ROW/2
        row_mid  = cur_y + row_h/2
        drawer.text(label_cx, row_mid - 10, f"排{ri+1}", font_size=13, bold=True)
        drawer.text(label_cx, row_mid + 10, f"W{part_w}", font_size=10, fill="#555")

        cur_y += row_h

    # 長度方向餘料：與切割明細相同算法（扣最後一刀鋸縫及尾端修邊），相鄰等長的排畫成一整塊
    for run in layout_scrap_runs(L):
        sx1 = x1 - trim*scale
        kx1 = sx1 - run["len"]*scale
        sy0 = row_top[run["start"]]
        sy1 = row_top[run["end"]] + row_h
        draw_scrap_block(drawer, kx1, sy0, sx1, sy1)
        scrap_ref = rows[run["start"]].get("scrap_ref")
        no_str = scrap_ref.get("_no", "") if scrap_ref else ""
        n = run["end"] - run["start"] + 1
        size = f"{run_width(run, part_w, kerf)}×{run['len']}mm" if n > 1 else f"{run['len']}mm"
        label = f"♻ {no_str} {size}" if no_str else f"♻ 餘料 {size}"
        if (sx1-kx1) > 55 and row_h > 28:
            fs2 = max(9, min(13, int(row_h/4)))
            drawer.text((kx1+sx1)/2, (sy0+sy1)/2, label,
                        font_size=fs2, bold=True, fill="#B45309")

    # 寬度方向餘料：與切割明細相同算法（扣一刀鋸縫及另一側修邊）
    wrem = layout_width_left(L)
    if wrem > 0:
        ky1 = cur_y + kerf*scale
        sy1 = ky1 + wrem*scale
        drawer.rect(x0, cur_y, x1, ky1, fill="#FFD966")
        draw_scrap_block(drawer, x0, ky1, x1, sy1)
        width_ref = L.get("width_scrap_ref")
        no_str = width_ref.get("_no", "") if width_ref else ""
        label = f"♻ {no_str} 寬度方向餘料 {wrem}mm" if no_str \
                else f"♻ 寬度方向餘料 {wrem}mm"
        if (sy1 - ky1) > 20:
            drawer.text((x0+x1)/2, (ky1+sy1)/2, label, font_size=12, bold=True, fill="#B45309")

    return scale, margin, (x0, y0, x1, y1)


def layout_to_pil_image(layout, zoom=1.0, scale_up=2):
    """
    把單一 layout 字典直接渲染成 PIL Image 物件（不需要任何 Tkinter 視窗）。
    用於單張下載與批次下載共用，確保兩者輸出品質、版面完全一致。
    """
    from PIL import Image, ImageDraw
    part_w = layout["part_w"]
    scale_base = (80 / part_w) * zoom   # 與render_layout的MIN_ROW_H=80一致
    LABEL_W_BOARD = 40
    GAP_1, LABEL_W_ROW, GAP_2 = 30, 52, 20
    margin_base = LABEL_W_BOARD + GAP_1 + LABEL_W_ROW + GAP_2
    TOP_MARGIN, BOT_MARGIN = 50, 30
    img_w = int(layout["board_l"] * scale_base) + margin_base + 60
    img_h = int(layout["board_w"] * scale_base) + TOP_MARGIN + BOT_MARGIN + 20

    img = Image.new("RGB", (img_w*scale_up, img_h*scale_up), "white")
    draw = ImageDraw.Draw(img)
    draw._image = img
    pil_drawer = PILDrawer(draw, lambda size, bold=False:
                           find_cjk_font(int(size*scale_up), bold))
    render_layout(pil_drawer, layout, zoom=zoom*scale_up)
    return img


class LayoutWindow(tk.Toplevel):
    """切割排列圖視窗：用 Canvas 繪製單張鋼板上每排零件的實際排列位置"""
    def __init__(self, parent, idx, layout):
        super().__init__(parent)
        self.title(f"切割排列圖　片次 {idx}")
        self.configure(bg=CLR_BG)
        self.layout = layout
        self.idx = idx   # 保留片次編號，供下載時組成預設檔名

        # 頂部資訊列
        info = tk.Frame(self, bg=CLR_HEADER, height=50)
        info.pack(fill="x")
        info.pack_propagate(False)
        spec = f'PL{layout["thick"]}×{layout["board_w"]}×{layout["board_l"]}'
        tk.Label(info, text=f"片次 {idx}　{spec}　共 {len(layout['rows'])} 排　"
                            f"裁切板寬 W{layout['part_w']}mm",
                 font=("Microsoft JhengHei", 14, "bold"),
                 bg=CLR_HEADER, fg="white").pack(side="left", padx=16, pady=10)
        tk.Button(info, text="💾 下載圖片", command=self._download_image,
                  font=("Microsoft JhengHei", 10, "bold"),
                  bg="#3D6B6B", fg="white", relief="flat",
                  padx=12, pady=4).pack(side="right", padx=16)

        # 圖例
        legend = tk.Frame(self, bg=CLR_BG)
        legend.pack(fill="x", padx=12, pady=(8, 0))
        for color, text in [("#4A90D9", "零件"), ("#FDE7C8", "♻ 餘料"), (LAYOUT_EDGE_CLR, "修邊"),
                            ("#FFD966", "切割縫(kerf)")]:
            box = tk.Frame(legend, bg=color, width=18, height=18,
                           highlightbackground="#DC6803" if color=="#FDE7C8" else color,
                           highlightthickness=2 if color=="#FDE7C8" else 0)
            box.pack(side="left", padx=(0, 5))
            tk.Label(legend, text=text, font=("Microsoft JhengHei", 11),
                     bg=CLR_BG).pack(side="left", padx=(0, 16))

        # 縮放控制
        zoom_frame = tk.Frame(legend, bg=CLR_BG)
        zoom_frame.pack(side="right")
        tk.Label(zoom_frame, text="縮放：", font=("Microsoft JhengHei", 10),
                 bg=CLR_BG).pack(side="left")
        self.zoom_var = tk.DoubleVar(value=1.0)
        for label, val in [("－", -0.25), ("100%", None), ("＋", 0.25)]:
            if val is None:
                tk.Button(zoom_frame, text=label, font=("Microsoft JhengHei", 9),
                          command=self._zoom_reset, relief="flat", bg="#E2E8F0",
                          padx=8).pack(side="left", padx=2)
            else:
                tk.Button(zoom_frame, text=label, font=("Microsoft JhengHei", 10, "bold"),
                          command=lambda v=val: self._zoom_adjust(v),
                          relief="flat", bg="#E2E8F0", width=3).pack(side="left", padx=2)

        # 可捲動 Canvas
        outer = tk.Frame(self, bg=CLR_BG)
        outer.pack(fill="both", expand=True, padx=12, pady=10)
        self.canvas = tk.Canvas(outer, bg="white", highlightthickness=1,
                                highlightbackground="#999")
        vsb = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        hsb = ttk.Scrollbar(outer, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)

        # 滑鼠滾輪縮放（Ctrl+滾輪）與一般捲動
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Control-MouseWheel>", self._on_ctrl_mousewheel)
        self.canvas.bind("<Shift-MouseWheel>", lambda e: self.canvas.xview_scroll(int(-e.delta/60), "units"))

        # 視窗大小：依板材比例與排數動態決定，並設定上限避免超出螢幕
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        n_rows = max(len(layout["rows"]), 1)
        # 預估理想視窗：每排約 60px 高度起跳，最少 700 寬
        ideal_w = min(int(screen_w * 0.85), 1400)
        ideal_h = min(max(700, 140 + n_rows * 70), int(screen_h * 0.85))
        self.geometry(f"{ideal_w}x{ideal_h}")
        self.minsize(700, 500)
        self.resizable(True, True)   # 可自由拉伸放大縮小

        self.after(50, self._draw)

    def _zoom_reset(self):
        self.zoom_var.set(1.0)
        self._draw()

    def _zoom_adjust(self, delta):
        new_zoom = max(0.25, min(4.0, self.zoom_var.get() + delta))
        self.zoom_var.set(new_zoom)
        self._draw()

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-event.delta/60), "units")

    def _on_ctrl_mousewheel(self, event):
        self._zoom_adjust(0.1 if event.delta > 0 else -0.1)

    def _draw(self):
        """畫面顯示用：使用 TkDrawer 把排列圖畫到 self.canvas 上"""
        scale, margin, dims = render_layout(TkDrawer(self.canvas), self.layout,
                                            zoom=self.zoom_var.get())
        L = self.layout
        canvas_w = int(L["board_l"] * scale) + margin + 60
        canvas_h = int(L["board_w"] * scale) + margin*2 + 50
        self.canvas.configure(scrollregion=(0, 0, canvas_w, canvas_h))

    def _download_image(self):
        """
        把目前這張排列圖匯出成 PNG 圖檔（單張下載）。
        實際繪圖呼叫共用的 layout_to_pil_image()，與批次下載用同一份邏輯，
        確保單張下載／批次下載／畫面顯示三者結果一致。
        """
        try:
            from PIL import Image
        except ImportError:
            messagebox.showerror("缺少套件",
                "下載圖片需要 Pillow 套件，請先安裝：\n\npip install Pillow")
            return

        safe_idx = str(self.idx).replace("*", "星").replace("/", "_")
        default_name = f"BH排列圖_{safe_idx}.png"
        path = filedialog.asksaveasfilename(
            title="下載排列圖",
            defaultextension=".png",
            initialfile=default_name,
            filetypes=[("PNG 圖片", "*.png"), ("所有檔案", "*.*")]
        )
        if not path:
            return

        try:
            img = layout_to_pil_image(self.layout, zoom=self.zoom_var.get())
            img.save(path, "PNG")
            messagebox.showinfo("下載完成", f"排列圖已儲存：\n{path}")
        except Exception as e:
            messagebox.showerror("下載失敗", f"圖片匯出失敗：{e}")


class PreviewWindow(tk.Toplevel):
    def __init__(self, parent, result):
        super().__init__(parent)
        self.title("配料結果預覽")
        self.geometry("1000x680")
        self.minsize(700, 420)      # 最小尺寸，避免縮太小導致表格內容互相擠壓看不到
        self.resizable(True, True)  # 可自由拉伸放大縮小
        self.configure(bg=CLR_BG)
        r = result
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=6, pady=6)

        # 注意：「餘料清單」分頁的建立過程會把編號（餘01/餘02...）寫回每個
        # 餘料物件的 "_no" 欄位，「切割明細」分頁的排列圖功能需要讀取這個編號，
        # 所以資料處理順序上要先跑 _tab_scraps，確保 _no 已寫入；
        # 但畫面上的頁籤顯示順序仍維持「切割明細」在前、「餘料清單」在後，不影響操作習慣。
        tabs_data = [
            ("配料清單",   self._tab_bh_rows,    r),
            ("BH拆板明細", self._tab_decomposed, r),
            ("BH合板清單", self._tab_bh_combine, r),
            ("採購清單",   self._tab_purchase,   r),
            ("餘料清單",   self._tab_scraps,     r),
            ("切割明細",   self._tab_cuts,       r),
        ]
        built_frames = {}
        for name, builder, data in tabs_data:
            frame = ttk.Frame(nb)
            built_frames[name] = frame
            builder(frame, data)
        for name in ["配料清單", "BH拆板明細", "BH合板清單", "採購清單", "切割明細", "餘料清單"]:
            nb.add(built_frames[name], text=f"  {name}  ")

    def _make_tree(self, parent, cols, widths):
        frame = tk.Frame(parent, bg=CLR_BG)
        frame.pack(fill="both", expand=True, padx=4, pady=4)
        vsb = ttk.Scrollbar(frame, orient="vertical")
        hsb = ttk.Scrollbar(frame, orient="horizontal")
        tree = ttk.Treeview(frame, columns=cols, show="headings",
                            yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)
        for col, w in zip(cols, widths):
            tree.heading(col, text=col)
            tree.column(col, width=w, anchor="center")
        tree.tag_configure("odd",   background=CLR_ROW_ODD)
        tree.tag_configure("even",  background=CLR_ROW_EVEN)
        tree.tag_configure("scrap", background=CLR_SCRAP)
        tree.tag_configure("comp",  background="#D6EAF8")
        tree.tag_configure("unused",background="#FFF3CD")
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return tree

    def _tab_bh_rows(self, frame, r):
        cols = ("構件編號","零件編號","斷面規格","長度(mm)","數量","材質")
        tree = self._make_tree(frame, cols, [80,80,200,90,70,70])
        for i, row in enumerate(r["bh_rows"]):
            tree.insert("", "end", values=(row["comp"],row["part"],row["spec"],
                        row["length"],row["qty"],row["mat"]),
                        tags=("odd" if i%2==0 else "even",))

    def _tab_decomposed(self, frame, r):
        cols = ("構件編號","零件編號","拆板零件","流水號","流水號(W/F)","斷面規格","長度(mm)","數量","材質","單重(kg)","總重(kg)")
        tree = self._make_tree(frame, cols, [80,80,90,80,110,200,80,60,60,80,80])
        for d in r["decomposed"]:
            cw = round(d["flange"]["unit_wt"]*2+d["web"]["unit_wt"],1)
            ct = round(cw*d["qty"],1)
            fl_sn = d["flange"].get("serial","")
            wb_sn = d["web"].get("serial","")
            wf_str = f"{wb_sn}W / {fl_sn}F" if (wb_sn or fl_sn) else ""
            tree.insert("","end", values=(d["comp"],"","","","",d["spec"],d["length"],d["qty"],d["mat"],cw,ct), tags=("comp",))
            fl  = d["flange"]
            fl_sn_wf = f"{fl_sn}F" if fl_sn else ""
            tree.insert("","end", values=(d["comp"],d.get("part",""),fl["name"],fl_sn,fl_sn_wf,
                        f'PL{fl["thick"]}×{fl["width"]}×{fl["length"]}',
                        fl["length"],fl["total"],fl["mat"],fl["unit_wt"],fl["total_wt"]), tags=("odd",))
            wb_ = d["web"]
            wb_sn_wf = f"{wb_sn}W" if wb_sn else ""
            tree.insert("","end", values=(d["comp"],d.get("part",""),wb_["name"],wb_sn,wb_sn_wf,
                        f'PL{wb_["thick"]}×{wb_["width"]}×{wb_["length"]}',
                        wb_["length"],wb_["total"],wb_["mat"],wb_["unit_wt"],wb_["total_wt"]), tags=("even",))

    def _tab_bh_combine(self, frame, r):
        """BH合板清單：構件編號、流水號(W/F)、斷面規格、長度、數量"""
        cols = ("構件編號", "流水號(W/F)", "斷面規格", "長度(mm)", "數量")
        tree = self._make_tree(frame, cols, [90, 120, 220, 90, 70])
        for d in r["decomposed"]:
            fl_sn = d["flange"].get("serial", "")
            wb_sn = d["web"].get("serial", "")
            fl_sn_wf = f"{fl_sn}F" if fl_sn else ""
            wb_sn_wf = f"{wb_sn}W" if wb_sn else ""
            # 翼板行（共2片/支）
            tree.insert("", "end", values=(
                d["comp"], fl_sn_wf,
                f'PL{d["flange"]["thick"]}×{d["flange"]["width"]}×{d["flange"]["length"]}',
                d["flange"]["length"], d["flange"]["total"]),
                tags=("odd",))
            # 腹板行（共1片/支）
            tree.insert("", "end", values=(
                d["comp"], wb_sn_wf,
                f'PL{d["web"]["thick"]}×{d["web"]["width"]}×{d["web"]["length"]}',
                d["web"]["length"], d["web"]["total"]),
                tags=("even",))

    def _tab_thick_list(self, frame, r):
        """板厚清單：流水號(W/F)、斷面規格、數量、材質、單重、總重
           同流水號(W/F)的合計在一起
        """
        cols = ("流水號(W/F)", "斷面規格", "數量", "材質", "單重(kg)", "總重(kg)")
        tree = self._make_tree(frame, cols, [120, 220, 70, 70, 90, 90])

        # 依流水號(W/F)分組，相同者合計數量和總重
        from collections import OrderedDict
        groups = OrderedDict()
        for d in r["decomposed"]:
            for part, suffix in [(d["flange"], "F"), (d["web"], "W")]:
                sn   = part.get("serial", "")
                sn_wf = f"{sn}{suffix}" if sn else ""
                spec = f'PL{part["thick"]}×{part["width"]}×{part["length"]}'
                key  = (sn_wf, spec, part["mat"], part["unit_wt"])
                if key not in groups:
                    groups[key] = {"qty": 0, "total_wt": 0.0}
                groups[key]["qty"]      += part["total"]
                groups[key]["total_wt"] += part["total_wt"]

        for i, ((sn_wf, spec, mat, unit_wt), val) in enumerate(groups.items()):
            tree.insert("", "end", values=(
                sn_wf, spec, val["qty"], mat,
                round(unit_wt, 1), round(val["total_wt"], 1)),
                tags=("odd" if i%2==0 else "even",))

    def _tab_purchase(self, frame, r):
        cols = ("NO","採購規格(mm)","數量","材質","單片重(kg)","總重(kg)")
        tree = self._make_tree(frame, cols, [50,260,60,60,100,100])
        from collections import OrderedDict
        merged = OrderedDict()
        for p in r["purchase_list"]:
            key = (p["spec"], p["mat"], p["unit_wt"])
            if key not in merged:
                merged[key] = {"qty": 0, "total_wt": 0}
            merged[key]["qty"]      += p["qty"]
            merged[key]["total_wt"] += p["total_wt"]
        total = 0
        total_qty = 0
        for i, ((spec, mat, unit_wt), val) in enumerate(merged.items()):
            tree.insert("","end", values=(i+1, spec, val["qty"], mat,
                        unit_wt, round(val["total_wt"],0)),
                        tags=("odd" if i%2==0 else "even",))
            total     += val["total_wt"]
            total_qty += val["qty"]
        tree.insert("","end",
                    values=("", f"採購合計：共 {total_qty} 片", "", "", "總重(kg)：", f"{total:,.0f}"),
                    tags=("comp",))

    def _tab_cuts(self, frame, r):
        # 提示文字先 pack（必須在 _make_tree 之前，否則 tree 的 expand=True 會把空間吃光）
        tip = tk.Label(frame, text="💡 雙擊片次可查看實際排列圖",
                       font=("Microsoft JhengHei", 9), bg=CLR_BG, fg="#1A3050")
        tip.pack(anchor="w", padx=6, pady=(2,0))

        cols = ("片次","板別","採購規格(mm)","材質","板重(kg)","構件（裁切尺寸）","數量","單重(kg)","餘料")
        tree = self._make_tree(frame, cols, [55,65,180,55,75,220,50,75,220])
        self._layout_map = {}
        for i, ct in enumerate(r["cut_details"]):
            tag = "scrap" if ct["is_scrap"] else ("odd" if i%2==0 else "even")
            item = tree.insert("","end", values=(ct["idx"],"F=翼" if ct["type"]=="F" else "W=腹",
                        ct["board_spec"],ct["mat"],ct["board_wt"],
                        f'{ct["comp"]}（{ct["part_spec"]}）',
                        ct["qty"],ct["unit_wt"],ct["leftover"]), tags=(tag,))
            if ct.get("layout"):
                self._layout_map[item] = (ct["idx"], ct["layout"])

        def on_double_click(event):
            item = tree.focus()
            if item in self._layout_map:
                idx, layout = self._layout_map[item]
                assign_scrap_numbers(r)   # 保險呼叫：確保開圖前餘NO一定已寫入
                LayoutWindow(self, idx, layout)
            else:
                messagebox.showinfo("提示", "此片次沒有排列圖")

        tree.bind("<Double-1>", on_double_click)

    def _tab_scraps(self, frame, r):
        # 工具列：匯出功能
        toolbar = tk.Frame(frame, bg=CLR_BG)
        toolbar.pack(fill="x", padx=6, pady=(4,0))
        tk.Button(toolbar, text="💾 匯出本次產生餘料 CSV",
                  command=lambda: self._export_scraps_csv(r, "new"),
                  bg="#6B46C1", fg="white", font=("Microsoft JhengHei", 9, "bold"),
                  relief="flat", padx=10, pady=3).pack(side="left", padx=(0,4))
        tk.Button(toolbar, text="💾 匯出未使用現有餘料 CSV",
                  command=lambda: self._export_scraps_csv(r, "unused"),
                  bg="#744210", fg="white", font=("Microsoft JhengHei", 9, "bold"),
                  relief="flat", padx=10, pady=3).pack(side="left", padx=(0,4))
        tk.Button(toolbar, text="📒 匯出本次庫存異動",
                  command=lambda: self._export_inventory_events(r),
                  bg="#276749", fg="white", font=("Microsoft JhengHei", 9, "bold"),
                  relief="flat", padx=10, pady=3).pack(side="left")

        cols = ("餘NO","來源","板別","餘料規格(mm)","材質","重量(kg)","狀態")
        tree = self._make_tree(frame, cols, [60,70,65,220,60,80,80])
        assign_scrap_numbers(r)   # 統一賦予餘NO，寫回每個餘料物件的 "_no" 欄位
        total_wt = 0
        for i, sc in enumerate(r["new_scraps"]):
            tree.insert("","end", values=(sc["_no"],sc["src"],
                        "F=翼" if sc["type"]=="F" else "W=腹",
                        sc["spec"],sc["mat"],sc["wt"],"本次產生"),
                        tags=("odd" if i%2==0 else "even",))
            total_wt += sc["wt"]
        unused = [s for s in r["existing_scraps"] if s["id"] not in r["used_scrap_ids"]]
        for s in unused:
            wt = round(calc_weight(s["width"],s["thick"],s["length"],
                                   r["params"]["density"]))
            tree.insert("","end", values=(f'*{s["id"]}',"—","—",
                        f'PL{s["thick"]}×{s["width"]}×{s["length"]}',
                        s["mat"],wt,"未使用"), tags=("unused",))
        # 餘料合計總重
        tree.insert("","end",
                    values=("", f"餘料合計：共 {len(r['new_scraps'])} 件", "", "", "總重(kg)：", f"{total_wt:,.0f}", ""),
                    tags=("comp",))

    def _export_scraps_csv(self, r, mode):
        """
        匯出餘料 CSV。
        mode="new"    → 本次配料產生的新餘料
        mode="unused" → 現有餘料中本次未使用到的部分
        兩種格式欄位一致，可直接再匯入到「現有餘料」分頁使用。
        """
        import csv
        density = r["params"]["density"]

        if mode == "new":
            title = "匯出本次產生餘料"
            default_name = f"本次產生餘料_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
            rows_data = []
            for sc in r["new_scraps"]:
                # 解析規格字串 PL厚×寬×長
                import re
                m = re.match(r"PL(\d+)×(\d+)×(\d+)", sc["spec"])
                if m:
                    rows_data.append({
                        "厚度(mm)": m.group(1),
                        "寬度(mm)": m.group(2),
                        "長度(mm)": m.group(3),
                        "數量(片)": 1,
                        "材質":     sc["mat"],
                        "重量(kg)": sc["wt"],
                        "來源片次": sc["src"],
                    })
        else:  # mode == "unused"
            title = "匯出未使用現有餘料"
            default_name = f"未使用現有餘料_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
            unused = [s for s in r["existing_scraps"] if s["id"] not in r["used_scrap_ids"]]
            if not unused:
                messagebox.showinfo("提示", "本次配料所有現有餘料均已使用，無未使用餘料可匯出。")
                return
            rows_data = []
            for s in unused:
                wt = round(calc_weight(s["width"], s["thick"], s["length"], density))
                rows_data.append({
                    "厚度(mm)": s["thick"],
                    "寬度(mm)": s["width"],
                    "長度(mm)": s["length"],
                    "數量(片)": s.get("qty", 1),
                    "材質":     s["mat"],
                    "重量(kg)": wt,
                    "來源片次": s["id"],
                })

        if not rows_data:
            messagebox.showinfo("提示", "沒有資料可匯出。")
            return

        path = filedialog.asksaveasfilename(
            title=title,
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("CSV 檔案","*.csv"),("所有檔案","*.*")])
        if not path:
            return

        try:
            with open(path, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=["厚度(mm)","寬度(mm)","長度(mm)","數量(片)","材質","重量(kg)","來源片次"])
                writer.writeheader()
                writer.writerows(rows_data)
            messagebox.showinfo("匯出完成",
                f"已成功匯出 {len(rows_data)} 筆餘料至：\n{path}\n\n"
                f"此檔案可直接用「匯入餘料CSV」按鈕載入到下次的現有餘料清單中。")
        except Exception as e:
            messagebox.showerror("匯出失敗", str(e))

    def _export_inventory_events(self, r):
        """匯出本次計算的餘料庫存異動；使用量為負數，新產生餘料為正數。"""
        import csv
        events = []
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        proj_no = r["params"].get("proj_no", "")
        for sc in r["existing_scraps"]:
            if sc["id"] in r["used_scrap_ids"]:
                events.append({
                    "時間": stamp, "案號": proj_no, "異動": "使用既有餘料",
                    "餘料編號": sc["id"],
                    "規格(mm)": f"PL{sc['thick']}×{sc['width']}×{sc['length']}",
                    "材質": sc["mat"], "數量變化": -1, "來源": "配料計算"
                })
        for sc in r["new_scraps"]:
            events.append({
                "時間": stamp, "案號": proj_no, "異動": "新增切割餘料",
                "餘料編號": sc.get("_no", ""), "規格(mm)": sc["spec"],
                "材質": sc["mat"], "數量變化": 1,
                "來源": f"片次 {sc['src']}"
            })
        if not events:
            messagebox.showinfo("提示", "本次沒有餘料庫存異動可匯出。", parent=self)
            return
        path = filedialog.asksaveasfilename(
            title="儲存本次餘料庫存異動紀錄", defaultextension=".csv",
            initialfile="BH_餘料庫存異動紀錄.csv",
            filetypes=[("CSV 檔案", "*.csv")], parent=self)
        if not path:
            return
        try:
            write_header = not os.path.exists(path) or os.path.getsize(path) == 0
            encoding = "utf-8-sig" if write_header else "utf-8"
            with open(path, "a", encoding=encoding, newline="") as f:
                writer = csv.DictWriter(f, fieldnames=list(events[0].keys()))
                if write_header:
                    writer.writeheader()
                writer.writerows(events)
            messagebox.showinfo("匯出完成", f"已追加 {len(events)} 筆庫存異動：\n{path}", parent=self)
        except Exception as e:
            messagebox.showerror("匯出失敗", str(e), parent=self)



class BHPeilianApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BH型鋼鋼板配料系統　v24.5")
        self.geometry("1100x780")
        self.minsize(900, 600)   # 強制最小視窗尺寸，避免按鈕列被擠出畫面
        self.configure(bg=CLR_BG)
        self.resizable(True, True)
        self._set_app_icon()
        self._build_ui()
        self._result       = None
        self._dirty        = False
        self._current_file = None   # 目前開啟的檔案路徑（None = 新檔案）

        # 關閉視窗時詢問
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # 啟動時詢問是否開啟上次檔案
        self.after(300, self._ask_restore)

    # ── 視窗標題更新 ──────────────────────────────────────────────
    def _update_title(self):
        if self._current_file:
            fname = os.path.basename(self._current_file)
            self.title(f"BH型鋼鋼板配料系統　v24.5　—　{fname}")
        else:
            self.title("BH型鋼鋼板配料系統　v24.5　—　新檔案")

    # ── 關閉確認 ──────────────────────────────────────────────────
    def _on_close(self):
        if getattr(self, "_calc_cancel_event", None) is not None:
            messagebox.showinfo("配料進行中", "請先取消計算並等待它結束，再關閉程式。", parent=self)
            return
        bh_rows = self._get_bh_rows() or []
        if len(bh_rows) > 0:
            ans = messagebox.askyesnocancel(
                "關閉確認",
                "目前有配料資料尚未儲存。\n\n"
                "「是」→ 儲存後關閉\n"
                "「否」→ 直接關閉（不儲存）\n"
                "「取消」→ 返回程式")
            if ans is None:
                return
            if ans:
                if not self._save_session():
                    return
            else:
                # 不儲存但做緊急暫存，下次開啟可還原
                self._autosave_tmp()
        self.destroy()

    # ── 儲存（有檔名直接存，無則另存新檔）────────────────────────
    def _save_session(self):
        if self._current_file:
            return self._write_session(self._current_file)
        else:
            return self._save_session_as()

    # ── 另存新檔 ──────────────────────────────────────────────────
    def _save_session_as(self):
        proj_no   = self.var_proj_no.get().strip()
        proj_name = self.var_proj_name.get().strip()
        default   = f"BH配料_{proj_no}_{proj_name}.json" if proj_no else "BH配料.json"
        path = filedialog.asksaveasfilename(
            title="儲存配料進度",
            defaultextension=".json",
            filetypes=[("BH配料進度檔","*.json"),("所有檔案","*.*")],
            initialfile=default)
        if not path:
            return False
        self._current_file = path
        self._update_title()
        return self._write_session(path)

    # ── 實際寫入 ──────────────────────────────────────────────────
    def _write_session(self, path):
        import json
        try:
            params  = self._get_params() or {}
            bh_rows = self._get_bh_rows() or []
            scraps  = self._get_scraps()  or []
            data = {
                "version":  "1.0",
                "saved_at": datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                "params":   params,
                "bh_rows":  bh_rows,
                "scraps":   scraps,
            }
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.status_var.set(f"✅ 已儲存：{os.path.basename(path)}")
            return True
        except Exception as e:
            messagebox.showerror("儲存失敗", str(e))
            return False

    # ── 開啟進度檔 ────────────────────────────────────────────────
    def _open_session(self):
        import json
        path = filedialog.askopenfilename(
            title="開啟配料進度檔",
            filetypes=[("BH配料進度檔","*.json"),("所有檔案","*.*")])
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
            self._current_file = path
            self._update_title()
            self._restore_session(data)
        except Exception as e:
            messagebox.showerror("開啟失敗", str(e))

    # ── 啟動時詢問 ────────────────────────────────────────────────
    def _ask_restore(self):
        import json
        # 尋找同資料夾最近一次的暫存（若有）
        base = os.path.dirname(os.path.abspath(sys.argv[0]))
        tmp  = os.path.join(base, "_bh_autosave.json")
        if not os.path.exists(tmp):
            return
        try:
            with open(tmp, encoding="utf-8") as f:
                data = json.load(f)
            saved_at = data.get("saved_at", "不明時間")
            bh_count = len(data.get("bh_rows", []))
            ans = messagebox.askyesno(
                "發現未完成的進度",
                f"上次程式關閉時有未儲存的資料：\n\n"
                f"  時間：{saved_at}\n"
                f"  BH構件：{bh_count} 筆\n\n"
                f"是否還原？")
            os.remove(tmp)   # 無論是否還原都刪掉暫存
            if ans:
                self._restore_session(data)
        except Exception:
            pass

    # ── 關閉時自動暫存（緊急備份）────────────────────────────────
    def _autosave_tmp(self):
        import json
        try:
            bh_rows = self._get_bh_rows() or []
            if not bh_rows:
                return
            base = os.path.dirname(os.path.abspath(sys.argv[0]))
            tmp  = os.path.join(base, "_bh_autosave.json")
            data = {
                "version":  "1.0",
                "saved_at": datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                "params":   self._get_params() or {},
                "bh_rows":  bh_rows,
                "scraps":   self._get_scraps() or [],
            }
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ── 還原資料 ──────────────────────────────────────────────────
    def _restore_session(self, data):
        p = data.get("params", {})
        param_map = {
            "proj_no":    "var_proj_no",
            "proj_name":  "var_proj_name",
            "proj_date":  "var_proj_date",
            "split_code": "var_split_code",
            "mat_name":   "var_mat",
            "new_kerf":   "var_new_kerf",
            "new_trim":   "var_new_trim",
            "scrap_kerf": "var_scrap_kerf",
            "scrap_trim": "var_scrap_trim",
            "bw_min":     "var_bw_min",
            "bw_max":     "var_bw_max",
            "bl_min":     "var_bl_min",
            "bl_max":     "var_bl_max",
            "w_min":      "var_w_min",
            "w_max":      "var_w_max",
            "cut_mode":   "var_cut_mode",
        }
        for key, attr in param_map.items():
            val = p.get(key)
            if val is not None and hasattr(self, attr):
                try:
                    getattr(self, attr).set(str(val))
                except Exception:
                    pass

        for item in self.bh_tree.get_children():
            self.bh_tree.delete(item)
        for i, row in enumerate(data.get("bh_rows", [])):
            tag = "odd" if i % 2 == 0 else "even"
            self.bh_tree.insert("", "end",
                values=(row.get("comp",""), row.get("part",""),
                        row.get("spec",""), row.get("length",""),
                        row.get("qty",""),  row.get("mat","")),
                tags=(tag,))

        for item in self.sc_tree.get_children():
            self.sc_tree.delete(item)
        for i, sc in enumerate(data.get("scraps", [])):
            tag = "odd" if i % 2 == 0 else "even"
            self.sc_tree.insert("", "end",
                values=(sc.get("id",""), sc.get("thick",""), sc.get("width",""),
                        sc.get("length",""), sc.get("qty",""), sc.get("mat","")),
                tags=(tag,))

        self.status_var.set(
            f"已還原進度（{data.get('saved_at','')}）　"
            f"BH {len(data.get('bh_rows',[]))} 筆　"
            f"餘料 {len(data.get('scraps',[]))} 筆")

    def _set_app_icon(self):
        """
        設定視窗左上角與工作列圖示（app_icon.ico）。
        圖示檔案需與本程式（或打包後的 .exe）放在同一目錄下，
        找不到檔案時靜默略過，不影響程式正常啟動。
        """
        try:
            icon_path = os.path.join(
                os.path.dirname(os.path.abspath(sys.argv[0])), "app_icon.ico")
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass   # Windows以外平台或找不到檔案時，保持預設圖示即可

    # ── UI 建構 ────────────────────────────────────────────────────
    def _build_ui(self):
        # 頂部標題（固定頂部）
        hdr = tk.Frame(self, bg=CLR_HEADER, height=50)
        hdr.pack(side="top", fill="x")
        hdr.pack_propagate(False)

        # 標題列圖示（從 app_icon.ico 或 app_icon.png 載入，放在文字前方）
        try:
            icon_path_png = os.path.join(
                os.path.dirname(os.path.abspath(sys.argv[0])), "app_icon.png")
            icon_path_ico = os.path.join(
                os.path.dirname(os.path.abspath(sys.argv[0])), "app_icon.ico")
            photo = None
            for p in [icon_path_png, icon_path_ico]:
                if os.path.exists(p):
                    from PIL import Image, ImageTk
                    im = Image.open(p).convert("RGBA").resize((34, 34), Image.LANCZOS)
                    photo = ImageTk.PhotoImage(im)
                    break
            if photo:
                lbl_icon = tk.Label(hdr, image=photo, bg=CLR_HEADER)
                lbl_icon.image = photo   # 保留參照防止被GC回收
                lbl_icon.pack(side="left", padx=(12,4), pady=8)
        except Exception:
            pass

        tk.Label(hdr, text="BH 型鋼鋼板配料系統", bg=CLR_HEADER,
                 fg="white", font=("Microsoft JhengHei", 16, "bold")).pack(side="left", padx=(0,16), pady=8)
        tk.Label(hdr, text="v24.5", bg=CLR_HEADER,
                 fg="#90CDF4", font=("Microsoft JhengHei", 10)).pack(side="right", padx=16)

        # 標題下方第二列：開啟進度 / 儲存進度 / 另存新檔
        sub_bar = tk.Frame(self, bg="#1A3050", height=36)
        sub_bar.pack(side="top", fill="x")
        sub_bar.pack_propagate(False)

        def sub_btn(text, cmd, bg):
            tk.Button(sub_bar, text=text, command=cmd,
                      bg=bg, fg="white", activebackground=bg,
                      font=("Microsoft JhengHei", 9, "bold"),
                      relief="flat", cursor="hand2",
                      padx=16, pady=4
                      ).pack(side="left", padx=(4,0), pady=4)

        sub_btn("📁 開啟進度", self._open_session,    "#6B46C1")
        sub_btn("💾 儲存進度", self._save_session,    "#6B46C1")
        sub_btn("💾 另存新檔", self._save_session_as, "#4A6C7A")

        # ⚠️ 重要：按鈕列與狀態列要「優先」pack 在 top/bottom，
        # 確保即使分頁內容很高，這兩排也一定保留可見空間，
        # 不會被中間的 Notebook（expand=True）擠出畫面外。

        # 右側兩欄（左：紫色「配料結果」、右：橘色「採購修正」），用 grid 讓各列左右對齊：
        #   第 0 列：匯入        ／ 採購修正（新採購清單）
        #   第 1 列：配料結果＋匯出 ／ 匯出（含信暐訂購單）
        #   第 2 列：預覽        ／ 新預覽
        #   第 3 列：空白表單    ／ （空白）
        #   第 4 列：填滿剩餘高度
        VIOLET, ORANGE = "#E4DFF1", "#F4E1C9"   # 兩欄皆用淺色底，標題改深色
        # 外層可捲動：視窗高度不足時出現捲軸（滑鼠移入側欄可用滾輪捲動）
        side_wrap = tk.Frame(self, bg=VIOLET)
        side_wrap.pack(side="right", fill="y")
        side_cv = tk.Canvas(side_wrap, bg=VIOLET, highlightthickness=0, bd=0)
        side_sb = ttk.Scrollbar(side_wrap, orient="vertical", command=side_cv.yview)
        side_cv.configure(yscrollcommand=side_sb.set)
        side_cv.pack(side="left", fill="y")
        side = tk.Frame(side_cv, bg=VIOLET)
        side_win = side_cv.create_window(0, 0, window=side, anchor="nw")

        def _side_fit(_=None):
            need_h = side.winfo_reqheight()
            have_h = side_cv.winfo_height()
            side_cv.configure(width=side.winfo_reqwidth())
            # 夠高時撐滿（第 4 列填色到底），不夠高時保持內容高度並顯示捲軸
            side_cv.itemconfigure(side_win, height=max(need_h, have_h))
            side_cv.configure(scrollregion=(0, 0, side.winfo_reqwidth(), max(need_h, have_h)))
            if need_h > have_h > 1:
                if not side_sb.winfo_ismapped():
                    side_sb.pack(side="right", fill="y")
            elif side_sb.winfo_ismapped():
                side_sb.pack_forget()
                side_cv.yview_moveto(0)

        side.bind("<Configure>", _side_fit)
        side_cv.bind("<Configure>", _side_fit)

        def _side_wheel(e):
            if side_sb.winfo_ismapped():
                side_cv.yview_scroll(int(-e.delta / 120), "units")
        side_wrap.bind("<Enter>", lambda e: side_cv.bind_all("<MouseWheel>", _side_wheel))
        side_wrap.bind("<Leave>", lambda e: side_cv.unbind_all("<MouseWheel>"))

        for c in (0, 1):
            side.grid_columnconfigure(c, minsize=110, uniform="side")
        side.grid_rowconfigure(4, weight=1)

        def cell(row, col):
            bg = VIOLET if col == 0 else ORANGE
            f = tk.Frame(side, bg=bg)
            f.grid(row=row, column=col, sticky="nsew")
            if 0 < row < 4:   # 列與列之間的分隔線
                tk.Frame(f, bg="#B9AED6" if col == 0 else "#D9B48F", height=1).pack(fill="x", padx=8, pady=(6, 2))
            return f

        def title(f, text, ghost=False):
            """欄標題（白字）；ghost=True 時文字與底色同色，只佔位以對齊另一欄"""
            bg = f.cget("bg")
            fg = "#3B2A6B" if bg == VIOLET else "#5A3510"
            tk.Label(f, text=text, bg=bg, fg=bg if ghost else fg,
                     font=("Microsoft JhengHei", 9, "bold")).pack(pady=(6, 0))

        def head(f, text):
            tk.Label(f, text=text, bg=f.cget("bg"), fg="#5B3F8C" if f.cget("bg") == VIOLET else "#7B4A1E",
                     font=("Microsoft JhengHei", 9, "bold")).pack(pady=(4, 2))

        def btn(f, text, cmd, bg, abg):
            tk.Button(f, text=text, command=cmd,
                      bg=bg, fg="white", activebackground=abg,
                      font=("Microsoft JhengHei", 10, "bold"),
                      relief="flat", cursor="hand2",
                      width=10, pady=5
                      ).pack(fill="x", padx=8, pady=3)

        # 第 0 列：匯入 ／ 採購修正
        f = cell(0, 0)
        head(f, "匯入")
        btn(f, "📂 配料CSV", self._import_csv,    "#5A7184", "#485B6B")
        btn(f, "📂 餘料CSV", self._sc_import_csv, "#5A7184", "#485B6B")
        f = cell(0, 1)
        head(f, "採購修正")
        btn(f, "📋 新採購清單", self._open_purchase_edit, "#A65C72", "#8A4A5E")

        # 第 1 列：匯出（左欄上方為「配料結果」標題，右欄放同高的隱藏標題以對齊）
        f = cell(1, 0)
        title(f, "配料結果")
        head(f, "匯出")
        btn(f, "🗂 排列圖PDF", self._export_layout_pdf,   "#7B68B5", "#655497")
        btn(f, "🖨 匯出 PDF",  self._export_pdf,          "#7B68B5", "#655497")
        btn(f, "💾 匯出Excel", self._export_xlsx,         "#7B68B5", "#655497")
        btn(f, "📝 信暐訂購單", self._open_order_original, "#7B68B5", "#655497")
        f = cell(1, 1)
        title(f, "配料結果", ghost=True)
        head(f, "匯出")
        btn(f, "🗂 排列圖PDF", self._new_export_layout, "#96603A", "#7A4D2E")
        btn(f, "🖨 匯出PDF",   self._new_export_pdf,    "#96603A", "#7A4D2E")
        btn(f, "💾 匯出Excel", self._new_export_xlsx,   "#96603A", "#7A4D2E")
        btn(f, "📝 信暐訂購單", self._open_order_select, "#96603A", "#7A4D2E")

        # 第 2 列：預覽 ／ 新預覽
        f = cell(2, 0)
        head(f, "預覽")
        btn(f, "🔍 配料結果", self._preview,        "#44797A", "#356263")
        btn(f, "📐 排列圖",   self._preview_layout, "#44797A", "#356263")
        f = cell(2, 1)
        head(f, "新預覽")
        btn(f, "🔍 配料結果", self._new_preview_result, "#5B7A99", "#4A6680")
        btn(f, "🗂 排列圖",   self._new_preview_layout, "#5B7A99", "#4A6680")

        # 第 3 列：空白表單
        f = cell(3, 0)
        head(f, "空白表單")
        btn(f, "📄 空白配料表", self._download_blank_csv, "#8A6A42", "#705634")
        btn(f, "📄 空白餘料表", self._sc_download_blank,  "#8A6A42", "#705634")
        cell(3, 1)

        # 第 4 列：填滿剩餘高度
        cell(4, 0)
        cell(4, 1)

        # 底部按鈕列（固定底部，最先保留空間）
        btn_bar = tk.Frame(self, bg=CLR_BG)
        btn_bar.pack(side="bottom", fill="x", padx=10, pady=8)
        self._calc_btn = self._btn(btn_bar, "▶  執行配料計算", self._run_calc,      CLR_BTN_MAIN, side="left")
        self._cancel_btn = self._btn(btn_bar, "取消計算", self._cancel_calc,
                                     CLR_BTN_DEL, side="left")
        self._cancel_btn.config(state="disabled")
        self._btn(btn_bar, "🗑 清除全部",    self._clear_all,     CLR_BTN_DEL,  side="right")

        # 狀態列（固定底部，緊接按鈕列上方先保留空間）
        self.status_var = tk.StringVar(value="就緒。請填寫前置參數及 BH 配料表後，點擊「執行配料計算」。")
        tk.Label(self, textvariable=self.status_var, bg="#E2E8F0",
                 anchor="w", font=("Microsoft JhengHei", 9)
                 ).pack(side="bottom", fill="x", ipady=4)

        # 主體使用 Notebook（最後 pack，自動吃掉「剩餘」空間，不會擠掉上面已保留的區塊）
        nb = ttk.Notebook(self)
        nb.pack(side="top", fill="both", expand=True, padx=10, pady=(6,0))

        self.tab_param  = ttk.Frame(nb)
        self.tab_bh     = ttk.Frame(nb)
        self.tab_scrap  = ttk.Frame(nb)
        nb.add(self.tab_param, text="  ① 前置參數  ")
        nb.add(self.tab_bh,    text="  ② BH配料表  ")
        nb.add(self.tab_scrap, text="  ③ 現有餘料  ")

        # 放大分頁標籤字型
        style = ttk.Style()
        style.configure("TNotebook.Tab",
                        font=("Microsoft JhengHei", 12, "bold"),
                        padding=[12, 6])

        # 「前置參數」分頁內容是多個堆疊的表單區塊，視窗縮小時總高度容易超出
        # 可視範圍，所以包一層可捲動容器（Canvas+Scrollbar+內嵌Frame的標準寫法），
        # 確保不管視窗多小，所有設定欄位都能透過捲軸捲動查看，不會被截斷。
        self.tab_param_inner = self._make_scrollable(self.tab_param)

        self._build_param_tab()
        self._build_bh_tab()
        self._build_scrap_tab()

        # 計算結果暫存
        self._result = None

    def _make_scrollable(self, parent):
        """
        在 parent 內建立一個可垂直捲動的容器，回傳供呼叫端把元件放進去的 inner frame。
        當 inner frame 實際內容高度超過可視區域時，垂直捲軸自動可用；
        滑鼠停在容器上時也支援滾輪捲動，使用體驗跟一般可捲動視窗一致。
        """
        canvas = tk.Canvas(parent, bg=CLR_BG, highlightthickness=0)
        vsb = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=CLR_BG)

        inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)

        def on_inner_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        inner.bind("<Configure>", on_inner_configure)

        def on_canvas_configure(event):
            # inner frame 寬度跟著 canvas 一起變化，內容才會隨視窗縮放正確排版
            canvas.itemconfig(inner_id, width=event.width)
        canvas.bind("<Configure>", on_canvas_configure)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-event.delta/60), "units")
        canvas.bind("<MouseWheel>", on_mousewheel)
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        return inner

    def _btn(self, parent, text, cmd, color, side="left", padx=(4,4), font_size=10):
        b = tk.Button(parent, text=text, command=cmd,
                  bg=color, fg="white", activebackground=color,
                  font=("Microsoft JhengHei", font_size, "bold"),
                  relief="flat", cursor="hand2", padx=12, pady=6)
        b.pack(side=side, padx=padx)
        return b

    def _calc_btn_state(self, disabled=False):
        btn = getattr(self, "_calc_btn", None)
        if btn:
            if disabled:
                btn.config(state="disabled", text="⏳ 計算中...", bg="#AAAAAA")
            else:
                btn.config(state="normal", text="▶  執行配料計算", bg=CLR_BTN_MAIN)
        cancel_btn = getattr(self, "_cancel_btn", None)
        if cancel_btn:
            cancel_btn.config(state="normal" if disabled else "disabled")

    def _cancel_calc(self):
        event = getattr(self, "_calc_cancel_event", None)
        if event is not None:
            event.set()
            self._cancel_btn.config(state="disabled")
            self.status_var.set("⏳ 已要求取消，正在結束目前的配料步驟…")

    # ── Tab 1：前置參數 ───────────────────────────────────────────
    def _build_param_tab(self):
        f = self.tab_param_inner
        # 工程資訊
        sec = self._section(f, "工程資訊", 0)
        self._lbl_entry(sec, "案號：",            "proj_no",    "",  0)
        self._lbl_entry(sec, "案名：",            "proj_name",  "",  1)
        self._lbl_entry(sec, "日期(dd.mm.yyyy)：", "proj_date",  "",  2)
        self._lbl_entry(sec, "拆鈑字串：",        "split_code", "A", 3)

        # 材質
        sec2 = self._section(f, "材質設定", 1)
        tk.Label(sec2, text="材質：", bg=CLR_WHITE,
                 font=("Microsoft JhengHei", 10)).grid(row=0, column=0, sticky="w", padx=8, pady=4)
        self.var_mat = tk.StringVar(value="黑鐵板（碳鋼）")
        cb = ttk.Combobox(sec2, textvariable=self.var_mat,
                          values=list(DENSITY.keys()), width=22, state="readonly")
        cb.grid(row=0, column=1, sticky="w", padx=4, pady=4)
        self.lbl_density = tk.Label(sec2, text=f"比重：{DENSITY['黑鐵板（碳鋼）']}",
                                    bg=CLR_WHITE, font=("Microsoft JhengHei", 10), fg="#4A6C7A")
        self.lbl_density.grid(row=0, column=2, sticky="w", padx=12)
        cb.bind("<<ComboboxSelected>>",
                lambda e: self.lbl_density.config(
                    text=f"比重：{DENSITY[self.var_mat.get()]}"))

        # 切割損耗
        sec3 = self._section(f, "切割損耗設定（mm）", 2)
        self._lbl_entry(sec3, "新板每刀損耗：",   "new_kerf",  "1", 0)
        self._lbl_entry(sec3, "新板頭尾損耗（每端）：", "new_trim",  "1", 1)
        self._lbl_entry(sec3, "餘料每刀損耗：",   "scrap_kerf","1", 2)
        self._lbl_entry(sec3, "餘料頭尾損耗（每端）：", "scrap_trim","0", 3)

        # 切割模式
        sec3b = self._section(f, "切割模式", 3)
        self.var_cut_mode = tk.StringVar(value="multi")
        tk.Radiobutton(sec3b, text="單段切割　（同一張板只切一種長度）",
                       variable=self.var_cut_mode, value="single",
                       bg=CLR_WHITE, font=("Microsoft JhengHei", 10),
                       activebackground=CLR_WHITE
                       ).grid(row=0, column=0, sticky="w", padx=8, pady=4)
        tk.Radiobutton(sec3b, text="多段切割　（同一張板可切不同長度，最省料）",
                       variable=self.var_cut_mode, value="multi",
                       bg=CLR_WHITE, font=("Microsoft JhengHei", 10),
                       activebackground=CLR_WHITE
                       ).grid(row=1, column=0, sticky="w", padx=8, pady=4)
        tk.Label(sec3b, text="※ 多段切割需確認切割機支援橫向多刀",
                 bg=CLR_WHITE, fg="#C53030",
                 font=("Microsoft JhengHei", 9)).grid(row=2, column=0, sticky="w", padx=12, pady=(0,4))

        # 採購限制（可手動修改）
        sec4 = self._section(f, "採購限制（可手動修改）", 4)
        self._lbl_entry(sec4, "板寬最小值(mm)：", "bw_min", str(BW_MIN), 0)
        self._lbl_entry(sec4, "板寬最大值(mm)：", "bw_max", str(BW_MAX), 1)
        self._lbl_entry(sec4, "板長最小值(mm)：", "bl_min", str(BL_MIN),  2)
        self._lbl_entry(sec4, "板長最大值(mm)：", "bl_max", str(BL_MAX),  3)
        self._lbl_entry(sec4, "最低重量(kg)：",   "w_min",  str(W_MIN),   4)
        self._lbl_entry(sec4, "最高重量(kg)：",   "w_max",  str(W_MAX),   5)
        tk.Button(sec4, text="↺ 恢復預設值", command=self._reset_purchase_limits,
                  font=("Microsoft JhengHei", 9), bg="#E2E8F0", fg="#2D3748",
                  relief="flat", cursor="hand2", padx=10, pady=3
                  ).grid(row=0, column=2, rowspan=5, sticky="n", padx=(16,8), pady=4)

    def _reset_purchase_limits(self):
        """把「採購限制」六個欄位重設回程式內建的預設值，並提示使用者已重設成功。"""
        self.var_bw_min.set(str(BW_MIN))
        self.var_bw_max.set(str(BW_MAX))
        self.var_bl_min.set(str(BL_MIN))
        self.var_bl_max.set(str(BL_MAX))
        self.var_w_min.set(str(W_MIN))
        self.var_w_max.set(str(W_MAX))
        self.status_var.set("採購限制已恢復預設值。")

    def _section(self, parent, title, row_idx):
        frm = tk.LabelFrame(parent, text=f"  {title}  ",
                             bg=CLR_WHITE, font=("Microsoft JhengHei", 10, "bold"),
                             fg=CLR_HEADER, relief="groove", bd=2)
        frm.grid(row=row_idx, column=0, sticky="ew", padx=16, pady=(8,0))
        parent.columnconfigure(0, weight=1)
        return frm

    def _lbl_entry(self, parent, label, key, default, row):
        tk.Label(parent, text=label, bg=CLR_WHITE,
                 font=("Microsoft JhengHei", 10)).grid(row=row, column=0, sticky="w", padx=8, pady=4)
        var = tk.StringVar(value=default)
        setattr(self, f"var_{key}", var)
        tk.Entry(parent, textvariable=var, width=20,
                 font=("Microsoft JhengHei", 10)).grid(row=row, column=1, sticky="w", padx=4, pady=4)

    # ── Tab 2：BH配料表 ────────────────────────────────────────────
    def _build_bh_tab(self):
        f = self.tab_bh
        # 工具列
        toolbar = tk.Frame(f, bg=CLR_BG)
        toolbar.pack(fill="x", pady=(6,2), padx=8)
        self._btn(toolbar, "＋ 新增列",   self._bh_add_row, CLR_BTN_ADD)
        self._btn(toolbar, "✕ 刪除選取", self._bh_del_row, CLR_BTN_DEL, padx=(4,0), font_size=12)

        # 說明
        tk.Label(f, text="✏ 雙擊儲存格可編輯　|　斷面格式：BH高×寬×腹厚×翼厚",
                 bg=CLR_BG, fg="#718096", font=("Microsoft JhengHei", 9)
                 ).pack(anchor="w", padx=10)

        cols = ("構件編號","零件編號","斷面規格","長度(mm)","數量","材質")
        self.bh_tree = self._make_tree(f, cols, [80,80,200,90,70,70])

        # 雙擊編輯
        self.bh_tree.bind("<Double-1>", self._bh_edit_cell)

    def _download_blank_csv(self):
        """下載空白配料表 CSV（格式範本，直接嵌入程式不需外部檔案）"""
        import base64
        # 空白配料表 CSV（Big5編碼，嵌入base64）
        BLANK_CSV_B64 = (
            "ICAgILB0rsayTbPmLCwsLCwNCiAgICCu17i5OiAsLCwsLA0KICAgIK7Xplc6"
            "ICwsLCyk6bTBOiwNCiwsLCwsDQogICAgILpjpfO9c7i5ICAgICwgILlzpfO9"
            "c7i5ICAsICDCX62xs1eu5iAgICAgICAgICAgICAgICAgICAgICwgICAgqvir1y"
            "wsILzGtnEsICAgp/e96A0K"
        )
        path = filedialog.asksaveasfilename(
            title="儲存空白配料表",
            defaultextension=".csv",
            initialfile="BH配料表.csv",
            filetypes=[("CSV 檔案","*.csv"),("所有檔案","*.*")])
        if not path:
            return
        try:
            data = base64.b64decode(BLANK_CSV_B64)
            with open(path, 'wb') as f:
                f.write(data)
            messagebox.showinfo("下載完成", f"空白配料表已儲存：\n{path}")
        except Exception as e:
            messagebox.showerror("下載失敗", str(e))

    def _make_tree(self, parent, cols, widths):
        style = ttk.Style()
        style.configure("Custom.Treeview", rowheight=24,
                        font=("Microsoft JhengHei", 10),
                        background=CLR_WHITE, fieldbackground=CLR_WHITE)
        style.configure("Custom.Treeview.Heading",
                        font=("Microsoft JhengHei", 10, "bold"),
                        background="#D0D8E4", foreground="#1A3050")
        style.map("Custom.Treeview",
                  background=[("selected", "#BEE3F8")])

        frame = tk.Frame(parent, bg=CLR_BG)
        frame.pack(fill="both", expand=True, padx=8, pady=(0,4))
        vsb = ttk.Scrollbar(frame, orient="vertical")
        hsb = ttk.Scrollbar(frame, orient="horizontal")
        tree = ttk.Treeview(frame, columns=cols, show="headings",
                             style="Custom.Treeview",
                             yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)
        for col, w in zip(cols, widths):
            tree.heading(col, text=col)
            tree.column(col, width=w, anchor="center")
        tree.tag_configure("odd",  background=CLR_ROW_ODD)
        tree.tag_configure("even", background=CLR_ROW_EVEN)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return tree

    def _bh_add_row(self):
        dlg = BHRowDialog(self, title="新增 BH 構件")
        if dlg.result:
            self._bh_insert(dlg.result)
        self._refresh_tags(self.bh_tree)

    def _bh_insert(self, d):
        n = len(self.bh_tree.get_children())
        tag = "odd" if n % 2 == 0 else "even"
        self.bh_tree.insert("", "end",
                             values=(d["comp"], d["part"], d["spec"],
                                     d["length"], d["qty"], d["mat"]),
                             tags=(tag,))

    def _bh_del_row(self):
        sel = self.bh_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "請先選取要刪除的列。")
            return
        for item in sel:
            self.bh_tree.delete(item)
        self._refresh_tags(self.bh_tree)

    def _bh_edit_cell(self, event):
        item = self.bh_tree.identify_row(event.y)
        col  = self.bh_tree.identify_column(event.x)
        if not item or not col:
            return
        col_idx = int(col.replace("#","")) - 1
        vals = list(self.bh_tree.item(item, "values"))
        col_name = self.bh_tree["columns"][col_idx]
        new_val  = tk.simpledialog.askstring(
            "編輯", f"修改「{col_name}」：", initialvalue=vals[col_idx], parent=self)
        if new_val is not None:
            vals[col_idx] = new_val
            self.bh_tree.item(item, values=vals)

    # ── Tab 3：現有餘料 ────────────────────────────────────────────
    def _build_scrap_tab(self):
        f = self.tab_scrap
        toolbar = tk.Frame(f, bg=CLR_BG)
        toolbar.pack(fill="x", pady=(6,2), padx=8)
        self._btn(toolbar, "＋ 新增餘料",    self._sc_add_row,    CLR_BTN_ADD)
        self._btn(toolbar, "✕ 刪除選取",    self._sc_del_row,    CLR_BTN_DEL, padx=(4,0), font_size=12)

        tk.Label(f, text="✏ 雙擊儲存格可編輯　|　格式：PL厚×寬×長×數量片(材質)",
                 bg=CLR_BG, fg="#718096", font=("Microsoft JhengHei", 9)
                 ).pack(anchor="w", padx=10)

        cols = ("編號","厚度(mm)","寬度(mm)","長度(mm)","數量(片)","材質")
        self.sc_tree = self._make_tree(f, cols, [60,90,90,90,80,80])
        self.sc_tree.bind("<Double-1>", self._sc_edit_cell)

    def _sc_add_row(self):
        dlg = ScrapDialog(self, title="新增現有餘料")
        if dlg.result:
            n   = len(self.sc_tree.get_children()) + 1
            tag = "odd" if n % 2 == 0 else "even"
            d   = dlg.result
            self.sc_tree.insert("", "end",
                                 values=(f"既{n:02d}", d["thick"], d["width"],
                                         d["length"], d["qty"], d["mat"]),
                                 tags=(tag,))
        self._refresh_tags(self.sc_tree)

    def _sc_del_row(self):
        sel = self.sc_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "請先選取要刪除的列。")
            return
        for item in sel:
            self.sc_tree.delete(item)
        self._refresh_tags(self.sc_tree)

    def _sc_edit_cell(self, event):
        item = self.sc_tree.identify_row(event.y)
        col  = self.sc_tree.identify_column(event.x)
        if not item or not col:
            return
        col_idx  = int(col.replace("#","")) - 1
        if col_idx == 0:
            return  # 編號不可改
        vals     = list(self.sc_tree.item(item, "values"))
        col_name = self.sc_tree["columns"][col_idx]
        new_val  = tk.simpledialog.askstring(
            "編輯", f"修改「{col_name}」：", initialvalue=vals[col_idx], parent=self)
        if new_val is not None:
            vals[col_idx] = new_val
            self.sc_tree.item(item, values=vals)

    def _sc_import_csv(self):
        """從 CSV 檔案匯入現有餘料，追加到現有清單後面（不覆蓋）"""
        path = filedialog.askopenfilename(
            title="選擇餘料 CSV 檔案",
            filetypes=[("CSV 檔案","*.csv"),("所有檔案","*.*")])
        if not path:
            return
        imported = 0
        errors   = []
        for enc in ["utf-8-sig", "big5", "cp950", "utf-8"]:
            try:
                import csv
                with open(path, encoding=enc, newline="") as f:
                    reader = csv.DictReader(f)
                    rows = list(reader)
                break
            except Exception:
                rows = None
        if not rows:
            messagebox.showerror("錯誤", "無法讀取 CSV，請確認檔案格式與編碼。")
            return

        for i, row in enumerate(rows, 1):
            try:
                # 支援兩種欄位名稱：中文標題或英文欄位名
                thick  = int(row.get("厚度(mm)", row.get("thick", "")))
                width  = int(row.get("寬度(mm)", row.get("width", "")))
                length = int(row.get("長度(mm)", row.get("length", "")))
                qty    = int(row.get("數量(片)", row.get("qty", 1)))
                mat    = str(row.get("材質",    row.get("mat", "黑鐵板（碳鋼）"))).strip()
                if min(thick, width, length, qty) <= 0:
                    raise ValueError("厚度、寬度、長度與數量都必須大於 0")
                n      = len(self.sc_tree.get_children()) + 1
                tag    = "odd" if n % 2 == 0 else "even"
                self.sc_tree.insert("", "end",
                                     values=(f"既{n:02d}", thick, width, length, qty, mat),
                                     tags=(tag,))
                imported += 1
            except Exception as e:
                errors.append(f"第{i}列：{e}")

        self._refresh_tags(self.sc_tree)
        msg = f"成功匯入 {imported} 筆餘料。"
        if errors:
            msg += f"\n另有 {len(errors)} 列未匯入，請開啟錯誤明細查看。"
        messagebox.showinfo("匯入完成", msg)
        if errors:
            self._show_import_errors("餘料 CSV 匯入錯誤", errors)

    def _sc_download_blank(self):
        """下載空白餘料CSV範本，讓使用者填寫後再匯入"""
        path = filedialog.asksaveasfilename(
            title="儲存空白餘料表",
            defaultextension=".csv",
            initialfile="現有餘料表.csv",
            filetypes=[("CSV 檔案","*.csv"),("所有檔案","*.*")])
        if not path:
            return
        try:
            import csv
            with open(path, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["厚度(mm)", "寬度(mm)", "長度(mm)", "數量(片)", "材質"])
                # 寫入兩列範例資料，幫助使用者了解填寫格式
                writer.writerow([16, 744, 9000, 1, "SN490B"])
                writer.writerow([28, 350, 6000, 2, "黑鐵板（碳鋼）"])
            messagebox.showinfo("下載完成", f"空白餘料表已儲存：\n{path}\n\n請填寫後再用「匯入餘料CSV」按鈕載入。")
        except Exception as e:
            messagebox.showerror("下載失敗", str(e))

    # ── 輔助 ──────────────────────────────────────────────────────
    def _refresh_tags(self, tree):
        for i, item in enumerate(tree.get_children()):
            tree.item(item, tags=("odd" if i%2==0 else "even",))

    def _get_params(self):
        try:
            params = {
                "proj_no":    self.var_proj_no.get().strip(),
                "proj_name":  self.var_proj_name.get().strip(),
                "proj_date":  self.var_proj_date.get().strip(),
                "split_code": self.var_split_code.get().strip() or "A",
                "mat_name":   self.var_mat.get(),
                "density":    DENSITY[self.var_mat.get()],
                "new_kerf":   int(self.var_new_kerf.get()),
                "new_trim":   int(self.var_new_trim.get()),
                "scrap_kerf": int(self.var_scrap_kerf.get()),
                "scrap_trim": int(self.var_scrap_trim.get()),
                "bw_min":     int(self.var_bw_min.get()),
                "bw_max":     int(self.var_bw_max.get()),
                "bl_min":     int(self.var_bl_min.get()),
                "bl_max":     int(self.var_bl_max.get()),
                "w_min":      int(self.var_w_min.get()),
                "w_max":      int(self.var_w_max.get()),
                "cut_mode":   self.var_cut_mode.get(),
            }
            if any(params[k] < 0 for k in ("new_kerf", "new_trim", "scrap_kerf", "scrap_trim")):
                raise ValueError("鋸縫與修邊尺寸不可小於 0。")
            if min(params["bw_min"], params["bw_max"],
                   params["bl_min"], params["bl_max"]) <= 0:
                raise ValueError("板寬與板長限制必須大於 0。")
            if params["w_min"] < 0 or params["w_max"] <= 0:
                raise ValueError("重量限制必須為正數（最低重量可為 0）。")
            if params["bw_min"] > params["bw_max"] or params["bl_min"] > params["bl_max"]:
                raise ValueError("板材最小尺寸不可大於最大尺寸。")
            if params["w_min"] > params["w_max"]:
                raise ValueError("最低重量不可大於最高重量。")
            return params
        except ValueError as e:
            messagebox.showerror("參數錯誤", f"數值欄位必須為整數：{e}")
            return None

    def _get_bh_rows(self):
        rows = []
        for item in self.bh_tree.get_children():
            v = self.bh_tree.item(item, "values")
            try:
                row = {
                    "comp": v[0], "part": v[1], "spec": v[2],
                    "length": int(v[3]), "qty": int(v[4]), "mat": v[5]
                }
                if row["length"] <= 0 or row["qty"] <= 0:
                    raise ValueError("長度與數量必須大於 0。")
                rows.append(row)
            except (ValueError, IndexError) as e:
                messagebox.showerror("資料錯誤", f"BH配料表資料有誤：{e}\n{v}")
                return None
        return rows

    def _get_scraps(self):
        scraps = []
        for item in self.sc_tree.get_children():
            v = self.sc_tree.item(item, "values")
            try:
                thick, width, length, qty = map(int, (v[1], v[2], v[3], v[4]))
                if min(thick, width, length, qty) <= 0:
                    raise ValueError("厚度、寬度、長度與數量都必須大於 0。")
            except (ValueError, IndexError) as e:
                messagebox.showerror("餘料資料錯誤", f"{e}\n{v}")
                return None

            # 拆成單片並給每片唯一編號，才能正確表示同一批餘料只用掉部分數量。
            for piece_no in range(1, qty + 1):
                piece_id = f"{v[0]}-{item}-{piece_no:02d}"
                scraps.append({
                    "id": piece_id, "thick": thick, "width": width,
                    "length": length, "qty": 1, "mat": v[5]
                })
        return scraps

    # ── 執行計算 ──────────────────────────────────────────────────
    def _run_calc(self):
        params = self._get_params()
        if not params:
            return
        bh_rows = self._get_bh_rows()
        if bh_rows is None or len(bh_rows) == 0:
            messagebox.showwarning("警告", "請先在「BH配料表」標籤頁輸入構件資料。")
            return

        scraps = self._get_scraps()
        if scraps is None:
            return

        # 拆板
        decomposed = []
        for row in bh_rows:
            d = decompose_bh(row, params["density"],
                             params["new_kerf"], params["new_trim"],
                             params["scrap_kerf"], params["scrap_trim"], scraps)
            if d is None:
                messagebox.showerror("格式錯誤",
                    f"無法解析 BH 斷面規格：{row['spec']}\n請確認格式為 BH高×寬×腹厚×翼厚")
                return
            decomposed.append(d)

        # 加入拆鈑流水號
        # 規則：前置串 + - + 依相同(厚×寬×長)共用同一號，按出現順序編排
        split_code = params.get("split_code", "A")

        def build_serials(items):
            """
            items: list of (thick, width, length, obj)
            相同 (thick, width, length) 共用同一流水號，按出現順序給號
            """
            dim_counter = {}  # (thick, width, length) → 編號
            idx = 1
            for thick, width, length, _ in items:
                key = (thick, width, length)
                if key not in dim_counter:
                    dim_counter[key] = idx
                    idx += 1
            for thick, width, length, obj in items:
                num = dim_counter[(thick, width, length)]
                obj["serial"] = f"{split_code}-{num}"

        build_serials([
            (d["flange"]["thick"], d["flange"]["width"], d["flange"]["length"], d["flange"])
            for d in decomposed
        ])
        build_serials([
            (d["web"]["thick"], d["web"]["width"], d["web"]["length"], d["web"])
            for d in decomposed
        ])

        # 整理零件列表
        parts_list = []
        for d in decomposed:
            parts_list.append(d["flange"])
            parts_list.append(d["web"])

        # 避免把超出可採購板材最大尺寸的零件硬塞進較短/較窄的板材。
        too_large = next((p for p in parts_list
                          if p["width"] + 2 * params["new_trim"] > params["bw_max"]
                          or p["length"] + 2 * params["new_trim"] > params["bl_max"]), None)
        if too_large:
            messagebox.showerror(
                "尺寸超出限制",
                f"零件 {too_large['name']} 尺寸為 {too_large['width']}×{too_large['length']} mm，"
                f"加上修邊後超出設定的板材最大尺寸 {params['bw_max']}×{params['bl_max']} mm。")
            return

        # 移到背景執行緒，避免計算時卡住視窗
        self.status_var.set("⏳ 配料最佳化：準備中…")
        self._calc_btn_state(disabled=True)
        self.update_idletasks()
        import threading
        self._calc_cancel_event = threading.Event()
        cancel_event = self._calc_cancel_event
        self._result = None

        def _progress(done, total):
            if cancel_event.is_set():
                return
            self.after(0, lambda d=done, t=total: self.status_var.set(
                f"⏳ 配料最佳化：{d}/{t} 次（{d * 100 // t}%）"))

        def _do_calc():
            try:
                result = plan_purchase(
                    parts_list, params["density"],
                    params["new_kerf"], params["new_trim"],
                    params["scrap_kerf"], params["scrap_trim"], scraps,
                    bw_min=params["bw_min"], bw_max=params["bw_max"],
                    bl_max=params["bl_max"], bl_min=params["bl_min"],
                    w_min=params["w_min"], w_max=params["w_max"],
                    cut_mode=params["cut_mode"], progress_callback=_progress,
                    cancel_event=cancel_event)
                if result is None:
                    self.after(0, _on_cancelled)
                else:
                    pl, cd, ns, ui = result
                    self.after(0, lambda: _on_done(pl, cd, ns, ui))
            except Exception as e:
                self.after(0, lambda: _on_error(str(e)))

        def _on_done(purchase_list, cut_details, new_scraps, used_scrap_ids):
            self._calc_btn_state(disabled=False)
            self._calc_cancel_event = None
            self._result = {
                "params": params, "bh_rows": bh_rows,
                "decomposed": decomposed,
                "purchase_list": purchase_list, "cut_details": cut_details,
                "new_scraps": new_scraps, "existing_scraps": scraps,
                "used_scrap_ids": set(used_scrap_ids)
            }
            self._purchase_edit_rows   = None
            self._purchase_edit_result = None
            total_buy = sum(p["total_wt"] for p in purchase_list)
            msg = (f"✅ 計算完成！\n\n"
                   f"  BH 構件：{len(bh_rows)} 項\n"
                   f"  採購新板：{len(purchase_list)} 片　合計 {total_buy:,} kg\n"
                   f"  切割片次：{len(cut_details)} 次\n"
                   f"  產生餘料：{len(new_scraps)} 件\n"
                   f"  使用既有餘料：{len(used_scrap_ids)} 件\n\n"
                   f"請點擊「💾 匯出 Excel」儲存報表。")
            messagebox.showinfo("配料完成", msg)
            self.status_var.set(
                f"計算完成｜採購 {len(purchase_list)} 片，{total_buy:,} kg｜"
                f"切割 {len(cut_details)} 次｜餘料 {len(new_scraps)} 件")

        def _on_error(err_msg):
            self._calc_btn_state(disabled=False)
            self._calc_cancel_event = None
            messagebox.showerror("計算失敗", err_msg)
            self.status_var.set("計算失敗，請檢查輸入資料。")

        def _on_cancelled():
            self._calc_btn_state(disabled=False)
            self._calc_cancel_event = None
            self.status_var.set("配料計算已取消。")

        threading.Thread(target=_do_calc, daemon=True).start()

    # ── 匯入 CSV ──────────────────────────────────────────────────
    def _show_import_errors(self, title, errors):
        if not errors:
            return
        win = tk.Toplevel(self)
        win.title(title)
        win.geometry("760x480")
        win.transient(self)
        tk.Label(win, text=f"共 {len(errors)} 筆資料未匯入；可選取並複製下方明細。",
                 anchor="w", font=("Microsoft JhengHei", 10, "bold")).pack(fill="x", padx=10, pady=8)
        body = tk.Frame(win)
        body.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        ybar = ttk.Scrollbar(body, orient="vertical")
        text = tk.Text(body, wrap="none", yscrollcommand=ybar.set,
                       font=("Consolas", 10), height=18)
        ybar.config(command=text.yview)
        text.pack(side="left", fill="both", expand=True)
        ybar.pack(side="right", fill="y")
        text.insert("1.0", "\n".join(errors))
        text.config(state="disabled")
        bar = tk.Frame(win)
        bar.pack(fill="x", padx=10, pady=(0, 10))

        def save_report():
            path = filedialog.asksaveasfilename(
                parent=win, title="儲存匯入錯誤明細", defaultextension=".txt",
                initialfile="CSV匯入錯誤明細.txt",
                filetypes=[("文字檔案", "*.txt"), ("所有檔案", "*.*")])
            if not path:
                return
            try:
                with open(path, "w", encoding="utf-8-sig") as f:
                    f.write("\n".join(errors))
                messagebox.showinfo("儲存完成", f"錯誤明細已儲存：\n{path}", parent=win)
            except Exception as e:
                messagebox.showerror("儲存失敗", str(e), parent=win)

        tk.Button(bar, text="儲存錯誤明細", command=save_report).pack(side="left")
        tk.Button(bar, text="關閉", command=win.destroy).pack(side="right")

    def _import_csv(self):
        import csv
        path = filedialog.askopenfilename(
            title="選擇 BH配料表 CSV",
            filetypes=[("CSV 檔案","*.csv"),("所有檔案","*.*")])
        if not path:
            return
        try:
            # 嘗試多種編碼讀取全部列
            rows_raw = None
            for enc in ("big5", "cp950", "utf-8-sig", "utf-8"):
                try:
                    with open(path, encoding=enc, newline="") as f:
                        rows_raw = list(csv.reader(f))
                    break
                except Exception:
                    continue
            if rows_raw is None:
                messagebox.showerror("錯誤", "無法讀取 CSV 檔案，請確認編碼格式。")
                return

            # ── 從 CSV 標題區讀取案號、案名、日期 ──
            # 格式：第2行含「案號:」、第3行含「案名:」與「日期:」
            import re
            for row in rows_raw[:8]:
                row_str = ",".join(str(c) for c in row)
                # 案號：「案號: XXXX」或「案號：XXXX」
                m = re.search(r'案號[：:]\s*([^\s,]+)', row_str)
                if m and m.group(1).strip():
                    self.var_proj_no.set(m.group(1).strip())
                # 案名：「案名: XXXX」
                m = re.search(r'案名[：:]\s*([^,日期]+)', row_str)
                if m and m.group(1).strip():
                    self.var_proj_name.set(m.group(1).strip())
                # 日期：「日期: XXXX」或「日期：XXXX」
                m = re.search(r'日期[：:]\s*([^\s,]+)', row_str)
                if m and m.group(1).strip():
                    self.var_proj_date.set(m.group(1).strip())

            # 搜尋標頭列（含「構件編號」或「零件編號」的那行）
            header_idx = None
            for i, row in enumerate(rows_raw):
                if any(c.strip() in ("構件編號", "零件編號") for c in row):
                    header_idx = i
                    break
            if header_idx is None:
                messagebox.showerror("錯誤", "找不到標頭列（構件編號、零件編號）\n請確認 CSV 格式正確。")
                return

            # 取標頭與資料列
            headers = [c.strip() for c in rows_raw[header_idx]]
            data_rows = rows_raw[header_idx + 1:]

            def get_col(row, *names):
                for name in names:
                    if name in headers:
                        idx = headers.index(name)
                        if idx < len(row):
                            return row[idx].strip()
                return ""

            # 先完整解析，成功後才清除舊資料
            parsed_rows = []
            errors = []
            for csv_line, row in enumerate(data_rows, header_idx + 2):
                if not any(c.strip() for c in row):
                    continue
                comp = get_col(row, "構件編號")
                if not comp:
                    errors.append(f"第 {csv_line} 列：缺少構件編號。")
                    continue
                part   = get_col(row, "零件編號")
                spec   = get_col(row, "斷面規格")
                length = get_col(row, "長度", "長度(mm)")
                qty    = get_col(row, "數量")
                mat    = get_col(row, "材質") or "A36"
                if not spec or not length:
                    errors.append(f"第 {csv_line} 列：缺少斷面規格或長度。")
                    continue
                try:
                    length_val = int(float(length))
                    qty_val    = int(float(qty)) if qty else 1
                    if length_val <= 0 or qty_val <= 0:
                        errors.append(f"第 {csv_line} 列：長度與數量必須大於 0。")
                        continue
                except ValueError as e:
                    errors.append(f"第 {csv_line} 列：長度或數量格式錯誤（{e}）。")
                    continue
                if not parse_bh(spec):
                    errors.append(f"第 {csv_line} 列：斷面規格格式錯誤：{spec}")
                    continue
                parsed_rows.append({"comp": comp, "part": part, "spec": spec,
                                    "length": length_val, "qty": qty_val, "mat": mat})

            if not parsed_rows:
                messagebox.showerror("錯誤", f"沒有有效資料可匯入（錯誤 {len(errors)} 筆）。")
                if errors:
                    self._show_import_errors("BH CSV 匯入錯誤", errors)
                return

            for item in self.bh_tree.get_children():
                self.bh_tree.delete(item)
            for r in parsed_rows:
                self._bh_insert(r)

            self._refresh_tags(self.bh_tree)
            msg = f"成功匯入 {len(parsed_rows)} 筆構件。"
            if errors:
                msg += f"\n另有 {len(errors)} 列未匯入，請開啟錯誤明細查看。"
            self.status_var.set(f"已匯入 {len(parsed_rows)} 筆 BH 構件資料。")
            messagebox.showinfo("匯入成功", msg)
            if errors:
                self._show_import_errors("BH CSV 匯入錯誤", errors)
        except Exception as e:
            messagebox.showerror("匯入失敗", str(e))

    # ── 匯出 Excel ────────────────────────────────────────────────
    def _export_xlsx(self):
        if not self._result:
            messagebox.showwarning("提示", "請先執行配料計算。")
            return
        r = self._result
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 活頁簿","*.xlsx")],
            initialfile=f"BH_配料報表_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
        if not path:
            return
        try:
            write_xlsx(
                path,
                r["params"]["proj_no"], r["params"]["proj_name"],
                r["params"]["proj_date"], r["params"]["mat_name"],
                r["params"]["density"],
                r["params"]["new_kerf"], r["params"]["new_trim"],
                r["params"]["scrap_kerf"], r["params"]["scrap_trim"],
                r["bh_rows"], r["decomposed"],
                r["purchase_list"], r["cut_details"],
                r["new_scraps"], r["existing_scraps"], r["used_scrap_ids"]
            )
            self.status_var.set(f"Excel 已儲存：{path}")
            messagebox.showinfo("儲存成功", f"報表已儲存至：\n{path}")
        except Exception as e:
            messagebox.showerror("儲存失敗", str(e))

    # ── 預覽結果 ──────────────────────────────────────────────────
    def _preview(self):
        if not self._result:
            messagebox.showwarning("提示", "請先執行配料計算。")
            return
        PreviewWindow(self, self._result)

    # ── 排列圖預覽 ────────────────────────────────────────────────
    def _preview_layout(self):
        if not self._result:
            messagebox.showwarning("提示", "請先執行配料計算。")
            return
        import tempfile, webbrowser
        r = self._result
        try:
            self.status_var.set("⏳ 正在產生排列圖預覽...")
            self.update_idletasks()
            bh_rows  = r.get("bh_rows", [])
            mats     = list(dict.fromkeys(row.get("mat","") for row in bh_rows if row.get("mat")))
            mat_name = "、".join(mats) if mats else r["params"]["mat_name"]
            serial_map = {}
            for d in r.get("decomposed", []):
                fl_sn = d["flange"].get("serial","")
                wb_sn = d["web"].get("serial","")
                serial_map[d["flange"]["name"]] = f"{fl_sn}F" if fl_sn else ""
                serial_map[d["web"]["name"]]    = f"{wb_sn}W" if wb_sn else ""
            tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            tmp.close()
            write_layout_pdf(
                tmp.name,
                r["params"]["proj_no"],
                r["params"]["proj_name"],
                mat_name,
                r["cut_details"],
                r["new_scraps"],
                serial_map=serial_map
            )
            os.startfile(tmp.name)
            self.status_var.set("排列圖預覽已開啟。")
        except Exception as e:
            messagebox.showerror("預覽失敗", str(e))

    # ── 排列圖快速預覽 ────────────────────────────────────────────
    def _show_layouts(self):
        if not self._result:
            messagebox.showwarning("提示", "請先執行配料計算。")
            return
        layouts = [ct for ct in self._result["cut_details"] if ct.get("layout")]
        if not layouts:
            messagebox.showinfo("提示", "目前沒有切割片次，無排列圖可預覽。")
            return
        # 直接點「排列圖」按鈕不會經過 PreviewWindow 的「餘料清單」分頁，
        # 所以這裡要先補上一次編號賦予，確保排列圖能正確顯示餘NO。
        assign_scrap_numbers(self._result)

        dlg = tk.Toplevel(self)
        dlg.title("選擇要預覽／下載的排列圖")
        dlg.geometry("620x480")
        dlg.minsize(480, 360)
        dlg.resizable(True, True)   # 可自由拉伸放大縮小
        dlg.configure(bg=CLR_BG)

        tk.Label(dlg, text=f"共 {len(layouts)} 個片次　"
                            f"勾選後可批次下載到指定資料夾，或雙擊單一列開啟預覽",
                 font=("Microsoft JhengHei", 10, "bold"),
                 bg=CLR_HEADER, fg="white", wraplength=600,
                 justify="left").pack(fill="x", ipady=8)

        # 全選 / 取消全選 工具列
        tool_bar = tk.Frame(dlg, bg=CLR_BG)
        tool_bar.pack(fill="x", padx=10, pady=(8, 0))
        tk.Button(tool_bar, text="☑ 全選", command=lambda: toggle_all(True),
                  font=("Microsoft JhengHei", 9), bg="#E2E8F0",
                  relief="flat", padx=10).pack(side="left", padx=(0, 4))
        tk.Button(tool_bar, text="☐ 取消全選", command=lambda: toggle_all(False),
                  font=("Microsoft JhengHei", 9), bg="#E2E8F0",
                  relief="flat", padx=10).pack(side="left")
        self._layout_sel_count_var = tk.StringVar(value="已選 0 項")
        tk.Label(tool_bar, textvariable=self._layout_sel_count_var,
                 font=("Microsoft JhengHei", 9), bg=CLR_BG, fg="#555"
                 ).pack(side="right")

        frame = tk.Frame(dlg, bg=CLR_BG)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        cols = ("選取", "片次", "板別", "採購規格", "排數", "構件")
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=12)
        widths = [44, 55, 60, 170, 50, 150]
        for c, w in zip(cols, widths):
            tree.heading(c, text=c)
            tree.column(c, width=w, anchor="center")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        # checked：用 set 記錄目前已勾選的片次 idx，作為勾選狀態的單一真實來源
        checked = set()

        for ct in layouts:
            n_rows = len(ct["layout"]["rows"])
            tree.insert("", "end", iid=ct["idx"],
                        values=("☐", ct["idx"], "F=翼" if ct["type"]=="F" else "W=腹",
                                ct["board_spec"], n_rows, ct["comp"]))

        def update_sel_count():
            self._layout_sel_count_var.set(f"已選 {len(checked)} 項")

        def toggle_check(item):
            if item in checked:
                checked.discard(item)
                tree.set(item, "選取", "☐")
            else:
                checked.add(item)
                tree.set(item, "選取", "☑")
            update_sel_count()

        def toggle_all(check_on):
            for item in tree.get_children():
                if check_on:
                    checked.add(item)
                    tree.set(item, "選取", "☑")
                else:
                    checked.discard(item)
                    tree.set(item, "選取", "☐")
            update_sel_count()

        def on_tree_click(event):
            # 只在點擊「選取」欄位（第一欄）時切換勾選，避免點擊其他欄位也誤觸發
            region = tree.identify("region", event.x, event.y)
            if region != "cell":
                return
            col = tree.identify_column(event.x)
            item = tree.identify_row(event.y)
            if not item:
                return
            if col == "#1":   # 第一欄＝「選取」
                toggle_check(item)

        def open_selected(event=None):
            item = tree.focus()
            if not item: return
            ct = next(c for c in layouts if c["idx"] == item)
            LayoutWindow(dlg, ct["idx"], ct["layout"])

        tree.bind("<Button-1>", on_tree_click)
        tree.bind("<Double-1>", open_selected)

        # 底部按鈕列：開啟單張預覽 + 批次下載PNG + 匯出合併PDF
        btn_bar = tk.Frame(dlg, bg=CLR_BG)
        btn_bar.pack(fill="x", padx=10, pady=(0, 10))
        tk.Button(btn_bar, text="🔍 開啟預覽", command=open_selected,
                  bg="#3D6B6B", fg="white", font=("Microsoft JhengHei", 10, "bold"),
                  relief="flat", pady=6).pack(side="left", fill="x", expand=True, padx=(0, 2))
        tk.Button(btn_bar, text="💾 批次下載 PNG",
                  command=lambda: self._batch_download_layouts(
                      [c for c in layouts if c["idx"] in checked]),
                  bg=CLR_BTN_MAIN, fg="white", font=("Microsoft JhengHei", 10, "bold"),
                  relief="flat", pady=6).pack(side="left", fill="x", expand=True, padx=(2, 2))
        tk.Button(btn_bar, text="📄 匯出合併 PDF",
                  command=lambda: self._export_layouts_pdf(
                      [c for c in layouts if c["idx"] in checked]),
                  bg="#6B46C1", fg="white", font=("Microsoft JhengHei", 10, "bold"),
                  relief="flat", pady=6).pack(side="left", fill="x", expand=True, padx=(2, 0))

    def _batch_download_layouts(self, selected_layouts):
        """
        把勾選的多個排列圖一次性下載到使用者選擇的資料夾，
        每個片次各自輸出一張 PNG，檔名以片次編號命名（特殊符號如 * 會轉成安全字元）。
        """
        if not selected_layouts:
            messagebox.showwarning("提示", "請至少勾選一個片次再下載。")
            return
        try:
            from PIL import Image
        except ImportError:
            messagebox.showerror("缺少套件",
                "下載圖片需要 Pillow 套件，請先安裝：\n\npip install Pillow")
            return

        folder = filedialog.askdirectory(title="選擇排列圖儲存資料夾")
        if not folder:
            return

        success, failed = [], []
        for ct in selected_layouts:
            safe_idx = str(ct["idx"]).replace("*", "星").replace("/", "_")
            out_path = os.path.join(folder, f"BH排列圖_{safe_idx}.png")
            try:
                img = layout_to_pil_image(ct["layout"], zoom=1.0)
                img.save(out_path, "PNG")
                success.append(safe_idx)
            except Exception as e:
                failed.append(f"{safe_idx}（{e}）")

        msg = f"成功下載 {len(success)} 張排列圖至：\n{folder}"
        if failed:
            msg += f"\n\n以下 {len(failed)} 張失敗：\n" + "\n".join(failed)
            messagebox.showwarning("部分下載失敗", msg)
        else:
            messagebox.showinfo("批次下載完成", msg)

    def _export_layouts_pdf(self, selected_layouts):
        """
        把勾選的多個排列圖合併成一份 PDF，每張各佔一頁。
        關鍵策略：PDF 頁面尺寸直接等於排列圖圖片的實際像素大小（px → pt，72dpi），
        不把圖片縮進固定頁面——確保文字大小與畫面顯示完全一致，不會因縮放而變小。
        """
        if not selected_layouts:
            messagebox.showwarning("提示", "請至少勾選一個片次再匯出。")
            return
        try:
            from PIL import Image as PILImage
            from reportlab.pdfgen import canvas as rl_canvas
            from reportlab.lib.utils import ImageReader
            import io
        except ImportError as e:
            messagebox.showerror("缺少套件", f"匯出PDF需要 Pillow 與 reportlab：\n{e}")
            return

        path = filedialog.asksaveasfilename(
            title="匯出排列圖 PDF",
            defaultextension=".pdf",
            initialfile=f"BH排列圖_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            filetypes=[("PDF 檔案", "*.pdf"), ("所有檔案", "*.*")]
        )
        if not path:
            return

        try:
            # DPI 設定：PDF 使用 72pt/inch，以 144dpi 渲染圖片（scale_up=2）
            # 這樣圖片像素 ÷ 2 = PDF 點數，等比例無失真，文字看起來跟螢幕一樣大
            RENDER_DPI = 144   # 144 = 72 * scale_up(2)
            PDF_PPI    = 72    # PDF 標準 72pt/inch

            c = rl_canvas.Canvas(path)

            for pg_idx, ct in enumerate(selected_layouts):
                # 渲染成 PIL Image（scale_up=2 對應 144dpi）
                img_pil = layout_to_pil_image(ct["layout"], zoom=1.0, scale_up=2)
                img_w_px, img_h_px = img_pil.size

                # 計算 PDF 頁面尺寸（px → pt，比例 = PDF_PPI / RENDER_DPI = 0.5）
                scale = PDF_PPI / RENDER_DPI
                page_w_pt = img_w_px * scale
                page_h_pt = img_h_px * scale

                # 設定本頁頁面尺寸（每頁可以不同大小）
                c.setPageSize((page_w_pt, page_h_pt))

                # 把 PIL Image 轉成 BytesIO 後貼到 PDF 頁面
                buf = io.BytesIO()
                img_pil.save(buf, format="PNG")
                buf.seek(0)

                # reportlab 座標系原點在左下角，圖片貼滿整頁
                c.drawImage(ImageReader(buf),
                            x=0, y=0,
                            width=page_w_pt, height=page_h_pt,
                            preserveAspectRatio=False)

                # 換頁（最後一頁不需要）
                c.showPage()

            c.save()
            messagebox.showinfo("匯出完成",
                f"已成功匯出 {len(selected_layouts)} 張排列圖至：\n{path}\n\n"
                f"每頁尺寸依排列圖實際大小自動調整，文字大小與畫面顯示一致。")

        except Exception as e:
            messagebox.showerror("匯出失敗", f"PDF 匯出失敗：{e}")

    # ── 匯出 PDF ──────────────────────────────────────────────────
    def _export_pdf(self):
        if not self._result:
            messagebox.showwarning("提示", "請先執行配料計算。")
            return
        r = self._result
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF 檔案","*.pdf")],
            initialfile=f"BH_配料報表_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf")
        if not path:
            return
        try:
            self.status_var.set("⏳ 正在產生 PDF...")
            self.update_idletasks()
            write_pdf(
                path,
                r["params"]["proj_no"], r["params"]["proj_name"],
                r["params"]["proj_date"], r["params"]["mat_name"],
                r["params"]["density"],
                r["params"]["new_kerf"], r["params"]["new_trim"],
                r["params"]["scrap_kerf"], r["params"]["scrap_trim"],
                r["bh_rows"], r["decomposed"],
                r["purchase_list"], r["cut_details"],
                r["new_scraps"], r["existing_scraps"], r["used_scrap_ids"]
            )
            self.status_var.set(f"PDF 已儲存：{path}")
            messagebox.showinfo("儲存成功", f"PDF 已儲存至：\n{path}")
        except Exception as e:
            messagebox.showerror("儲存失敗", str(e))

    # ── 匯出排列圖 PDF ────────────────────────────────────────────
    def _export_layout_pdf(self):
        if not self._result:
            messagebox.showwarning("提示", "請先執行配料計算。")
            return
        r = self._result
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF 排列圖","*.pdf")],
            initialfile=f"BH_排列圖_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf")
        if not path:
            return
        try:
            self.status_var.set("⏳ 正在產生排列圖 PDF...")
            self.update_idletasks()
            # 從 BH 配料表讀取材質
            bh_rows = r.get("bh_rows", [])
            mats = list(dict.fromkeys(row.get("mat", "") for row in bh_rows if row.get("mat")))
            mat_name = "、".join(mats) if mats else r["params"]["mat_name"]

            # 建立 comp名稱 → serial+W/F 對應表（供排列圖使用）
            serial_map = {}
            for d in r.get("decomposed", []):
                fl_sn = d["flange"].get("serial","")
                wb_sn = d["web"].get("serial","")
                serial_map[d["flange"]["name"]] = f"{fl_sn}F" if fl_sn else ""
                serial_map[d["web"]["name"]]    = f"{wb_sn}W" if wb_sn else ""

            write_layout_pdf(
                path,
                r["params"]["proj_no"],
                r["params"]["proj_name"],
                mat_name,
                r["cut_details"],
                r["new_scraps"],
                serial_map=serial_map
            )
            self.status_var.set(f"排列圖 PDF 已儲存：{path}")
            messagebox.showinfo("儲存成功", f"排列圖 PDF 已儲存至：\n{path}")
        except Exception as e:
            messagebox.showerror("儲存失敗", str(e))

    # ── 採購修正 ──────────────────────────────────────────────────
    def _open_purchase_edit(self):
        if not self._result:
            messagebox.showwarning("提示", "請先執行配料計算。")
            return
        self._purchase_edit_win = PurchaseEditWindow(self, self._result)

    def _new_preview_result(self):
        if not self._result:
            messagebox.showwarning("提示", "請先執行配料計算。")
            return
        r = self._get_purchase_edit_result()
        if r is None: return
        PreviewWindow(self, r)

    def _new_preview_layout(self):
        if not self._result:
            messagebox.showwarning("提示", "請先執行配料計算。")
            return
        r = self._get_purchase_edit_result()
        if r is None: return
        import tempfile, os
        p = r["params"]
        bh_rows  = r.get("bh_rows", [])
        mats     = list(dict.fromkeys(row.get("mat","") for row in bh_rows if row.get("mat")))
        mat_name = "、".join(mats) if mats else p["mat_name"]
        serial_map = {}
        for d in r.get("decomposed", []):
            fl_sn = d["flange"].get("serial","")
            wb_sn = d["web"].get("serial","")
            serial_map[d["flange"]["name"]] = f"{fl_sn}F" if fl_sn else ""
            serial_map[d["web"]["name"]]    = f"{wb_sn}W" if wb_sn else ""
        try:
            tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            tmp.close()
            write_layout_pdf(tmp.name, p["proj_no"], p["proj_name"],
                             mat_name, r["cut_details"], r["new_scraps"],
                             serial_map=serial_map,
                             modified_specs=r.get("modified_specs"))
            os.startfile(tmp.name)
            self.status_var.set("排列圖預覽已開啟。")
        except Exception as e:
            messagebox.showerror("預覽失敗", str(e))

    def _new_export_layout(self):
        r = self._get_purchase_edit_result()
        if r is None: return
        p   = r["params"]
        now = datetime.now().strftime("%Y%m%d_%H%M")
        bh_rows  = r.get("bh_rows",[])
        mats     = list(dict.fromkeys(row.get("mat","") for row in bh_rows if row.get("mat")))
        mat_name = "、".join(mats) if mats else p["mat_name"]
        serial_map = {}
        for d in r.get("decomposed",[]):
            fl_sn = d["flange"].get("serial","")
            wb_sn = d["web"].get("serial","")
            serial_map[d["flange"]["name"]] = f"{fl_sn}F" if fl_sn else ""
            serial_map[d["web"]["name"]]    = f"{wb_sn}W" if wb_sn else ""
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF","*.pdf")],
            initialfile=f"BH_採購修正_排列圖_{now}.pdf")
        if not path: return
        try:
            write_layout_pdf(path, p["proj_no"], p["proj_name"],
                             mat_name, r["cut_details"], r["new_scraps"],
                             serial_map=serial_map,
                             modified_specs=r.get("modified_specs"))
            messagebox.showinfo("完成", f"已儲存：\n{path}")
        except Exception as e:
            messagebox.showerror("失敗", str(e))

    def _new_export_pdf(self):
        r = self._get_purchase_edit_result()
        if r is None: return
        p   = r["params"]
        now = datetime.now().strftime("%Y%m%d_%H%M")
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF","*.pdf")],
            initialfile=f"BH_採購修正_{now}.pdf")
        if not path: return
        try:
            write_pdf(path, p["proj_no"], p["proj_name"], p["proj_date"],
                      p["mat_name"], p["density"],
                      p["new_kerf"], p["new_trim"], p["scrap_kerf"], p["scrap_trim"],
                      r["bh_rows"], r["decomposed"], r["purchase_list"],
                      r["cut_details"], r["new_scraps"],
                      r["existing_scraps"], r["used_scrap_ids"])
            messagebox.showinfo("完成", f"已儲存：\n{path}")
        except Exception as e:
            messagebox.showerror("失敗", str(e))

    def _new_export_xlsx(self):
        r = self._get_purchase_edit_result()
        if r is None: return
        p   = r["params"]
        now = datetime.now().strftime("%Y%m%d_%H%M")
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")],
            initialfile=f"BH_採購修正_{now}.xlsx")
        if not path: return
        try:
            write_xlsx(path, p["proj_no"], p["proj_name"], p["proj_date"],
                       p["mat_name"], p["density"],
                       p["new_kerf"], p["new_trim"], p["scrap_kerf"], p["scrap_trim"],
                       r["bh_rows"], r["decomposed"], r["purchase_list"],
                       r["cut_details"], r["new_scraps"],
                       r["existing_scraps"], r["used_scrap_ids"])
            messagebox.showinfo("完成", f"已儲存：\n{path}")
        except Exception as e:
            messagebox.showerror("失敗", str(e))

    def _open_order_original(self):
        if not self._result:
            messagebox.showwarning("提示", "請先執行配料計算。")
            return
        OrderSelectWindow(self, self._result)

    def _open_order_select(self):
        if not self._result:
            messagebox.showwarning("提示", "請先執行配料計算。")
            return
        r = self._get_purchase_edit_result() or self._result
        OrderSelectWindow(self, r)

    def _get_purchase_edit_result(self):
        """取得採購修正結果：優先用已完成的，其次用開著的視窗"""
        # 1. 已按「完成」儲存的結果
        saved = getattr(self, "_purchase_edit_result", None)
        if saved:
            return saved
        # 2. 視窗還開著
        win = getattr(self, "_purchase_edit_win", None)
        if win and win.winfo_exists():
            return win._make_modified_result()
        # 3. 都沒有
        messagebox.showwarning("提示",
            "請先開啟「新採購清單」視窗，修改後按「✅ 完成」。")
        return None

    def _clear_all(self):
        if not messagebox.askyesno("確認", "確定要清除所有資料？"):
            return
        for item in self.bh_tree.get_children():
            self.bh_tree.delete(item)
        for item in self.sc_tree.get_children():
            self.sc_tree.delete(item)
        self._result = None
        self.status_var.set("已清除全部資料。")


# ═══════════════════════════════════════════════════════════════════════
# 對話框：新增 BH 構件
# ═══════════════════════════════════════════════════════════════════════

class BHRowDialog(tk.simpledialog.Dialog):
    def body(self, master):
        fields = [
            ("構件編號：",   "comp",   "C1"),
            ("零件編號：",   "part",   "P1"),
            ("斷面規格：",   "spec",   "BH700×350×14×28"),
            ("長度(mm)：",   "length", "9700"),
            ("數量(支)：",   "qty",    "4"),
            ("材質：",       "mat",    "A36"),
        ]
        self.vars = {}
        for i, (label, key, default) in enumerate(fields):
            tk.Label(master, text=label, font=("Microsoft JhengHei", 10)
                     ).grid(row=i, column=0, sticky="w", padx=8, pady=3)
            var = tk.StringVar(value=default)
            self.vars[key] = var
            tk.Entry(master, textvariable=var, width=28,
                     font=("Microsoft JhengHei", 10)).grid(row=i, column=1, padx=4, pady=3)
        return None

    def apply(self):
        try:
            self.result = {
                "comp":   self.vars["comp"].get().strip(),
                "part":   self.vars["part"].get().strip(),
                "spec":   self.vars["spec"].get().strip(),
                "length": int(self.vars["length"].get()),
                "qty":    int(self.vars["qty"].get()),
                "mat":    self.vars["mat"].get().strip(),
            }
        except ValueError:
            messagebox.showerror("錯誤", "長度與數量必須為整數。")
            self.result = None


class ScrapDialog(tk.simpledialog.Dialog):
    def body(self, master):
        fields = [
            ("厚度(mm)：",   "thick",  "18"),
            ("寬度(mm)：",   "width",  "620"),
            ("長度(mm)：",   "length", "9600"),
            ("數量(片)：",   "qty",    "1"),
            ("材質：",       "mat",    "A36"),
        ]
        self.vars = {}
        for i, (label, key, default) in enumerate(fields):
            tk.Label(master, text=label, font=("Microsoft JhengHei", 10)
                     ).grid(row=i, column=0, sticky="w", padx=8, pady=3)
            var = tk.StringVar(value=default)
            self.vars[key] = var
            tk.Entry(master, textvariable=var, width=20,
                     font=("Microsoft JhengHei", 10)).grid(row=i, column=1, padx=4, pady=3)
        return None

    def apply(self):
        try:
            self.result = {
                "thick":  int(self.vars["thick"].get()),
                "width":  int(self.vars["width"].get()),
                "length": int(self.vars["length"].get()),
                "qty":    int(self.vars["qty"].get()),
                "mat":    self.vars["mat"].get().strip(),
            }
        except ValueError:
            messagebox.showerror("錯誤", "數值欄位必須為整數。")
            self.result = None


# ═══════════════════════════════════════════════════════════════════════
# 採購修正視窗
# ═══════════════════════════════════════════════════════════════════════

class PurchaseEditWindow(tk.Toplevel):
    """新採購清單：讀取配料結果採購清單，可自行修改板厚/板寬/板長，重新產生完整報表"""

    def __init__(self, parent, result):
        super().__init__(parent)
        self.parent = parent
        self.result = result
        self.title("新採購清單")
        self.geometry("1040x580")
        self.configure(bg=CLR_BG)
        self.resizable(True, True)

        # 若主視窗已有保留的 rows（上次修改），直接沿用；否則重新建立
        saved_rows = getattr(parent, "_purchase_edit_rows", None)

        self.rows = self._build_rows() if saved_rows is None else saved_rows

        self._build()

    # ── 採購列 ────────────────────────────────────────────────────
    def _build_rows(self):
        """依 規格 + 材質 + 單重 合併採購清單；cts 為這一列包含的切割片次（依片次順序）"""
        import re
        from collections import OrderedDict
        merged = OrderedDict()
        for p in self.result["purchase_list"]:
            key = (p["spec"], p["mat"], p["unit_wt"])
            merged[key] = merged.get(key, 0) + p["qty"]
        rows = []
        for (spec, mat, unit_wt), qty in merged.items():
            m = re.match(r"PL(\d+)×(\d+)×(\d+)", spec)
            thick  = int(m.group(1)) if m else 0
            width  = int(m.group(2)) if m else 0
            length = int(m.group(3)) if m else 0
            key = self._board_key(thick, width, length, mat)
            cts = [ct["idx"] for ct in self.result["cut_details"]
                   if not ct.get("is_scrap") and self._ct_key(ct) == key]
            rows.append({
                "thick": thick, "width": width, "length": length,
                "orig_width": width, "orig_length": length,
                "qty": qty, "mat": mat, "orig_spec": spec, "note": "", "cts": cts
            })
        return rows

    def _split_row(self, i, k):
        """拆分採購列：第 i 列保留前 k 片（片次），其餘片數另成一列並恢復原本配出來的長寬"""
        row = self.rows[i]
        if not (1 <= k < row["qty"]):
            return
        rest = dict(row, width=row["orig_width"], length=row["orig_length"], qty=row["qty"] - k,
                    cts=list(row.get("cts", []))[k:], note="", lim_ignored=None)
        row["qty"] = k
        row["cts"] = list(row.get("cts", []))[:k]
        self.rows.insert(i + 1, rest)
        # 拆出的原尺寸列若已有相同的列（先前拆分留下的）就併入，但不併回剛拆分的第 i 列
        self._merge_row(i + 1, skip=i)

    def _merge_row(self, i, skip=None):
        """
        第 i 列與其他相同的採購列（同原始規格、同材質、同目前尺寸）合併：
        修改後與另一列相同時自動併回。skip：不與該列合併。回傳合併後第 i 列所在的索引
        """
        a = self.rows[i]
        sig = (a["orig_spec"], a["mat"], a["width"], a["length"])
        j = next((n for n, b in enumerate(self.rows) if n != i and n != skip
                  and (b["orig_spec"], b["mat"], b["width"], b["length"]) == sig), None)
        if j is None:
            return i
        lo, hi = min(i, j), max(i, j)
        keep, drop = self.rows[lo], self.rows[hi]
        keep["qty"] += drop["qty"]
        keep["cts"] = list(keep.get("cts", [])) + list(drop.get("cts", []))
        if not keep.get("note") and drop.get("note"):
            keep["note"] = drop["note"]
        del self.rows[hi]
        return lo

    def _row_matcher(self, row):
        """判斷片次是否屬於這一列：有 cts 依片次，否則依規格 + 材質"""
        if "cts" in row:
            cts = set(row["cts"])
            return lambda ct: ct["idx"] in cts
        key = self._row_key(row)
        return lambda ct: self._ct_key(ct) == key

    # ── 重量計算 ──────────────────────────────────────────────────
    def _wt(self, row):
        return round(calc_weight(row["width"], row["thick"], row["length"],
                                 self.result["params"]["density"]), 0)

    def _spec(self, row):
        return f'PL{row["thick"]}×{row["width"]}×{row["length"]}'

    def _limit_issues(self, row):
        """採購限制檢查（板寬 / 板長 / 單片重），回傳各欄的超出說明（未超出為空字串）"""
        p  = self.result["params"]
        wt = int(self._wt(row))

        def chk(label, v, lo, hi, unit):
            if v < lo:
                return f"{label} {v:,}{unit} 低於最小值 {lo:,}{unit}"
            if v > hi:
                return f"{label} {v:,}{unit} 超過最大值 {hi:,}{unit}"
            return ""

        return {
            "width":  chk("板寬", row["width"], p.get("bw_min", BW_MIN), p.get("bw_max", BW_MAX), "mm"),
            "length": chk("板長", row["length"], p.get("bl_min", BL_MIN), p.get("bl_max", BL_MAX), "mm"),
            "wt":     chk("單片重", wt, p.get("w_min", W_MIN), p.get("w_max", W_MAX), "kg"),
        }

    @staticmethod
    def _limit_msgs(iss):
        return [m for m in (iss["width"], iss["length"], iss["wt"]) if m]

    # 「忽略」超出採購限制：記住忽略當下的尺寸，之後再改尺寸會重新檢查
    @staticmethod
    def _lim_key(row):
        return f'{row["width"]}x{row["length"]}'

    def _is_ignored(self, row):
        return row.get("lim_ignored") == self._lim_key(row)

    def _ignore(self, row):
        row["lim_ignored"] = self._lim_key(row)

    def _active_limit_msgs(self, row):
        """尚未忽略的超出項目"""
        return [] if self._is_ignored(row) else self._limit_msgs(self._limit_issues(row))

    def _target_rows(self):
        """選取的列（未選取則為全部）"""
        sel = [int(i) for i in self.tree.selection() if i.isdigit()]
        return [self.rows[i] for i in sel] if sel else list(self.rows)

    def _ignore_selected(self):
        for row in self._target_rows():
            if self._active_limit_msgs(row):
                self._ignore(row)
        self._refresh()

    def _unignore_selected(self):
        for row in self._target_rows():
            row["lim_ignored"] = None
        self._refresh()

    @staticmethod
    def _board_key(t, w, l, mat):
        """採購列與切割片次的對應鍵：板規格 + 材質（同規格不同材質為不同採購列）"""
        return (t, w, l, mat)

    def _row_key(self, row):
        import re
        mo = re.match(r"PL(\d+)×(\d+)×(\d+)", row["orig_spec"])
        return self._board_key(int(mo.group(1)), int(mo.group(2)), int(mo.group(3)), row["mat"]) if mo else None

    def _ct_key(self, ct):
        import re
        m = re.match(r"PL(\d+)×(\d+)×(\d+)", ct["board_spec"])
        return self._board_key(int(m.group(1)), int(m.group(2)), int(m.group(3)), ct.get("mat")) if m else None

    def _check_size(self, row, orig_spec=None):
        """檢查修改後板寬/板長是否足夠裁切零件，回傳 (ok, msg)"""
        p        = self.result["params"]
        new_kerf = p["new_kerf"]
        new_trim = p["new_trim"]
        cur_w, cur_l = row["width"], row["length"]
        if cur_w >= row.get("orig_width", cur_w) and cur_l >= row.get("orig_length", cur_l):
            return True, ""

        # 只檢查這一列包含的片次
        mine = self._row_matcher(row)
        errors = []
        for ct in self.result["cut_details"]:
            layout = ct.get("layout")
            if ct.get("is_scrap") or not layout or not mine(ct):
                continue
            num_rows = len([r for r in layout.get("rows", []) if r.get("parts")])
            # 板寬需容納所有排；板長需容納最長的一整排（各段零件 + 段間鋸縫 + 兩端修邊）
            need_w = num_rows * layout["part_w"] + max(0, num_rows - 1) * new_kerf + new_trim * 2
            need_l = max((layout_row_used(layout, r) for r in layout.get("rows", [])), default=0) + new_trim
            if cur_w < need_w:
                errors.append(f"板寬不足！需要 {need_w}mm，目前 {cur_w}mm")
            if cur_l < need_l:
                errors.append(f"板長不足！需要 {need_l}mm，目前 {cur_l}mm")
        if errors:
            return False, "；".join(dict.fromkeys(errors))
        return True, ""

    # ── 介面 ──────────────────────────────────────────────────────
    def _build(self):
        hdr = tk.Frame(self, bg=CLR_HEADER, height=44)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="新採購清單　── 雙擊板寬/板長可修改尺寸；雙擊數量可拆分出要修改的片數",
                 bg=CLR_HEADER, fg="white",
                 font=("Microsoft JhengHei", 11, "bold")).pack(side="left", padx=12, pady=8)

        frame = tk.Frame(self, bg=CLR_BG)
        frame.pack(fill="both", expand=True, padx=8, pady=4)

        cols = ("NO","板厚(mm)","板寬(mm)","板長(mm)","數量","材質","單片重(kg)","總重(kg)","備註","原始規格","片次")
        vsb  = ttk.Scrollbar(frame, orient="vertical")
        hsb  = ttk.Scrollbar(frame, orient="horizontal")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings",
                                 yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        for col, w in zip(cols, [45,80,80,90,55,70,100,100,160,180,120]):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="center")
        self.tree.tag_configure("odd",  background=CLR_ROW_ODD)
        self.tree.tag_configure("even", background=CLR_ROW_EVEN)
        self.tree.tag_configure("tot",  background="#D6EAF8", font=("Microsoft JhengHei", 9, "bold"))
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", self._on_dbl)

        # 超出採購限制說明（Treeview 無法單格變色，超出的數字以 ⚠ 標示、整列紅字）
        self._lim_lbl = tk.Label(self, text="", bg=CLR_BG, fg="#C53030", justify="left",
                                 anchor="w", font=("Microsoft JhengHei", 9))
        self._lim_lbl.pack(fill="x", padx=10)
        self._ign_lbl = tk.Label(self, text="", bg=CLR_BG, fg="#B7791F", justify="left",
                                 anchor="w", font=("Microsoft JhengHei", 9))
        self._ign_lbl.pack(fill="x", padx=10)

        btn_bar = tk.Frame(self, bg=CLR_BG)
        btn_bar.pack(fill="x", padx=10, pady=8)

        def bb(text, cmd, bg, side="left", padx=(0,4)):
            tk.Button(btn_bar, text=text, command=cmd,
                      bg=bg, fg="white", font=("Microsoft JhengHei", 10, "bold"),
                      relief="flat", padx=14, pady=6, cursor="hand2"
                      ).pack(side=side, padx=padx)

        bb("↩ 重設",  self._reset,   "#718096", padx=(0,8))
        bb("🔍 診斷", self._debug,   "#4A6C7A", padx=(0,8))
        bb("忽略超出限制", self._ignore_selected,   "#B7791F", padx=(0,4))
        bb("取消忽略",     self._unignore_selected, "#A0AEC0", padx=(0,8))
        self._btn_confirm = tk.Button(btn_bar, text="✅ 完成",
                  command=self._confirm,
                  bg="#4A6741", fg="white", font=("Microsoft JhengHei", 10, "bold"),
                  relief="flat", padx=14, pady=6, cursor="hand2")
        self._btn_confirm.pack(side="left", padx=(0,4))

        self._refresh()  # 初始化時檢查尺寸，若有問題立即鎖定

    def _refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree.tag_configure("ok",      background=CLR_ROW_ODD)
        self.tree.tag_configure("ok2",     background=CLR_ROW_EVEN)
        self.tree.tag_configure("tot",     background="#D6EAF8", font=("Microsoft JhengHei", 9, "bold"))
        self.tree.tag_configure("error",   background="#FED7D7")
        self.tree.tag_configure("changed", background="#C6EFCE")
        self.tree.tag_configure("over",    foreground="#C53030")
        self.tree.tag_configure("ign",     foreground="#B7791F")   # 已忽略的超出採購限制：琥珀色
        total_wt  = 0
        total_qty = 0
        has_error = False
        lims = []
        ignored = []
        for i, row in enumerate(self.rows):
            wt  = self._wt(row)
            tot = wt * row["qty"]
            ok, msg = self._check_size(row, row["orig_spec"])
            changed = (row["width"]  != row.get("orig_width",  row["width"]) or
                       row["length"] != row.get("orig_length", row["length"]))
            iss = self._limit_issues(row)
            ign_iss = {"width": "", "length": "", "wt": ""}
            all_msgs = self._limit_msgs(iss)
            if all_msgs and self._is_ignored(row):
                ignored.append(f"• NO {i+1}　{self._spec(row)}：{'；'.join(all_msgs)}")
                iss, ign_iss = ign_iss, iss   # 已忽略的項目不再標紅，改標琥珀色
            lim_msgs = self._limit_msgs(iss)
            if lim_msgs:
                lims.append(f"• NO {i+1}　{self._spec(row)}：{'；'.join(lim_msgs)}")
            if not ok:
                tag = ("error",)
                has_error = True
            elif changed:
                tag = ("changed",)
            else:
                tag = ("ok" if i%2==0 else "ok2",)
            if lim_msgs:
                tag = tag + ("over",)
            elif self._limit_msgs(ign_iss):
                tag = tag + ("ign",)
            mark = lambda k, v: f"⚠{v}" if iss[k] else f"◇{v}" if ign_iss[k] else v
            self.tree.insert("", "end", iid=str(i), tags=tag,
                             values=(i+1, row["thick"], mark("width", row["width"]),
                                     mark("length", row["length"]),
                                     row["qty"], row["mat"], mark("wt", int(wt)), int(tot),
                                     row.get("note",""), row["orig_spec"], "、".join(row.get("cts", []))))
            total_wt  += tot
            total_qty += row["qty"]
        self.tree.insert("", "end", tags=("tot",),
                         values=("", "合計", "", "", total_qty, "",
                                 "總重(kg)：", f"{total_wt:,.0f}", "", "", ""))
        if hasattr(self, "_lim_lbl"):
            txt = ""
            if lims:
                txt += ("⚠ 以下項目超出採購限制（紅字 ⚠），請確認是否可採購"
                        "（可選取列後按「忽略超出限制」；未選取則全部忽略）：\n" + "\n".join(lims))
            self._lim_lbl.config(text=txt)
            self._ign_lbl.config(text=("已忽略的超出採購限制（琥珀色 ◇）：\n" + "\n".join(ignored)) if ignored else "")
        if hasattr(self, "_btn_confirm"):
            if has_error:
                self._btn_confirm.config(state="disabled", bg="#AAAAAA",
                                         text="✅ 完成（有錯誤）")
            else:
                self._btn_confirm.config(state="normal", bg="#4A6741",
                                         text="✅ 完成")
        self.parent._purchase_edit_rows = self.rows
    def _on_dbl(self, event):
        if self.tree.identify("region", event.x, event.y) != "cell": return
        col_n = int(self.tree.identify_column(event.x).replace("#",""))
        if col_n not in (3, 4, 5, 9): return   # 板寬=3, 板長=4, 數量=5, 備註=9（板厚不可調整）
        iid = self.tree.identify_row(event.y)
        if not iid or not iid.isdigit(): return
        idx = int(iid)
        row = self.rows[idx]
        field_map = {3: ("板寬", "width"), 4: ("板長", "length"), 5: ("數量", "qty"), 9: ("備註", "note")}
        label, key = field_map[col_n]
        if key == "qty":
            # 數量只能改小：改小時拆分成兩列，其餘片數保留原本配出來的長寬
            if row["qty"] <= 1:
                messagebox.showinfo("提示", "此列只有 1 片，無法再拆分。", parent=self)
                return
            val = tk.simpledialog.askinteger(
                "修改數量（拆分）",
                f"目前 {row['qty']} 片（片次：{'、'.join(row.get('cts', []))}）。\n"
                f"請輸入要保留在此列的片數（1 ~ {row['qty']}），\n其餘片數會另成一列並恢復原本配出來的長寬：",
                initialvalue=row["qty"], minvalue=1, maxvalue=row["qty"], parent=self)
            if val is None or val == row["qty"]:
                return
            self._split_row(idx, val)
            self._refresh()
            return
        if key == "note":
            val = tk.simpledialog.askstring(
                "修改備註", "請輸入備註（如：和調整板片大小說明）：",
                initialvalue=row.get("note",""), parent=self)
            if val is None: return
            row["note"] = val
            self._refresh()
            return
        val = tk.simpledialog.askinteger(
            f"修改{label}", f"請輸入新的{label}（mm）：",
            initialvalue=row[key], minvalue=1, maxvalue=99999, parent=self)
        if val is None: return
        # 檢查縮小警告（修改前先比較）
        row[key] = val
        row = self.rows[self._merge_row(idx)]   # 與其他相同的列合併後，以合併後的列檢查
        self._refresh()
        ok, msg = self._check_size(row, row["orig_spec"])
        lim_msgs = self._active_limit_msgs(row)
        if not ok:
            row[key] = val  # 保留修改但警告
            extra = ("\n\n另外超出採購限制：\n" + "\n".join(lim_msgs)) if lim_msgs else ""
            messagebox.showwarning("尺寸不足",
                f"⚠ 鐵板片數不足，無法裁切！\n\n{msg}\n\n請調大板寬或板長。{extra}", parent=self)
        elif lim_msgs:
            if messagebox.askyesno("超出採購限制",
                    f"⚠ {self._spec(row)} 超出採購限制：\n\n" + "\n".join(lim_msgs)
                    + "\n\n若確認此規格可採購，是否忽略並繼續？\n（選「否」返回修改）",
                    icon="warning", parent=self):
                self._ignore(row)
                self._refresh()

    def _debug(self):
        import re
        msg = "=== 診斷報告 ===\n\n"
        msg += "purchase_list rows:\n"
        for row in self.rows:
            msg += f"  orig={row['orig_spec']} cur_w={row['width']} cur_l={row['length']}\n"

        msg += "\ndim_map keys (thick,width,length):\n"
        for row in self.rows:
            mo = re.match(r"PL(\d+)×(\d+)×(\d+)", row["orig_spec"])
            if mo:
                key = (int(mo.group(1)), int(mo.group(2)), int(mo.group(3)))
                msg += f"  {key}\n"

        msg += "\ncut_details board_spec → parsed key:\n"
        for ct in self.result["cut_details"]:
            if ct.get("is_scrap"): continue
            m = re.match(r"PL(\d+)×(\d+)×(\d+)", ct["board_spec"])
            if m:
                key = (int(m.group(1)), int(m.group(2)), int(m.group(3)))
                msg += f"  {ct['board_spec']} → {key}\n"
            else:
                msg += f"  {ct['board_spec']} → ❌ no match\n"

        msg += "\n_check_size 結果:\n"
        for row in self.rows:
            ok, err = self._check_size(row, row["orig_spec"])
            msg += f"  {row['orig_spec']} → {'✅' if ok else '❌ '+err}\n"

        msg += "\n_make_modified_result leftover:\n"
        r = self._make_modified_result()
        for ct in r["cut_details"]:
            if not ct.get("is_scrap"):
                msg += f"  {ct['board_spec']} leftover={ct.get('leftover','?')}\n"
        messagebox.showinfo("診斷", msg)

    def _reset(self):
        # 清除主視窗記憶
        self.parent._purchase_edit_rows   = None
        self.parent._purchase_edit_result = None
        self.rows = self._build_rows()
        self._refresh()

    def _confirm(self):
        """儲存修改後的結果到主視窗，關閉視窗"""
        # 整體尺寸檢查 — 有不足直接拒絕，不給選擇
        errors = []
        for row in self.rows:
            ok, msg = self._check_size(row, row["orig_spec"])
            if not ok:
                errors.append(f"• PL{row['thick']}×{row['width']}×{row['length']}：{msg}")
        if errors:
            err_txt = "\n".join(errors)
            messagebox.showerror("無法完成 — 尺寸不足",
                f"以下項目尺寸不足（紅色底色），無法裁切：\n\n{err_txt}\n\n請調大板寬或板長後再完成。",
                parent=self)
            return
        lims, lim_rows = [], []
        for row in self.rows:
            ms = self._active_limit_msgs(row)
            if ms:
                lim_rows.append(row)
                lims.append(f"• {self._spec(row)}：{'；'.join(ms)}")
        if lims:
            if not messagebox.askyesno(
                    "超出採購限制",
                    "以下項目超出採購限制：\n\n" + "\n".join(lims)
                    + "\n\n若確認可採購，是否全部忽略並套用？\n（選「否」返回修改）",
                    icon="warning", parent=self):
                return
            for row in lim_rows:
                self._ignore(row)
            self._refresh()
        self.parent._purchase_edit_result = self._make_modified_result()
        messagebox.showinfo("完成", "新採購清單已套用！\n可點擊「新預覽」查看結果。", parent=self)
        self.destroy()

    def _preview_result(self):
        r = self._make_modified_result()
        PreviewWindow(self, r)

    def _preview_layout(self):
        import tempfile, os
        r = self._make_modified_result()
        p = r["params"]
        bh_rows  = r.get("bh_rows", [])
        mats     = list(dict.fromkeys(row.get("mat","") for row in bh_rows if row.get("mat")))
        mat_name = "、".join(mats) if mats else p["mat_name"]
        serial_map = {}
        for d in r.get("decomposed", []):
            fl_sn = d["flange"].get("serial","")
            wb_sn = d["web"].get("serial","")
            serial_map[d["flange"]["name"]] = f"{fl_sn}F" if fl_sn else ""
            serial_map[d["web"]["name"]]    = f"{wb_sn}W" if wb_sn else ""
        try:
            tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            tmp.close()
            write_layout_pdf(tmp.name, p["proj_no"], p["proj_name"],
                             mat_name, r["cut_details"], r["new_scraps"],
                             serial_map=serial_map,
                             modified_specs=r.get("modified_specs"))
            os.startfile(tmp.name)
        except Exception as e:
            messagebox.showerror("預覽失敗", str(e), parent=self)

    def _make_modified_result(self):
        import copy
        r = copy.deepcopy(self.result)
        density = r["params"]["density"]

        # 記錄被修改過的板規格（新尺寸 + 材質），排列圖 PDF 以綠色標示
        r["modified_specs"] = {
            self._board_key(row["thick"], row["width"], row["length"], row["mat"])
            for row in self.rows
            if row["width"] != row.get("orig_width", row["width"])
            or row["length"] != row.get("orig_length", row["length"])}

        # 更新 purchase_list
        r["purchase_list"] = []
        for row in self.rows:
            wt = int(self._wt(row))
            r["purchase_list"].append({"spec": self._spec(row), "qty": row["qty"], "mat": row["mat"],
                                       "unit_wt": wt, "total_wt": wt * row["qty"]})

        # 片次 → 採購列
        matchers = [(self._row_matcher(row), row) for row in self.rows]

        def row_of(ct):
            return next((row for f, row in matchers if f(ct)), None)

        new_scraps = []
        old_scraps = r.get("new_scraps", [])

        def keep_old(ct):
            """未重新計算的片次（使用現有餘料、或找不到對應採購規格）保留原本產生的餘料"""
            new_scraps.extend(sc for sc in old_scraps if sc.get("src") == ct["idx"])

        for ct in r["cut_details"]:
            layout = ct.get("layout")
            row = None if ct.get("is_scrap") or not layout else row_of(ct)
            if row is None:
                keep_old(ct)
                continue
            new_bw, new_bl = row["width"], row["length"]
            ct["board_spec"] = self._spec(row)
            ct["board_wt"]   = int(calc_weight(new_bw, row["thick"], new_bl, density))
            # 與配料計算相同算法重算餘料；排列圖跟著修改後尺寸更新
            layout["board_w"] = new_bw
            layout["board_l"] = new_bl
            lo = board_leftovers(row["thick"], new_bw, new_bl, layout["part_w"], layout["trim"], layout["kerf"],
                                 [layout_row_used(layout, rv) for rv in layout["rows"]],
                                 density, ct["idx"], ct["type"], ct["mat"], new_scraps)
            layout["width_scrap_ref"] = lo["width_obj"]
            for rv, obj in zip(layout["rows"], lo["row_objs"]):
                rv["scrap_ref"] = obj
            ct["leftover"] = lo["left_spec"]

        r["new_scraps"] = new_scraps
        return r

    def _export(self, mode):
        r   = self._make_modified_result()
        p   = r["params"]
        now = datetime.now().strftime("%Y%m%d_%H%M")
        serial_map = {}
        for d in r.get("decomposed", []):
            fl_sn = d["flange"].get("serial","")
            wb_sn = d["web"].get("serial","")
            serial_map[d["flange"]["name"]] = f"{fl_sn}F" if fl_sn else ""
            serial_map[d["web"]["name"]]    = f"{wb_sn}W" if wb_sn else ""

        if mode == "xlsx":
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")],
                initialfile=f"BH_新採購清單_{now}.xlsx", parent=self)
            if not path: return
            try:
                write_xlsx(path, p["proj_no"], p["proj_name"], p["proj_date"],
                           p["mat_name"], p["density"],
                           p["new_kerf"], p["new_trim"], p["scrap_kerf"], p["scrap_trim"],
                           r["bh_rows"], r["decomposed"], r["purchase_list"],
                           r["cut_details"], r["new_scraps"],
                           r["existing_scraps"], r["used_scrap_ids"])
                messagebox.showinfo("完成", f"已儲存：\n{path}", parent=self)
            except Exception as e:
                messagebox.showerror("失敗", str(e), parent=self)

        elif mode == "pdf":
            path = filedialog.asksaveasfilename(
                defaultextension=".pdf", filetypes=[("PDF","*.pdf")],
                initialfile=f"BH_新採購清單_{now}.pdf", parent=self)
            if not path: return
            try:
                write_pdf(path, p["proj_no"], p["proj_name"], p["proj_date"],
                          p["mat_name"], p["density"],
                          p["new_kerf"], p["new_trim"], p["scrap_kerf"], p["scrap_trim"],
                          r["bh_rows"], r["decomposed"], r["purchase_list"],
                          r["cut_details"], r["new_scraps"],
                          r["existing_scraps"], r["used_scrap_ids"])
                messagebox.showinfo("完成", f"已儲存：\n{path}", parent=self)
            except Exception as e:
                messagebox.showerror("失敗", str(e), parent=self)

        elif mode == "layout":
            bh_rows  = r.get("bh_rows",[])
            mats     = list(dict.fromkeys(row.get("mat","") for row in bh_rows if row.get("mat")))
            mat_name = "、".join(mats) if mats else p["mat_name"]
            path = filedialog.asksaveasfilename(
                defaultextension=".pdf", filetypes=[("PDF 排列圖","*.pdf")],
                initialfile=f"BH_新採購清單_排列圖_{now}.pdf", parent=self)
            if not path: return
            try:
                write_layout_pdf(path, p["proj_no"], p["proj_name"],
                                 mat_name, r["cut_details"], r["new_scraps"],
                                 serial_map=serial_map,
                                 modified_specs=r.get("modified_specs"))
                messagebox.showinfo("完成", f"已儲存：\n{path}", parent=self)
            except Exception as e:
                messagebox.showerror("失敗", str(e), parent=self)


# ═══════════════════════════════════════════════════════════════════════
# 訂購單輸出
# ═══════════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════════════
# 訂購單範本（base64 嵌入，不需外部檔案）
# ═══════════════════════════════════════════════════════════════════════
_ORDER_TEMPLATE_B64 = """UEsDBBQABgAIAAAAIQB0NlqmegEAAIQFAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACsVM1OAjEQvpv4DpteDVvwYIxh4YB6VBLwAWo7sA3dtukMCG/vbEFiDEIIXLbZtvP9TGemP1w3rlhBQht8JXplVxTgdTDWzyvxMX3tPIoCSXmjXPBQiQ2gGA5ub/rTTQQsONpjJWqi+CQl6hoahWWI4PlkFlKjiH/TXEalF2oO8r7bfZA6eAJPHWoxxKD/DDO1dFS8rHl7q+TTelGMtvdaqkqoGJ3VilioXHnzh6QTZjOrwQS9bBi6xJhAGawBqHFlTJYZ0wSI2BgKeZAzgcPzSHeuSo7MwrC2Ee/Y+j8M7cn/rnZx7/wcyRooxirRm2rYu1w7+RXS4jOERXkc5NzU5BSVjbL+R/cR/nwZZV56VxbS+svAJ3QQ1xjI/L1cQoY5QYi0cYDXTnsGPcVcqwRmQly986sL+I19QodWTo9qLpErJ2GPe4yfW3qcQkSeGgnOF/DTom10JzIQJLKwb9JDxb5n5JFzsWNoZ5oBc4Bb5hk6+AYAAP//AwBQSwMEFAAGAAgAAAAhALVVMCP0AAAATAIAAAsACAJfcmVscy8ucmVscyCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACskk1PwzAMhu9I/IfI99XdkBBCS3dBSLshVH6ASdwPtY2jJBvdvyccEFQagwNHf71+/Mrb3TyN6sgh9uI0rIsSFDsjtnethpf6cXUHKiZylkZxrOHEEXbV9dX2mUdKeSh2vY8qq7iooUvJ3yNG0/FEsRDPLlcaCROlHIYWPZmBWsZNWd5i+K4B1UJT7a2GsLc3oOqTz5t/15am6Q0/iDlM7NKZFchzYmfZrnzIbCH1+RpVU2g5abBinnI6InlfZGzA80SbvxP9fC1OnMhSIjQS+DLPR8cloPV/WrQ08cudecQ3CcOryPDJgosfqN4BAAD//wMAUEsDBBQABgAIAAAAIQBm2OK1vgMAAB0JAAAPAAAAeGwvd29ya2Jvb2sueG1srFbdbqM4GL1fad8BcU+xwRBATUfhTxupHVWdTHtTqXLBKaiAs8Y0qap5g7ndfYyVdq5m32e18xr7mZCkbVarbGejxGB/5nCOv/PZOX63qivtgYm25M1Yx0dI11iT8bxs7sb6x1lqeLrWStrktOING+uPrNXfnfz4w/GSi/tbzu81AGjasV5IuQhMs80KVtP2iC9YA5E5FzWV0BV3ZrsQjOZtwZisK9NCyDVrWjb6GiEQh2Dw+bzMWMyzrmaNXIMIVlEJ9NuiXLQbtDo7BK6m4r5bGBmvFwBxW1alfOxBda3OguldwwW9rUD2CjvaSsDXhR9G0FibN0Fo71V1mQne8rk8AmhzTXpPP0Ymxi+WYLW/BochEVOwh1LlcMtKuG9k5W6x3B0YRt+NhsFavVcCWLw3ojlbbpZ+cjwvK3a5tq5GF4v3tFaZqnStoq1M8lKyfKyPoMuX7MWA6BZhV1YQtYhtE9082dr5XGg5m9OukjMw8gYeKsN1fctRM8EYk0oy0VDJIt5I8OGg63s912NHBQeHaxfs564UDAoL/AVaoaVZQG/bcyoLrRPVWI+C648tyL8Op2dMXMd82VQcCuz6mTPpfhn8B2/STAk2QfGa1fr+tXogJ4KN/86l0OB+Gp9CDj7QB8gI5D0fCnYKS47tmyYTAb55St2JP0oiYoRxmBrEThPDS72J4aA4iknkRakTfQIxwg0yTjtZDMlW0GOdQGb3Qmd0tYlgFHRlvqPxhIaPoa6vmk3skxKstrXLki3bnS1UV1tdlU3Ol2PdsV0Q9bjpwv2yj1yVuSyUqbC9HfuJlXcF0LWQ56kKEJaiNdafYi92iZ+GhuVMQoM4ODZCn4SGZ8dO6lp2iCy3p2M+49PvnsCrv2pN7/g/v/767ZfPf33549tvv8NmrfbXfp11TQTqTWKa4z6Pm4czWmXgc3XpJ/oYWb6awVbytJX9FSxWAklM0GSEfGKgxHYM4vmW4RHbMiISW4kzSuIkdFSK1BkQ/B87Ye/0YHO4KJYFFXImaHYPR9IFm4e0BU+tBQHf52RDxwuRDRRJisFP2EdGGLrEcOLUdkY4jhIn3ZFV8udv3Ic8s3+aUdlBjary7PuBatNhdDs4Xw8M2XpRfsFFrNZ9ePrfJn4A9RU7cHJ6eeDE6P3Z7OzAuafJ7OYq7Y30j2rNV9mIMfGRnUwM24YSJ6N0BNWNHMMmIxI5JEwwGu2yUS2zh7clwyLmxi7R83N82CxUchR4MPzJ0Vomh9ALGyn6vfm3aCd/AwAA//8DAFBLAwQUAAYACAAAACEAkgeU7AQBAAA/AwAAGgAIAXhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxzIKIEASigAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArJLLasQwDEX3hf6D0b5xMn1QhnFm0VKYbZt+gHCUOExiB1t95O9rUjrJwJBusjFIwvceibvbf3et+CQfGmcVZEkKgqx2ZWNrBe/Fy80jiMBoS2ydJQUDBdjn11e7V2qR46dgmj6IqGKDAsPcb6UM2lCHIXE92TipnO+QY+lr2aM+Yk1yk6YP0s81ID/TFIdSgT+UtyCKoY/O/2u7qmo0PTv90ZHlCxYy8NDGBUSBviZW8FsnkRHkZfvNmvYcz0KT+1jK8c2WGLI1Gb6cPwZDxBPHqRXkOFmEuV8TRmOrnww2doI5tZYucrdqKAx6Kt/Yx8zPszFv/8HIs9jnPwAAAP//AwBQSwMEFAAGAAgAAAAhAIK7I1UiPQAAsOEBABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWysXV2PJDdyfDfg/7DYd+121ydLkGTIOhxswDAOdz77eTQ7kga3uyPvjO5ONvzfHUw2M3IkVU+RmYbtDZ+b0eyprMhgMln1xT/9/cP7V3+9+/R4//Dxy9fnN6fXr+4+3j68u//4/Zev//wfv/8svX71+HTz8d3N+4ePd1++/vnu8fU/ffWP//DF3x4+/eXxh7u7p1dg+Pj45esfnp5+/Pzt28fbH+4+3Dy+efjx7iP+P989fPpw84T/89P3bx9//HR3804GfXj/djidlrcfbu4/vi4Mn386wvHw3Xf3t3e/e7j96cPdx6dC8unu/c0T5v/4w/2Pj5Xtw+0Rug83n/7y04+f3T58+BEU396/v3/6WUhfv/pw+/m/fv/x4dPNt+/xu/9+nm5uX/39E/57wP+M9WvkP//VN324v/308Pjw3dMbML8tc/71z9/ebm9vbpXp17//EM15evvp7q/3+QKSauib0nlWroFkYyfZomT5z/Xp85/u3335+n+/mYf562+W3312xn99Nv1+TJ/987htn81fr9s36Zt1nr7++v9ef/WFxMkfPn31xY8339/96e7pzz/+4dOr7+6f/uPhD/gPEKuv3371xVv91Lt7BET+I7z6dPfdl6+/Pn/+53nIH5FP/Of93d8eDX71dPPtn+7e390+3WFO59ev/ufh4cOfbm/ypV5xD+j/+e85ft//4j/8U477f7v5+eGnp0wMghPG5Dvi24eHv+T/6F/Beso/Qr4jz+rm9un+r3ff3L0H2R/HGXfVf8tEgT//87jxx+Th9YfZaf9e7iT8Dd7dfXfz0/unPz787V/u7r//4Qlfv70BoYTi5+9+/t3d4y3uDUzgzZRpbx/e46fjf7/6cJ/vcYT2zd/l37/dv3v64cvXy5tlHNJ5yJN6+jn/CfCZ258enx4+/Ff5hPytlQGhIQz498JwxlfN4y8Yvr17fPr9fZ7eVTbEhrDh38o2vhmWrZdvuvDh3wvf+GY7LYd/Hf4KMh/8exk//2I8Znrwty0XLvxb56J/4+Ms64UF/15Y1jentKVpXHnJjtNB0eUH4l/HpLYLC/69sKQ34y/j6PikzriBSljmO6lEHUZf4hHoh/t37+5K+F4JzbNGN4Bevl/E5nG2GulnhnrPpGqInxnjklblXjs+mxrZZ4b2r2+8Af/Pg9F5rqGegf6x1mGZj98t5xriGdT795c3DK7F0SnVWD8z2Ld+bTnXWM/gyuyuRVQN9DMjPf/WY0pZo3pgOJ5/469zTWxVbRk8mIl+fRH4Iu352hdlfhYhv1K/gxcjS2ahm3lxd797rpduyKg3tcz1ig0ZBaSXYa4XUFBIihkWva4ZdaWZYalSJWgn1Ry/c4ZFAyWjX0loC1PVq2FhzP1Gymmh1MhcTF62d9HRmFw0KDPaTz0tk6saBtehlLiuvMWPJZ8hO8dyw2QUEbyrBklGnRo9rBoaGflVelg1QjK6EDp0elg1PjLqU+ohu6HLn5+hcVyrVw2DlQLWqtar6tfKWNrVzKQBkxEDpluxU42X2ViWva+fsx8pXtc4k9bVwJztyIXlmjE5fkPOakwEhUj2rEZFUJdkz9mUXH4pI6Tf7MxqTwQ5JHvO5uQyM4adS7LnocamIM/ksvkpkzM26DdWCw0RMlRFQzHBI9nzUG96QQGSPQ8aJBl1SvY8VB0R5JfsedAIycgv2fOo8ZFRn2RjjV9DI6NKctRez6OGwcg81CjZ86j6lcsy9S+zY7HnUQMmI79kz6Mu7xc63j3JXrNDkXtJ0OXrWyV7VaMiKCDqVzUqgkIke1XbIqhLsld1KoL0buxNsKsaFUEOVVzVrQi6MLkke1UHI8g1OQ1K44Vckr2qM1qzSblMrsNlr6ne9IIigjdVH7RmpEHSVglZU9URQVcE8eBSZ03VzwqqwtRfDVlT1S5BV2Z4pR6xJg2NjCrJUcleU01sazIFlbaayLrV3LNujKVdzdw0YDLyS/a6aUHvlEsu9crspIzzSQspBV4+36raGKzFNoEBoQ9OrbgJrJyuEjxYtfAmsEpRUyEeLDXZFag3Zq96n09qWgp0SCQIqncp8MLlUnAwaclZoG+CDFNjj1wqjglqFfqUnctlgh06DioGSYYh4YydtLqPkKGGTGNV+zRxPyLDK1p5tLJ9mhgvGVbVcFS3TxOjJcMrs7xW4T5NDJQMK81RTT+fJgbFZEovbaoOGurbZOr/u7o6MYAyZAD1C8Sk8TOYksxeajlj27TGm8BeaR/U1IAyZl8VRHpZBcZI+6A2B19garlt0j6otwHL/k7r8YoAaDR4BHqUc1CPA1oGok/aB3U+IP2tvdemH2vClDLnk/ZBXdV5yA7HIe3DpoIgMELah023KwX2SvuwMVAyDJD2YVMTJTBA2odNtU1gp7QPGwMlw2ZpH3NnS9mYH0+mRNMo7eNJM9V4enkj8zyeNIAEBkj7eOJ2tynd7Eo7CzNngb3SPtPmCIy4F2baHIEx0j7T9AisEtom7TN9jkC9T7uT8kybI9Aj7TO9jsALl0/aZ/ofgb4JMkyNl/JJ+0xnNWeH45D2eVZBEBgSzrO6qDnDXmmf0fF3ESuBAdI+z2qiBAZI+zyrtgnslPZ5ZqBk2Cztc+7qKNI+m5aNxvL5edaWDUCztbDn2mdtnjgLDJD2WXspzqsp8exKuyngCOyV9pU2R2DEvbDS5giMkfaVpkdgn7Sv9DkC/dK+0uYI9CjnSq8jMETaE/2PQM8EUzZQ5YYTeOHySXuis0rZ4VxIewoy6aSCIDAinNNJrbFADZnGgkw6sZUvwyuiebQgk07qjQVeKF3thmxjTRlemeW1gkxiG6vASnO4IINmbo0009rRKu1JWzvOyfSI7upqYjeoQAZQt/dL2nMxnEyNZ28K+FC9qAVeptDc4n5SmzMIDLgXQFTv/wIrp6vWDirt7hJYFarJtYNFW7oE6n3a3Yp0UpsDcmfT30m9Drgocr6exJP6H5CyEIC/Axv/jnYlntj9KjBC2jErbSY7ZYfTL+0Dzj/U3jiBIeGcu0NLv93JNIrOb9qkHXPT5j+BfmkHpbYCCvRLOyhrciywT9oxtubDAlulHaNMUJiKTmNX+Ym9qifTWrqvq+whPdkm0tOvZebwPcMuUlvj2Z0CCziDwF5pH9XmgOdZD0/3GSYQ6WUVGCPto5oefIEpGbdJ+6g+ByyMmP6+xWFUm1NgTTk9yjmq1wEXRc4n7TiKVqVJoG+CKnOjPcDz69NOxwv4+KmqeaM56tLh2kGlgiAwQtpH7ScFvQ2ZRmkftYl0EBgg7aN2k4IyZBsVPIyWDDulfRwYKBk2S/vI8zujaWttbUEftU0Ev+zlbVR8iAGUIQOo2/uN2rkx2BrPrrSzgIPP86I2u3YcXKv3vcCIe2GmzREYI+0zTY/AqlBt0j7T5wi8sDikfabNEehRzpleR+CFyyftM/2PQN8EdbE4m6M/roIMTr/poY7ZnI7pkfaZx2IEhoQzT8nMGWrINEr7zOMyAgOkfc69ppcDh6bt1HOuaOZ5HIGd0j7nwzx1ZsZ0Hy3I4DAjg8K0v7ZK+6xtImB8eRsVH1KbITBA2mft3EDDC93zrrSbAo7AyxSapT3R5giMuBcSbY7AGGlPND0C+6Q90ecI9Et7os0R6FHORK8jMETaE/2PQN8EtQKQjJfySXuis0rmFE2PtCftREVvlqkUOx6kASKtLwrslfakTaigvO6Hj671U25ALQIq8DI1j7Qn7UTFLPsPjaZ86KfOrGMbFV+uhaFkOmJbpT1pmwgYX95GRe+dZhSBAdKetHMDNQxe9z1px4fqRS2wU9oxuGapAgOkHUT1shYYIu2gqte6wC5px9Dqcwp0Szto6hKuQIdygqB6nQIjpB3VqxqtBbomqD234OJi0SXtYKorR0DXNirGM0gyDAln7UgFff+xUgxmoJgm1N8Sq6OPb8KTiC4CCnYjXI4nQp20ExUFy+5tVIyt+bDAeiGOunaMYlCYjthGaQcN9S13d9Tkt9Mhg88zgGyHan+tHZQqOVLueWkKLOCMAnulfVSbMwqMuBdGtTngNC2prm1UUOm1Ftgn7aCpN4NAv7SPanNA6NtGBYEmRYEh0j6q/wG/bxsVBAxT46V80j6qs8If0LWNivEMEnvgxvP4u1E7UkHPNX3rNioGq8oIrLdafzEQlIwX04vqcO2gVG0TeGWWVzpkQMNAMUeFDj/sBQSaCEfTEdsq7aO2ieCGMrsRe9I+arMGPv/snE9vrR08Kjm2xrPr2lnAGQX2SvtMmyMwQtpn2hyBldMn7TNNj8A+aZ/pcwT6pX2mzRFY59WxjTrO9DoCQ6R9pv8R6Jsgw9R4KZ+0z3RWc3Y4lwl2FGTwB1RBEBgSztqRCvr+bdRx0SbUAgOkfdFeVFCGbKOCR5OjwE5pX/L5ICnIgLFjGxWjNBEupiO2VdoXbRMB48vbqPiQBpBABlC3tC/auTHaGs+utJsCjsBeaU+0OQIj7gU8rbZeVoEx0p5oegT2SXuizxHol/ZEmyPQo5yJXkdgiLQn+h+BvglqBSAZL+WT9kRnlbLDcUh70k7UUWBIOGtHKjj7t1ExWP2wwABpT9qLCnYjXI6CTNJOVFBe776/5tqTPosVNB3bqGPSs0OAZunf1vyIsZqpkulc3ddVbdbAULPV4CjIJO3cmKTcc7nue1PAh+qCp8BOacfgmqUKDLgXQFTv/wJDpH2SZ+qLDyiwS9oxtNYAC3RLO2hq8BToUE4Q1KRYYIS0g6n6nwJ9E6wVAHBR5lzSDqaqeYCUgQ7XPp21E7XAiHDGaykuTgWcFJnWggwGM1DM01gdtXZQMl5ML6qjIANKRot9sHybpIKGgWKOCh0uyICAQWE6YhtdO2hU386mc3VXV8/arDEJZAD1unbwaPzYGs/uFFjAmQT2SvuoNgc8z7JUd187iPSyCoyR9lFND77AlIybmh8xVFOiQL+0oyBW73yBHuXEip9cFDlX8+OERaaS2ifPd1SMwMUwNV7KJ+2jOivwu7ZRMV4FQWCEtI/akQr6/m1UDGagmKe2eqR91Ce7gj1kGxWvh2G02AfSN0r7qM92BaMx3Ue3UTFKE+FoOmJbpX3UNpHy4puXLPOozRr4/LN9+G5pH7VzY7I1nl1pZwEHn+dFbW1+nBbaHIER98JCmyMwRtoXmh6Bfa59oc8R6Jf2hTZHoEfaF3odgSGufaH/EeiboGbGxXgpn7QvdFaLOXjT49oX7USdBIaEs3akgrN/GxWD1Q8LrHPr30YFpZoogVW4+gsyoNTkKPDKLK8UZEDDQDFHhY679kXPDoHLVHQaU8yibSLTYlo69nVVmzXw+ZBtVPDo0srWeHanYAo4Ai9XoFnaE22OwIh7IdHmCIyR9kTTI7BP2hN9jkC/tCfaHIEe5Uz0OgJDpD3R/wj0TZBharyUT9oTnVUyB296pD1pJ+okMCSctSMVnBSZ5oJM0ibUSWCAtCftRQVlyDYqeNQtC+yU9qRPfwVjxzYqRmkiTKYjttW1J20TAePL26j4kC6tBDKAul170s4NvBWJ7aR70o4P1ZVxgZ3SjsH1ShYYcC+AqN7/BYZIO6jqtS6wS9oxtPqcAt3SDppqcwp0KCcIqtcpMELawVSjtUDfBGvFCFyUOd+LkvAk1UvFCG/7cm2jYjyDxB64cTQ/gpMhY96+0yrt4GGg2PfyeF7Opb2oYA/ZRgUPo8U+8b7NLYOGgWKOCh127XhjG4PCdMS2vjjprG0iYGRs7euqNmvg8yHbqODR+LE1nt0psICDZ/ryoja/825UmwOeZ89M6K61g0gvq8AYaR/V9ODpw0x+je/EHvliQYF+aR/V5mBeXMIhjJqfvgUCTYoCQ6QdLVdVOQV6pB3tDuQK2kbFr1bNG83Bmw7XjmdIqyAIjHAqo3akgr5/GxWD1UUJ9Lt2UDJeTC+qYxsVlIyWDPtcO2gYKOao0HFpH/XsELgocM3vxNM2EdCY3Yi9t+JN2qGKJ4c/O87c69rBoy7c1nh2pZ0FHDymhdmoWdoX2hyBEffCQpsjMEbaF5oegX2ufaHPEeiX9oU2R6BHORd6HYEh0r7Q/wh0TVAPTeM5PpQ5n2tf6KwWc/CmR9oX7UTF/GJOo4JIl1ICNWTaniEDHvXDAgOkfdHnv4I9ZBsVD1bS5CiwU9oXfforGDu2UTFKE+FiOmJbpX3RNhEwvnwaFR9SmyGQetgt7Yt2buA8Dt3zrrSbAo7AyxSapT3R5giMkPZEmyMwRtoTTY/APmlP9DkC/dKeaHMEepQz0esIDJH2RP8j0DdBlblkvJRP2hOd1WYO3vRI+6adqDg+9mwXrH8RumlHKjj7t1ExWP2wwABp3/QFPmA3wtW/jQoe1TaBndK+6dNfwWhM99EOGYzSRLiZPpFWad+0TQRnAF8+jYoPaUYRGCDtm3ZuoM2C66A9aceH6oKnwE5px+B6JQsMkHYQ1fu/wBBpB1W91gV2STuG1mJRgW5pRydLtTkFOpQTBHXtVmCEtIOpRmuBvgnWCgC4KHMuaQdT1TxAykCHtGM8g8QeuHHU2sHJkDGnb1pr7ehOYqCYJlRH8yMoGS+mF9VRkAEloyXDPmkHDQPFHBU6XJABAYPCdMQ2SjtoqG+mc3VfV7VZA0OfFah7XTt4NH5sjWd3CizgYJHNbNTq2jFYr6TACGmf1OaA3uyOux40ACq91gL7pH1SnwNCe7Sw+8pNanPAaPrtO2rtINCkKDBE2if1P+BntPZsBoDAhGnQNioqO6p5kzl40yPtk3aigjXmmAaI1EUJVDfQVpABj6qMQL9rByXjxfSieqR90k5UsHefRsVYBoo5KnRc2ic9O4R6nanotG3nYiz1zXSu7uuqNmtgaMg2Kng0fmyNZ3cKLOAsAntd+0KbIzBC2hfaHIExrn2h6RHYJ+0LfY5Av2tfaHMEekzxQq8jMETaF/ofgb4JMkyNl/K59oXOajEHb3qkfdFO1EVgSDhrRyo4+7dRMVhVRmCAtC/6/Fewm/6P/oIMeNRBCex07Ys+/RWMtFDHpX3Rs0Mg6N5GxVhdxi2mc3VfV7VZY1lMs8a5/0ED4NGlla3x7E7BFHAE9kr7RpsjMOJe2GhzBMZI+0bTI7BP2jf6HIF+ad9ocwR6lHOj1xEYIu0b/Y9A1wS153bZjJfySftGZ7WZh8D3SPumnaiYX8w2KojUGgvUkGl07Zs2oYLyuh8++FBf8Kg3FniZmse1b9qJugjslPZNn/4Kmo5tVIzSRLiZPpHWgsymbSJgfHkbFR/SpZVA6mH3sn7Tzo3V1nj2pB0fqhe1wE5pX89qcwoMkHYQ1TRVYIi0g6pe6wK7pB1Daw2wQLe0g6banAIdygmCaooLjJB2MNVoLdA3wSpz4KKXckk7mKqzWvFuYs/jwTCeQWI6UWG6urdRwcmQMadvWmvt4GGgZHhFNA9KOygZL+Zcj0PaQclosU+8byuEgIaBYo4KHXbtIDBBwdJfo7SvZ20TAXx5GxUfYgBl6Jd2UGr82BrPrrSzgLMK7JX2SW0OeGJsDoj0sgqMkfZJTQ++wJSMm54hg6GaEgX6pX1Sm7MK9CjnpF4HXAxE1zNkwKTRKtA3QYap8VI+aZ/UWWGurm1UjFdBEBjhVCbtSAU9RaZZ2idtQl0FBkj7pL2ooDTC5Uhkk3aigrJ7GxVjGSjmqNBxaZ/07BC4TEWnMcVM2iYCGsbWvq5qswY+H7KNCh6VHFvj2Z0CCzirwF5pX2hzBEbcCwttjsAYaV9oegRWhWqT9oU+R6Bf2hfaHIEe5VzodQReuHzSvtD/CPRN0IQpZc4n7QudFRbELte+aCfqKjAknLUjFZx2572tIIPB6ocFBkj7os9/BbsRLoe0L9qJCsrubVSMZaCYo0LHpX3Vs0PravpEWl37qm0ioDG7ETtHlvAhXVoJZAD1FmRAqUsrW+PZlXZTwBHYK+0bbY7AiHtho80RGCPtG02PwD5p3+hzBPqlfaPNEehRzo1eR2CItG/0PwJ9E2SYGi/lk/aNzmozB286au3rpp2oBYaEs3akgrN/GxWDdfUkMEDaN32BD9hDtlHBo9om8MosrzweDDQMFHNU6Li0b3p2CFzd26gYq8u4zXSu7uuqNmusm+1Q7d9GBU+tT+NN0VwH7U0BH6oLngI7pR2D65UsMOBeAFG9rAWGSDuoqukpsEvaMbQWiwp0SztoavAU6FBOEFSvU2CEtOMd4bW3sEDPBAc9NA0uypxL2sFUNQ+QtrhD2jFeg0RgRDgP2pEKeopMa0EGgzVQBPqlHZQaLwIvlI5aO94Fz2jJsE/aQaN6JrDSHD2yBAIGhemIbXTtoFF9G8xz43d1ddBmDQx9VqDude3gUcmxNZ7dKbCAg1cOcWXc2teONw3plRQYcS9ManNAb3bHXX3toNJrLbBP2rGbdnkICgjpc/of0QQaDR6BHuWc1OuAliLnKsiASeVOoG+CDFPjpXzSPqmzwmuqXNuoGM8gsQduHNuo4GTImNM3zdI+aRMqKEO2UcGjJkpggLRP2okKdqNtbTVujGWgmKNCh107CDQRTqYjtlXaJ20TwWvLXt5GxYcYQBlSD7ulfdLODTwZ9uXTqPiQXlSBlyk0S/tKmyMwQtpX2hyBldMn7StNj8A+aV/pcwReWBzSvtLmCPQo50qvI/DC5ZP2lf5HoG+CKnOr8VI+aV/prNAt46m140nIKggCQ8JZO1JB37+Nigcqq4sSWOfmiT3tRQV7yDYqeFTbBF6Z5ZWCDGgYKOao0HFpX/XsELi6t1ExVn3Nap4bv2uZV23WwNCQbVTw6NLK1nh2p2AKOAJ7pX2jzREYcS9stDkCY6R9o+kR2CftG32OQL+0b7Q5Aj3KudHrCAyR9o3+R6BvgiZMg7ZR8TAa1bzNHLzpKchs2okK1pjTqCBSFyVQQ6ZtGxU8qjICA6R90xf4gD1kGxU8mhwFdkr7pk9/BSMD5bi0b3p2CI8YMkv/xtXDpm0ioHl5GxUf0jW4QOpht2vftHMDR0G4DtqTdnyoLngK7JR2DK5XssAAaQdRvf8LDJF2nGupAlBgl7RjaK0BFuiWdtBUm1OgQzlBUL1OgRHSDqYarQX6JlhlDlz0Ui7XDqaqeYCUgQ5p3wbtRC0wJJy1IxWc/duoGMxAsS/w6XftoGS8mF5UR60dlIwW+8T7NkkFDQPFHBU6LO0gYFCYjtjGggxoqG+mc3VfV7VZYxsy9Es7eFRypNxzodydAgs4m8BeaZ/U5oAnxuaASC+rwBhpn9T04AuY/Nqe146hmhIF+qV9UpsDctNv3/EMGRBoUhQYIu2T+p9NoEfaJz00Da6gbVQwqeZN5uBNj7RP2okK1phjGiBiyJjTN621dvCoygist4VD2id9gQ/YmRQ90j5pJyq6lbq3UTGWemaOCh2X9knPDoGLdrtV2idtEwENvci+rmqzBj4fso0KHo0fW+PZnQILOJvAXmlfaXMERticlTZHYIy0rzQ9AqtCNfW1Y8GvPkegX9pX2hyBHuVc6XUEhkj7Sv8j0DdBhqnxUj7XvtJZrebgTY+040jaZZcc9Zhnu2Ddp1FBxJAxp2+apR1nVzi3kG1UTE1NlMBqCPv72kGpblnglQR0pdYOGgaKOSp0XNpXPTsELlPRaVw9rNomgmLay9uo+BADKEPqYW9BBpQaP7bGsyvtpoAjsFfaN9ocgRHSvtHmCIyR9o2mR2CftG/0OQL90r7R5gj0KOdGryMwRNo3+h+BvgkyTI2X8kn7Rme1mYM3PdK+aSfqJjAknLUjFZwUmVZpP5/wXxdtv+ArsnnwUQOZqNrjC/bLeyaqAnfBfQKfB9d4ueBKdLQJMg+rGTHj7g3VPLgmrYxfPpmaP1XXWRfsF/pMVFfm55Ot+exJff4UL7HgTrHPTLyuggPuj8zKiyw4RPAzL6+84C7JzzzVAF2wW/QzD4NJsENVM1s1QhccIfyZivEr2DlJG7gUQpf4Y5JqvjLm+rxD/jOBCZmMYwJcG1fzN1CC2lPAoN2qmel6A+LhFDDou34yaciWayYysWOfj9/mrTORCRtztOiwyweFHjbKuHvjNQ82+me6Xa9or3Z45NHPAqrX62ciCpKtDe1Pg5Wf80lwdwqY1CFlpmeNQt2r38zEiyw4KAXMxjIJ7kwBs3FJggNSwGxMkmCXus7GKQmOSQGzcU+CnZNk4M7GizlTwGyc2WxeztOVAmZtaz2fBIekgFkbXDNr/7ZsHk0NElzn11+9z6S0YIIvpI76fSZlMhV8ZaZXyjyZyISNOYLUkAJmPZSU6ZjjGqv4eTDz2my6Yve1d9ZOEIzOmAHVnwJmbQw5n6ScVC/YzrMU8qd4iQV3p4DVmCTBIXfIakyS4KAUsBrLJLgzBazGJQkOSAGrMUmCXeq6GqckOCYFrMY9CfZNUk9pIySNF3OmgNU4s9U8eL4rBaza/prnGLN5m5m4SBOsAdTWdJmZqEGCI1IAnhKj1SXBESkAz1shaca9KWDVh8/i55ujSg0pYNXDS5mieyM3D2ZeW82z6/dTwKodI3l0yGZuJtJ4QlWITnx3GviUXuKCe1PA+USTVHBECgCTFqoKjkkB4FLLVHBfCsBYTaEF+1MAeDSYCvaoKxjUKRUckgJApdWqgp2TNIFrvJgvBWBiqopYIrtOyp5BYEIm45gA14bZ/A2UoOZCEEabsLGvGHKsAnKtpap1wQEpIFcrDOn1ktW1VQCITNiYI03HUwAobIh0b/ji4mlHS8Yvb/nmT5mAsv20/U+8yaQUJFtd2k8BpnR0FtydAgaaJDAFmSQw8SILDkoBAy0TvsM0RzZ19pwxlilUcEAKGGiSzoJd6jrQKYGNoek6S4sfTvdUsHOSJnCNF3OmgIHODJPkKr9nFQACioXgkBQwaGNt/ptSgtpTwKDdtGCyryLypIBB22ozacjJ2kxE7RNc/5KNewEgMmFjjj41pIBBD0PlefVvB2Mw89pgunGvaK92oOSvflY67y4EgYiCZKtL+9MwpSO8r5mXuPURCmeM5nUVHHKHzMYkCQ5KAbOxTII7VwGzcUmCA1LAbEySYJe6zsYpCY5ZBczGPQl2TtIGLoXQmQJm48xmc8ioKwXM2naLYLfHixxPzMlMtNmCNYAaC0FgogYJviKsR7eDQUoLJvhC6tkLACmTqeDeFDDrQ2/zH9JY+MMdQXgHvBalgPu3gzHY6J/p2r2ivWxJweiY7WAQMZ5sdWl/GrZ0JPhyNdpTwGpMkuCQFLAakyQ4KAWsxjIJ7kwBq3FJggNSwGpMkmCXuq7GKQmOSQGrcU+CnZM0gWu8mDMFrMaZreYwUlcKWLU993wWHBPg2qibWR3bwRhNDRIckQJWfTRunh7rVK4UsGqbbia9fjbhaiEIDxvWipLg+pOPpwA8LNhQ9G8H44cwr62mu3dfexNbUs6CGVD9q4DE/pRBKk01Z+9tB+NTupQquDcFYLSWIguOuEPApBe54JgUAC698gX3pQCMVZdUsD8FgEeDqWCPuoJBDXbBISlgONE9FeybJHuTwRZ0jvcMKlVFYK7ye1IACEzIZBwT4OzZxTdQgpoLQRhtwsa+AslRCAKpiR5zNsqTAvBoanXeBV9JVtdSAAbrTmHBzSkAw0yImL7h1o4gEBn9M929uykAI0xA2S5ex14ASClItrq0Pw1TOhoEd6eAgSYJj0R51uLU3xQKJl5kwUEpYKBlwndwb6/tAQ5njGUKFRyQAgaaJPDTJOGbHp9+fn/35Wv8TY4u5cHAJCo4JgUMdE/4Cgpi7yRN4JrzW75VACZGVRzNg/a7UsDIXl08sydoOxhMDCDBGkCNhSAwUYMEXxHWw9Ez6qN2z/gCJlFXChjZqwvS/u1gDGbYCG5PASPPZ4HO1JIa9yXwnCXmtdF09+5r78iWFIx+FlDdqwAQMZ5sdWl/GqZ0hKMvvMTNhSCMZk4THGKSZmOSBAelgNlYJsHVxLZtB+NX0yUJDkgBszFJgl0GezZOSXBMCpiNexLsnKQJXOPFnClgNs5sNgebulLAzF5dnOEKOvgCJhNA5nRT+ypgZpsuToZdb7U8nAJmduuCNGY7GER03oJ7VwGzPpT3DFLT0Xm4EIRhTJyz6RtuXgXMbIcBqTlEt1uBmdmSghEx28EgYjzZ6tJ+CrClI8Hdq4BkTJLgkBSQjEkSHJQCkrFMgjtTQDIuSXBACkjGJAl2qWsyTklwTApIxj0Jdk7SBi69mDMFJOPMknlwf1cKSOzVRW9M0OlgMHGRJlgDqHUVkNimC9br3vpwCkj66N7cEBRzOhhEXJwJ7k0BeMlbLeSD1Fj44ykArwAjhemHaU4Bie0w6HIyuya7KSCxJQUjYraDQaTxhFUbo2A3BeBTupQquDcFYLRe14IjUgCY9CIXHJMCsPzUK19wXwrAWC1UFexPAeBRk1SwR13BoAa74JAUACp1TwU7J2kC13gxXwrAxHRlCkyJ6EkBKDSYkLEHmjwdQWA1AWR6dptXAWAyYWNf0eTYCwCpiR7TrespBIHUxI594G9j8QVEJmzMcazjTaGgMCFi+oZbUwCIjP6Z7t4r2suWFJSjnpXOuwtBIGI82erS/jRM6WgU3J0CRpokMAWZJDDxIgsOSgEjLRO+w5Sy2wpBGMsUKjggBYw0SeA3ZxZ69gLAwCQqOCYFjHRPo2BXChh5rB1sUdvBoKIqjuZFAF0pYGSvLoiDtoPBZALInG5qTwEj23TBauoKnhQw6rucziBlEnWlgJG9uqPgzlUABhu9M8exGlIAflVdBYDO1JJa09HIdhjwmF2TvVUAPsXsI5imuT8FjOxPGW11aT8FmNIRRvASN+8FjLMxSYJDVgGzMUmCg1LAbCyT4M5VwGxckuCAFDAbkyTYpa6zcUqCY1LAbNyTYOckKYSz8WLOVcBsnNlsDjZ1pYCZvbqj4JgAZ88uWClB7SlgZpsumK63Wh4tBIGIFkzw5Ue7UsDMXl18Qf92MAabsDHHsRpSwMzzWaAztaTWFDCzHWacTXfvFe1lSwpGxGwHg4jLNltd2p+GLR0J7l4FJGOSBIfcIcmYJMFBKSAZyyS4MwUk45IEB6SAZEySYJe6JuOUBMekgGTck2DnJE3gGi/mTAHJOLNkXizQlQISe3VHwTEBzp5dsFKC2lNAYpvuKLjOz7MKSOzWBWnMdjCI6LwFX5nptaZQEJmwMcexGlJA4vks0NG8NxeC8IZ2XU7gtenaM3xFe9mSgq+O2Q4GkS76J1td2p0GPqVLqYJ7UwBG63UtOOIOAZNe5IJjUgC4tBxZcF8KwFh1SQX7UwB4NJgK9qgrGNQpFRySAkClpc+CnZO0gUsh9KWA6URnBkyL3ZMCQGBCxh5o8uwFgNUEkHlRVHMKAJMJG/sKKUcKAKmJHtOt61kFgNTEjn2ocKPzBpEJG3Mc63gKmPBktKrcwP2ngzGY+ocn9bycAjCCASWYitldCAIp48lWl/ZTgCkdTYK7U8BIkwSmoJ45MPEiCw5KASMt0yS4MwWMdEngsYc7+6/iSJMETnNmoWcvAAxMooJjUsBI94SvcB5eAIMJXOPFnClgpDPDVzCzdKWAkb26k+AQjzOyZxesjtPBGE0NEhywCgCpiR7TretKASN7dfEF/aeDMdiEjTmO1ZACUP7WFIBidvdTSzEX5jWUYw6kgJEtKZPgiBQwsj9lstWl/RRgSkcYwdzVvBeA0cxpgkPukNmYJMFBKWA2lklwZwqYjUsSHLAKmI1JEuwy2LNxSoJjUsBi3JNg1yQXHmufBF/YnClgMc5sMQebulLAwl5dzDFoOxhMtNmCNYAam0LBRA0SHJECFj5JGF9ghK3/pZDnaWGvbsFXZnqtEITBrGEIrkSHm0JBwcS5mL7h1kIQiKh/i3l3wb72LmxJwehnAdXvHxf2p0y2urQ/DVs6Enz5I7angGRMkuCQFJCMSRIclAKSsUyCO1NAMi5JcEAKSMYkCXapazJOSXBMCkjGPQl2TpJCmIwXc6aAZJxZMgebulJAYq8u3vUedDoYTLTZgrtTQGKbLlhjtoNBRJ8t+DI91yogsVcXX9C/HYzBJmzMcayGVUDi+SzQMcc1p4DEdpgpme7eK9rLlhSMiNkOBpHGE95rxCjYnQY+pZe44N4UgNG6oio4IgWASS9ywTEpAFxqmQruSwEYq4Wqgv0pAO+O0vJkwR51BYM6pYJDUgCo1D0V7JykCVzjxXwpABNTZwbMVX5PCgCBCRnTq3v27AWA1QSQOd3UvBeAF4WZsLEvoXLsBYDURI/p1vWkAJCa2LFvXmjcCwCRCRtzHOt4CgCFCRHTN9yaAkBk9M90917RXrakYHTMdjCIGE+2urQ/DVM6wgNmWYpsXgVgNK+r4JAUMNIk4RtM2+74Bo/nn3H4ASms/ck54OKVF9yZAka6JHCyocBz2400SeA0ZxY6fymT6Gicku99AZgYBVGwKwWMPNYOZnoxZwoY6czwpGPfdjAITMjYA02uFDCyZxffYAOosRCE0dQgwfUGdMUiX0KFL4g5HQwiEzsZX5nptUIQiIzemeNYDSlg5PksPLe6fzsYg43+me7eK9rLlhSMfnaSqrsQBCJ6eltd2p+GKR3hESC8xO0pYDEmSXBICliMSRIctApYjGUS3JkCFuOSBAesAhZjkgS71HUxTklwzCpgMe5JsHOSJnCNF3OmgMU4s8UcbOpaBSzs1cXzaJ4d5ux/DiKYaLMFawC1poCFbbpgNXUFTwpY+CRhkJo+F8deAIiYTAX3poCFzxEGqenoPLwXgGFMnIvpG25eBSxshwGp2TXZOxqGpxNx2SaYitmfAhb2p+DUBKNgPwXY0pHgyzTaU0AyJklwSApIxiQJDkoByVgmwZ0pIBmXJDggBSRjkgS71DUZpyQ4JgUk454E+ybJ3mSc1Yk6HQwqOrNkXkPQlQISe3VBHLQdDCYu0gR3p4DENl2wxpwOBhF9tuDL9FyFoMReXZzGMsm0tRCU+BxhEBkLfzwFJJ7PAoUpM7TPhXktme7eK9rLlhR8dcx2MIg0nrDTzSjYnQY+pZe44N4UgN10zWkFR6QAMKlJKjgmBYBLxaHgvhSAsVqNLNifAsCjwVSwR13BoAa74JAUACp1TwU7J2kC13gx3yoAE9OKAxo2KBE9KQAEJmRMr65rLwCsJoDM6abmvQAwmbCxL6FyrAJAaqLHdOt6UgBITezYNy80yi6ITNiY41jHC0GgsCFC8966CkATkNE/0917RXvZkoLRMdvBIGI82erSL6bx9vbh/eNX/w8AAP//AAAA//+snNtu3EYWRX/FEPKQPEzsVku+QRYQqeXmtSV1k5hnw3HGwWDiwPZ4Zv5+SNY55K5alCw7fMllcVd1q3axWLV5pLNP79+9+7x58/nN+dnHD/959PHV0ero0ac/3/zxqfuvl8fd/7z//Opo/eTo0dt/f/r84V/Zu9//0ZPuwn9XJ2/evvz1f5t3n96++6NjT34+OTo/e9t38kvfy6ujk+5f3YVPHf5y/uTs8Zfzs8dvTXIxSh4buQTZgFyBvAbZgmQgOUgBUoJUIDXIDuQa5AbkFmQPcgBpQFoljztLR1+PF/G17yXxdZX4OkpGX0E2IFcgr0G2IBlIDlKAlCAVSA2yA7kGuQG5BdmDHEAakFZJ5Ov6ob5+fv/7239efLj75j0d796+z1dH3UeON+9xYrIrRo9TsEnBLgXXKbhJwW0K9ik4pKBJQSsgGreTmXE7fvHz8enMUjcMXfPhz6+ve32v3f2hQ7dOhs4k66Nx7EA2IFcgr0G2IBlIDlKAlCAVSA2yA7kGuQG5BdmDHEAakNbIST+qkc+dn3ieHXfqb3ue9b10vsodcZLYGhTdP8d75jRWXFLxNFZs7FO0k2ex5MokT8fZ8xpkC5KB5CAFSAlSgdQgu0CeykA8j3+Gaw7Ei1hxQ8Uq2T7czkiSJ9HevtuzcawOIA1Ia+Q5ZlL3I33/TJoW176bfiqdn/12vr/++48//a3/5w+//HD609nj3/qdUvpEDS2mH+MyBZsAhq887L+uAngxzZIAVk9GsgXJQHKQAqQEqUBqI6vx03dGjkdybWRaF2+MDPf08GPdGjkdW+0DORGPQRqQ1gg9fraMx3035vGwHrx/8/Hdr0ePPr77rdsbP3v5S7+x/vR7t2U+umMKpM/b0KFMgRRsApApEIBMgQB0CoBkIDlIAVKCVCC1EZkCRmQKGJEpYESmgBGZAoHoFABpQFojnALPl5kCfTczUyC43j2kvpynG4PQQjxOwSYA8TgA8TgA9RgkA8lBCpASpAKpjYjHRsRjI+KxEfHYiHgcSH8OHR+vq+QeOYwa31g1IK2R4RtG24QXy7jed3O/6+m+IbQQ11OwCUBcD0BcD0BdB8lAcpACpASpQGoj4roRcd2IuG5EXDcirgfSue6OHkAakNYIPe4GaJEn+NDP/S4ne78LayI2g2yMiNFGxGkjajVRRpQTFUQlUUVUOxLHHYnljsRzR2K6I3HdkNpO1BC1jmac75+3aar14FPAtHdbWbzSb970wS6rerKnv7Am6nzoZCIb06jzQaPOBxI5D5RZT6LKiQqikqgiqh2p8/Yl1HlD6rwhdd6QOh9Q5DxQY19CVK2jGef74GQJ5y2Aucf55Kh2sQpN1PmUbEyjzgeNOh9I5DxQZj1FzkNVUFUSVUS1I3XeulfnDanzhtR5Q+p8QJHzQI19ich5U804P5eMfc89HxKk4cB2xz2fHHAvVqGJOp+SjWnU+aBR5wOJnAfKrKfIeagKqkqiiqh2pM5b9+q8IXXekDofkBpv+Vx36p52dsnO+GBfocvmxq0dUetoyEaizd2qT4eWWAQsZbpnEUiSjIvho18d6VQInejyH4hOhUB0KgQSTQWgzD4tmgpQFVSVRBVR7UingnWvU8GQTgVDOhUC0qkQiLh8sA+MjIeqddWM8X2Ys4TxFgrdY3waUF2sQpvuX9MLroDU+kDU+kDUeutIshvve0IZUU5UEJVEFVHtSK2376XWG1LrDan1nt35wOyt9z58n1aB5KR0mETTKhB6kvnRumpmMvSpzxKTwdKj+yZDGuGtkFxdGtLJEEQ6GQLRyYBwbet962SAKqeqICqJKqLakU4GBnqu0snASM9Uug4EUbQOADXWLrLeVDPW92nPEtZbanSf9Wl0t0JidWlIrQ8itT4QtR6h2tb7VuuhyqkqiEqiiqh2pNYzyHOVWs8oz1RqfRDF60Bylj5Ys+ihMLbzpaF11cxk6EOgJSZDCJPu2xiu0oxvhWjr0pBOhiDSyRCITgakb1vvWycDVDlVBVFJVBHVjnQyMPFzlU4GZn6m0skQRNE6ANRYu2gdMNWM9X02tIT1IWO61/o06Fsh37o0pNYHkVofiFqPCG7rfav1UOVUFUQlUUVUO1LrGfu5Sq1n8GcqtT6I4nUgSVYO1ixaB8Z20zpgiJPheKE4cOjn/jhwleaB1kY3h4amUdgYkclgRCaDdySbQ6KMKCcqiEqiiqh2JJPBkWwOHclkcCSbQwzD3shaiypWSdhycNHUeUPUOmKZwPBGboGVYejnK5MhjQitTTQZQuylkyEQnQzICL0jnQzMCKnKiQqikqgiqh3pZGBG6CqdDMwITSUrg5H1hA6OppqIhqh1NCy2cWXcXES4Pp4vBZqrepzy4WNPCX35uTASeRtE6m0g6i1SQO9IvWUKSFVOVBCVRBVR7Ui9ZQroKvWWKaCp1FvL8qZxOJhoPa14DVHrwz6MTeztQiHgsYeAk7GWbsnx3kRqbBCpsQj5rJUmO0QZUU5UEJVEFVHtSI1lyOcqNZYhH4ZhbyR+nCeh6WESjcd7otbRzOP8wSHfN1RC9nlxXwo9rS0XjiTSI9o4UuutL316A2294aTKiHKigqgkqohqop2h7nHpXlzzZ7xx1TRpbqnau2oarwNRQ9Q6GoYwvqsfHOs9uHJz10/Ovuj1eCjMOrT1j7unL7uB8KKspP7sxvXrUX/z9GU3JHfoD67vtqNTxJWkxc0kGvevEYpHYe4c++2l+sd2jO2evnfW6k+aMcsk2hBdEb0m2hJlRDlRQVQSVUQ10Y7omuiG6JZoT3QgaojaCMVuzx1dv8NtO7mq26jgHzWT20CbY6ArotdEW6KMKCcqiEqiiqgm2hFdE90Q3RLtiQ5EDVEbobiqf+5sOuv2t5T1951+pa7fJaPz/YcOjZxsQHYg1yA3ILcge5ADSAPSKokHsT9Q4Ez3l2v81+Gccm+Rv2umndIl0Yboiug10ZYoI8qJCqKSqCKqiXZE10Q3RLdEe6IDUUPUOuKJvs8LFqj8H7q5v/TfJPryKq39n5Gkxf/+QdpNWv3vGin/J9oSZUQ5UUFUElVENdHOUPdKZtxVpL8HMDMo6S8CzEjwmwBzmvRXAfwLyiaUqCFqHXETul7oaDn0ExIjlv7bRTlrgGyMyEnDiBw0jOghkygjyokKopKoIqodySHTkcSEjuSQ6UhiQkcSHxjSQnCihqh1NOPxg8+U9+dC63DO698XfDnH79KFi+pxSjbWgXocNOoxKj221kxsz4hyooKoJKqIakfqMUtEXKUes0TEVeqxDaTex0CND/ekah3NePzgw+RXPA4HyOBx+t6vyyv7XZR6nJKNadRj1IKYJrqPUR6SUZUTFUQlUUVUO1KPWQviKvWYtSCuUo+Dqt/W3V3sb+2kFLAhah2xOnC9UDHI0I+v3ukrP7uorocSBXnjZxp1HUUfpolcR4VHRlVOVBCVRBVR7UhdZ9GHq9R1Fn24Sl0PKjH0YKrIY6haV814vFDVx3qq+vhynr7Js4vqcZCrx6jusFa6erO6w0TR6s3qDqoKopKoIqodqces7nCVeszqDlepx0EVeQzUWENRtY5mPF6omGM9FXN8OU9f0NlF9TjI1WMUbVgr9ZhFGyaKPGbRBlUFUUlUEdWO1GMWbbhKPWbRhqvU46CKPAZqrGHksalmPJ6Lvvp0+EG/wy1/1sASq2EXllbrd6+Y0id0Sjam0bU6aNRjq2+Q93TWLPIYqpyqgqgkqohqR+qxfaLutA2px4Z0p21IPQ4o8hiosS8ReWwqenyyUDHG0I8/j9O6fLso9zHIxoh4bEQ8NqLPY6KMKCcqiEqiiqh2JB47Eo8diceOxGNDYrGR+JVdWpc/icZXdkStI76y6/8+0BLlWEM/bnpagW8X1fQQ4MnibRo1HcUVpolMZ3EFVTlRQVQSVUS1IzWdxRWuUtNZXGEqNT2ItOLORFpkRdQ6mrH4L6Vw09rdp632m7VdSpL+LSm7qlUWhtRkVFmYRu9slFRsvW+prSPKiQqikqgiqh2pyayycJWazCoLU6nJQRTf2WmtvTWLbB/bje8mXTVj+0Lh2InXXQx/FyN9UWVXI9tDA7UdNRjWTG23KgZ5aHvfajtUOVUFUUlUEdWO1HbWYLhKbWcNhqnU9iCK7m2gxtppNa2jGZMXSsdONB1Lf63/wq5GJiMfM5Eu4MjHvCM1GZFZRlVOVBCVRBVR7UhNZj7mKjWZ+Zip1OQgiu/ttH7emkX39thuurcNzdi+UGB2ooEZKuXtamQ7IjMTqe2IzLwjtZ2RGVU5UUFUElVEtSO1nZGZq9R2RmamUtuDKLq3gRprF93bppoxeaF87CQENiEVXaUBmV2NTEZCZiI1GQmZd6QmMyGjKicqiEqiiqh2pCYzIXOVmsyEzFRqchDF93ZaE2/Nont7bDfd24ZmbF8oMutewsh2Lc3M7Gpke2ggtZMmUtuDRp/bDM28b31uMzSjqiAqiSqi2pHaztDMVWo7QzNTqe0WkHW7rykNT6vfrVn3Ims6iFm7CbWu4rvyk4VStKGf8V1mGqPZ1cj2EPqo7cjRrJnazhzN+1bbmaNRVRCVRBVR7UhtZ47mKrWdOZqp1HZLwyZ0MFGXd08mm2pCrauGHXD8Fy/nYrTvq3M/8SRtLIc2EnkbROptIHpLIz/zjnQlR1iWUZUTFUQlUUVUO1JvmZ+5Sr1lfmYq9dZSsGkcDibSOnei1oedde6nC8VnQz/9DTwaa0SNNSTGGhFjjchN6x2JsUQZUU5UEJVEFVHtSIx1JKGZIzHWkYRmGIa9kfgRnda5T6LxNiZqHfERffrg0OwbSgOHTuM6d0cSnBFtHKn1FjWp90Bbbyh17kQ5UUFUElVENdHOkNa582e8cdU0aW6p2rtKShOIGqLWEUsTTufCs/69OF58PLzOfegzrnNfr192I3FXobs3kEL3rkE3KHdVunuDbpW7u9J9Eo270giFJ9fj6e/T/x8AAP//AAAA//90k8tu2zAQRX+F4AfU4ksvWAYMB0GzMFpIFrpW7bFERCEJimnRfH1HTtAuMt5Jc8WZc+9Q2xeIIxxgnhd29q8uNVyWfLf9V2YRrg3fi7oXfPO5LuteUnVd95qoP5j6aIh6a+quIOonU/dU/UHJ+qioyS0qnaJmn1DpSaUt0R3ZTChUqGatQDKRU1ZEgUpJKhUqFaXIDHMkCSQSSIqgk0ggqTD3Ev3gFql1IYGkCPYKN6yoFbcG2QyldAYDNfQSkE2RTjV20+Qcjd20otLR620ir43OUSG3oHEL+kaw+X/Fd9sweQfJnr9HdvUuPV0arjhLfwI03PmDd78gLta7FSNE69K3kPB1YZOP9g1PDPMBXIIIePJmIwwjHIc4Wvxohiv+QNkXUeWlyVSliiwrygL9smjH6Z6WfLhz6qdPyb/cEScYLhBXUQld5Vkuq0oKaapSc3TnkXIVjRClEJlUuZSZLt6v2UrdQXoNLAwBYmffMICKs+U8zPhU5tjBppP/Ch/cnGEA6HxY42j4PLgLfhsAndUWs4hPl1sc71CPt+lsmO3oftg0fQSEOGuwm98+Pi8TQNr9BQAA//8DAFBLAwQUAAYACAAAACEAG9IFPWEHAADNIAAAEwAAAHhsL3RoZW1lL3RoZW1lMS54bWzsWVuPGzUUfkfiP4zmPc1tJpdVU5Rrl3Z3W3XToj56Eyfjrmcc2c5uo6oSal9AQkhIBcEDEjzxgBBIIFGBED+mqBWUH8GxZ5KxN05vbBGg3UirjPOd48/nHB+fOT7/1u2YekeYC8KSll8+V/I9nIzYmCTTln99OCg0fE9IlIwRZQlu+Qss/LcuvPnGebQlIxxjD+QTsYVafiTlbKtYFCMYRuIcm+EEfpswHiMJj3xaHHN0DHpjWqyUSrVijEjiewmKQe2VyYSMsPfbL+89/uzb3x7+/PuXH/gXlnP0KUyUSKEGRpTvqxmwJaix48OyQoiF6FLuHSHa8mG6MTse4tvS9ygSEn5o+SX95xcvnC+irUyIyg2yhtxA/2VymcD4sKLn5NOD1aRBEAa19kq/BlC5juvX+7V+baVPA9BoBCtNudg665VukGENUPrVobtX71XLFt7QX13j3A7Vx8JrUKo/WMMPBl2wooXXoBQfruHDTrPTs/VrUIqvreHrpXYvqFv6NSiiJDlcQ5fCWrW7XO0KMmF02wlvhsGgXsmU5yiIhlV0qSkmLJGbYi1GtxgfAEABKZIk8eRihidoBMHcRZQccOLtkGkEgTdDCRMwXKqUBqUq/FefQH/THkVbGBnSihcwEWtDio8nRpzMZMu/BFp9A/L44cNH9354dO/HR/fvP7r3bTa3VmXJbaNkaso9/eqjPz9/1/vj+y+ePvg4nfokXpj4J9+8/+SnX5+lHlacm+LxJ989+eG7x59++PvXDxza2xwdmPAhibHw9vCxd43FsEAHf3zAX05iGCFiSaAIdDtU92VkAfcWiLpwHWyb8AaHLOMCXpzfsrjuR3wuiWPmy1FsAXcZox3GnQa4rOYyLDycJ1P35Hxu4q4hdOSau4sSy8H9+QzSK3Gp7EbYonmVokSiKU6w9NRv7BBjx+puEmLZdZeMOBNsIr2bxOsg4jTJkBxYgZQLbZMY/LJwEQRXW7bZveF1GHWtuoePbCRsC0Qd5IeYWma8iOYSxS6VQxRT0+A7SEYukvsLPjJxfSHB01NMmdcfYyFcMlc4rNdw+mXIMG6379JFbCO5JIcunTuIMRPZY4fdCMUzJ2eSRCb2bXEIIYq8q0y64LvM3iHqGfyAko3uvkGw5e7nJ4LrkFxNSnmAqF/m3OHLi5jZ+3FBJwi7skybx1Z2bXPijI7OfGqF9g7GFB2jMcbe9bcdDDpsZtk8J30pgqyyjV2BdQnZsaqeEyywp+ua9RS5Q4QVsvt4yjbw2V2cSDwLlMSIb9K8B163QhdOOWcqvUJHhyZwj0AVCPHiNMoVATqM4O5v0no1QtbZpZ6FO14X3PLfi+wx2Je3XnZfggx+aRlI7C9smyGi1gR5wAwRFBiudAsilvtzEXWuarG5U25ib9rcDVAYWfVOTJLnFj8nyp7wnyl73AXMKRQ8bsV/p9TZlFK2TxQ4m3D/wbKmh+bJVQwnyXrOOqtqzqoa/39f1Wzay2e1zFktc1bLuN6+Xkstk5cvUNnkXR7d84k3tnwmhNJ9uaB4R+iuj4A3mvEABnU7SvckVy3AWQRfswaThZtypGU8zuQ7REb7EZpBa6isG5hTkameCm/GBHSM9LDuqOITunXfaR7vsnHa6SyXVVczNaFAMh8vhatx6FLJFF2r5927lXrdD53qLuuSgJJ9GRLGZDaJqoNEfTkIXngWCb2yU2HRdLBoKPVLVy29uDIFUFt5BV65PXhRb/lhkHaQoRkH5flY+SltJi+9q5xzqp7eZExqRgCU2MsIyD3dVFw3Lk+tLg21F/C0RcIIN5uEEYYRvAhn0Wm23E/T183cpRY9ZYrlbshp1Buvw9cqiZzIDTQxMwVNvOOWX6uGcLkyQrOWP4GOMXyNZxA7Qr11ITqF25eR5OmGf5XMMuNC9pCIUoPrpJNmg5hIzD1K4pavlr+KBproHKK5lSuQEP615JqQVv5t5MDptpPxZIJH0nS7MaIsnT5Chk9zhfNXLf7qYCXJ5uDu/Wh87B3QOb+GIMTCelkZcEwEXByUU2uOCdyErRJZHn8nDqYs7ZpXUTqG0nFEZxHKThQzmadwnURXdPTTygbGU7ZmMOi6CQ+m6oD926fu849qZTkjaeZnppVV1KnpTqav75A3WOWHqMUqTd36nVrkua65zHUQqM5T4jmn7gscCAa1fDKLmmK8noZVzs5GbWqnWBAYlqhtsNvqjHBa4lVPfpA7GbXqgFjWlTrw9c25eavNDm5B8ujB/eGcSqFdCXfWHEHRl95ApmkDtshtmdWI8M2bc9Ly75TCdtCthN1CqRH2C0E1KBUaYbtaaIdhtdwPy6Vep3IXDhYZxeUwvbUfwBUGXWR393p87f4+Xt7SnBuxuMj0/XxRE9f39+WK6/5+qG7mfY9A0rlTqwya1WanVmhW24NC0Os0Cs1urVPo1br13qDXDRvNwV3fO9LgoF3tBrV+o1Ard7uFoFZS9BvNQj2oVNpBvd3oB+27WRkDK0/TR2YLMK/mdeEvAAAA//8DAFBLAwQUAAYACAAAACEAUlyT7IcGAADDOAAADQAAAHhsL3N0eWxlcy54bWzkW8uL20YYvxf6PwgFeijV6mHJr9je7q5XEEhDaLbQQ2HRSmNbRNK4o/GunRIoFEou6an0Ab0GCu0hh0J76X+TbNv/ot+MJEvO2ruyV/IjudgzI+mb33yv+eabmdb+2PeEc0RCFwdtUd1TRAEFNnbcoN8WPzsxpboohNQKHMvDAWqLExSK+53332uFdOKhRwOEqAAkgrAtDigdNmU5tAfIt8I9PEQBPOlh4lsUqqQvh0OCLCdkH/merClKVfYtNxAjCk3fzkPEt8jj0VCysT+0qHvmei6dcFqi4NvNe/0AE+vMA6hjVbdsYaxWiSaMSdIJb73Sj+/aBIe4R/eArox7PddGV+E25IZs2SkloLwaJdWQFW1m7GOyIiVdJujcZeITO61g5Js+DQUbjwLaFivTJiF6cs9pi3pFFCKhHGEH2HQqfSjc+ejOHWVPUU6lu1/MVtnTD74cYXpXiv729+GlU+njU0mUkw4z1NVadQH5LG1OYlnCtVnCfoTn8pdnUcGJ6z++iAoMnhwzpNPq4SDlC1DiatB8HOCLwGSPQPeBWeytTit8IpxbHrRofIiWj6L65Q8v//nj5eVP3/332/fsSc/yXW8Sv8sa7IFFQjCI6OtKlUOIaK6Ncv0NzL/+fPnizzmA9aUAN0rjBMcRpDwuBu+bojtxfRQKD9CF8Cn2rWCu9OS5kor4hD1MBAqODQxGLYjDvL8Q1Mj1vKnJakwLoaHTAu9GEQlMqAhx+WQyhP4DcMSRZvH3bni7T6yJqhmZD2TeYad1hokDjj9xFhrzFlFbp+WhHgWFJm5/wP4pHsLvGaYUvGOn5bhWHweWxyws+SL7JcwYMDm0RQePwBMDWZvzzw0cNEbggqqcqTLrJO4j+YIOYDpY9D5Hw8Hk7ABgJ6hzdRANMP/4rkX7zo3uBmlvnfRK185r1WPHdXMnxhZ5rlymX5Y4Zp3oRqHk6nyO11rarjeiGwmjV5xJ3mJ9zjFh305Ll+gg58R9da7IK58lopO8s/fy8c+bupjbglYJslYcRqEB0YJQsZQ+llahKVNLiQG3Loq4AdCWxbjL+t0Nje52kURBoIuJJTasH6VHE7cYX7yYhTW5jTzvEVvEft6byaaNe5lcF+RKWcqAJdVYEVbjcTFaC0eVTsvy3H7gowDyQohQ12apJRuqiPBF+bjXac2QjRJ0EV11IV3BGg69Cctd8d6jGkBIa4d8fZ/WDxIcadNDgimyKc/9KrDuXhZqhgPGQqQsS5DlZ8TdDGP1BnS9PGeFcW9VFmdw64s5nNDPshoEwlmdYSZLHKaMG2DiPgGpZKScU+63AXUr0ZXJAraeEItjAEC9TiqJym9WSmAKGwYZZ+EjH7JIvFOYCbcejPwzREy+YcPSnPP0fss4rBUOknmrKy45Y5pse2Oep4f2woVesPGwbY/YZRYm2hUhLuLiDkBsbD8Xy3BABetiGVyMg6qcsw0LreLo7UYPmbiZsj3kXOej1vjOZ05nDq6KByi748xVOGKw/ZO6unmNzREgavDO9vNSA9vb9skylfeWhUe5YrvdRj9V4l1k/RULzOOJS1hO5tKTlTi9mbXvDgQ1ZUxkBcddpUxjBWPUINza9unhhnn2rTcR4YJYwxM0huwjPw917Zp5UXy9RnspBK9aRnS1wHiKAbzGQKsYwGs0/WIAw2nadfmqQgAX4bhKVdkyUrylAl5jYFKMyq4xVVkM4DVmBZcCDM5qwa7ftT5hM8FB7qmglL2k3H69lN5zO+lSei/D3pbLdebIHJVhY4WD3PzOWg5ObiSdyXfYYU89c3Rh5uDCdANeYHct2uKrv77+99nvmdjlbOR61A2iHXXIQ8//QJjaMSuQ5siFww9fqQ3NqOlmQzKPG6qkK1VNOqgeVaTjA7Wrd2vdeq1uPuX3TZJjFTGK18+/ef3s21d/P0+AgJPKAKnwAxJTJDA4Z5yex1DYU8ruevGTGtPhgogc1LNGHj2ZPmyLafkT5LgjH/DHbz10zzHlJNpiWr7Prj6o/AoPrHLuh3BTAf6FEXFhwMeHtUb32NSkunJYl/QKMqSGcdiVDP3osNs1G4qmHD3N3Di7xX0zfkEOjkeoejP04FYaiQcbg3+UtrXFTCWCz/kHsLPYG1pVOTBURTIrCgiratWlerViSKahat2qfnhsmEYGu7HivTRFVtXohhsDbzQp3MDx3CCRVSKhbCsICarXDEJOJCGntw87/wMAAP//AwBQSwMEFAAGAAgAAAAhAA/sD6dSAgAAeQcAABQAAAB4bC9zaGFyZWRTdHJpbmdzLnhtbKRVX0/aUBR/X7LvcHPfaUvVRZtSH0yW7GHJHrYP0EAVErhl3GLmKwaRIIGHgcGxCct01TEmkky0ML/Lwr0tT/sKO4BLzP5ostuHJj09/f0595xTffVVKok2rQxN2CSCw5KCkUWidixBNiL4xfPHoWWMqGOSmJm0iRXBWxbFq8bDBzqlDoJvCY3guOOkNVmm0biVMqlkpy0Cb9btTMp04DGzIdN0xjJjNG5ZTiopq4rySE6ZCYJR1M4SJ4IXgTZLEi+z1to8oCrY0GnC0GckGk2bUSAHFGplNi1sjK/biB9U0aQ0RPxjCQW5Nhp7I8SbRTRpVBHLdxCrDJAuO4aejoN0JxF9lkHrNnGexCJ4ASNnKw2QxF6zyY1/LBu6PCWdERuBm0MIBf0h3Fm9KwLFLo58t8SqZd/taeNB3a/t8fNLv9sD6L9d83z/wg0ah9o/cv4Ig14Qy/ePePNQC4eXZGVFVleEZHstVtu5T8DkjRec9O7OYrm+D6rusDK++hCcuzfqRURPWnneaYshVP3u/rz6Ijj8bTXonwq14HEFSsZbwxnIrd7ktcGkUBE623p3PCoLlalQFtfAti+FXOQagetpIhBhifeugs9nbHeH7eWh338MD+4ZOV685q87IqSqBBsFvAspX5D8wUgUZFFi+a+suz2dze/1NnRb89R3878qMA2x6ieIiNhdkubjzZpnE++dkGneGrDCbzuJfTmB8O1DGw+84NuxEFHQduGMYAH774caehpSQ8pySFGBxS/uwobRFPX/SiLDr9P4CQAA//8DAFBLAwQUAAYACAAAACEAO20yS8EAAABCAQAAIwAAAHhsL3dvcmtzaGVldHMvX3JlbHMvc2hlZXQxLnhtbC5yZWxzhI/BisIwFEX3A/5DeHuT1oUMQ1M3IrhV5wNi+toG25eQ9xT9e7McZcDl5XDP5Tab+zypG2YOkSzUugKF5GMXaLDwe9otv0GxOOrcFAktPJBh0y6+mgNOTkqJx5BYFQuxhVEk/RjDfsTZsY4JqZA+5tlJiXkwyfmLG9Csqmpt8l8HtC9Ote8s5H1Xgzo9Uln+7I59Hzxuo7/OSPLPhEk5kGA+okg5yEXt8oBiQet39p5rfQ4Epm3My/P2CQAA//8DAFBLAwQUAAYACAAAACEAuIYweugFAADMfQAAJwAAAHhsL3ByaW50ZXJTZXR0aW5ncy9wcmludGVyU2V0dGluZ3MxLmJpbuyYTW8bVRSGT9yKD7FB6oYFC5RVF0WkbdoCQki2xy6uYo8VO6Wsool9nQyZeKzxuEpaumPBkjUL2LIAiQ1rFvwAxJ4NIPb8A3jPvZ7YTsZ2nFaCxeuRZ25mzpxz7nPPx3V+97/s3ZSq/Pr1s8eD1sEXu5/KSp+1q1de+k3+/qz801qhIK/KV69tvtKVNXlZHhXWcH1UuIJzUTZXU7tQem38VK8FHeP0Dz73a60ZM16tsbMu1wufX42Of25+88sipdfwUPW5r7OQ2XmBrlPV/5jAKut9HTHXqrcf6HRelzcKT+Wm3JaK3JESvvcQ8W/jegt3bmG0Ke/h+i5GJdmA3B2MiuJhVMb9KqRu26OIv59BY60/GKWlsC/FnbYv25WWt7UlO/0wMUMdNYOBSVrhEyPFTambbhi0TwZGWmnQ7wZJd3fbdE46kemKn4SmnwZpGPdlq9jwWuVis7JbLt+6tyHlOIqTetw1UvcbMGKGcTSyknc3NrqDULzRIDLH0vAbFRWOgtSI34DxfeMnXZNINYn7aTsuBZ1D8Ufp2OVmEvZTk3imF4yi1IoPm3D3wJgUkEpxfBhh5Pd66vAgCvv74ler0hz1OwfZH7UjWPEw3471qBEnR0Ek95NgcBB2htbrh3gWJ1Lp7pt2rGdr2Oq1o2AvMsXEBJjZ/igKkv8m8KJjkYebXj2LrQ+/rb31JlyJxrXmyY9/Xftj+3rthzvff/DOn7e/a+L+DXyvqLt4N/t4EktHRnIkRvqSSgvXFEeIv/ZlaMWeyroMMD6U96UpAe4b2cZ3iHcjvKvSMeTX8Xwd1yFi0R3vy67cteMuNIR4euOcRKbxY2hOoTWBLwHOh7n6fOnhWE3P2Tk5Pxfp+Ah+BNK13lQxt3jsWd4M53s0T0sTs1O+CGJrIV+ryuxbuQCUM0p5Xs+z442JvhivM21N0BjCr0Wr3obMQLYwux7kLuOzUneRoO+fj6tPxB11qaPeeYgzjeQUNW/6vfPRNo9Vw+bA3oL1mBd3F9F4EWYl63+K8xGqtctGzYZV6E3PYhlBHxEYyYnMzvzixNrw8diu7vnVWT0nMm2rkrpMjGW2lhGqIsojzLKBPNL6eJmKoTY00yc6lNbzaGrBqyfWG9V0E9X1MtrKiPch/NJom13JRdp8eWwzJAKRE3neiv0A1vfQSUJ0oElNdHfzcv6sZOM016NcBpn+JvKpJbO+u2oWnK6LZ2vV/M6wA1qajUXb7w7G+amea9dy/W8L/riKnUlnq94GrcHpms12yEzWeaB9waDz5K2Cm89EXtcusB1ikX7ncZyr0RHNNNZQQ1+kpx5YDWwGHYNbhOPsbmOIOyGoGnBc1E/m1xOtmjH2Cpqn2k/d7sSHTte983ufrkUfc41Pe5S+4bpWHvmzfpfxps7H7Vh0DppL2rH3luyHJnGu0TSr5yKWK/A7gBWd79m3z9fhWWvzY8rx1QjP77bLMi3jPr2zCWzEJ8i86Xo1G/utcWUtgl+mo27jP8STyZvalc7uQFs296LxTjVvnbV+p3ZVVbeu+bQ3y3VmFrTyZ7U2v9q6zNRepHtntaMVQneQdUSKRmKef7Or4/I/QPxEc/ZMzsp0VmXatZfjdxrmq7a6U9V0tb2fszDxK+sMuj4lG++dlfflTqeuxTE8072N5o7u0i7jm/atxFYM92vlwGpK7W+YLHaVQQ/0l5Gs2P1LAskdqZ375TO/F0z3Lc3HnvWhA6tZJdU5qhfLusqkXi3uq4t+wUzyQuuo5rCSmXSli1VYzXCNobwaUQXvbAc0f6+n/1/ghwRIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIgARIQAn8CwAA//8DAFBLAwQUAAYACAAAACEAm2fmDd0AAAD8AQAAEAAAAHhsL2NhbGNDaGFpbi54bWxkkctqwzAQRfeB/oOYfSNLbR4tloMJdFmySD5AyNPYoIeRRGn/vkqJTZLZCHR0uZyL6t2Ps+wbYxqCVyCWFTD0JnSDPys4HT+et8BS1r7TNnhU8IsJds3Tojbamn2vB89Kg08K+pzHd86T6dHptAwj+vLyFaLTuVzjmacxou5Sj5id5bKq1tyVAmhqw6KCw0oCG4oEMHs5+ZV/znwiB7m6JucMIa0sPf9tU6aVZdw9EW+ElL0PmQ0ha0IefVrxSjIvhEx7Z0NBnAVxJsrEmAjf+PL535o/AAAA//8DAFBLAwQUAAYACAAAACEAq/J2wkoBAABfAgAAEQAIAWRvY1Byb3BzL2NvcmUueG1sIKIEASigAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjJJfS8MwFMXfBb9DyXub/qGbDW0HUwaCA8ENxbeQ3G3FJg1JtNu3N2232qEPPuaec3/33EvyxVHU3hdoUzWyQFEQIg8ka3gl9wXablb+HfKMpZLTupFQoBMYtChvb3KmCGs0POtGgbYVGM+RpCFMFehgrSIYG3YAQU3gHNKJu0YLat1T77Gi7IPuAcdhOMMCLOXUUtwBfTUS0RnJ2YhUn7ruAZxhqEGAtAZHQYR/vBa0MH829MrEKSp7Um6nc9wpm7NBHN1HU43Gtm2DNuljuPwRfls/vfSr+pXsbsUAlTlnhGmgttGlaaMoidymOZ5UuwvW1Ni1O/auAr48lcvHNegc/xYcrQ8/IIF7Lg4Zwl+U1+T+YbNCZRzGMz/M/DjbhClJMxLP37u5V/1dvKEgztP/SZyTNCVJNiFeAGWf+/pLlN8AAAD//wMAUEsDBBQABgAIAAAAIQDvkXixqAEAABYDAAAQAAgBZG9jUHJvcHMvYXBwLnhtbCCiBAEooAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJySsW4UMRCGeyTeYeU+572AInTyOoouoBQgTrpLeuOdvbPitS17srrjGWihoqFAoooEDVDwNoHwGJndVS57CRXdzPyj39+MRxyua5s1EJPxrmDjUc4ycNqXxi0Ldrp4sfeMZQmVK5X1Dgq2gcQO5eNHYhZ9gIgGUkYWLhVshRgmnCe9glqlEcmOlMrHWiGlccl9VRkNx15f1OCQ7+f5AYc1giuh3AtbQ9Y7Thr8X9PS65YvnS02gYClOArBGq2QppSvjI4++Qqz52sNVvChKIhuDvoiGtzIXPBhKuZaWZiSsayUTSD4XUGcgGqXNlMmJikanDSg0ccsmbe0tn2WvVEJWpyCNSoa5ZCw2rY+6WIbEkb5+/vnq18f/376Ijjpfa0Lh63D2DyV466Bgt3G1qDnIGGXcGHQQnpdzVTEfwCPh8AdQ4/b41z9+HD9/t2fbz+vL78+oOzmpvfuvTD1dVBuQ8I2emnceToNC3+sEG53ulsU85WKUNI3bHe+LYgTWme0rcl0pdwSytueh0J7AWf9mcvxwSh/ktPnDmqC3x20vAEAAP//AwBQSwECLQAUAAYACAAAACEAdDZapnoBAACEBQAAEwAAAAAAAAAAAAAAAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQC1VTAj9AAAAEwCAAALAAAAAAAAAAAAAAAAALMDAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQBm2OK1vgMAAB0JAAAPAAAAAAAAAAAAAAAAANgGAAB4bC93b3JrYm9vay54bWxQSwECLQAUAAYACAAAACEAkgeU7AQBAAA/AwAAGgAAAAAAAAAAAAAAAADDCgAAeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHNQSwECLQAUAAYACAAAACEAgrsjVSI9AACw4QEAGAAAAAAAAAAAAAAAAAAHDQAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1sUEsBAi0AFAAGAAgAAAAhABvSBT1hBwAAzSAAABMAAAAAAAAAAAAAAAAAX0oAAHhsL3RoZW1lL3RoZW1lMS54bWxQSwECLQAUAAYACAAAACEAUlyT7IcGAADDOAAADQAAAAAAAAAAAAAAAADxUQAAeGwvc3R5bGVzLnhtbFBLAQItABQABgAIAAAAIQAP7A+nUgIAAHkHAAAUAAAAAAAAAAAAAAAAAKNYAAB4bC9zaGFyZWRTdHJpbmdzLnhtbFBLAQItABQABgAIAAAAIQA7bTJLwQAAAEIBAAAjAAAAAAAAAAAAAAAAACdbAAB4bC93b3Jrc2hlZXRzL19yZWxzL3NoZWV0MS54bWwucmVsc1BLAQItABQABgAIAAAAIQC4hjB66AUAAMx9AAAnAAAAAAAAAAAAAAAAAClcAAB4bC9wcmludGVyU2V0dGluZ3MvcHJpbnRlclNldHRpbmdzMS5iaW5QSwECLQAUAAYACAAAACEAm2fmDd0AAAD8AQAAEAAAAAAAAAAAAAAAAABWYgAAeGwvY2FsY0NoYWluLnhtbFBLAQItABQABgAIAAAAIQCr8nbCSgEAAF8CAAARAAAAAAAAAAAAAAAAAGFjAABkb2NQcm9wcy9jb3JlLnhtbFBLAQItABQABgAIAAAAIQDvkXixqAEAABYDAAAQAAAAAAAAAAAAAAAAAOJlAABkb2NQcm9wcy9hcHAueG1sUEsFBgAAAAANAA0AZAMAAMBoAAAAAA=="""

def write_order_xlsx(path, proj_name, proj_no, order_date, items):
    """直接填入原始訂購單範本（第一份，row 1~25）"""
    import base64, io
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment

    wb = load_workbook(io.BytesIO(base64.b64decode(_ORDER_TEMPLATE_B64)))
    ws = wb['世界油箱']

    FONT = "標楷體"
    FS   = 14

    def fill(row, col, val, align="left", vertical="center", wrap=None):
        cell = ws.cell(row, col, val)
        cell.font      = Font(name=FONT, size=FS)
        cell.alignment = Alignment(horizontal=align, vertical=vertical, wrap_text=wrap)

    # ── 工程資訊（Row 3）
    fill(3, 1,
        f"工程名稱:{proj_name}                       "
        f"工程編號:{proj_no}                                        "
        f"訂購日期:{order_date}")

    # ── 資料行（Row 6~24，固定19行）
    from collections import OrderedDict
    merged = OrderedDict()
    for it in items:
        key = (it["mat"], it["thick"], it["width"], it["length"],
               it.get("note",""))
        if key not in merged:
            merged[key] = 0
        merged[key] += it["qty"]

    MAX_ROWS = 19   # row 6~24
    notes = []
    for i, ((mat, thick, width, length, note), qty) in enumerate(merged.items()):
        if i >= MAX_ROWS:
            break
        r = 6 + i
        fill(r,  1, i+1,   align="center")  # 項次直接填數字
        fill(r,  3, mat,    align="center")
        fill(r,  4, "PL")
        fill(r,  5, thick)
        fill(r,  6, "X",    align="center")
        fill(r,  7, width)
        fill(r,  8, "X",    align="center")
        fill(r,  9, length)
        fill(r, 14, qty)
        fill(r, 15, "片",   align="center")
        if note:
            notes.append(f"項次{i+1}：{note}")

    # 範本的 R 欄是合併的「備註」區塊，逐列寫入會落在合併儲存格內（唯讀而出錯）；
    # 與網頁版相同，集中寫入 R8 備註欄
    if notes:
        fill(8, 18, "備註:\n" + "\n".join(notes), vertical="top", wrap=True)

    wb.save(path)


class OrderSelectWindow(tk.Toplevel):
    """訂購單選擇視窗：勾選要加入訂購單的採購項目"""

    def __init__(self, parent, result):
        super().__init__(parent)
        self.parent = parent
        self.result = result
        self.title("信暐訂購單 - 選擇項目")
        self.geometry("780x560")
        self.configure(bg=CLR_BG)
        self.resizable(True, True)
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=CLR_HEADER, height=44)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="信暐訂購單 ── 勾選要加入的採購項目",
                 bg=CLR_HEADER, fg="white",
                 font=("Microsoft JhengHei", 11, "bold")).pack(side="left", padx=12, pady=8)

        # 工程資訊輸入
        info_frame = tk.LabelFrame(self, text="訂購資訊", bg=CLR_BG,
                                   font=("Microsoft JhengHei", 9, "bold"))
        info_frame.pack(fill="x", padx=10, pady=(6,2))

        p = self.result["params"]
        fields = [
            ("工程名稱：", p.get("proj_name","")),
            ("工程編號：", p.get("proj_no","")),
            ("訂購日期：", datetime.now().strftime("%Y/%m/%d")),
        ]
        self.info_vars = {}
        for i, (label, default) in enumerate(fields):
            tk.Label(info_frame, text=label, bg=CLR_BG,
                     font=("Microsoft JhengHei", 9)).grid(row=0, column=i*2, sticky="e", padx=(8,2), pady=4)
            var = tk.StringVar(value=default)
            self.info_vars[label] = var
            tk.Entry(info_frame, textvariable=var, width=18,
                     font=("Microsoft JhengHei", 9)).grid(row=0, column=i*2+1, padx=(0,8), pady=4)

        # 全選/全不選
        ctrl = tk.Frame(self, bg=CLR_BG)
        ctrl.pack(fill="x", padx=10, pady=(2,0))
        tk.Button(ctrl, text="全選", command=self._select_all,
                  bg="#4A6C7A", fg="white", font=("Microsoft JhengHei", 9),
                  relief="flat", padx=8, pady=2, cursor="hand2").pack(side="left")
        tk.Button(ctrl, text="全不選", command=self._deselect_all,
                  bg="#718096", fg="white", font=("Microsoft JhengHei", 9),
                  relief="flat", padx=8, pady=2, cursor="hand2").pack(side="left", padx=(4,0))

        # 項目列表
        frame = tk.Frame(self, bg=CLR_BG)
        frame.pack(fill="both", expand=True, padx=10, pady=4)

        cols = ("選取","材質","板厚(mm)","板寬(mm)","板長(mm)","數量","單片重(kg)","總重(kg)","備註")
        vsb  = ttk.Scrollbar(frame, orient="vertical")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings",
                                 yscrollcommand=vsb.set, height=12)
        vsb.config(command=self.tree.yview)
        for col, w in zip(cols, [55,80,80,80,90,60,100,100,160]):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="center")
        self.tree.tag_configure("odd",  background=CLR_ROW_ODD)
        self.tree.tag_configure("even", background=CLR_ROW_EVEN)
        self.tree.tag_configure("sel",  background="#C6EFCE")
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<ButtonRelease-1>", self._on_click)
        self.tree.bind("<Double-1>",        self._on_dbl_note)

        # 建立資料（合計相同規格）
        import re
        from collections import OrderedDict
        density = self.result["params"]["density"]
        merged  = OrderedDict()
        for p in self.result["purchase_list"]:
            m = re.match(r"PL(\d+)×(\d+)×(\d+)", p["spec"])
            if not m: continue
            thick,width,length = int(m.group(1)),int(m.group(2)),int(m.group(3))
            key = (p["mat"], thick, width, length)
            if key not in merged:
                merged[key] = {"qty":0,"mat":p["mat"],"thick":thick,"width":width,"length":length}
            merged[key]["qty"] += p["qty"]

        self.items  = []
        self.checks = {}
        for i, ((mat,thick,width,length), val) in enumerate(merged.items()):
            wt  = round(calc_weight(width, thick, length, density), 0)
            tot = int(wt * val["qty"])
            iid = str(i)
            self.tree.insert("", "end", iid=iid, tags=("sel",),
                             values=("☑", mat, thick, width, length,
                                     val["qty"], int(wt), tot, ""))
            self.items.append({"mat":mat,"thick":thick,"width":width,
                                "length":length,"qty":val["qty"],"note":""})
            self.checks[iid] = True

        # 底部按鈕
        btn_bar = tk.Frame(self, bg=CLR_BG)
        btn_bar.pack(fill="x", padx=10, pady=8)
        tk.Button(btn_bar, text="❌ 取消", command=self.destroy,
                  bg="#718096", fg="white", font=("Microsoft JhengHei", 10, "bold"),
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(side="left")
        tk.Button(btn_bar, text="👁 預覽訂購單",
                  command=self._preview,
                  bg="#4A6C7A", fg="white", font=("Microsoft JhengHei", 10, "bold"),
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(side="left", padx=(8,0))
        tk.Button(btn_bar, text="📥 信暐訂購單 Excel",
                  command=self._export,
                  bg="#4A6741", fg="white", font=("Microsoft JhengHei", 10, "bold"),
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(side="right")

    def _on_dbl_note(self, event):
        col = self.tree.identify_column(event.x)
        if col != "#9": return   # 備註欄
        iid = self.tree.identify_row(event.y)
        if not iid or not iid.isdigit(): return
        idx = int(iid)
        cur = self.items[idx].get("note","")
        val = tk.simpledialog.askstring(
            "編輯備註", "請輸入備註：",
            initialvalue=cur, parent=self)
        if val is None: return
        self.items[idx]["note"] = val
        vals = list(self.tree.item(iid, "values"))
        vals[8] = val
        self.tree.item(iid, values=vals)

    def _on_click(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid: return
        col = self.tree.identify_column(event.x)
        if col != "#1": return   # 只有第一欄切換
        self.checks[iid] = not self.checks.get(iid, True)
        tag  = "sel" if self.checks[iid] else ("odd" if int(iid)%2==0 else "even")
        mark = "☑" if self.checks[iid] else "☐"
        vals = list(self.tree.item(iid, "values"))
        vals[0] = mark
        self.tree.item(iid, values=vals, tags=(tag,))

    def _select_all(self):
        for iid in self.checks:
            self.checks[iid] = True
            vals = list(self.tree.item(iid, "values"))
            vals[0] = "☑"
            self.tree.item(iid, values=vals, tags=("sel",))

    def _deselect_all(self):
        for iid in self.checks:
            self.checks[iid] = False
            vals = list(self.tree.item(iid, "values"))
            vals[0] = "☐"
            tag = "odd" if int(iid)%2==0 else "even"
            self.tree.item(iid, values=vals, tags=(tag,))

    def _get_selected(self):
        selected = [self.items[int(iid)] for iid, checked in self.checks.items() if checked]
        return selected

    def _preview(self):
        selected = self._get_selected()
        if not selected:
            messagebox.showwarning("提示", "請至少勾選一項。", parent=self)
            return
        import tempfile, os
        proj_name  = self.info_vars["工程名稱："].get()
        proj_no    = self.info_vars["工程編號："].get()
        order_date = self.info_vars["訂購日期："].get()
        try:
            tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
            tmp.close()
            write_order_xlsx(tmp.name, proj_name, proj_no, order_date, selected)
            os.startfile(tmp.name)
        except Exception as e:
            messagebox.showerror("預覽失敗", str(e), parent=self)

    def _export(self):
        selected = self._get_selected()
        if not selected:
            messagebox.showwarning("提示", "請至少勾選一項。", parent=self)
            return
        proj_name  = self.info_vars["工程名稱："].get()
        proj_no    = self.info_vars["工程編號："].get()
        order_date = self.info_vars["訂購日期："].get()
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel","*.xlsx")],
            initialfile=f"訂購單_{proj_no}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            parent=self)
        if not path: return
        try:
            write_order_xlsx(path, proj_name, proj_no, order_date, selected)
            messagebox.showinfo("完成", f"訂購單已儲存：\n{path}", parent=self)
            self.destroy()
        except Exception as e:
            messagebox.showerror("失敗", str(e), parent=self)


# ═══════════════════════════════════════════════════════════════════════
import tkinter.simpledialog

if __name__ == "__main__":
    app = BHPeilianApp()
    app.mainloop()
