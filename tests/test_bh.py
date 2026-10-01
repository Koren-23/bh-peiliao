"""
BH配料系統 自動化測試

    python -m unittest discover -s tests -v

- 回歸測試：曾經發生過的 bug，確保不再出現
- 一致性：切割明細餘料文字 = 餘料清單，各排餘料與排列圖算法相同
- 網頁版 / 桌面版比對：以 Node 執行 BH配料系統.html 的計算核心，
  與 BH_GUI.py 跑相同的隨機案例，結果必須完全相同（找不到 Node 時略過）
"""
import copy
import importlib.util
import json
import os
import random
import re
import shutil
import subprocess
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_spec = importlib.util.spec_from_file_location("bh_gui", os.path.join(ROOT, "BH_GUI.py"))
bh = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(bh)
PE = bh.PurchaseEditWindow
SPEC_RE = re.compile(r"PL(\d+)×(\d+)×(\d+)")

BASE_P = {"density": 7.85, "new_kerf": 5, "new_trim": 5, "scrap_kerf": 5, "scrap_trim": 5,
          "bw_min": 1524, "bw_max": 2499, "bl_min": 8000, "bl_max": 18499,
          "w_min": 2600, "w_max": 12500, "cut_mode": "multi"}


def plan(P, bh_rows, scraps=(), it=8):
    parts = []
    for r in bh_rows:
        d = bh.decompose_bh(r, P["density"], P["new_kerf"], P["new_trim"], P["scrap_kerf"], P["scrap_trim"], [])
        parts += [d["flange"], d["web"]]
    pl, cds, ns, used = bh.plan_purchase(
        parts, P["density"], P["new_kerf"], P["new_trim"], P["scrap_kerf"], P["scrap_trim"],
        copy.deepcopy(list(scraps)), iterations=it,
        bw_min=P["bw_min"], bw_max=P["bw_max"], bl_min=P["bl_min"], bl_max=P["bl_max"],
        w_min=P["w_min"], w_max=P["w_max"], cut_mode=P["cut_mode"])
    return {"params": P, "purchase_list": pl, "cut_details": cds, "new_scraps": ns}


def pe_rows(res):
    """與網頁版 buildPeRows 相同：依 規格 + 材質 + 單重 合併採購清單"""
    merged = {}
    for p in res["purchase_list"]:
        k = (p["spec"], p["mat"], p["unit_wt"])
        merged[k] = merged.get(k, 0) + p["qty"]
    rows = []
    for (s, mat, _), q in merged.items():
        t, w, l = map(int, SPEC_RE.match(s).groups())
        rows.append({"thick": t, "width": w, "length": l, "orig_width": w, "orig_length": l,
                     "qty": q, "mat": mat, "orig_spec": s, "note": ""})
    return rows


def pe(res, rows):
    """不開視窗，直接使用採購修正視窗的計算方法"""
    w = PE.__new__(PE)
    w.result, w.rows = res, rows
    return w


def bh_row(comp, spec, length, qty, mat="SN490B"):
    return {"comp": comp, "part": "", "spec": spec, "length": length, "qty": qty, "mat": mat}


def leftover_specs(left):
    return bh.leftover_specs(left)


def random_cases(n, seed=7):
    rnd = random.Random(seed)
    specs = ["BH400×200×8×13", "BH500×250×9×16", "BH600×300×12×20", "BH350×175×7×11", "BH800×300×14×26"]
    cases = []
    for i in range(n):
        P = dict(BASE_P, new_kerf=rnd.choice([1, 3, 5]), new_trim=rnd.choice([0, 1, 5, 10]),
                 scrap_kerf=rnd.choice([1, 5]), scrap_trim=rnd.choice([0, 5]),
                 cut_mode=rnd.choice(["single", "multi"]))
        mats = ["SN490B"] if rnd.random() < .6 else ["SN490B", "SS400"]
        rows = [bh_row(rnd.choice(["C", "FB", "G"]) + str(j), rnd.choice(specs), rnd.randrange(1500, 15000, 10),
                       rnd.randint(1, 6), rnd.choice(mats)) for j in range(rnd.randint(2, 7))]
        scraps = [{"id": f"S{j}", "thick": rnd.choice([8, 9, 12, 13, 16, 20]), "width": rnd.randrange(150, 2000, 10),
                   "length": rnd.randrange(2000, 12000, 10), "qty": rnd.randint(1, 2), "mat": rnd.choice(mats)}
                  for j in range(rnd.randint(0, 4))]
        cases.append({"P": P, "bh": rows, "scraps": scraps, "it": 6})
    return cases


class TestConsistency(unittest.TestCase):
    """切割明細、餘料清單、排列圖三者一致"""

    def test_random_cases(self):
        for n, c in enumerate(random_cases(40)):
            res = plan(c["P"], c["bh"], c["scraps"], c["it"])
            for ct in res["cut_details"]:
                with self.subTest(case=n, idx=ct["idx"]):
                    L = ct["layout"]
                    self.assertIsNotNone(L, "每個片次都要有排列資料")
                    objs = [s["spec"] for s in res["new_scraps"] if s["src"] == ct["idx"]]
                    self.assertEqual(leftover_specs(ct["leftover"]), objs, "切割明細餘料 ≠ 餘料清單")
                    for rw in L["rows"]:
                        used = bh.layout_row_used(L, rw)
                        self.assertLessEqual(used + L["trim"], L["board_l"], "排超出板長")
                        self.assertEqual(bh.layout_row_left(L, rw), max(0, L["board_l"] - used - L["trim"] - L["kerf"]))

    def test_unchanged_purchase_edit_equals_original(self):
        """採購修正未改尺寸時，結果與原始配料完全相同（原 bug：單段切割餘料差一個鋸縫）"""
        for n, c in enumerate(random_cases(40, seed=11)):
            res = plan(c["P"], c["bh"], c["scraps"], c["it"])
            mod = pe(res, pe_rows(res))._make_modified_result()
            with self.subTest(case=n, mode=c["P"]["cut_mode"]):
                self.assertEqual([x["leftover"] for x in mod["cut_details"]], [x["leftover"] for x in res["cut_details"]])
                self.assertEqual([x["spec"] for x in mod["new_scraps"]], [x["spec"] for x in res["new_scraps"]])
                self.assertEqual(mod["modified_specs"], set())


class TestRegressions(unittest.TestCase):

    def test_single_mode_padded_length_is_scrap(self):
        """單段切割：板長補到最小值時，多出的長度要成為餘料（原 bug：直接消失）"""
        P = dict(BASE_P, cut_mode="single")
        res = plan(P, [bh_row("A", "BH400×200×8×13", 5000, 2)], it=1)
        for ct in res["cut_details"]:
            L = ct["layout"]
            self.assertEqual(L["board_l"], 8000)
            self.assertTrue(any(f"×{8000 - 5 - 5000 - 5 - 5}" in s for s in leftover_specs(ct["leftover"])),
                            f"{ct['board_spec']} 缺少長度方向餘料：{ct['leftover']}")

    def test_single_mode_shorter_column_is_scrap(self):
        """單段切割：同一張板較短的排，尾端也要成為餘料"""
        P = dict(BASE_P, cut_mode="single")
        res = plan(P, [bh_row("A", "BH400×200×8×13", 12000, 1), bh_row("B", "BH400×200×8×13", 9000, 1)], it=1)
        web = [ct for ct in res["cut_details"] if ct["type"] == "W"][0]
        self.assertIn("PL8×374×2995", leftover_specs(web["leftover"]))   # 12010 - (5 + 9000) - 5 - 5

    def test_scrap_board_kerf_and_length(self):
        """現有餘料：寬度餘料扣鋸縫、長度方向剩料要記錄（原 bug：沒扣鋸縫、長度剩料消失）"""
        P = dict(BASE_P, cut_mode="single")
        scraps = [{"id": "S1", "thick": 8, "width": 1000, "length": 6000, "qty": 1, "mat": "SN490B"}]
        res = plan(P, [bh_row("A", "BH400×200×8×13", 4000, 1)], scraps, it=1)
        ct = [c for c in res["cut_details"] if c["is_scrap"]][0]
        # 寬：1000 - (374 + 5×2) - 5 = 611；長：6000 - (5 + 4000) - 5 - 5 = 1985
        self.assertEqual(leftover_specs(ct["leftover"]), ["PL8×611×6000", "PL8×374×1985"])

    def test_narrow_scrap_not_used(self):
        """比零件窄的現有餘料不可使用（原 bug：至少排 1 排）"""
        scraps = [{"id": "S1", "thick": 8, "width": 300, "length": 9000, "qty": 1, "mat": "SN490B"}]
        res = plan(BASE_P, [bh_row("A", "BH400×200×8×13", 4000, 1)], scraps, it=1)
        self.assertFalse(any(c["is_scrap"] for c in res["cut_details"]))

    def test_material_rows_independent(self):
        """同規格不同材質：只改其中一個材質，另一個不受影響（原 bug：修改被蓋掉）"""
        P = dict(BASE_P, cut_mode="single")
        res = plan(P, [bh_row("X", "BH400×200×8×13", 9000, 4, "SN490B"),
                       bh_row("Y", "BH400×200×8×13", 9000, 4, "SS400")], it=1)
        rows = pe_rows(res)
        tgt = next(r for r in rows if r["mat"] == "SN490B" and r["thick"] == 8)
        tgt["length"] += 500
        mod = pe(res, rows)._make_modified_result()
        boards = {(c["mat"], c["board_spec"]) for c in mod["cut_details"] if c["board_spec"].startswith("PL8×")}
        self.assertEqual(boards, {("SN490B", f"PL8×{tgt['width']}×{tgt['length']}"),
                                  ("SS400", f"PL8×{tgt['width']}×{tgt['orig_length']}")})
        self.assertEqual(mod["modified_specs"], {(8, tgt["width"], tgt["length"], "SN490B")})

    def test_check_size_only_same_board(self):
        """尺寸檢查只比對使用該規格的片次（原 bug：拿同寬的長板來比，誤判板長不足）"""
        P = dict(BASE_P, cut_mode="single")
        res = plan(P, [bh_row("A", "BH400×200×8×13", 12000, 2), bh_row("B", "BH400×200×8×13", 3000, 2)], it=1)
        rows = pe_rows(res)
        for row in rows:
            row["width"] -= 1   # 只縮 1mm，必定仍放得下（板寬至少多出最小板寬限制或一刀鋸縫）
            ok, msg = pe(res, rows)._check_size(row)
            if not ok:
                self.assertNotIn("板長不足", msg, f"{row['orig_spec']}：{msg}")

    def test_check_size_detects_short_row(self):
        """板長縮到比一整排還短時要判定不足"""
        res = plan(BASE_P, [bh_row("A", "BH400×200×8×13", 4000, 4)], it=1)
        row = pe_rows(res)[0]
        row["length"] = 4005   # 一排需要 5 + 4000 + 5 = 4010
        ok, msg = pe(res, [row])._check_size(row)
        self.assertFalse(ok)
        self.assertIn("板長不足", msg)

    def test_flag_keeps_star(self):
        """過輕補寬後又過重時，標記為 *H（原 bug：H 蓋掉 *）"""
        # 翼板 PL13×?×9010：補寬到 2000kg 時板寬進位到 10mm，重量會略超過 2001kg
        P = dict(BASE_P, cut_mode="single", w_min=2000, w_max=2001)
        res = plan(P, [bh_row("A", "BH400×200×8×13", 9000, 1)], it=1)
        flags = [c["idx"][1:] for c in res["cut_details"]]
        self.assertTrue(any(f == "*H" for f in flags), flags)

    def test_type_by_suffix(self):
        """板別以名稱結尾 -F / -W 判斷（原 bug：構件編號含 F 時腹板被當成翼板）"""
        res = plan(dict(BASE_P, cut_mode="single"), [bh_row("FB1", "BH400×200×8×13", 9000, 1)], it=1)
        types = {c["part_spec"].split("×")[0]: c["type"] for c in res["cut_details"]}
        self.assertEqual(types, {"PL13": "F", "PL8": "W"})


NODE = shutil.which("node") or next((p for p in [r"C:\Program Files\nodejs\node.exe"] if os.path.exists(p)), None)


@unittest.skipIf(NODE is None, "找不到 Node.js，略過網頁版比對")
class TestWebDesktopParity(unittest.TestCase):

    def test_parity(self):
        cases = random_cases(60, seed=3)
        # 加入採購修正的修改（改尺寸）一起比對
        rnd = random.Random(5)
        py_out = []
        for c in cases:
            res = plan(c["P"], c["bh"], c["scraps"], c["it"])
            rows = pe_rows(res)
            c["edits"] = []
            for row in rnd.sample(rows, min(2, len(rows))):
                st = {"width": row["width"] + rnd.choice([-50, 0, 100]), "length": row["length"] + rnd.choice([-300, 0, 500])}
                c["edits"].append({"spec": row["orig_spec"], "mat": row["mat"], "set": st})
                row.update(st)
            mod = pe(res, rows)._make_modified_result()
            dump = lambda r: {"board": [x["board_spec"] for x in r["cut_details"]], "left": [x["leftover"] for x in r["cut_details"]],
                              "type": [x["type"] for x in r["cut_details"]], "scraps": [s["spec"] for s in r["new_scraps"]],
                              "buy": [p["spec"] for p in r["purchase_list"]]}
            py_out.append({"orig": dump(res), "mod": dump(mod), "check": [list(pe(res, rows)._check_size(r)) for r in rows]})
        with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf8") as f:
            json.dump(cases, f, ensure_ascii=False)
        try:
            out = subprocess.run([NODE, os.path.join(ROOT, "tests", "web_runner.js"), f.name],
                                 capture_output=True, check=True).stdout.decode("utf8")
        finally:
            os.unlink(f.name)
        js_out = json.loads(out)
        for n, (js, py) in enumerate(zip(js_out, py_out)):
            for k in ("orig", "mod"):
                for fld in py[k]:
                    with self.subTest(case=n, part=f"{k}.{fld}"):
                        self.assertEqual(js[k][fld], py[k][fld])
            with self.subTest(case=n, part="checkSize"):
                self.assertEqual(js["check"], py["check"])


if __name__ == "__main__":
    unittest.main()
