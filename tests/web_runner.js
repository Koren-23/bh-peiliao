// 以 Node 執行網頁版（BH配料系統.html）的計算核心，供 tests/test_bh.py 與桌面版比對。
// 用法：node tests/web_runner.js <cases.json>   → 將結果 JSON 輸出到 stdout
"use strict";
const fs = require("fs");
const path = require("path");

const html = fs.readFileSync(path.join(__dirname, "..", "BH配料系統.html"), "utf8");
const core = html.match(/<script id="core-src">([\s\S]*?)<\/script>/)[1];
const lines = html.split("\n");
const line = start => { const l = lines.find(x => x.startsWith(start)); if (!l) throw new Error("找不到：" + start); return l; };
const grab = start => {
  const i = html.indexOf(start); if (i < 0) throw new Error("找不到：" + start);
  let d = 0, j = html.indexOf("{", i);
  for (; j < html.length; j++) { if (html[j] === "{") d++; else if (html[j] === "}" && --d === 0) break; }
  return html.slice(i, j + 1);
};
const ui = [
  line("const specRe"), line("const layoutRowUsed"), line("const layoutRowLeft"),
  grab("function mergePurchase"), grab("function buildPeRows"), grab("function peSplitRow"), grab("function peMergeRow"),
  "const peWt = (r, row) => pyRound(calcWeight(row.width, row.thick, row.length, r.params.density), 0);",
  "const peSpec = row => `PL${row.thick}×${row.width}×${row.length}`;",
  line("const peBoardKey"), line("const peRowKey"), line("const ctBoardKey"), grab("const peRowCts") + ";",
  grab("function checkSize"), grab("function makeModifiedResult"),
  "const LOSS_CATS = " + html.slice(html.indexOf("const LOSS_CATS = [") + "const LOSS_CATS = ".length, html.indexOf("];", html.indexOf("const LOSS_CATS = [")) + 1) + ";",
  grab("function lossAnalysis"),
].join("\n");
const T = (0, eval)(core + "\n" + ui + "\n({planPurchase, decomposeBH, buildPeRows, peSplitRow, makeModifiedResult, checkSize, lossAnalysis});");
const lossDump = r => { const a = T.lossAnalysis(r); return {
  total: Object.fromEntries(Object.entries(a.total).map(([k, v]) => [k, Math.round(v * 10) / 10])),
  reasons: a.rows.map(x => x.reasons)}; };

const dump = r => ({
  board: r.cut_details.map(c => c.board_spec), left: r.cut_details.map(c => c.leftover),
  type: r.cut_details.map(c => c.type), scraps: r.new_scraps.map(s => s.spec),
  buy: r.purchase_list.map(p => p.spec),
});

const cases = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const out = cases.map(c => {
  const parts = [];
  c.bh.forEach(r => { const d = T.decomposeBH(r, c.P.density); parts.push(d.flange, d.web); });
  const res = T.planPurchase(parts, c.P, JSON.parse(JSON.stringify(c.scraps)), c.it);
  res.params = c.P;
  const rows = T.buildPeRows(res);
  (c.edits || []).forEach(e => {
    const i = rows.findIndex(x => x.orig_spec === e.spec && x.mat === e.mat);
    if (e.split) T.peSplitRow(rows, i, e.split);   // 拆分：第 i 列保留前 split 片
    Object.assign(rows[i], e.set);
  });
  const mod = T.makeModifiedResult(res, rows);
  return {orig: dump(res), mod: dump(mod), check: rows.map(row => T.checkSize(res, row)),
          loss: [lossDump(res), lossDump(mod)],
          rows: rows.map(row => [row.orig_spec, row.mat, row.width, row.length, row.qty, row.cts])};
});
process.stdout.write(JSON.stringify(out));
