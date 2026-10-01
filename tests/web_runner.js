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
  grab("function mergePurchase"), grab("function buildPeRows"),
  "const peWt = (r, row) => pyRound(calcWeight(row.width, row.thick, row.length, r.params.density), 0);",
  "const peSpec = row => `PL${row.thick}×${row.width}×${row.length}`;",
  line("const peBoardKey"), line("const peRowKey"), line("const ctBoardKey"),
  grab("function checkSize"), grab("function makeModifiedResult"),
].join("\n");
const T = (0, eval)(core + "\n" + ui + "\n({planPurchase, decomposeBH, buildPeRows, makeModifiedResult, checkSize});");

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
  (c.edits || []).forEach(e => { const row = rows.find(x => x.orig_spec === e.spec && x.mat === e.mat); Object.assign(row, e.set); });
  const mod = T.makeModifiedResult(res, rows);
  return {orig: dump(res), mod: dump(mod), check: rows.map(row => T.checkSize(res, row))};
});
process.stdout.write(JSON.stringify(out));
