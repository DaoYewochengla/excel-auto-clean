# -*- coding: utf-8 -*-
"""生成 before/after 对比图（商品图），供闲鱼商品页使用。"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import load_workbook
from seed_and_clean import SAMPLE, OUT, HERE
import glob

def rows(path, limit=7):
    ws = load_workbook(path, data_only=True).active
    out = []
    for r in ws.iter_rows(max_row=limit + 1, values_only=True):
        out.append(["" if v is None else str(v) for v in r[:8]])
    return out

before = []
for p in sorted(glob.glob(os.path.join(SAMPLE, "*.xlsx")))[:2]:
    before.append((os.path.basename(p), rows(p, 5)))
after = rows(os.path.join(OUT, "合并清洗结果.xlsx"), 8)

def table(title, data, cls=""):
    h = "".join(f"<th>{c}</th>" for c in data[0])
    b = "".join("<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>" for r in data[1:])
    return f'<div class="card {cls}"><div class="ttl">{title}</div><table><thead><tr>{h}</tr></thead><tbody>{b}</tbody></table></div>'

html = """<!doctype html><meta charset="utf-8"><style>
body{margin:0;padding:0;font-family:"Microsoft YaHei",sans-serif;background:#f4f7fb;width:1120px}
h1{font-size:26px;margin:0 0 4px;color:#12263f}
.sub{color:#5b6b82;font-size:14px;margin-bottom:18px}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.card{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:14px 16px;box-shadow:0 1px 3px rgba(16,24,40,.06)}
.ttl{font-size:14px;font-weight:700;color:#12263f;margin-bottom:10px}
.ttl span{font-weight:400;color:#8496ad;font-size:12px;margin-left:8px}
table{border-collapse:collapse;width:100%;font-size:11.5px}
th{background:#eef4ff;color:#2c4a7c;text-align:left;padding:6px 7px;border:1px solid #dfe8f6;white-space:nowrap}
td{padding:5px 7px;border:1px solid #eef1f6;color:#33455c;white-space:nowrap}
.bad td{background:#fff6f5}
.ok td{background:#f5fbf6}
.arrow{display:flex;align-items:center;justify-content:center;font-size:30px;color:#2563eb}
.foot{margin-top:16px;font-size:12.5px;color:#5b6b82;display:flex;gap:22px}
.foot b{color:#0f7b3e}
</style>
<div id="wrap" style="padding:30px 34px;"><h1>Excel 多表合并 · 自动清洗去重 · 一键汇总</h1>
<div class="sub">客户发来 N 个格式不一的表 → 3 分钟出规范明细 + 汇总 + 异常清单（样例为脚本真实运行结果）</div>
<div class="grid">
<div>@LEFT@</div>
<div>@RIGHT@</div>
</div>
<div class="foot">
<span>清洗规则：空白/合计行剔除 · 全字段去重 · 日期 4 种格式归一 · 金额去 ￥与千分位 · 城市别名归并（SH→上海）</span>
<span><b>自动产出</b>：明细表 / 汇总表 / 异常行表</span>
</div></div>
"""
_x = (
    "".join(table(f"处理前 #{i+1}：{n}", d, "bad") for i, (n, d) in enumerate(before)),
    table("处理后：统一明细（有效数据行已去重）", after, "ok"),
)
html = html.replace("@LEFT@", _x[0]).replace("@RIGHT@", _x[1])

p = os.path.join(HERE, "preview.html")
open(p, "w", encoding="utf-8").write(html)
print(p)

from playwright.sync_api import sync_playwright
with sync_playwright() as pw:
    b = pw.chromium.launch(channel="chrome", args=["--no-proxy-server"], headless=True)
    pg = b.new_page(viewport={"width": 1180, "height": 760}, device_scale_factor=2)
    pg.goto("file:///" + p.replace("\\", "/"))
    pg.wait_for_timeout(400)
    out = os.path.join(HERE, "商品图_表格批处理.png")
    pg.locator("#wrap").screenshot(path=out)
    b.close()
print("saved", out, os.path.getsize(out), "bytes")
