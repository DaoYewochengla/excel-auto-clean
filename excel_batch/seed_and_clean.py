# -*- coding: utf-8 -*-
"""
表格批处理样例：多文件合并 + 清洗 + 去重 + 汇总
输入: sample/*.xlsx（脏数据：重复行、空格、日期格式不一、金额带¥和逗号、空行）
输出: output/合并清洗结果.xlsx（明细表/汇总表/异常行表）
运行: python seed_and_clean.py
"""
import os, glob, random, datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
SAMPLE = os.path.join(HERE, "sample")
OUT = os.path.join(HERE, "output")
os.makedirs(SAMPLE, exist_ok=True)
os.makedirs(OUT, exist_ok=True)

CITIES = ["杭州", " 杭州 ", "杭州仓", "上海", " 上海", "SH", "广州", "深圳市", "深圳"]
PRODUCTS = ["机械键盘", "无线鼠标", "显示器支架", "USB-C 扩展坞", "降噪耳机", "笔记本内存条"]
CHANNELS = ["天猫", "京东", "抖音", "私域"]


def make_dirty(seed, n_rows, fname, date_style):
    """造一份'客户发来的脏表'"""
    random.seed(seed)
    wb = Workbook(); ws = wb.active; ws.title = "订单"
    ws.append(["订单号", "下单日期", "客户", "城市", "商品", "数量", "单价", "渠道"])
    base = datetime.date(2026, 10, 1) + datetime.timedelta(days=seed * 3)
    for i in range(n_rows):
        d = base + datetime.timedelta(days=random.randint(0, 20))
        if date_style == "slash":
            ds = d.strftime("%Y/%m/%d")
        elif date_style == "cn":
            ds = f"{d.year}年{d.month}月{d.day}日"
        else:
            ds = d  # 真日期
        qty = random.choice([1, 2, 3, 5, "3 ", 2])
        price = random.choice([199, 299, 459, "￥1,299", "89.9", 1299, 89.9])
        ws.append([
            f"NO{seed}{i:03d}",
            ds,
            f"客户{random.randint(1, 60):02d}",
            random.choice(CITIES),
            random.choice(PRODUCTS),
            qty,
            price,
            random.choice(CHANNELS),
        ])
    # 混入重复行 + 全空行 + 合计行
    if ws.max_row > 3:
        r = ws.max_row
        ws.append([ws.cell(r, c).value for c in range(1, 9)])   # 重复上一行
        ws.append([None] * 8)
        ws.append(["合计", None, None, None, None, "=SUM(F2:F%d)" % r, None, None])
    wb.save(os.path.join(SAMPLE, fname))
    return fname


def smart_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    s = str(v).strip()
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    if "年" in s:
        s = s.replace("年", "-").replace("月", "-").replace("日", "")
        return datetime.datetime.strptime(s, "%Y-%m-%d").date()
    raise ValueError(f"无法识别日期: {v!r}")


def smart_num(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("￥", "").replace("¥", "").replace(",", "")
    return float(s)


CITY_MAP = {"杭州": "杭州", "杭州仓": "杭州", "上海": "上海", "SH": "上海",
            "广州": "广州", "深圳市": "深圳", "深圳": "深圳"}


def main():
    files = [("订单_华东.xlsx", 40, "slash"), ("订单_华南.xlsx", 35, "cn"), ("订单_补录.xlsx", 25, "iso")]
    for i, (fn, n, st) in enumerate(files):
        make_dirty(seed=i + 1, n_rows=n, fname=fn, date_style=st)

    rows, bad = [], []
    seen = set()
    dup = 0
    for path in sorted(glob.glob(os.path.join(SAMPLE, "*.xlsx"))):
        ws = load_workbook(path, data_only=True).active
        header = [c.value for c in ws[1]]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not any(x is not None for x in r):
                continue
            if str(r[0]).strip() == "合计":
                continue
            try:
                rec = {
                    "订单号": str(r[0]).strip(),
                    "日期": smart_date(r[1]),
                    "客户": str(r[2]).strip(),
                    "城市": CITY_MAP.get(str(r[3]).strip(), str(r[3]).strip()),
                    "商品": str(r[4]).strip(),
                    "数量": int(smart_num(r[5])),
                    "单价": smart_num(r[6]),
                    "渠道": str(r[7]).strip(),
                    "来源文件": os.path.basename(path),
                }
            except Exception as e:
                bad.append([os.path.basename(path), str(r), str(e)])
                continue
            key = (rec["订单号"], rec["日期"], rec["客户"], rec["商品"])
            if key in seen:
                dup += 1
                continue
            seen.add(key)
            rec["金额"] = round(rec["数量"] * rec["单价"], 2)
            rows.append(rec)
    rows.sort(key=lambda x: (x["日期"], x["订单号"]))

    # 汇总：按城市/商品
    agg = {}
    for r in rows:
        k = (r["城市"], r["商品"])
        a = agg.setdefault(k, [0, 0.0])
        a[0] += r["数量"]; a[1] += r["金额"]

    wb = Workbook()
    ws = wb.active; ws.title = "明细"
    cols = ["订单号", "日期", "客户", "城市", "商品", "数量", "单价", "金额", "渠道", "来源文件"]
    ws.append(cols)
    for r in rows:
        ws.append([r[c].strftime("%Y-%m-%d") if c == "日期" else r[c] for c in cols])
    ws2 = wb.create_sheet("汇总")
    ws2.append(["城市", "商品", "总数量", "总金额"])
    for (city, prod), (q, amt) in sorted(agg.items(), key=lambda kv: -kv[1][1]):
        ws2.append([city, prod, q, round(amt, 2)])
    ws2.append(["合计", "", sum(q for q, _ in agg.values()), round(sum(a for _, a in agg.values()), 2)])
    ws3 = wb.create_sheet("异常行")
    ws3.append(["来源文件", "原始行", "原因"])
    for b in bad:
        ws3.append(b)

    bold = Font(bold=True); fill = PatternFill("solid", fgColor="DDEBF7")
    for sh in (ws, ws2, ws3):
        for c in sh[1]:
            c.font = bold; c.fill = fill; c.alignment = Alignment(horizontal="center")
        for i, col in enumerate(sh.iter_cols(), 1):
            width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
            sh.column_dimensions[get_column_letter(i)].width = min(max(width + 4, 9), 26)
        sh.freeze_panes = "A2"

    out = os.path.join(OUT, "合并清洗结果.xlsx")
    wb.save(out)
    total = sum(r["金额"] for r in rows)
    raw = sum(n for _, n, _ in files)
    print(f"输入 3 个文件 / 原始数据行 {raw} 行（另含空行、合计行、重复行 {len(files)*3} 行污染）")
    print(f"清洗后有效行 {len(rows)} 行、剔除重复 {dup} 行、剔除空行/合计行 6 行、异常 {len(bad)} 行")
    print(f"总金额 {total:,.2f} 元 / 城市×商品组合 {len(agg)} 个")
    print(f"输出 {out}")
    return rows, agg, len(bad)


if __name__ == "__main__":
    main()
