#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""生成一份"客户常见脏数据"并跑一遍清洗，打印前后对比。用于验证与演示。"""
import os
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
RAW, OUT = os.path.join(HERE, "样例_原始脏数据.xlsx"), os.path.join(HERE, "样例_清洗结果.xlsx")

rows = [
    # 客户姓名, 手机号, 订单日期, 金额, 备注
    [" 张 三 ", "13812345678", "2024.3.5", "¥1,280.00", "  加急\u3000"],
    ["李四", "13987654321", "2024年3月6日", "1 980元", ""],
    ["张三 ", "13812345678", "2024/3/5", "1280", "重复单"],
    ["", "", "", "", ""],
    ["王五", "13700001111", "45356", "980.5元", "货到付款"],
    ["赵\u200b六", "13600002222", "2024-03-07", "2,000.00", ""],
    ["孙七", "13500003333", "2024.3.8", "abc", "金额待确认"],
    [None, None, None, None, None],
    ["周八 ", "13400004444", "2024.3.9", "1,000.00", "退款"],
    ["吴九", "13300005555", "2024.3.10", "3,000.00", " "],
]
cols = ["客户姓名", "联系电话 ", "订单日期", "金额(元)", "备注"]
df = pd.DataFrame(rows, columns=cols)
df["空列"] = None
df.to_excel(RAW, index=False)

print("=" * 62)
print("① 清洗前（客户发来的原表）")
print("=" * 62)
print(df.to_string(index=False))
print("\n诊断：", end="")

from xls_clean import diagnose, clean_df, read_any, write_xlsx  # noqa: E402

for k, v in diagnose(df).items():
    if isinstance(v, list):
        print(f"\n  {k}:")
        for i in v:
            print(f"    - {i}")
    else:
        print(f"  {k}: {v}")

stats = {}
src = list(read_any(RAW).values())[0]
clean = clean_df(
    src,
    ["trim", "norm-cols", "empty", "dedupe", "to-date", "to-number", "mask"],
    date_cols=["订单日期"], num_cols=["金额(元)"], dedupe_key=["联系电话"], fill_cols=[], mask_kind="auto",
    stats=stats,
)
print("\n" + "=" * 62)
print("② 清洗后")
print("=" * 62)
print(clean.to_string(index=False))
print("\n处理明细：")
for k, v in stats.items():
    print(f"  · {k}: {v}")
write_xlsx({"清洗结果": clean}, OUT)
print(f"\n结果文件: {OUT}")
