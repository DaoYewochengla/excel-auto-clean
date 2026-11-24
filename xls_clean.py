#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
xls_clean —— Excel/CSV 批量清洗工具（可商用版）

支持操作：
  trim          去首尾空格 / 全角空格 / 不可见字符(零宽、\xa0)
  norm-cols     列名标准化（去空格、全角转半角、去重复后缀）
  empty         删除空行、整列为空白的列
  dedupe        按指定列（或整行）去重，保留首次出现
  to-date       日期列统一为 YYYY-MM-DD（兼容 2024.1.5 / 2024年1月5日 / Excel序列号）
  to-number     数值列清洗：去千分位、货币符号(¥$￥)、百分号、单位后缀
  mask          脱敏：手机号/身份证/邮箱/姓名
  fillna        指定列空值填充（默认空字符串）

子命令：
  clean   清洗单个/多个同构表 -> 输出一个清洗后的 Excel（多 sheet）
  merge   把目录下所有 Excel/CSV 同构表纵向合并成一张总表
  split   按某列的唯一值拆成多个 sheet（或拆成多个文件）
  report  只出诊断报告，不改数据

用法示例：
  python xls_clean.py clean 原始表.xlsx -o 清洗后.xlsx \
      --ops trim,norm-cols,empty,dedupe,to-date,to-number \
      --date-cols 下单日期,发货日期 --number-cols 金额,数量 --dedupe-key 手机号
"""
import argparse
import os
import re
import sys
import unicodedata
from datetime import datetime, timedelta

import pandas as pd

ZERO_WIDTH = dict.fromkeys(map(ord, "\u200b\u200c\u200d\ufeff"), None)
PHONE_RE = re.compile(r"(?<!\d)(1[3-9]\d)\d{4}(\d{4})(?!\d)")
ID_RE = re.compile(r"(?<!\d)(\d{6})\d{8}(\d{3}[\dXx])(?!\d)")
MAIL_RE = re.compile(r"([\w.+-]{2})[\w.+-]*(@[\w-]+\.[\w.]+)")
MONEY_RE = re.compile(r"[¥$￥€,，元块]|USD|CNY|RMB", re.I)


# ---------------------------------------------------------------- 基础工具
def s(v):
    """安全转字符串"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v)


def clean_text(v):
    t = s(v).translate(ZERO_WIDTH).replace("\u3000", " ").replace("\xa0", " ")
    return re.sub(r"[ \t]+", " ", t).strip()


def norm_col(c):
    c = clean_text(c)
    c = unicodedata.normalize("NFKC", c)          # 全角 -> 半角
    c = re.sub(r"[\s\-_/]+", "", c)               # 去空格与分隔符
    return c or "未命名列"


def dedupe_cols(cols):
    """列名去重：A, A -> A, A_2"""
    seen, out = {}, []
    for c in cols:
        if c in seen:
            seen[c] += 1
            out.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 1
            out.append(c)
    return out


def to_date_series(col):
    """把一列尽量转成日期；解析失败保留原值"""
    if pd.api.types.is_datetime64_any_dtype(col):
        return col.dt.strftime("%Y-%m-%d")
    txt = col.map(lambda v: clean_text(v).replace("年", "-").replace("月", "-").replace("日", "")
                  .replace(".", "-").replace("/", "-"))
    # 纯数字且像 Excel 序列号（1~60000）
    num = pd.to_numeric(txt, errors="coerce")
    serial = num.between(1, 60000)
    out = pd.to_datetime(txt.where(~serial), errors="coerce", format="mixed")
    base = pd.Timestamp("1899-12-30")
    out_serial = num.where(serial).map(lambda n: base + timedelta(days=float(n)) if pd.notna(n) else pd.NaT)
    res = out.fillna(out_serial).dt.strftime("%Y-%m-%d")
    return res.where(res.notna(), col.map(s))


def to_number_series(col):
    def one(v):
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return v
        t = clean_text(v)
        if not t:
            return pd.NA
        pct = t.endswith("%")
        t = MONEY_RE.sub("", t).replace("％", "").replace(" ", "")
        t = t.replace("(", "-").replace(")", "")      # 会计负数 (123)
        try:
            n = float(t)
        except ValueError:
            return pd.NA
        return round(n / 100, 6) if pct else n
    return col.map(one)


def apply_mask(v, kind):
    t = s(v)
    if not t:
        return t
    if kind == "phone":
        return PHONE_RE.sub(r"\1****\2", t)
    if kind == "id":
        return ID_RE.sub(r"\1********\2", t)
    if kind == "email":
        return MAIL_RE.sub(r"\1***\2", t)
    if kind == "name":
        return t[0] + "*" * max(1, len(t) - 1) if len(t) > 1 else t
    return t


# ---------------------------------------------------------------- 读 / 写
def read_any(path):
    """读入 -> {sheet名: DataFrame}；CSV 编码自动探测"""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".csv", ".txt"):
        last = None
        for enc in ("utf-8-sig", "gb18030", "utf-8"):
            try:
                return {os.path.splitext(os.path.basename(path))[0]:
                        pd.read_csv(path, dtype=object, encoding=enc)}
            except Exception as e:      # noqa: BLE001
                last = e
        raise last
    return pd.read_excel(path, sheet_name=None, dtype=object)


def write_xlsx(sheets, out):
    os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        for name, df in sheets.items():
            safe = re.sub(r"[\[\]:*?/\\]", "_", str(name))[:31] or "Sheet1"
            df.to_excel(w, sheet_name=safe, index=False)
    return out


# ---------------------------------------------------------------- 清洗主体
def clean_df(df, ops, date_cols, num_cols, dedupe_key, fill_cols, mask_kind, stats):
    df = df.copy()
    if "trim" in ops:
        obj = df.columns
        df[obj] = df[obj].map(clean_text)
        df.columns = list(obj)
        stats["去空格/不可见字符"] = f"{len(df)} 行 × {len(df.columns)} 列"

    if "norm-cols" in ops:
        old = list(df.columns)
        df.columns = dedupe_cols([norm_col(c) for c in old])
        stats["列名标准化"] = " | ".join(f"{a}→{b}" for a, b in zip(old, df.columns) if a != b) or "无需改动"

    if "empty" in ops:
        n0, c0 = len(df), len(df.columns)
        keep = [c for c in df.columns if df[c].map(clean_text).ne("").any()]
        df = df[keep]
        not_empty = df.map(clean_text).ne("").any(axis=1)
        df = df[not_empty]
        stats["删除空行/空列"] = f"行 {n0}→{len(df)}，列 {c0}→{len(df.columns)}"

    if "fillna" in ops and fill_cols:
        for c in fill_cols:
            if c in df.columns:
                df[c] = df[c].map(lambda v: "" if pd.isna(v) else v)
        stats["空值填充"] = ", ".join(fill_cols)

    if "to-date" in ops and date_cols:
        for c in date_cols:
            if c in df.columns:
                df[c] = to_date_series(df[c])
        stats["日期统一"] = ", ".join(c for c in date_cols if c in df.columns)

    if "to-number" in ops and num_cols:
        for c in num_cols:
            if c in df.columns:
                df[c] = to_number_series(df[c])
        stats["数值清洗"] = ", ".join(c for c in num_cols if c in df.columns)

    if "mask" in ops:
        hits = 0
        for c in df.columns:
            sample = " ".join(df[c].map(s).head(200))
            kind = None
            if mask_kind and mask_kind != "auto":
                kind = mask_kind
            elif PHONE_RE.search(sample):
                kind = "phone"
            elif ID_RE.search(sample):
                kind = "id"
            elif MAIL_RE.search(sample):
                kind = "email"
            if kind:
                df[c] = df[c].map(lambda v, k=kind: apply_mask(v, k))
                hits += 1
        stats["脱敏列数"] = str(hits)

    if "dedupe" in ops:
        n0 = len(df)
        if dedupe_key:
            keys = [k for k in dedupe_key if k in df.columns]
            if keys:
                df = df.drop_duplicates(subset=keys, keep="first")
            else:
                df = df.drop_duplicates()
        else:
            df = df.drop_duplicates()
        stats["去重"] = f"{n0}→{len(df)} 行（删除 {n0 - len(df)} 条）"

    df = df.reset_index(drop=True)
    return df


def diagnose(df):
    out = {"行数": len(df), "列数": len(df.columns), "完全重复行": int(df.duplicated().sum()),
           "空单元格占比": f"{df.isna().mean().mean() * 100:.1f}%"}
    issues = []
    for c in df.columns:
        col = df[c].map(s)
        nonempty = col[col.str.strip().ne("")]
        if len(nonempty) == 0:
            issues.append(f"[{c}] 整列为空")
            continue
        junk = int(nonempty.map(lambda t: t != clean_text(t)).sum())
        if junk:
            issues.append(f"[{c}] {junk} 个值含首尾空格/不可见字符")
        uniq = nonempty.nunique()
        if uniq and uniq / len(nonempty) < 0.98 and uniq > 1:
            issues.append(f"[{c}] {len(nonempty) - uniq} 个重复值，疑似主键不唯一")
        sample = " ".join(nonempty.head(50))
        if PHONE_RE.search(sample) or ID_RE.search(sample):
            issues.append(f"[{c}] 含手机号/身份证，需脱敏")
        if MONEY_RE.search(sample):
            stripped = nonempty.map(lambda t: MONEY_RE.sub("", t).replace(" ", ""))
            ok = pd.to_numeric(stripped, errors="coerce").notna().mean()
            if 0.5 < ok < 0.9:
                issues.append(f"[{c}] 疑似数值列但有 {int(len(nonempty) * (1 - ok))} 个值无法解析（混入符号/文字）")
    out["问题"] = issues or ["未发现明显问题"]
    return out


# ---------------------------------------------------------------- 命令
def cmd_clean(a):
    ops = [o.strip() for o in a.ops.split(",") if o.strip()]
    sheets_in = {}
    for p in a.inputs:
        for name, df in read_any(p).items():
            if a.sheet and a.sheet not in name:
                continue
            key = name if len(a.inputs) == 1 else f"{os.path.splitext(os.path.basename(p))[0]}"
            sheets_in[key] = df
    out, total = {}, {}
    for name, df in sheets_in.items():
        stats = {}
        out[name] = clean_df(df, ops, a.date_cols or [], a.number_cols or [],
                             a.dedupe_key or [], a.fill_cols or [], a.mask, stats)
        total[name] = stats
        print(f"\n=== {name} ===")
        for k, v in stats.items():
            print(f"  {k}: {v}")
    path = write_xlsx(out, a.output)
    print(f"\n已输出 -> {path}")
    return path


def cmd_merge(a):
    frames = {}
    for p in a.inputs:
        for name, df in read_any(p).items():
            frames.setdefault(("__all__" if a.single else name), []).append(df)
    merged = {k: pd.concat(v, ignore_index=True) for k, v in frames.items()}
    if not a.single:
        merged = {"__all__": pd.concat(merged.values(), ignore_index=True)}
    df = list(merged.values())[0]
    df.columns = dedupe_cols([norm_col(c) for c in df.columns])
    path = write_xlsx({"合并结果": df}, a.output)
    print(f"合并 {sum(len(v) for v in frames.values())} 个表 -> {len(df)} 行 {len(df.columns)} 列\n-> {path}")
    return path


def cmd_split(a):
    if not a.inputs:
        sys.exit("需要输入文件")
    src = list(read_any(a.inputs[0]).values())[0]
    if a.by not in src.columns:
        sys.exit(f"列不存在: {a.by}（现有列: {list(src.columns)[:15]}）")
    groups = {"全部": src}
    for val, part in src.groupby(src[a.by].map(s), dropna=False):
        groups[str(val)[:31] or "空值"] = part
    if a.outdir:
        os.makedirs(a.outdir, exist_ok=True)
        for name, part in list(groups.items())[1:]:
            safe = re.sub(r"[\\/:*?\"<>|]", "_", name)
            part.to_excel(os.path.join(a.outdir, f"{safe}.xlsx"), index=False)
        print(f"按 [{a.by}] 拆成 {len(groups) - 1} 个文件 -> {a.outdir}")
        return a.outdir
    path = write_xlsx(groups, a.output or "拆分结果.xlsx")
    print(f"按 [{a.by}] 拆成 {len(groups) - 1} 个 sheet -> {path}")
    return path


def cmd_report(a):
    for p in a.inputs:
        for name, df in read_any(p).items():
            print(f"\n=== {p} / {name} ===")
            for k, v in diagnose(df).items():
                if isinstance(v, list):
                    print(f"  {k}:")
                    for i in v:
                        print(f"    - {i}")
                else:
                    print(f"  {k}: {v}")


def main():
    ap = argparse.ArgumentParser(description="Excel/CSV 批量清洗工具")
    sub = ap.add_subparsers(dest="cmd", required=True)

    c = sub.add_parser("clean", help="清洗同构表")
    c.add_argument("inputs", nargs="+")
    c.add_argument("-o", "--output", default="清洗结果.xlsx")
    c.add_argument("--ops", default="trim,norm-cols,empty,dedupe,to-date,to-number")
    c.add_argument("--date-cols", nargs="*")
    c.add_argument("--number-cols", nargs="*")
    c.add_argument("--dedupe-key", nargs="*")
    c.add_argument("--fill-cols", nargs="*")
    c.add_argument("--mask", default="auto", help="auto/phone/id/email/name/off")
    c.add_argument("--sheet", help="只处理名字包含该关键字的 sheet")
    c.set_defaults(func=cmd_clean)

    m = sub.add_parser("merge", help="多表纵向合并")
    m.add_argument("inputs", nargs="+")
    m.add_argument("-o", "--output", default="合并结果.xlsx")
    m.add_argument("--single", action="store_true", help="所有 sheet 合并为一张")
    m.set_defaults(func=cmd_merge)

    sp = sub.add_parser("split", help="按列拆分")
    sp.add_argument("inputs", nargs="+")
    sp.add_argument("--by", required=True)
    sp.add_argument("-o", "--output", default="拆分结果.xlsx")
    sp.add_argument("--outdir", help="拆成多个文件而非多个 sheet")
    sp.set_defaults(func=cmd_split)

    r = sub.add_parser("report", help="只做诊断报告")
    r.add_argument("inputs", nargs="+")
    r.set_defaults(func=cmd_report)

    a = ap.parse_args()
    a.func(a)


if __name__ == "__main__":
    main()
