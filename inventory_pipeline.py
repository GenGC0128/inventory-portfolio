#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
wide-table-etl — 多来源 Excel 横表自动转标准长表（作品集演示版）

对应真实业务场景：多个生产基地每天上报格式各异的库存横表，
需要自动定位数据区域、宽转长、统一口径后汇总出数。

核心设计：
1. 每类模板一条 BaseRule 配置：坐标定位规则 + 日期模式 + 口径映射
2. 动态边界定位：日期列起点 / identify 关键字 / 表尾标志自动识别，
   报表列数增减、行数增减均不影响解析
3. Excel 序列日期、datetime 日期、行内日期三种模式统一处理
4. 口径归一化词典外置 config/normalization.json，业务规则与代码分离
5. 单文件异常隔离 + failed.csv，与异构 PDF 解析项目同一套架构思路

用法：
    python make_sample_data.py
    python inventory_pipeline.py --input data/reports --config config/normalization.json --out data/output
"""
from __future__ import annotations

import argparse
import json
import logging
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger("wide-etl")

LONG_COLS = ["date", "base", "identify", "quantity", "spec", "supplier"]

# ---------------------------------------------------------------------------
# 1. 模板规则配置
# ---------------------------------------------------------------------------

@dataclass
class BaseRule:
    """一类横表模板的解析规则。新增来源 = 新增一条配置。"""
    name: str
    date_row: int                 # 日期所在行
    id_row: int                   # identify（结存/入库/出库 等）所在行
    supplier_col: int             # 供应商所在列
    spec_col: int                 # 规格所在列
    date_mode: str = "datetime"   # "datetime" | "serial"
    data_row_start: int = 3       # 数据起始行（含表头行之下）
    end_marker: str = "total"     # supplier_col 中出现该词视为表尾（不区分大小写）
    id_keyword: Optional[str] = None   # date_row 中用于定位右边界的关键字（不含则到行尾）
    project_mode: bool = False    # 项目表：日期在行内第 0 列，无 identify 概念


RULES: dict[str, BaseRule] = {
    "base_a": BaseRule(name="base_a", date_row=0, id_row=1,
                       supplier_col=1, spec_col=2, date_mode="datetime",
                       data_row_start=3, end_marker="total", id_keyword="qty"),
    "base_b": BaseRule(name="base_b", date_row=0, id_row=1,
                       supplier_col=0, spec_col=2, date_mode="datetime",
                       data_row_start=2, end_marker="total"),
    "base_c": BaseRule(name="base_c", date_row=2, id_row=3,
                       supplier_col=0, spec_col=2, date_mode="serial",
                       data_row_start=4, end_marker=""),
    "projects": BaseRule(name="projects", date_row=0, id_row=0,
                         supplier_col=0, spec_col=0, date_mode="datetime",
                         data_row_start=1, project_mode=True),
}

EXCEL_SERIAL_ORIGIN = "1899-12-30"  # Excel 序列日期基准（含 1900 闰年 bug 修正）


# ---------------------------------------------------------------------------
# 2. 解析阶段：动态定位 + 宽转长
# ---------------------------------------------------------------------------

def _find_left_border(df: pd.DataFrame, rule: BaseRule) -> int:
    """在 date_row 上找第一个日期值所在列，作为数据区左边界。"""
    import datetime as _dt
    row = df.iloc[rule.date_row]
    for col in range(len(row)):
        v = row.iloc[col]
        if pd.isna(v):
            continue
        if rule.date_mode == "datetime" and isinstance(v, (pd.Timestamp, _dt.datetime, _dt.date)):
            return col
        if rule.date_mode == "serial" and isinstance(v, (int, float, np.integer, np.floating)):
            return col
    raise ValueError(f"[{rule.name}] date_row={rule.date_row} 上未找到日期起点")


def _find_right_border(df: pd.DataFrame, rule: BaseRule, left: int) -> int:
    """优先按 id_keyword 定位右边界；无关键字则到行尾。返回开区间上界。"""
    if rule.id_keyword:
        row = df.iloc[rule.date_row]
        for col in range(left, len(row)):
            if rule.id_keyword.lower() in str(row.iloc[col]).lower():
                return col
    return len(df.columns)


def _find_down_border(df: pd.DataFrame, rule: BaseRule, up: int) -> int:
    """supplier_col 中出现 end_marker（或空值）即视为表尾。返回开区间上界。"""
    col = df.iloc[:, rule.supplier_col]
    for row in range(up, len(df)):
        v = col.iloc[row]
        if pd.isna(v):
            return row
        if rule.end_marker and rule.end_marker.lower() in str(v).lower():
            return row
    return len(df)


def extract_wide(df: pd.DataFrame, rule: BaseRule) -> pd.DataFrame:
    """按规则把横表数据区转成长表。所有来源输出统一 LONG_COLS schema。"""
    if rule.project_mode:
        return _extract_project(df, rule)

    left = _find_left_border(df, rule)
    right = _find_right_border(df, rule, left)
    up = rule.data_row_start
    down = _find_down_border(df, rule, up)
    if down <= up or right <= left:
        return pd.DataFrame(columns=LONG_COLS)

    records = []
    for x in range(up, down):
        supplier = df.iloc[x, rule.supplier_col]
        spec = df.iloc[x, rule.spec_col]
        if pd.isna(supplier) or pd.isna(spec):
            continue
        for y in range(left, right):
            value = df.iloc[x, y]
            if pd.isna(value) or isinstance(value, str):
                continue
            raw_date = df.iloc[rule.date_row, y]
            if rule.date_mode == "serial":
                date = pd.to_datetime(float(raw_date), unit="D", origin=EXCEL_SERIAL_ORIGIN)
            else:
                date = pd.to_datetime(raw_date, errors="coerce")
            records.append({
                "date": date, "base": rule.name,
                "identify": str(df.iloc[rule.id_row, y]),
                "quantity": float(value),
                "spec": str(spec), "supplier": str(supplier),
            })
    return pd.DataFrame(records, columns=LONG_COLS)


def _extract_project(df: pd.DataFrame, rule: BaseRule) -> pd.DataFrame:
    """项目表：行首即日期，后面列为数量，identify 固定为 project。"""
    records = []
    for x in range(rule.data_row_start, len(df)):
        raw_date = df.iloc[x, 0]
        date = pd.to_datetime(raw_date, errors="coerce")
        if pd.isna(date):
            continue
        for y in range(1, len(df.columns)):
            value = df.iloc[x, y]
            if pd.isna(value) or isinstance(value, str):
                continue
            records.append({
                "date": date, "base": rule.name, "identify": "project",
                "quantity": float(value),
                "spec": str(df.iloc[rule.data_row_start - 1, y]),  # 表头行作 spec
                "supplier": str(df.iloc[0, y]),
            })
    return pd.DataFrame(records, columns=LONG_COLS)


# ---------------------------------------------------------------------------
# 3. 口径归一化（词典外置）
# ---------------------------------------------------------------------------

def load_normalization(path: Path) -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def normalize(df: pd.DataFrame, norm: dict) -> pd.DataFrame:
    """按外置词典统一 spec/density/size/category/supplier 口径。"""
    if df.empty:
        return df
    out = df.copy()

    def first_match(text: str, keywords: list) -> Optional[str]:
        for kw in keywords:
            if re.search(re.escape(kw), text, flags=re.IGNORECASE):
                return kw
        return None

    out["spec_std"] = out["spec"].map(lambda s: first_match(s, norm["spec_aliases"]) or "other")
    out["density"] = out["spec_std"].map(
        lambda s: "loose" if s in norm["loose_specs"] else "dense")
    out["size"] = out["spec"].map(lambda s: first_match(s, norm["size_aliases"]) or "block")
    out["size_std"] = out["size"].map(
        lambda s: "small" if s in norm["small_sizes"] else "block")
    out["category"] = out["spec_std"].map(
        lambda s: "granular" if s in norm["granular_specs"]
        else ("recycled" if s in norm["recycled_specs"] else "virgin"))
    out["supplier_std"] = out["supplier"].map(
        lambda s: first_match(s, norm["supplier_aliases"]) or "other")
    return out


# ---------------------------------------------------------------------------
# 4. 汇总调度
# ---------------------------------------------------------------------------

def run_pipeline(input_dir: Path, config_path: Path, out_dir: Path) -> dict:
    norm = load_normalization(config_path)
    out_dir.mkdir(parents=True, exist_ok=True)
    failures, frames, report = [], [], []

    for xlsx in sorted(input_dir.rglob("*.xlsx")):
        source = xlsx.parent.name
        rule = RULES.get(source)
        if rule is None:
            failures.append({"file": xlsx.name, "source": source, "reason": "unknown source dir"})
            continue
        try:
            df = pd.read_excel(xlsx, header=None)
            long_df = extract_wide(df, rule)
            long_df = normalize(long_df, norm)
            if long_df.empty:
                failures.append({"file": xlsx.name, "source": source, "reason": "empty result"})
                continue
            frames.append(long_df)
            report.append({"file": xlsx.name, "source": source, "rows": len(long_df)})
        except Exception as exc:
            failures.append({"file": xlsx.name, "source": source,
                             "reason": f"{type(exc).__name__}: {exc}"})

    result = {"summary_rows": 0, "failed_files": len(failures)}
    if frames:
        summary = pd.concat(frames, ignore_index=True, sort=False)
        summary.to_excel(out_dir / "summary.xlsx", index=False)
        result["summary_rows"] = len(summary)
        log.info("summary: %d rows -> %s", len(summary), out_dir / "summary.xlsx")
    if failures:
        pd.DataFrame(failures).to_csv(out_dir / "failed.csv", index=False)
        log.warning("%d failed files, see failed.csv", len(failures))
    pd.DataFrame(report).to_csv(out_dir / "report.csv", index=False)
    return result


def main() -> None:
    ap = argparse.ArgumentParser(description="wide-table-etl demo pipeline")
    ap.add_argument("--input", required=True, type=Path)
    ap.add_argument("--config", required=True, type=Path)
    ap.add_argument("--out", required=True, type=Path)
    args = ap.parse_args()
    stats = run_pipeline(args.input, args.config, args.out)
    log.info("done: %s", stats)


if __name__ == "__main__":
    main()
