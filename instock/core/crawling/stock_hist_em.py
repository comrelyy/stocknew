#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
A 股实时行情 + 历史 K 线：东财 → 新浪 → 腾讯 多源容错。

主源：东方财富（保留原 40 列 spot / 11 列 hist 契约，stockfetch 依赖此顺序）。
备用源：
  - spot：新浪 `ak.stock_zh_a_spot`。新浪只覆盖价量与少量估值字段，
    剩余财务字段（每股收益/营收/净利润/总市值等）会被填充为 NaN —— 仅做应急兜底。
  - hist：新浪 `ak.stock_zh_a_daily`、腾讯 `ak.stock_zh_a_hist_tx`，均归一化为东财 11 列。
"""
import logging
from typing import Optional, cast

import akshare as ak
import pandas as pd

from instock.core.crawling._retry import with_retry, with_retry_lite, reorder
from instock.core.crawling._fallback import Source, try_sources

logger = logging.getLogger(__name__)

# 旧 crawling 实现 stock_zh_a_spot_em 返回的 40 列（按位置契约，stockfetch 依赖此顺序）
_SPOT_COLS = [
    "代码", "名称", "最新价", "涨跌幅", "涨跌额", "成交量", "成交额", "振幅",
    "换手率", "量比", "今开", "最高", "最低", "昨收", "涨速", "5分钟涨跌",
    "60日涨跌幅", "年初至今涨跌幅", "市盈率动", "市盈率TTM", "市盈率静", "市净率",
    "每股收益", "每股净资产", "每股公积金", "每股未分配利润", "加权净资产收益率",
    "毛利率", "资产负债率", "营业收入", "营业收入同比增长", "归属净利润",
    "归属净利润同比增长", "报告期", "总股本", "已流通股份", "总市值", "流通市值",
    "所处行业", "上市时间",
]

_HIST_COLS = [
    "日期", "开盘", "收盘", "最高", "最低", "成交量", "成交额",
    "振幅", "涨跌幅", "涨跌额", "换手率",
]

# 新浪 stock_zh_a_spot 实际返回 14 列：代码、名称、最新价、涨跌额、涨跌幅、买入、卖出、
# 昨收、今开、最高、最低、成交量、成交额、时间戳。
# 与东财 40 列 schema 对齐的字段：价/量/开高低/昨收/涨跌（10 列）；缺：换手率、量比、
# 振幅（可派生）、所有 PE/PB/财务/总市值/流通市值/行业/上市时间（25+ 列 → NaN）。


def _strip_market_prefix(s: pd.Series) -> pd.Series:
    return s.astype(str).str.replace(r"^(sh|sz|bj)", "", regex=True)


# ====================== Spot 数据源 ======================

@with_retry_lite
def _spot_em() -> pd.DataFrame:
    df = ak.stock_zh_a_spot_em()
    return reorder(df, _SPOT_COLS)


@with_retry_lite
def _spot_sina_raw() -> pd.DataFrame:
    return ak.stock_zh_a_spot()


def _spot_sina_to_em(df: pd.DataFrame) -> pd.DataFrame:
    """新浪 spot → 东财 40 列 schema。

    新浪当前版（akshare 1.18+）返回 14 列中文列名，可对齐东财 10 列价量字段。
    振幅由 (最高-最低)/昨收 派生。剩余 29 列（换手率/量比/PE/PB/财务/市值/行业等）
    全部置 NaN —— 这是新浪源应急兜底的预期取舍：保住价/量主线，下游财务过滤的策略
    在 fallback 期间会被自动跳过（NaN 比较返回 False）。
    """
    df = df.copy()

    if "代码" in df.columns:
        df["代码"] = _strip_market_prefix(df["代码"])

    # 新浪 成交量 单位是股，东财契约是手 → /100
    if "成交量" in df.columns:
        df["成交量"] = pd.to_numeric(df["成交量"], errors="coerce") / 100.0

    # 振幅可派生
    if {"最高", "最低", "昨收"}.issubset(df.columns):
        high = cast(pd.Series, pd.to_numeric(df["最高"], errors="coerce"))
        low = cast(pd.Series, pd.to_numeric(df["最低"], errors="coerce"))
        pre = cast(pd.Series, pd.to_numeric(df["昨收"], errors="coerce"))
        df["振幅"] = ((high - low) / pre * 100).round(4)

    # 补齐缺失列为 NaN
    missing_cols = [c for c in _SPOT_COLS if c not in df.columns]
    for col in missing_cols:
        df[col] = pd.NA

    logger.warning(
        "[新浪 spot fallback] 已生效；%d 个东财字段已置 NaN（换手率/PE/财务/市值等）",
        len(missing_cols),
    )
    return df[_SPOT_COLS].copy()


@with_retry
def stock_zh_a_spot_em() -> Optional[pd.DataFrame]:
    """全市场 A 股实时行情：东财（主）→ 新浪（兜底，财务列为 NaN）。"""
    return try_sources(
        Source("东财", _spot_em),
        Source("新浪", _spot_sina_raw, normalize=_spot_sina_to_em),
    )


# ====================== Hist 日 K 数据源 ======================

def _add_market_prefix(code: str) -> str:
    """6 位代码 → 带 sh/sz/bj 前缀，供新浪/腾讯接口使用。"""
    code = str(code)
    if code.startswith(("sh", "sz", "bj")):
        return code
    if code.startswith(("60", "68", "9")):  # 60x/688/9xx 沪市
        return f"sh{code}"
    if code.startswith(("0", "3", "2")):  # 000/001/002/003/300/301/200 深市
        return f"sz{code}"
    if code.startswith(("4", "8")):  # 北交所
        return f"bj{code}"
    return f"sh{code}"


@with_retry_lite
def _hist_em(symbol, period, start_date, end_date, adjust) -> pd.DataFrame:
    df = ak.stock_zh_a_hist(
        symbol=symbol, period=period,
        start_date=start_date, end_date=end_date, adjust=adjust,
    )
    if df is None or df.empty:
        return df
    return reorder(df, _HIST_COLS)


def _normalize_ohlcv_to_em(
    df: pd.DataFrame,
    vol_unit_is_share: bool,
    turnover_is_fraction: bool = False,
) -> pd.DataFrame:
    """把 OHLCV 列（英文或中文）归一化到东财 _HIST_COLS 11 列。

    Args:
        df: 必须至少含 date/open/close/high/low/volume/amount 列（英文或对应中文）。
        vol_unit_is_share: True = 源 volume 单位是股（如新浪），需 /100 换算成手。
        turnover_is_fraction: True = 源 turnover 是 0-1 小数（如新浪），需 *100 换算成百分比。
    """
    df = df.copy()
    rename = {
        "date": "日期", "open": "开盘", "close": "收盘", "high": "最高",
        "low": "最低", "volume": "成交量", "amount": "成交额",
        "turnover": "换手率",
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})

    required = ["日期", "开盘", "收盘", "最高", "最低", "成交量", "成交额"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"hist fallback 源缺少必需列：{missing}；实际列：{list(df.columns)}")

    if vol_unit_is_share:
        df["成交量"] = pd.to_numeric(df["成交量"], errors="coerce") / 100.0

    # 派生列：涨跌额 / 涨跌幅 / 振幅
    close = cast(pd.Series, pd.to_numeric(df["收盘"], errors="coerce"))
    prev_close = close.shift(1)
    high = cast(pd.Series, pd.to_numeric(df["最高"], errors="coerce"))
    low = cast(pd.Series, pd.to_numeric(df["最低"], errors="coerce"))
    df["涨跌额"] = (close - prev_close).round(4)
    df["涨跌幅"] = ((close - prev_close) / prev_close * 100).round(4)
    df["振幅"] = ((high - low) / prev_close * 100).round(4)

    if "换手率" not in df.columns:
        df["换手率"] = pd.NA
    elif turnover_is_fraction:
        df["换手率"] = (cast(pd.Series, pd.to_numeric(df["换手率"], errors="coerce")) * 100).round(4)

    return df[_HIST_COLS].copy()


@with_retry_lite
def _hist_sina(symbol, period, start_date, end_date, adjust) -> pd.DataFrame:
    if period != "daily":
        raise ValueError(f"新浪 hist fallback 仅支持 daily, 收到 {period}")
    prefixed = _add_market_prefix(symbol)
    df = ak.stock_zh_a_daily(
        symbol=prefixed,
        start_date=start_date,
        end_date=end_date,
        adjust=adjust or "qfq",
    )
    if df is None or df.empty:
        return df
    return _normalize_ohlcv_to_em(df, vol_unit_is_share=True, turnover_is_fraction=True)


@with_retry_lite
def _hist_tx(symbol, period, start_date, end_date, adjust) -> pd.DataFrame:
    if period != "daily":
        raise ValueError(f"腾讯 hist fallback 仅支持 daily, 收到 {period}")
    prefixed = _add_market_prefix(symbol)
    df = ak.stock_zh_a_hist_tx(
        symbol=prefixed,
        start_date=start_date,
        end_date=end_date,
        adjust=adjust or "qfq",
    )
    if df is None or df.empty:
        return df
    # 腾讯返回列名 akshare 已是中文（日期/开盘/收盘/最高/最低/成交量/成交额），单位也是手
    return _normalize_ohlcv_to_em(df, vol_unit_is_share=False)


@with_retry
def stock_zh_a_hist(
    symbol: str = "000001",
    period: str = "daily",
    start_date: str = "19700101",
    end_date: str = "20500101",
    adjust: str = "",
) -> Optional[pd.DataFrame]:
    """A 股日 K 线：东财（主）→ 新浪 → 腾讯。"""
    return try_sources(
        Source("东财", _hist_em),
        Source("新浪", _hist_sina),
        Source("腾讯", _hist_tx),
        symbol=symbol, period=period,
        start_date=start_date, end_date=end_date, adjust=adjust,
    )
