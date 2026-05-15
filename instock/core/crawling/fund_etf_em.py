#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETF 行情：东方财富（主）+ 新浪 ETF hist（兜底）。

ETF 实时行情 `fund_etf_spot_em` akshare 没有等价新浪/腾讯接口，保留单源。
ETF 日 K `fund_etf_hist_em` 可在 `fund_etf_hist_sina` 兜底。
"""
import logging
from typing import Optional, cast

import akshare as ak
import pandas as pd

from instock.core.crawling._retry import with_retry, with_retry_lite, reorder
from instock.core.crawling._fallback import Source, try_sources

logger = logging.getLogger(__name__)

# 旧实现 fund_etf_spot_em 返回 14 列；注意"涨跌额"在"涨跌幅"之前（与 tablestructure 不完全匹配，
# 但此处保持与原项目行为一致，避免引入 silent data 重写）
_ETF_SPOT_COLS = [
    "代码", "名称", "最新价", "涨跌额", "涨跌幅", "成交量", "成交额",
    "开盘价", "最高价", "最低价", "昨收", "换手率", "流通市值", "总市值",
]

_ETF_HIST_COLS = [
    "日期", "开盘", "收盘", "最高", "最低", "成交量", "成交额",
    "振幅", "涨跌幅", "涨跌额", "换手率",
]


@with_retry
def fund_etf_spot_em() -> pd.DataFrame:
    df = ak.fund_etf_spot_em()
    return reorder(df, _ETF_SPOT_COLS)


# ====================== ETF Hist 多源 ======================

@with_retry_lite
def _etf_hist_em(symbol, period, start_date, end_date, adjust) -> pd.DataFrame:
    df = ak.fund_etf_hist_em(
        symbol=symbol, period=period,
        start_date=start_date, end_date=end_date, adjust=adjust,
    )
    if df is None or df.empty:
        return df
    return reorder(df, _ETF_HIST_COLS)


def _add_etf_market_prefix(code: str) -> str:
    """ETF 代码 → 带 sh/sz 前缀。51x/56x/58x = 沪市，15x/16x/18x = 深市。"""
    code = str(code)
    if code.startswith(("sh", "sz")):
        return code
    if code.startswith(("5",)):
        return f"sh{code}"
    if code.startswith(("1",)):
        return f"sz{code}"
    return f"sh{code}"


@with_retry_lite
def _etf_hist_sina(symbol, period, start_date, end_date, adjust) -> pd.DataFrame:
    if period != "daily":
        raise ValueError(f"新浪 ETF hist fallback 仅支持 daily, 收到 {period}")
    prefixed = _add_etf_market_prefix(symbol)
    df = ak.fund_etf_hist_sina(symbol=prefixed)
    if df is None or df.empty:
        return df

    # 新浪 fund_etf_hist_sina 返回列：date, open, high, low, close, volume
    # 需归一化为东财 11 列：日期/开盘/收盘/最高/最低/成交量/成交额/振幅/涨跌幅/涨跌额/换手率
    df = df.copy()
    df = df.rename(columns={
        "date": "日期", "open": "开盘", "close": "收盘",
        "high": "最高", "low": "最低", "volume": "成交量",
    })

    required = ["日期", "开盘", "收盘", "最高", "最低", "成交量"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"新浪 ETF hist 缺少必需列：{missing}；实际列：{list(df.columns)}")

    # 按日期过滤
    df["日期"] = pd.to_datetime(df["日期"]).dt.strftime("%Y-%m-%d")
    if start_date and start_date != "19700101":
        sd = pd.to_datetime(start_date).strftime("%Y-%m-%d")
        df = df[df["日期"] >= sd]
    if end_date and end_date != "20500101":
        ed = pd.to_datetime(end_date).strftime("%Y-%m-%d")
        df = df[df["日期"] <= ed]

    if df.empty:
        return df

    # 派生列
    close = cast(pd.Series, pd.to_numeric(df["收盘"], errors="coerce"))
    prev_close = close.shift(1)
    high = cast(pd.Series, pd.to_numeric(df["最高"], errors="coerce"))
    low = cast(pd.Series, pd.to_numeric(df["最低"], errors="coerce"))
    open_ = cast(pd.Series, pd.to_numeric(df["开盘"], errors="coerce"))
    vol = cast(pd.Series, pd.to_numeric(df["成交量"], errors="coerce"))

    # 新浪 ETF 单位是股 → 转换成手以匹配东财契约（stockfetch 之后会再 *100 回到股）
    df["成交量"] = vol / 100.0
    # 成交额：新浪不返回，用 close*volume 估算（仅在 fallback 时使用）
    df["成交额"] = (close * vol).round(2)
    df["涨跌额"] = (close - prev_close).round(4)
    df["涨跌幅"] = ((close - prev_close) / prev_close * 100).round(4)
    df["振幅"] = ((high - low) / prev_close * 100).round(4)
    df["换手率"] = pd.NA

    # 防止未使用变量警告（open_ 仅用于完整性校验）
    _ = open_

    return df[_ETF_HIST_COLS].copy()


@with_retry
def fund_etf_hist_em(
    symbol: str = "159707",
    period: str = "daily",
    start_date: str = "19700101",
    end_date: str = "20500101",
    adjust: str = "",
) -> Optional[pd.DataFrame]:
    """ETF 日 K：东财（主）→ 新浪（兜底，成交额估算自 close*volume，换手率为 NaN）。"""
    return try_sources(
        Source("东财", _etf_hist_em),
        Source("新浪", _etf_hist_sina),
        symbol=symbol, period=period,
        start_date=start_date, end_date=end_date, adjust=adjust,
    )
