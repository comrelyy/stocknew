#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
东方财富网-数据中心-大宗交易
P0 改造：薄包装 akshare 的同名函数，保留旧列序，加 tenacity 重试。
"""
import akshare as ak
import pandas as pd

from instock.core.crawling._retry import with_retry, reorder

# 旧版 stock_dzjy_mrtj 返回 12 列（按位置契约）。
# 注意：akshare 当前版返回顺序里 [涨跌幅, 收盘价] 顺序与旧版 [收盘价, 涨跌幅] 不同，
# 这里通过 reorder 重新按旧列序排列。
_MRTJ_COLS = [
    "序号", "交易日期", "证券代码", "证券简称",
    "收盘价", "涨跌幅", "成交价", "折溢率",
    "成交笔数", "成交总量", "成交总额", "成交总额/流通市值",
]


@with_retry
def stock_dzjy_mrtj(
    start_date: str = "20220105", end_date: str = "20220105",
) -> pd.DataFrame:
    df = ak.stock_dzjy_mrtj(start_date=start_date, end_date=end_date)
    if df is None or df.empty:
        return df
    return reorder(df, _MRTJ_COLS)
