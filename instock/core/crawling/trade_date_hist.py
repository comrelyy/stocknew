#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
交易日历-历史数据
P0 改造：薄包装 akshare 的同名函数，加 tenacity 重试。
"""
import akshare as ak
import pandas as pd

from instock.core.crawling._retry import with_retry


@with_retry
def tool_trade_date_hist_sina() -> pd.DataFrame:
    """返回单列 DataFrame: trade_date"""
    df = ak.tool_trade_date_hist_sina()
    # stockfetch 通过列名 'trade_date' 访问，无需重排
    return df
