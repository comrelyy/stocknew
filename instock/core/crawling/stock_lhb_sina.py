#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
新浪财经-龙虎榜
P0 改造：薄包装 akshare 的同名函数。

注意：akshare 1.18+ 把 stock_lhb_ggtj_sina 的参数从 `recent_day` 改名为 `symbol`，
此 wrapper 保留旧签名 `recent_day` 以兼容上层 stockfetch.fetch_stock_top_data。
"""
import akshare as ak
import pandas as pd

from instock.core.crawling._retry import with_retry, reorder

# stockfetch 把 9 列 [date, code, name, ranking_times, sum_buy, sum_sell, net_amount,
# buy_seat, sell_seat] 减去 date 后按位置赋值 8 列。
# akshare 当前返回列名："股票代码、股票名称、上榜次数、累积购买额、累积卖出额、净额、买入席位数、卖出席位数"
_GGTJ_COLS = [
    "股票代码", "股票名称", "上榜次数",
    "累积购买额", "累积卖出额", "净额",
    "买入席位数", "卖出席位数",
]


@with_retry
def stock_lhb_ggtj_sina(recent_day: str = "5") -> pd.DataFrame:
    # akshare 新版用 symbol 参数；值仍是 "5"/"10"/"30"/"60"
    df = ak.stock_lhb_ggtj_sina(symbol=recent_day)
    if df is None or df.empty:
        return df
    return reorder(df, _GGTJ_COLS)
