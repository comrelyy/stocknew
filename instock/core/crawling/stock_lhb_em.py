#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
东方财富网-数据中心-龙虎榜单
P0 改造：薄包装 akshare 的同名函数，加 tenacity 重试。
"""
import akshare as ak
import pandas as pd

from instock.core.crawling._retry import with_retry


@with_retry
def stock_lhb_jgmmtj_em(
    start_date: str = "20220906", end_date: str = "20220906",
) -> pd.DataFrame:
    """
    机构买卖每日统计。stockfetch.fetch_stock_top_entity_data 通过列名
    `代码`、`买方机构数` 访问，无需位置契约，直接透传 akshare 即可。
    """
    return ak.stock_lhb_jgmmtj_em(start_date=start_date, end_date=end_date)
