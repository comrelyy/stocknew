#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
东方财富网-数据中心-年报季报-分红送配
P0 改造：薄包装 akshare 的同名函数，保留原列序，加 tenacity 重试。
"""
import akshare as ak
import pandas as pd

from instock.core.crawling._retry import with_retry, reorder

_FHPS_COLS = [
    "代码", "名称",
    "送转股份-送转总比例", "送转股份-送转比例", "送转股份-转股比例",
    "现金分红-现金分红比例", "现金分红-股息率",
    "每股收益", "每股净资产", "每股公积金", "每股未分配利润",
    "净利润同比增长", "总股本",
    "预案公告日", "股权登记日", "除权除息日", "方案进度", "最新公告日期",
]


@with_retry
def stock_fhps_em(date: str = "20210630") -> pd.DataFrame:
    df = ak.stock_fhps_em(date=date)
    if df is None or df.empty:
        return df
    return reorder(df, _FHPS_COLS)
