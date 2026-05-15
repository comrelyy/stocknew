#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Crawling 重试装饰器 + 列重排工具。

P0 数据源替换层：所有 crawling/* 模块统一使用此模块的 @with_retry 装饰外层包装函数，
保证 akshare 上游函数偶发网络异常时自动重试。
"""
import logging
from typing import Iterable

import pandas as pd
from tenacity import (
    retry,
    retry_if_exception_type,
    stop_after_attempt,
    wait_exponential,
    before_sleep_log,
)
import requests

logger = logging.getLogger(__name__)

# 触发重试的异常类型：网络层、JSON 解析失败、akshare 内部偶发 KeyError/ValueError
_RETRYABLE = (
    requests.RequestException,
    ConnectionError,
    ValueError,
    KeyError,
)

with_retry = retry(
    stop=stop_after_attempt(4),
    wait=wait_exponential(multiplier=1.5, min=1, max=20),
    retry=retry_if_exception_type(_RETRYABLE),
    reraise=True,
    before_sleep=before_sleep_log(logger, logging.WARNING),
)

# 多源 fallback 内部使用的轻量级重试：每个备用源只试 2 次、退避更短，
# 总开销可控（N 源 × 2 次 ≈ 单源 4 次），保证整体超时不爆炸。
with_retry_lite = retry(
    stop=stop_after_attempt(2),
    wait=wait_exponential(multiplier=1.0, min=0.5, max=5),
    retry=retry_if_exception_type(_RETRYABLE),
    reraise=True,
    before_sleep=before_sleep_log(logger, logging.WARNING),
)


def reorder(df: pd.DataFrame, expected_cols: Iterable[str]) -> pd.DataFrame:
    """按 expected_cols 选取并重排 DataFrame 的列。

    若 akshare 上游列名有改动导致期望列不存在，抛 KeyError 让重试机制处理，
    多次失败后向上层抛出便于告警。
    """
    expected = list(expected_cols)
    missing = [c for c in expected if c not in df.columns]
    if missing:
        raise KeyError(
            f"akshare 返回列与本地契约不一致，缺失列：{missing}；"
            f"实际列：{list(df.columns)}"
        )
    return df[expected].copy()
