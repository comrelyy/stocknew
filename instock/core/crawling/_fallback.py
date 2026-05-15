#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
多数据源容错框架。

设计目标：
  - 每个 crawling 接口声明一个或多个备用数据源（Source）；
  - try_sources() 依次尝试：失败/空数据 → 切换下一个源；
  - 每个源内部已经被 @with_retry_lite 装饰，单源最多重试 2 次；
  - 任何源成功即返回（经过该源自带的 normalize 归一化到主源 schema）；
  - 全部失败时，向上抛出最后一次异常，由调用方落到 stockfetch 的 except 分支。

使用示例:
    from instock.core.crawling._fallback import Source, try_sources
    from instock.core.crawling._retry import with_retry_lite

    @with_retry_lite
    def _spot_em():
        return ak.stock_zh_a_spot_em()

    @with_retry_lite
    def _spot_sina():
        return ak.stock_zh_a_spot()

    def stock_zh_a_spot_em():
        return try_sources(
            Source("东财", _spot_em),
            Source("新浪", _spot_sina, normalize=_sina_spot_to_em),
        )
"""
import logging
from typing import Callable, Optional

import pandas as pd

logger = logging.getLogger(__name__)


class Source:
    """单个数据源候选。

    Attributes:
        name: 源名称，仅用于日志（"东财"/"新浪"/"腾讯"）。
        fetch: 不带参数即可调用的 callable，返回 DataFrame；应已被 @with_retry_lite 装饰。
        normalize: 可选的列归一化函数，把该源返回的 DataFrame 转换为主源 schema。
                   主源（第一个 Source）无需 normalize；备用源若 schema 与主源不同则必须提供。
    """

    __slots__ = ("name", "fetch", "normalize")

    def __init__(
        self,
        name: str,
        fetch: Callable[..., pd.DataFrame],
        normalize: Optional[Callable[[pd.DataFrame], pd.DataFrame]] = None,
    ):
        self.name = name
        self.fetch = fetch
        self.normalize = normalize


def try_sources(*sources: Source, **fetch_kwargs) -> Optional[pd.DataFrame]:
    """按声明顺序尝试每个 Source，第一个成功（非空）即返回。

    Args:
        *sources: 至少一个 Source。第一个为主源，其余为 fallback。
        **fetch_kwargs: 透传给每个源的 fetch 函数。所有源的签名必须兼容这组 kwargs。

    Returns:
        归一化后的 DataFrame；如果所有源都失败，向上抛出最后一次异常。
        如果所有源都返回空，返回 None（语义同单源原行为）。
    """
    if not sources:
        raise ValueError("try_sources 至少需要一个 Source")

    last_err: Optional[BaseException] = None
    saw_empty = False
    for src in sources:
        try:
            df = src.fetch(**fetch_kwargs)
        except Exception as e:
            last_err = e
            logger.warning("[数据源 %s] 失败: %s，尝试下一个源", src.name, str(e)[:160])
            continue

        if df is None or (hasattr(df, "empty") and df.empty):
            saw_empty = True
            logger.warning("[数据源 %s] 返回空数据，尝试下一个源", src.name)
            continue

        if src.normalize is not None:
            try:
                df = src.normalize(df)
            except Exception as e:
                last_err = e
                logger.warning(
                    "[数据源 %s] 归一化失败: %s，尝试下一个源", src.name, str(e)[:160]
                )
                continue

        logger.info("[数据源 %s] 命中", src.name)
        return df

    # 全部源都被尝试过：若至少一个源返回空，视作"无数据"返回 None；
    # 否则全部抛异常，抛出最后一次异常给上层
    if saw_empty and last_err is None:
        return None
    if last_err is not None:
        raise last_err
    return None
