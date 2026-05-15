#!/usr/bin/env python3
"""
契约检查：调用 P0 改造后的 crawling/* 包装函数，断言列名/列序与 stockfetch.py
依赖的旧契约一致。

跑法：.venv/bin/python tests/check_akshare_contract.py
"""
import os
import sys
import warnings

warnings.filterwarnings("ignore")

cpath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, cpath)

# 通过 wrapper 调用（已 reorder + retry）
from instock.core.crawling import (
    stock_hist_em as she,
    fund_etf_em as fee,
    trade_date_hist as tdh,
    stock_fund_em as sff,
    stock_fhps_em as sfe,
    stock_lhb_em as sle,
    stock_lhb_sina as sls,
    stock_dzjy_em as sde,
)

# 每个测试：(label, callable, 期望列名 list 或 None=只要不抛异常)
CASES = [
    ("stock_zh_a_spot_em",
        lambda: she.stock_zh_a_spot_em(),
        # 40 列前缀检查（取前 14 列足够断言契约）
        ["代码", "名称", "最新价", "涨跌幅", "涨跌额", "成交量",
         "成交额", "振幅", "换手率", "量比", "今开", "最高", "最低", "昨收"]),
    ("stock_zh_a_hist(000001)",
        lambda: she.stock_zh_a_hist(symbol="000001", period="daily",
                                     start_date="20240101", end_date="20240110",
                                     adjust="qfq"),
        ["日期", "开盘", "收盘", "最高", "最低", "成交量", "成交额",
         "振幅", "涨跌幅", "涨跌额", "换手率"]),
    ("fund_etf_spot_em",
        lambda: fee.fund_etf_spot_em(),
        ["代码", "名称", "最新价", "涨跌额", "涨跌幅", "成交量", "成交额",
         "开盘价", "最高价", "最低价", "昨收", "换手率", "流通市值", "总市值"]),
    ("fund_etf_hist_em(510300)",
        lambda: fee.fund_etf_hist_em(symbol="510300", period="daily",
                                     start_date="20240101", end_date="20240110",
                                     adjust="qfq"),
        ["日期", "开盘", "收盘", "最高", "最低", "成交量", "成交额",
         "振幅", "涨跌幅", "涨跌额", "换手率"]),
    ("tool_trade_date_hist_sina",
        lambda: tdh.tool_trade_date_hist_sina(),
        ["trade_date"]),
    ("stock_individual_fund_flow_rank(今日)",
        lambda: sff.stock_individual_fund_flow_rank(indicator="今日"),
        ["代码", "名称", "最新价", "今日涨跌幅",
         "今日主力净流入-净额", "今日主力净流入-净占比",
         "今日超大单净流入-净额", "今日超大单净流入-净占比",
         "今日大单净流入-净额", "今日大单净流入-净占比",
         "今日中单净流入-净额", "今日中单净流入-净占比",
         "今日小单净流入-净额", "今日小单净流入-净占比"]),
    ("stock_sector_fund_flow_rank(今日,行业)",
        lambda: sff.stock_sector_fund_flow_rank(indicator="今日",
                                                 sector_type="行业资金流"),
        ["名称", "今日涨跌幅",
         "今日主力净流入-净额", "今日主力净流入-净占比",
         "今日超大单净流入-净额", "今日超大单净流入-净占比",
         "今日大单净流入-净额", "今日大单净流入-净占比",
         "今日中单净流入-净额", "今日中单净流入-净占比",
         "今日小单净流入-净额", "今日小单净流入-净占比",
         "今日主力净流入最大股"]),
    ("stock_fhps_em(20231231)",
        lambda: sfe.stock_fhps_em(date="20231231"),
        ["代码", "名称",
         "送转股份-送转总比例", "送转股份-送转比例", "送转股份-转股比例",
         "现金分红-现金分红比例", "现金分红-股息率",
         "每股收益", "每股净资产", "每股公积金", "每股未分配利润",
         "净利润同比增长", "总股本",
         "预案公告日", "股权登记日", "除权除息日", "方案进度", "最新公告日期"]),
    ("stock_lhb_jgmmtj_em(20240301~20240401)",
        lambda: sle.stock_lhb_jgmmtj_em(start_date="20240301",
                                         end_date="20240401"),
        # stockfetch 仅按名访问 `代码` 和 `买方机构数`
        None),
    ("stock_lhb_ggtj_sina(5)",
        lambda: sls.stock_lhb_ggtj_sina(recent_day="5"),
        ["股票代码", "股票名称", "上榜次数", "累积购买额",
         "累积卖出额", "净额", "买入席位数", "卖出席位数"]),
    ("stock_dzjy_mrtj(20240301)",
        lambda: sde.stock_dzjy_mrtj(start_date="20240301", end_date="20240301"),
        ["序号", "交易日期", "证券代码", "证券简称",
         "收盘价", "涨跌幅", "成交价", "折溢率",
         "成交笔数", "成交总量", "成交总额", "成交总额/流通市值"]),
]


def run():
    passed, failed, network_blocked = [], [], []
    for label, fn, expected in CASES:
        print(f"\n=== {label} ===")
        try:
            df = fn()
            cols = list(df.columns) if df is not None else []
            print(f"  返回 {len(df) if df is not None else 0} 行 × {len(cols)} 列")
            if expected is None:
                if df is None or len(df) == 0:
                    failed.append(f"{label}: 返回空")
                    print("  ⚠️ 返回空")
                else:
                    must_have = ["代码", "买方机构数"] if "jgmmtj" in label else []
                    miss = [c for c in must_have if c not in cols]
                    if miss:
                        failed.append(f"{label}: 缺列 {miss}")
                        print(f"  ❌ 缺列 {miss}")
                    else:
                        passed.append(label)
                        print(f"  ✅ 关键列存在: {must_have}")
            else:
                if cols[:len(expected)] != expected:
                    diff = [(i, e, cols[i] if i < len(cols) else "MISSING")
                            for i, e in enumerate(expected)
                            if i >= len(cols) or cols[i] != e]
                    failed.append(f"{label}: 列序不匹配 {diff[:3]}")
                    print(f"  ❌ 列序不匹配，前 3 个差异: {diff[:3]}")
                else:
                    passed.append(label)
                    print(f"  ✅ 列序匹配")
        except Exception as e:
            msg = f"{type(e).__name__}: {e}"
            # 网络层问题（被东财限流/封 IP）单独归类
            if any(s in msg.lower() for s in ("connection", "remotedisconnect",
                                                "timeout", "max retries")):
                network_blocked.append(f"{label}: {msg[:120]}")
                print(f"  🌐 上游网络受阻: {msg[:120]}")
            else:
                failed.append(f"{label}: {msg[:200]}")
                print(f"  ❌ 异常: {msg[:200]}")

    print("\n\n========== 总结 ==========")
    print(f"✅ PASS:            {len(passed)}/{len(CASES)}")
    print(f"❌ FAIL（契约/逻辑）: {len(failed)}")
    for f in failed:
        print(f"   - {f}")
    print(f"🌐 网络受阻：         {len(network_blocked)}")
    for n in network_blocked:
        print(f"   - {n}")
    # P0 验收标准：契约/逻辑 0 失败（网络问题不计入）
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(run())
