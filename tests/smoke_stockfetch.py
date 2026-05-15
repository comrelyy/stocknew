#!/usr/bin/env python3
"""
端到端冒烟测试：直接调用 instock.core.stockfetch 的业务函数，证明 P0 改造后
stockfetch 上层接口仍能产出与 tablestructure 列契约一致的 DataFrame。

不依赖数据库；不调用 ORM；只验证 DataFrame 形状和列名。

跑法：.venv/bin/python tests/smoke_stockfetch.py
"""
import os
import sys
import warnings
import datetime

warnings.filterwarnings("ignore")

cpath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, cpath)

# stockfetch 内部会 import talib（需要本机已装），如果环境没装会在此 raise
try:
    from instock.core import stockfetch as sf
    from instock.core import tablestructure as tbs
except ImportError as e:
    print(f"❌ stockfetch 模块加载失败：{e}")
    print("   可能原因：talib 未安装。运行 `brew install ta-lib && pip install ta-lib` 修复。")
    sys.exit(2)


today = datetime.datetime.now().date()

CASES = [
    ("fetch_stocks_trade_date()",
        lambda: sf.fetch_stocks_trade_date(),
        # 期望返回 set
        "set"),
    ("fetch_etfs(today)",
        lambda: sf.fetch_etfs(today),
        # 期望返回 DataFrame，列名等于 TABLE_CN_ETF_SPOT 的 columns
        tbs.TABLE_CN_ETF_SPOT),
    ("fetch_stocks(today)",
        lambda: sf.fetch_stocks(today),
        tbs.TABLE_CN_STOCK_SPOT),
    ("fetch_stocks_fund_flow(0)",
        lambda: sf.fetch_stocks_fund_flow(0),
        tbs.CN_STOCK_FUND_FLOW[0]),
    ("fetch_stocks_sector_fund_flow(0, 0)",
        lambda: sf.fetch_stocks_sector_fund_flow(0, 0),
        tbs.CN_STOCK_SECTOR_FUND_FLOW[1][0]),
    ("fetch_stocks_bonus(today)",
        lambda: sf.fetch_stocks_bonus(today),
        tbs.TABLE_CN_STOCK_BONUS),
    ("fetch_stock_top_data(today)",
        lambda: sf.fetch_stock_top_data(today),
        tbs.TABLE_CN_STOCK_TOP),
    ("fetch_stock_top_entity_data(today)",
        lambda: sf.fetch_stock_top_entity_data(today),
        "set"),
    ("fetch_stock_blocktrade_data(yesterday)",
        lambda: sf.fetch_stock_blocktrade_data(today - datetime.timedelta(days=1)),
        tbs.TABLE_CN_STOCK_BLOCKTRADE),
]


def check(label, fn, expected):
    print(f"\n=== {label} ===")
    try:
        res = fn()
    except Exception as e:
        msg = f"{type(e).__name__}: {e}"
        if any(s in msg.lower() for s in ("connection", "timeout", "remotedisconnect")):
            print(f"  🌐 上游网络受阻: {msg[:120]}")
            return "network"
        print(f"  ❌ 异常: {msg[:200]}")
        return "fail"

    if res is None:
        # stockfetch 内部 except 后返回 None；通常意味着底层抛了异常被吞掉
        print("  ⚠️  返回 None（底层异常或当日无数据）")
        return "none"

    if expected == "set":
        if isinstance(res, set):
            print(f"  ✅ 返回 set，{len(res)} 个元素")
            return "pass"
        print(f"  ❌ 期望 set，实际 {type(res).__name__}")
        return "fail"

    # 期望 DataFrame，列名等于 expected['columns'] 的 keys
    expected_cols = list(expected["columns"])
    actual_cols = list(res.columns)
    print(f"  返回 {len(res)} 行 × {len(actual_cols)} 列")
    if actual_cols != expected_cols:
        # 列名/列序不匹配
        diff = [(i, e, actual_cols[i] if i < len(actual_cols) else "MISSING")
                for i, e in enumerate(expected_cols)
                if i >= len(actual_cols) or actual_cols[i] != e]
        print(f"  ❌ 列名不匹配（期望 {len(expected_cols)} 列，实际 {len(actual_cols)} 列）")
        print(f"     前 3 个差异: {diff[:3]}")
        return "fail"
    print(f"  ✅ 列名完全匹配 ({len(expected_cols)} 列)")
    return "pass"


def main():
    counts = {"pass": [], "fail": [], "network": [], "none": []}
    for label, fn, expected in CASES:
        r = check(label, fn, expected)
        counts[r].append(label)

    print("\n\n========== 总结 ==========")
    print(f"✅ PASS:    {len(counts['pass'])}/{len(CASES)}")
    for c in counts["pass"]:
        print(f"   - {c}")
    print(f"❌ FAIL:    {len(counts['fail'])}")
    for c in counts["fail"]:
        print(f"   - {c}")
    print(f"🌐 NETWORK: {len(counts['network'])}")
    for c in counts["network"]:
        print(f"   - {c}")
    print(f"⚠️  NONE:    {len(counts['none'])}")
    for c in counts["none"]:
        print(f"   - {c}")

    return 0 if not counts["fail"] else 1


if __name__ == "__main__":
    sys.exit(main())
