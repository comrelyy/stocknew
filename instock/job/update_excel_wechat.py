"""
A股持仓追踪 - 多数据源容错版（v3）
==================================================
更新内容：
  - 数据采集失败时自动切换数据源（东财 → 新浪 → 腾讯）
  - 即使东财IP被封也能正常运行
  - 增加请求间隔，避免被反爬

依赖安装：
    pip install akshare openpyxl pandas requests

配置：
  在下方 ===配置区=== 修改：
  - Excel路径
  - 股票列表（含成本价）
  - Server酱 SendKey
"""

import os
import re
import akshare as ak
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, date, timedelta
import sys
import requests
import time


# ==================== 配置区（请修改）====================

# 用脚本所在目录作为锚点，避免在不同 CWD 下执行时找不到文件
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(_SCRIPT_DIR, "A股持仓追踪模板_v2.xlsx")

# (代码, 名称, 市场, 成本价, 持仓数量, 止损价, 止盈价)
STOCKS = [
    ("002138", "顺络电子", "sz", 35.00, 1000, 33.00, 47.00),
    ("600198", "大唐电信", "sh", 9.50,  2000, 8.50,  10.00),
    ("600848", "上海临港", "sh", 10.50, 1500, 8.50,  16.00),
    ("600839", "四川长虹", "sh", 9.00,  3000, 8.50,  16.00),
    ("600481", "双良节能", "sh", 7.30,  200, 6.50,  9.30),
    ("002421", "达实智能", "sz", 4.27,  0, 3.50,  6.8),
    ("000066", "中国长城", "sz", 20.77,  0, 19.30,  25.70),
    ("000725", "京东方A", "sz", 4.69,  0, 3.50,  11.00),
    ("002703", "浙江世宝", "sz", 21.29,  0, 18.50,  31.00)
]

SERVERCHAN_KEY = "SCT349699TvpKVWLxyl9BnxjljvpO5MMr2"

ALERT_THRESHOLDS = {
    "price_drop_pct": -3.0,
    "price_rise_pct": 5.0,
    "main_flow_out_pct": -5.0,
    "main_flow_in_pct": 10.0,
    "turnover_high": 10.0,
    "near_stop_loss_pct": 3.0,
    "limit_up_down": 9.5,
}

# 请求间隔（秒）- 避免被反爬
REQUEST_DELAY = 5.0

# 目标日期 / 时间范围。命令行参数优先级高于此处；留空则取今天。
# 支持三种写法：
#   单日：  "2026-05-19"          → 含资金流，并推送微信
#   近N天： "前60天"/"近60天"/"60天"/"60"  → 每日价格+指标，仅推送最后一天
#   区间：  "2026-01-01到2026-04-01"      （也兼容 至 / ~ / : 分隔）
# 命令行示例：
#   python update_excel_wechat.py "前60天"
#   python update_excel_wechat.py 2026-01-01 2026-04-01
TARGET_DATE = "2026-05-22"

# 范围回补时，为算准首日 MA60/RSI 而向前多取的自然日数（垫底窗口）
RANGE_LOOKBACK_DAYS = 130


# ==================== 多数据源采集 ====================

# 运行时解析后的实际日期（date 对象）。单日模式=目标日期；范围模式=区间内最后一个交易日。
# 由 run_single()/run_range() 在运行时覆盖，用于微信推送标题与"持仓汇总"最新价。
_RUN_DATE = date.today()


def _parse_one_date(raw):
    """把 'yyyy-MM-dd' 解析为 date，失败抛 ValueError。"""
    return datetime.strptime(raw.strip(), "%Y-%m-%d").date()


def resolve_date_spec():
    """解析目标日期 / 范围。命令行参数 > 配置区 TARGET_DATE > 今天。
    返回 ('single', date) 或 ('range', start_date, end_date)。"""
    args = [a.strip() for a in sys.argv[1:] if a.strip()]
    # 命令行传两个日期 = 区间
    if len(args) >= 2:
        try:
            d1, d2 = _parse_one_date(args[0]), _parse_one_date(args[1])
            return ("range", min(d1, d2), max(d1, d2))
        except ValueError:
            print(f"⚠️ 区间参数格式错误：{args[:2]}，应为两个 yyyy-MM-dd，回退到今天")
            return ("single", date.today())
    raw = args[0] if args else TARGET_DATE
    return _parse_date_spec_string(raw)


def _parse_date_spec_string(raw):
    """把单个字符串解析为单日或范围。"""
    raw = (raw or "").strip()
    if not raw:
        return ("single", date.today())

    # 近N天："前60天" / "近60天" / "过去60日" / "60天" / "60"
    m = re.fullmatch(r"(?:前|近|过去|最近)?\s*(\d+)\s*(?:天|日)?", raw)
    if m:
        n = int(m.group(1))
        end = date.today()
        return ("range", end - timedelta(days=n), end)

    # 区间："2026-01-01到2026-04-01"（分隔符：到/至/~/～/:）
    parts = re.split(r"\s*(?:到|至|~|～|:)\s*", raw)
    if len(parts) == 2:
        try:
            d1, d2 = _parse_one_date(parts[0]), _parse_one_date(parts[1])
            return ("range", min(d1, d2), max(d1, d2))
        except ValueError:
            pass

    # 单日
    try:
        return ("single", _parse_one_date(raw))
    except ValueError:
        print(f"⚠️ 日期/范围无法识别：'{raw}'，回退到今天")
        return ("single", date.today())


def get_trading_days_in_range(start_date, end_date):
    """返回 [start, end] 内的交易日（升序）。取不到交易日历则回退为区间内的周一至周五。"""
    try:
        cal = ak.tool_trade_date_hist_sina()
        all_days = pd.to_datetime(cal['trade_date']).dt.date
        days = sorted(d for d in all_days if start_date <= d <= end_date)
        if days:
            return days
    except Exception as e:
        print(f"  [交易日历获取失败，回退为工作日] {str(e)[:60]}")
    days, cur = [], start_date
    while cur <= end_date:
        if cur.weekday() < 5:
            days.append(cur)
        cur += timedelta(days=1)
    return days


def is_trading_day():
    if _RUN_DATE.weekday() >= 5:
        return False
    try:
        trade_cal = ak.tool_trade_date_hist_sina()
        trade_dates = set(pd.to_datetime(trade_cal['trade_date']).dt.date)
        return _RUN_DATE in trade_dates
    except Exception:
        return True


# 模块级缓存：避免在每只股票的循环里重复下载全市场行情触发东财风控
_em_spot_cache = {'data': None, 'tried': False}


def _fetch_em_spot_cached():
    """一次性拉取东财全市场行情，所有股票复用同一份结果。"""
    if _em_spot_cache['tried']:
        return _em_spot_cache['data']
    _em_spot_cache['tried'] = True
    try:
        df = ak.stock_zh_a_spot_em()
        if df is not None and not df.empty:
            _em_spot_cache['data'] = df
            print(f"  ✓ 已缓存东财批量行情 ({len(df)} 只)")
            return df
    except Exception as e:
        print(f"  [东财批量失败] {str(e)[:80]}")
    return None


def get_spot_data(code):
    """获取行情 - 多源容错。目标日期为今天时用东财实时，否则从日线取对应日期。"""
    # 优先级1：东财批量实时行情（仅当目标日期是今天才有意义，实时接口拿不到历史）
    if _RUN_DATE == date.today():
        spot_df = _fetch_em_spot_cached()
        if spot_df is not None:
            try:
                row = spot_df[spot_df['代码'] == code]
                if not row.empty:
                    r = row.iloc[0]
                    turnover = float(r['换手率']) if pd.notna(r['换手率']) else None
                    return {
                        "收盘价": float(r['最新价']),
                        "涨跌%": float(r['涨跌幅']),
                        "成交额(亿)": round(float(r['成交额']) / 1e8, 2),
                        "换手率%": turnover,
                        "source": "东财"
                    }
            except Exception as e:
                print(f"      [东财解析失败] {str(e)[:80]}")

    # 优先级2：从新浪日线取目标日期那一行（每股一查，支持回补历史）
    try:
        time.sleep(REQUEST_DELAY)
        symbol = f"sh{code}" if code.startswith('6') else f"sz{code}"
        end_date = _RUN_DATE.strftime("%Y%m%d")
        start_date = (_RUN_DATE - pd.Timedelta(days=15)).strftime("%Y%m%d")
        df = ak.stock_zh_a_daily(symbol=symbol, start_date=start_date, end_date=end_date, adjust="")
        if df is not None and len(df) >= 2:
            last = df.iloc[-1]
            prev = df.iloc[-2]
            last_date = pd.to_datetime(last['date']).date()
            if last_date != _RUN_DATE:
                print(f"      [提示] {code} 无 {_RUN_DATE} 数据，使用最近交易日 {last_date}")
            close = float(last['close'])
            prev_close = float(prev['close'])
            change_pct = round((close - prev_close) / prev_close * 100, 2)
            amount = float(last['amount']) if 'amount' in df.columns else float(last['volume']) * close
            return {
                "收盘价": close,
                "涨跌%": change_pct,
                "成交额(亿)": round(amount / 1e8, 2),
                "换手率%": None,
                "source": "新浪日线"
            }
    except Exception as e:
        print(f"      [新浪日线推算失败] {str(e)[:80]}")

    return None


def get_fund_flow(code, market):
    """获取资金流向 - 仅东财提供，新浪/腾讯不提供完整资金流"""
    try:
        time.sleep(REQUEST_DELAY)
        fund = ak.stock_individual_fund_flow(stock=code, market=market)
        # 选取目标日期那一行；找不到则退回最近一行
        match = fund[pd.DatetimeIndex(fund['日期']).date == _RUN_DATE]
        latest = match.iloc[-1] if not match.empty else fund.iloc[-1]
        return {
            "主力净流入(万)": round(float(latest['主力净流入-净额']) / 1e4, 2),
            "主力占比%": round(float(latest['主力净流入-净占比']), 2),
            "散户净流入(万)": round(float(latest['小单净流入-净额']) / 1e4, 2),
            "散户占比%": round(float(latest['小单净流入-净占比']), 2),
            "总流入(万)": round((float(latest['主力净流入-净额']) +
                              float(latest['中单净流入-净额']) +
                              float(latest['小单净流入-净额'])) / 1e4, 2)
        }
    except Exception as e:
        print(f"      [资金流向失败] {str(e)[:80]}")
        return None


def get_hist_data(code):
    """获取历史K线 - 多源容错"""
    end_date = _RUN_DATE.strftime("%Y%m%d")
    start_date = (_RUN_DATE - pd.Timedelta(days=120)).strftime("%Y%m%d")
    
    # 优先级1：新浪（无频率限制）
    try:
        time.sleep(REQUEST_DELAY)
        symbol = f"sh{code}" if code.startswith('6') else f"sz{code}"
        df = ak.stock_zh_a_daily(symbol=symbol, start_date=start_date, end_date=end_date, adjust="qfq")
        if not df.empty:
            return df, "新浪"
    except Exception as e:
        print(f"      [新浪K线失败] {str(e)[:80]}")
    
    # 优先级2：东财
    try:
        time.sleep(REQUEST_DELAY)
        df = ak.stock_zh_a_hist(symbol=code, period="daily",
                                start_date=start_date, end_date=end_date, adjust="qfq")
        if not df.empty:
            return df, "东财"
    except Exception as e:
        print(f"      [东财K线失败] {str(e)[:80]}")
    
    return None, None


def _indicators_from_closes(closes):
    """从收盘价序列（截至某交易日）计算 MA5/10/20/60 与 RSI。"""
    result = {}
    closes = pd.Series(closes).astype(float).values
    if len(closes) == 0:
        return result
    if len(closes) >= 5:  result["MA5"]  = round(closes[-5:].mean(), 2)
    if len(closes) >= 10: result["MA10"] = round(closes[-10:].mean(), 2)
    if len(closes) >= 20: result["MA20"] = round(closes[-20:].mean(), 2)
    if len(closes) >= 60: result["MA60"] = round(closes[-60:].mean(), 2)
    if len(closes) >= 15:
        diff = pd.Series(closes).diff()
        up = diff.clip(lower=0).rolling(14).mean()
        down = -diff.clip(upper=0).rolling(14).mean()
        if down.iloc[-1] > 0:
            rsi = 100 - (100 / (1 + up.iloc[-1] / down.iloc[-1]))
            result["RSI"] = round(rsi, 2)
    return result


def calculate_indicators(hist_df, source):
    """从K线计算技术指标（单日模式用，取整段K线末尾）"""
    if hist_df is None or hist_df.empty:
        return {}
    # 不同数据源字段名不同
    close_col = '收盘' if '收盘' in hist_df.columns else 'close'
    return _indicators_from_closes(hist_df[close_col].values)


def _blank_result(code, name, the_date, cost, qty, stop_loss, take_profit):
    """构建一行的空白结果字典（单日与范围模式共用）。the_date 为该行对应的日期。"""
    return {
        "日期": the_date,
        "代码": code, "名称": name, "成本价": cost, "数量": qty,
        "止损价": stop_loss, "止盈价": take_profit,
        "收盘价": None, "涨跌%": None, "成交额(亿)": None, "换手率%": None,
        "总流入(万)": None, "主力净流入(万)": None, "主力占比%": None,
        "散户净流入(万)": None, "散户占比%": None,
        "MA5": None, "MA10": None, "MA20": None, "MA60": None, "RSI": None,
        "浮盈%": None, "浮盈金额": None, "距止损%": None,
        "警报": [], "事件": [],
    }


def finalize_metrics(result, cost, qty, stop_loss, take_profit):
    """根据收盘价等计算浮盈/距止损，并生成警报与事件（原地写入 result）。单日与范围模式共用。"""
    if result["收盘价"]:
        result["浮盈%"] = round((result["收盘价"] - cost) / cost * 100, 2)
        result["浮盈金额"] = round((result["收盘价"] - cost) * qty, 2)
        result["距止损%"] = round((result["收盘价"] - stop_loss) / result["收盘价"] * 100, 2)

    alerts, events = [], []
    if result["涨跌%"] is not None:
        if result["涨跌%"] >= ALERT_THRESHOLDS["limit_up_down"]:
            events.append("🟢涨停")
        elif result["涨跌%"] <= -ALERT_THRESHOLDS["limit_up_down"]:
            alerts.append("🔴跌停！")
        elif result["涨跌%"] <= ALERT_THRESHOLDS["price_drop_pct"]:
            alerts.append(f"⚠️大跌 {result['涨跌%']:.2f}%")
        elif result["涨跌%"] >= ALERT_THRESHOLDS["price_rise_pct"]:
            events.append(f"📈大涨 {result['涨跌%']:.2f}%")

    if result["主力占比%"] is not None:
        if result["主力占比%"] <= ALERT_THRESHOLDS["main_flow_out_pct"]:
            alerts.append(f"🚨主力净流出 {result['主力占比%']:.2f}%")
        elif result["主力占比%"] >= ALERT_THRESHOLDS["main_flow_in_pct"]:
            events.append(f"💰主力净流入 {result['主力占比%']:.2f}%")

    if result["换手率%"] is not None and result["换手率%"] > ALERT_THRESHOLDS["turnover_high"]:
        events.append(f"⚡高换手 {result['换手率%']:.2f}%")

    if result["距止损%"] is not None and 0 < result["距止损%"] < ALERT_THRESHOLDS["near_stop_loss_pct"]:
        alerts.append(f"🆘距止损仅 {result['距止损%']:.2f}%")

    if result["收盘价"] and result["收盘价"] <= stop_loss:
        alerts.append(f"🆘🆘跌破止损 {stop_loss}")

    if result["收盘价"] and result["收盘价"] >= take_profit:
        events.append(f"🎯到达止盈 {take_profit}")

    result["警报"] = alerts
    result["事件"] = events
    return result


def get_stock_data(code, name, market, cost, qty, stop_loss, take_profit):
    """获取单只股票完整数据（单日模式：行情 + 资金流 + 指标）"""
    print(f"\n  📊 {name} ({code})")

    result = _blank_result(code, name, _RUN_DATE, cost, qty, stop_loss, take_profit)

    # 1. 实时行情
    spot = get_spot_data(code)
    if spot:
        result.update({k: v for k, v in spot.items() if k != "source"})
        print(f"    ✓ 行情 ({spot['source']}): {result['收盘价']} ({result['涨跌%']:+.2f}%)")

    # 2. 资金流向（仅东财）
    fund = get_fund_flow(code, market)
    if fund:
        result.update(fund)
        print(f"    ✓ 资金: 主力 {result['主力占比%']:+.2f}%")

    # 3. 技术指标
    hist, source = get_hist_data(code)
    if hist is not None:
        indicators = calculate_indicators(hist, source)
        result.update(indicators)
        print(f"    ✓ 指标 ({source}): MA20={result.get('MA20')} RSI={result.get('RSI')}")

    # 4. 计算盈亏 + 警报
    finalize_metrics(result, cost, qty, stop_loss, take_profit)
    return result


# ==================== 范围回补：整段日线 → 按日展开 ====================

def _normalize_daily(df):
    """把新浪/东财的日线 DataFrame 归一化为 ['d'(date), 'close', 'amount']，按日期升序。"""
    if df is None or df.empty:
        return None
    cols = df.columns
    if '收盘' in cols:  # 东财 stock_zh_a_hist
        out = pd.DataFrame({
            'd': pd.to_datetime(df['日期']).dt.date,
            'close': df['收盘'].astype(float),
            'amount': df['成交额'].astype(float) if '成交额' in cols else pd.NA,
        })
    else:  # 新浪 stock_zh_a_daily
        amount = df['amount'].astype(float) if 'amount' in cols else df['volume'].astype(float) * df['close'].astype(float)
        out = pd.DataFrame({
            'd': pd.to_datetime(df['date']).dt.date,
            'close': df['close'].astype(float),
            'amount': amount,
        })
    return out.sort_values('d').reset_index(drop=True)


def get_daily_history(code, start_date, end_date):
    """拉取覆盖 [start-垫底窗口, end] 的整段日线，新浪→东财容错。
    返回 (raw_norm, qfq_norm, source)：raw 用于价格列（不复权=当日实际价），qfq 用于指标。"""
    fetch_start = (start_date - timedelta(days=RANGE_LOOKBACK_DAYS)).strftime("%Y%m%d")
    end = end_date.strftime("%Y%m%d")
    symbol = f"sh{code}" if code.startswith('6') else f"sz{code}"

    # 优先级1：新浪
    try:
        time.sleep(REQUEST_DELAY)
        raw = ak.stock_zh_a_daily(symbol=symbol, start_date=fetch_start, end_date=end, adjust="")
        time.sleep(REQUEST_DELAY)
        qfq = ak.stock_zh_a_daily(symbol=symbol, start_date=fetch_start, end_date=end, adjust="qfq")
        rn = _normalize_daily(raw)
        if rn is not None and not rn.empty:
            return rn, _normalize_daily(qfq), "新浪"
    except Exception as e:
        print(f"    [新浪日线失败] {str(e)[:80]}")

    # 优先级2：东财
    try:
        time.sleep(REQUEST_DELAY)
        raw = ak.stock_zh_a_hist(symbol=code, period="daily", start_date=fetch_start, end_date=end, adjust="")
        time.sleep(REQUEST_DELAY)
        qfq = ak.stock_zh_a_hist(symbol=code, period="daily", start_date=fetch_start, end_date=end, adjust="qfq")
        rn = _normalize_daily(raw)
        if rn is not None and not rn.empty:
            return rn, _normalize_daily(qfq), "东财"
    except Exception as e:
        print(f"    [东财日线失败] {str(e)[:80]}")

    return None, None, None


def build_range_rows(stock_info, raw_df, qfq_df, trading_days):
    """用整段日线为区间内每个交易日生成一行（价格 + 指标，无资金流）。"""
    code, name, market, cost, qty, stop_loss, take_profit = stock_info
    rows = []
    for d in trading_days:
        day = raw_df[raw_df['d'] == d]
        if day.empty:
            continue  # 当日停牌或无数据
        i = day.index[0]
        close = float(raw_df.at[i, 'close'])
        result = _blank_result(code, name, d, cost, qty, stop_loss, take_profit)
        result["收盘价"] = close

        # 涨跌%：与该日之前最近一个交易日相比
        prev = raw_df[raw_df['d'] < d]
        if not prev.empty:
            prev_close = float(prev.iloc[-1]['close'])
            if prev_close:
                result["涨跌%"] = round((close - prev_close) / prev_close * 100, 2)

        amt = raw_df.at[i, 'amount']
        if pd.notna(amt):
            result["成交额(亿)"] = round(float(amt) / 1e8, 2)

        # 指标：用 qfq 序列中截至该日的部分
        if qfq_df is not None and not qfq_df.empty:
            sl = qfq_df[qfq_df['d'] <= d]
            result.update(_indicators_from_closes(sl['close'].values))

        finalize_metrics(result, cost, qty, stop_loss, take_profit)
        rows.append(result)
    return rows


# ==================== 微信推送 ====================

def send_wechat(title, content):
    if SERVERCHAN_KEY == "你的SendKey":
        print("    ⏭️  跳过微信推送（未配置 SERVERCHAN_KEY）")
        return False
    url = f"https://sctapi.ftqq.com/{SERVERCHAN_KEY}.send"
    data = {"title": title, "desp": content}
    try:
        resp = requests.post(url, data=data, timeout=10)
        result = resp.json()
        if result.get("code") == 0:
            print("    ✅ 微信推送成功")
            return True
        else:
            print(f"    ⚠️ 微信推送失败：{result.get('message')}")
            return False
    except Exception as e:
        print(f"    ⚠️ 微信推送异常：{e}")
        return False


def build_wechat_content(all_data):
    today = _RUN_DATE.strftime("%Y-%m-%d")
    all_alerts, all_events = [], []
    total_profit, total_cost = 0, 0
    
    for d in all_data:
        if d["浮盈金额"]: total_profit += d["浮盈金额"]
        if d["成本价"] and d["数量"]: total_cost += d["成本价"] * d["数量"]
        for a in d["警报"]: all_alerts.append(f"**{d['名称']}**：{a}")
        for e in d["事件"]: all_events.append(f"**{d['名称']}**：{e}")
    
    if all_alerts:
        title = f"🚨 持仓警报 {today}"
    elif total_profit > 0:
        title = f"📈 持仓快讯 {today} +{total_profit:,.0f}"
    else:
        title = f"📊 持仓快讯 {today} {total_profit:,.0f}"
    
    profit_pct = (total_profit / total_cost * 100) if total_cost else 0
    content = f"## 📊 {today} 持仓快讯\n\n"
    content += f"**今日组合浮盈/亏：{total_profit:+,.0f} 元 ({profit_pct:+.2f}%)**\n\n"
    content += "### 📋 持仓快照\n\n"
    content += "| 股票 | 价格 | 涨跌 | 主力% | 浮盈% |\n"
    content += "|------|------|------|-------|-------|\n"
    for d in all_data:
        price = f"{d['收盘价']:.2f}" if d['收盘价'] else "-"
        chg = f"{d['涨跌%']:+.2f}%" if d['涨跌%'] is not None else "-"
        main = f"{d['主力占比%']:+.2f}%" if d['主力占比%'] is not None else "-"
        profit = f"{d['浮盈%']:+.2f}%" if d['浮盈%'] is not None else "-"
        content += f"| {d['名称']} | {price} | {chg} | {main} | {profit} |\n"
    
    if all_alerts:
        content += f"\n### 🚨 风险警报 ({len(all_alerts)}条)\n\n"
        for a in all_alerts: content += f"- {a}\n"
    if all_events:
        content += f"\n### 📌 关键事件 ({len(all_events)}条)\n\n"
        for e in all_events: content += f"- {e}\n"
    
    content += "\n---\n💡 上传Excel到Claude → 发送'今日复盘'获取完整分析"
    return title, content


# ==================== Excel更新 ====================

def append_to_excel(all_data):
    try:
        wb = load_workbook(EXCEL_PATH)
    except FileNotFoundError:
        print(f"❌ Excel未找到：{EXCEL_PATH}")
        return False
    
    if "每日行情" not in wb.sheetnames:
        print("❌ 找不到'每日行情'Sheet")
        return False
    
    ws = wb["每日行情"]

    # 幂等：收集已存在的 (日期, 代码)，跳过重复行（范围回补可安全重复运行）
    existing = set()
    for row_idx in range(2, ws.max_row + 1):
        d_cell = ws.cell(row=row_idx, column=1).value
        c_cell = ws.cell(row=row_idx, column=2).value
        if d_cell is not None and c_cell is not None:
            existing.add((str(d_cell)[:10], str(c_cell)))

    last_row = ws.max_row + 1
    written, skipped = 0, 0
    for d in all_data:
        date_str = d["日期"].strftime("%Y-%m-%d") if hasattr(d["日期"], "strftime") else str(d["日期"])
        key = (date_str, str(d['代码']))
        if key in existing:
            skipped += 1
            continue
        events_str = " | ".join(d["事件"] + d["警报"]) if (d["事件"] or d["警报"]) else "正常"
        row_data = [
            date_str, d['代码'], d['名称'], d['收盘价'], d['涨跌%'],
            d['成交额(亿)'], d['换手率%'], d['总流入(万)'],
            d['主力净流入(万)'], d['主力占比%'],
            d['散户净流入(万)'], d['散户占比%'],
            d['MA5'], d['MA10'], d['MA20'], d['MA60'], d['RSI'],
            events_str
        ]
        for col_idx, value in enumerate(row_data, 1):
            if value is not None and value != "":
                ws.cell(row=last_row, column=col_idx, value=value)
        last_row += 1
        written += 1
        existing.add(key)

    # 持仓汇总：用每只股票"最新一天"的收盘价更新（范围模式取区间最后一个有数据日）
    if "持仓汇总" in wb.sheetnames:
        latest_close = {}
        for d in all_data:
            if d['收盘价'] is None:
                continue
            cur = latest_close.get(d['代码'])
            if cur is None or d['日期'] > cur[0]:
                latest_close[d['代码']] = (d['日期'], d['收盘价'])
        ws_summary = wb["持仓汇总"]
        for row_idx in range(2, ws_summary.max_row + 1):
            code_cell = ws_summary.cell(row=row_idx, column=3).value
            if code_cell and str(code_cell) in latest_close:
                ws_summary.cell(row=row_idx, column=10, value=latest_close[str(code_cell)][1])

    wb.save(EXCEL_PATH)
    print(f"    ✅ Excel已更新（新增 {written} 行，跳过重复 {skipped} 行）")
    return True


# ==================== 主程序 ====================

def _print_summary_line(d):
    events = " ".join(d["事件"] + d["警报"]) if (d["事件"] or d["警报"]) else "正常"
    price = f"{d['收盘价']:.2f}" if d['收盘价'] else "  -  "
    chg = d['涨跌%'] if d['涨跌%'] is not None else 0
    main_pct = d['主力占比%'] if d['主力占比%'] is not None else 0
    profit = d['浮盈%'] if d['浮盈%'] is not None else 0
    print(f"  {d['名称']:6} {price:>7} 涨跌{chg:+6.2f}% 主力{main_pct:+6.2f}% 浮盈{profit:+6.2f}% | {events}")


def _print_all_failed():
    print(f"\n❌ 所有股票数据采集都失败！")
    print(f"💡 可能原因：")
    print(f"   1. 你的IP被东方财富/新浪封禁了 - 等待30分钟-2小时")
    print(f"   2. 网络问题 - 检查能否访问 finance.sina.com.cn")
    print(f"   3. akshare 库需要升级: pip install akshare --upgrade")


def run_single(target_date):
    """单日模式：行情 + 资金流 + 指标，并推送微信。"""
    global _RUN_DATE
    _RUN_DATE = target_date
    is_today = (_RUN_DATE == date.today())
    print(f"📅 {_RUN_DATE.strftime('%Y-%m-%d %A')}")
    if not is_today:
        print(f"📌 回补历史数据模式（目标日期：{_RUN_DATE}）")
    print("=" * 60)

    if not is_trading_day():
        print("⚠️ 目标日期非交易日，退出")
        return

    if is_today:
        now = datetime.now()
        if now.hour < 15:
            print(f"⚠️ 当前 {now.strftime('%H:%M')}，建议15:30后运行")

    print(f"\n📡 [1/3] 采集数据（多源容错）")
    all_data = []
    for stock_info in STOCKS:
        try:
            all_data.append(get_stock_data(*stock_info))
        except Exception as e:
            print(f"  ❌ {stock_info[1]} 完全失败: {e}")

    success_count = sum(1 for d in all_data if d["收盘价"] is not None)
    if success_count == 0:
        _print_all_failed()
        return

    print(f"\n📋 [2/3] 数据摘要 ({success_count}/{len(STOCKS)} 成功)")
    print("-" * 60)
    for d in all_data:
        _print_summary_line(d)

    print(f"\n💾 [3/3] 写入Excel + 推送")
    append_to_excel(all_data)
    title, content = build_wechat_content(all_data)
    send_wechat(title, content)

    print(f"\n🎉 完成！")
    print(f"📂 Excel: {EXCEL_PATH}")


def run_range(start_date, end_date):
    """范围模式：逐只股票拉整段日线，按交易日展开（价格 + 指标），仅推送最后一天。"""
    global _RUN_DATE
    print(f"📌 范围回补模式：{start_date} → {end_date}")
    print("=" * 60)

    trading_days = [d for d in get_trading_days_in_range(start_date, end_date) if d <= date.today()]
    if not trading_days:
        print("⚠️ 区间内没有可用交易日（或都晚于今天），退出")
        return
    _RUN_DATE = trading_days[-1]  # 推送与"持仓汇总"以最后一个交易日为准
    print(f"📅 区间内 {len(trading_days)} 个交易日：{trading_days[0]} ~ {trading_days[-1]}")

    print(f"\n📡 [1/3] 逐只股票拉取日线并按日展开")
    all_data = []
    for stock_info in STOCKS:
        code, name = stock_info[0], stock_info[1]
        print(f"\n  📊 {name} ({code})")
        try:
            raw_df, qfq_df, source = get_daily_history(code, start_date, end_date)
            if raw_df is None or raw_df.empty:
                print(f"    ❌ 无日线数据")
                continue
            rows = build_range_rows(stock_info, raw_df, qfq_df, trading_days)
            print(f"    ✓ 日线({source})：展开 {len(rows)} 个交易日")
            all_data.extend(rows)
        except Exception as e:
            print(f"    ❌ {name} 失败: {str(e)[:80]}")

    if not all_data:
        _print_all_failed()
        return

    # 排序：先按日期，再按股票在 STOCKS 中的原始顺序
    order = {s[0]: i for i, s in enumerate(STOCKS)}
    all_data.sort(key=lambda d: (d["日期"], order.get(d["代码"], 999)))

    n_stocks = len(set(d['代码'] for d in all_data))
    print(f"\n📋 [2/3] 数据摘要（共 {len(all_data)} 行，{n_stocks} 只股票）")
    print("-" * 60)
    last_rows = [d for d in all_data if d["日期"] == _RUN_DATE]
    print(f"  最后交易日 {_RUN_DATE}：")
    for d in last_rows:
        _print_summary_line(d)

    print(f"\n💾 [3/3] 写入Excel + 推送最后一天")
    append_to_excel(all_data)
    if last_rows:
        title, content = build_wechat_content(last_rows)
        send_wechat(title, content)
    else:
        print("    ⏭️  最后交易日无数据，跳过推送")

    print(f"\n🎉 完成！")
    print(f"📂 Excel: {EXCEL_PATH}")


def main():
    spec = resolve_date_spec()
    print("=" * 60)
    print(f"📊 A股持仓追踪 v3 (多数据源容错)")
    if spec[0] == "single":
        run_single(spec[1])
    else:
        run_range(spec[1], spec[2])


if __name__ == "__main__":
    main()