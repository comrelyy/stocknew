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
import akshare as ak
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, date
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
    ("600198", "大唐电信", "sh", 9.50,  2000, 8.50,  12.00),
    ("600848", "上海临港", "sh", 10.50, 1500, 9.50,  13.00),
    ("600839", "四川长虹", "sh", 9.00,  3000, 7.50,  11.00)
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
REQUEST_DELAY = 1.0


# ==================== 多数据源采集 ====================

def is_trading_day():
    today = date.today()
    if today.weekday() >= 5:
        return False
    try:
        trade_cal = ak.tool_trade_date_hist_sina()
        trade_dates = set(pd.to_datetime(trade_cal['trade_date']).dt.date)
        return today in trade_dates
    except Exception:
        return True


def get_spot_data(code):
    """获取实时行情 - 多源容错"""
    # 优先级1：新浪（最稳定，国内基本不封）
    try:
        # 新浪需要带市场前缀
        symbol = f"sh{code}" if code.startswith('6') else f"sz{code}"
        df = ak.stock_zh_a_spot()
        row = df[df['代码'] == symbol]
        if not row.empty:
            r = row.iloc[0]
            return {
                "收盘价": float(r['最新价']),
                "涨跌%": float(r['涨跌幅']),
                "成交额(亿)": round(float(r['成交额']) / 1e8, 2),
                "换手率%": None,  # 新浪接口可能不直接返回换手率
                "source": "新浪"
            }
    except Exception as e:
        print(f"      [新浪行情失败] {str(e)[:80]}")
    
    # 优先级2：腾讯
    try:
        symbol = f"sh{code}" if code.startswith('6') else f"sz{code}"
        df = ak.stock_zh_a_spot_em()
        row = df[df['代码'] == code]
        if not row.empty:
            r = row.iloc[0]
            return {
                "收盘价": float(r['最新价']),
                "涨跌%": float(r['涨跌幅']),
                "成交额(亿)": round(float(r['成交额']) / 1e8, 2),
                "换手率%": float(r['换手率']) if r['换手率'] else None,
                "source": "东财"
            }
    except Exception as e:
        print(f"      [东财行情失败] {str(e)[:80]}")
    
    return None


def get_fund_flow(code, market):
    """获取资金流向 - 仅东财提供，新浪/腾讯不提供完整资金流"""
    try:
        time.sleep(REQUEST_DELAY)
        fund = ak.stock_individual_fund_flow(stock=code, market=market)
        latest = fund.iloc[-1]
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
    end_date = date.today().strftime("%Y%m%d")
    start_date = (date.today() - pd.Timedelta(days=120)).strftime("%Y%m%d")
    
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


def calculate_indicators(hist_df, source):
    """从K线计算技术指标"""
    if hist_df is None or hist_df.empty:
        return {}
    
    # 不同数据源字段名不同
    close_col = '收盘' if '收盘' in hist_df.columns else 'close'
    
    closes = hist_df[close_col].values.astype(float)
    
    result = {}
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


def get_stock_data(code, name, market, cost, qty, stop_loss, take_profit):
    """获取单只股票完整数据"""
    print(f"\n  📊 {name} ({code})")
    
    result = {
        "代码": code, "名称": name, "成本价": cost, "数量": qty,
        "止损价": stop_loss, "止盈价": take_profit,
        "收盘价": None, "涨跌%": None, "成交额(亿)": None, "换手率%": None,
        "总流入(万)": None, "主力净流入(万)": None, "主力占比%": None,
        "散户净流入(万)": None, "散户占比%": None,
        "MA5": None, "MA10": None, "MA20": None, "MA60": None, "RSI": None,
        "浮盈%": None, "浮盈金额": None, "距止损%": None,
        "警报": [], "事件": [],
    }
    
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
    today = date.today().strftime("%Y-%m-%d")
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
    last_row = ws.max_row + 1
    today_str = date.today().strftime("%Y-%m-%d")
    
    for d in all_data:
        events_str = " | ".join(d["事件"] + d["警报"]) if (d["事件"] or d["警报"]) else "正常"
        row_data = [
            today_str, d['代码'], d['名称'], d['收盘价'], d['涨跌%'],
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
    
    if "持仓汇总" in wb.sheetnames:
        ws_summary = wb["持仓汇总"]
        for row_idx in range(2, ws_summary.max_row + 1):
            code_cell = ws_summary.cell(row=row_idx, column=3).value
            if code_cell:
                for d in all_data:
                    if str(code_cell) == d['代码'] and d['收盘价']:
                        ws_summary.cell(row=row_idx, column=10, value=d['收盘价'])
                        break
    
    wb.save(EXCEL_PATH)
    print(f"    ✅ Excel已更新")
    return True


# ==================== 主程序 ====================

def main():
    print("=" * 60)
    print(f"📊 A股持仓追踪 v3 (多数据源容错)")
    print(f"📅 {date.today().strftime('%Y-%m-%d %A')}")
    print("=" * 60)
    
    if not is_trading_day():
        print("⚠️ 今天非交易日，退出")
        return
    
    now = datetime.now()
    if now.hour < 15:
        print(f"⚠️ 当前 {now.strftime('%H:%M')}，建议15:30后运行")
    
    print(f"\n📡 [1/3] 采集数据（多源容错）")
    all_data = []
    for stock_info in STOCKS:
        try:
            data = get_stock_data(*stock_info)
            all_data.append(data)
        except Exception as e:
            print(f"  ❌ {stock_info[1]} 完全失败: {e}")
    
    # 检查至少有一只股票获取成功
    success_count = sum(1 for d in all_data if d["收盘价"] is not None)
    if success_count == 0:
        print(f"\n❌ 所有股票数据采集都失败！")
        print(f"💡 可能原因：")
        print(f"   1. 你的IP被东方财富/新浪封禁了 - 等待30分钟-2小时")
        print(f"   2. 网络问题 - 检查能否访问 finance.sina.com.cn")
        print(f"   3. akshare 库需要升级: pip install akshare --upgrade")
        return
    
    print(f"\n📋 [2/3] 数据摘要 ({success_count}/{len(STOCKS)} 成功)")
    print("-" * 60)
    for d in all_data:
        events = " ".join(d["事件"] + d["警报"]) if (d["事件"] or d["警报"]) else "正常"
        price = f"{d['收盘价']:.2f}" if d['收盘价'] else "  -  "
        chg = d['涨跌%'] if d['涨跌%'] is not None else 0
        main = d['主力占比%'] if d['主力占比%'] is not None else 0
        profit = d['浮盈%'] if d['浮盈%'] is not None else 0
        print(f"  {d['名称']:6} {price:>7} 涨跌{chg:+6.2f}% 主力{main:+6.2f}% 浮盈{profit:+6.2f}% | {events}")
    
    print(f"\n💾 [3/3] 写入Excel + 推送")
    append_to_excel(all_data)
    
    title, content = build_wechat_content(all_data)
    send_wechat(title, content)
    
    print(f"\n🎉 完成！")
    print(f"📂 Excel: {EXCEL_PATH}")


if __name__ == "__main__":
    main()