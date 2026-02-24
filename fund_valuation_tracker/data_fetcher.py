"""
数据获取模块
"""
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent.parent))

import requests
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
from utils.query_data_from_choice import fetcher


class FundDataFetcher:
    """基金数据获取器"""

    def __init__(self):
        self.base_url = "http://61.152.230.191/api/qt/stock/kline/get"
        self.fetcher = fetcher  # 使用全局fetcher实例

    def get_fund_holdings(self, fund_code, report_date):
        """获取基金持仓"""
        holdings = self.fetcher.query_stock_holdings([fund_code], report_date)
        return holdings

    def get_stock_position_ratio(self, fund_code, report_date):
        """获取基金股票仓位"""
        stock_pct_dict = {
            "codes": fund_code,
            "indicators": "PRTSTOCKTONAV",
            "options": f"ReportDate={report_date}"
        }
        stock_pct_data = self.fetcher.query_from_choice("css", stock_pct_dict)
        return stock_pct_data.Data[fund_code][0]

    def get_previous_close(self, stock_codes, trade_date):
        """获取股票昨收价"""
        prices = self.fetcher.get_stock_close(stock_codes, trade_date, trade_date, if_adj=False)
        price_dict = dict(zip(prices['股票代码'], prices['收盘价']))
        return price_dict

    def get_stock_intraday_kline(self, stock_code, date_str):
        """
        获取个股日内5分钟K线
        :param stock_code: 如 '002475.SZ'
        :param date_str: 如 '20260202'
        :return: list of dict [{time, price}, ...]
        """
        # 解析股票代码
        code, exchange = stock_code.split('.')
        market = '0' if exchange == 'SZ' else '1'
        secid = f"{market}.{code}"

        params = {
            'secid': secid,
            'klt': '5',
            'fqt': '0',
            'beg': date_str,
            'end': date_str,
            'fields1': 'f1,f2,f3,f4,f5',
            'fields2': 'f51,f52,f53'
        }

        try:
            response = requests.get(self.base_url, params=params, timeout=5)
            data = response.json()

            if data.get('rc') != 0 or not data.get('data', {}).get('klines'):
                return []

            klines = data['data']['klines']
            result = []

            for kline in klines:
                parts = kline.split(',')
                time_str = parts[0]  # '2026-02-02 09:35'
                price = float(parts[2])  # 收盘价
                result.append({
                    'time': time_str,
                    'price': price
                })

            return result

        except Exception as e:
            print(f"获取 {stock_code} K线失败: {e}")
            return []

    def batch_get_intraday_klines(self, stock_codes, date_str, max_workers=10):
        """
        并发获取多只股票的日内K线
        :return: dict {股票代码: [{time, price}, ...]}
        """
        result = {}

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_stock = {
                executor.submit(self.get_stock_intraday_kline, code, date_str): code
                for code in stock_codes
            }

            for future in as_completed(future_to_stock):
                stock_code = future_to_stock[future]
                try:
                    kline_data = future.result()
                    result[stock_code] = kline_data
                except Exception as e:
                    print(f"处理 {stock_code} 时出错: {e}")
                    result[stock_code] = []

        return result


def get_today_str():
    """获取当前日期字符串，格式: 20260202"""
    return datetime.now().strftime('%Y%m%d')


def get_last_trade_date():
    """获取上一交易日"""
    today = datetime.now()
    one_month_ago = today - timedelta(days=30)

    # 获取近一个月的交易日
    trading_dates = fetcher.get_trading_dates(
        one_month_ago.strftime('%Y-%m-%d'),
        today.strftime('%Y-%m-%d')
    )

    # 找到今天之前的最后一个交易日
    today_str = today.strftime('%Y-%m-%d')
    valid_dates = [d for d in trading_dates if d.strftime('%Y-%m-%d') < today_str]

    if valid_dates:
        return valid_dates[-1].strftime('%Y-%m-%d')
    else:
        # 兜底：返回昨天
        return (today - timedelta(days=1)).strftime('%Y-%m-%d')