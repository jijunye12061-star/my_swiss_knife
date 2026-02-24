"""
估值计算模块
"""

class FundValuationCalculator:
    """基金估值计算器"""

    def __init__(self, holdings_data, prev_close_dict, stock_position_ratio):
        """
        :param holdings_data: DataFrame, 持仓数据
        :param prev_close_dict: dict, {股票代码: 昨收价}
        :param stock_position_ratio: float, 股票仓位占比(%)
        """
        self.holdings_data = holdings_data
        self.prev_close_dict = prev_close_dict
        self.stock_position_ratio = stock_position_ratio

        # 计算持仓占比缩放系数
        total_holdings = holdings_data['持仓占比'].sum()
        self.scale_factor = stock_position_ratio / total_holdings if total_holdings > 0 else 0

    def calculate_stock_return(self, stock_code, current_price):
        """
        计算个股涨跌幅
        :return: float, 涨跌幅(%)
        """
        prev_close = self.prev_close_dict.get(stock_code)
        if prev_close is None or prev_close == 0:
            return 0.0

        return (current_price - prev_close) / prev_close * 100

    def calculate_fund_valuation(self, intraday_klines):
        """
        计算基金日内估值走势
        :param intraday_klines: dict, {股票代码: [{time, price}, ...]}
        :return: list of dict, [{time, valuation_change}, ...]
        """
        # 1. 收集所有时间点
        all_times = set()
        for klines in intraday_klines.values():
            for kline in klines:
                all_times.add(kline['time'])

        all_times = sorted(list(all_times))

        # 2. 构建每只股票的时间->价格映射
        stock_price_map = {}
        for stock_code, klines in intraday_klines.items():
            price_map = {}
            last_price = self.prev_close_dict.get(stock_code, 0)

            for kline in klines:
                price_map[kline['time']] = kline['price']
                last_price = kline['price']

            # 对于缺失的时间点，用前值填充
            filled_prices = {}
            current_price = self.prev_close_dict.get(stock_code, 0)

            for time_point in all_times:
                if time_point in price_map:
                    current_price = price_map[time_point]
                filled_prices[time_point] = current_price

            stock_price_map[stock_code] = filled_prices

        # 3. 计算每个时间点的基金估值涨跌幅
        result = []
        for time_point in all_times:
            weighted_return = 0.0

            for _, row in self.holdings_data.iterrows():
                stock_code = row['股票代码']
                holding_ratio = row['持仓占比']

                # 获取该时间点的价格
                current_price = stock_price_map.get(stock_code, {}).get(time_point, 0)

                # 计算个股涨跌幅
                stock_return = self.calculate_stock_return(stock_code, current_price)

                # 加权累加（缩放后）
                weighted_return += stock_return * holding_ratio * self.scale_factor / 100

            result.append({
                'time': time_point,
                'valuation_change': round(weighted_return, 4)
            })

        return result

    def get_summary_stats(self, valuation_data):
        """
        获取汇总统计信息
        :param valuation_data: list, 估值数据
        :return: dict, 统计信息
        """
        if not valuation_data:
            return {}

        changes = [d['valuation_change'] for d in valuation_data]

        return {
            'current': changes[-1] if changes else 0,
            'max': max(changes) if changes else 0,
            'min': min(changes) if changes else 0,
            'max_time': valuation_data[changes.index(max(changes))]['time'] if changes else '',
            'min_time': valuation_data[changes.index(min(changes))]['time'] if changes else ''
        }