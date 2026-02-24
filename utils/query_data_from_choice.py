from EmQuantAPI import *
import pandas as pd
import time
from functools import wraps
from datetime import datetime
from utils.constants import finance
from typing import List, Union
from dotenv import load_dotenv
from pathlib import Path
import os

# 加载.env文件
current_dir = Path(__file__).parent.parent
env_path = os.path.join(current_dir, '.env')
load_dotenv(env_path)


def rate_limit(max_requests=700, time_window=60):
    """
    装饰器：限制函数调用频率
    max_requests: 最大请求次数
    time_window: 时间窗口（秒）
    """

    def decorator(func):
        # 使用列表存储最近的调用时间
        calls = []

        @wraps(func)
        def wrapper(*args, **kwargs):
            now = datetime.now()

            # 移除时间窗口之外的调用记录
            calls[:] = [call for call in calls
                        if (now - call).total_seconds() < time_window]

            # 如果达到限制，则等待
            if len(calls) >= max_requests:
                sleep_time = time_window - (now - calls[0]).total_seconds()
                if sleep_time > 0:
                    print(f"Rate limit reached. Sleeping for {sleep_time:.2f} seconds...")
                    time.sleep(sleep_time)
                calls.clear()  # 清空计数

            # 记录这次调用
            calls.append(now)

            # 执行原函数
            return func(*args, **kwargs)

        return wrapper

    return decorator


class ChoiceDataUtils:
    """Choice数据查询工具类，目前支持如下函数：
    - query_from_choice: 从Choice接口查询数据，支持csd和css和edb和ctr和sector
    - get_fund_nav: 获取基金净值数据
    - get_target_fund_nav: 获取基金指定日期的净值数据
    - get_index_nav: 获取指数净值数据
    - get_trading_dates: 获取指定日期范围内的交易日期
    - get_stock_close: 获取股票收盘价数据
    """
    _instance = None
    _initialized = False

    def __new__(cls, username=None, password=None):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self, username=None, password=None):
        # 避免重复初始化
        if self._initialized:
            return

        self.username = username or os.getenv('DB_USERNAME')
        self.password = password or os.getenv('DB_PASSWORD')
        loginresult = c.start(f"username={self.username},"
                              f"password={self.password},ForceLogin=1")

        if loginresult.ErrorCode != 0:
            print("login in fail")
        else:
            ChoiceDataUtils._initialized = True

    @staticmethod
    @rate_limit(max_requests=650, time_window=60)
    def query_from_choice(query_type='csd', query_dict=None):
        """
        从Choice接口查询数据
        :param query_type: str，查询类型，csd或css
        :param query_dict: dict，查询参数字典，包含codes, indicators, start_date, end_date, options等参数
        :return: DataFrame，查询结果
        """
        if query_type == 'csd':
            data = c.csd(query_dict['codes'], query_dict['indicators'], query_dict['start_date'], query_dict['end_date'],
                         query_dict['options'])
        elif query_type == 'css':
            data = c.css(query_dict['codes'], query_dict['indicators'], query_dict['options'])
        elif query_type == 'edb':
            data = c.edb(query_dict['codes'], query_dict['options'])
        elif query_type == 'ctr':
            data = c.ctr(query_dict['codes'], query_dict['indicators'], query_dict['options'])
        elif query_type == 'sector':
            data = c.sector(query_dict['codes'], query_dict['tradedate'])
        else:
            raise ValueError("查询类型错误，目前只支持csd和css和edb和ctr和sector")
        return data

    @staticmethod
    def get_fund_nav(
            fund_codes: List[str],
            start_dt: Union[str, datetime, pd.Timestamp],
            end_dt: Union[str, datetime, pd.Timestamp],
            nav_type: str = 'adj',
            ret_flag: bool = False
    ) -> pd.DataFrame:
        """
        获取基金净值数据
        Args:
            fund_codes: 基金代码列表
            start_dt: 开始日期
            end_dt: 结束日期
            nav_type: 净值类型，可选值: 'adj'(复权净值), 'raw'(原始净值), 'acc'(累计净值)
            ret_flag: 是否返回收益率，默认为False
        Returns:
            DataFrame with columns: ['基金代码', '交易日期', '单位净值']
        Raises:
            ValueError: 当输入参数无效时
        """
        if not fund_codes:
            raise ValueError("基金代码列表不能为空")

        nav_indicators = finance.NAV_INDICATORS

        if nav_type not in nav_indicators:
            raise ValueError(f"nav_type必须是以下之一: {', '.join(nav_indicators.keys())}")

        # 统一日期格式
        start_dt = pd.to_datetime(start_dt).strftime('%Y-%m-%d')
        end_dt = pd.to_datetime(end_dt).strftime('%Y-%m-%d')

        # 获取字段
        query_fields = nav_indicators[nav_type][0]
        if ret_flag:
            query_fields += ',ADJUSTEDNAVRATE'

        try:
            nav_data = c.csd(
                ','.join(fund_codes),
                query_fields,
                start_dt,
                end_dt,
                "period=1,adjustflag=1,curtype=1,order=1,market=CNSESH,isPandas=1"
            )
        except Exception as e:
            raise RuntimeError(f"获取基金数据失败: {str(e)}")

        if nav_data.empty:
            return pd.DataFrame(columns=['基金代码', '交易日期', '单位净值'])

        # 数据处理
        nav_data.reset_index(inplace=True)
        nav_data = nav_data.rename(columns={
            'CODES': '基金代码',
            'DATES': '交易日期',
            nav_indicators[nav_type][0]: '单位净值',
            'ADJUSTEDNAVRATE': '复权净值收益率' if ret_flag else None
        })
        nav_data['交易日期'] = pd.to_datetime(nav_data['交易日期'])

        return nav_data

    @staticmethod
    def get_target_fund_nav(fund_code, trade_dt, nav_type='adj'):
        """
        获取基金指定日期的净值数据
        :param fund_code:  str，基金代码
        :param trade_dt:  str，交易日期，格式为 'YYYY-MM-DD'
        :param nav_type:  str，净值类型，可选值: 'adj'(复权净值), 'raw'(原始净值), 'acc'(累计净值)
        :return:  float，基金净值
        """
        nav_indicators = {
                'adj': ('NAVADJ', '复权净值'),
                'raw': ('NAVUNIT', '原始净值'),
                'acc': ('NAVACCUM', '累计净值')
            }
        data = c.css(fund_code, nav_indicators[nav_type][0], f"TradeDate={trade_dt}")

        nav = data.Data[fund_code][0]
        return nav

    @staticmethod
    def get_index_nav(index_code, start_dt, end_dt):
        """
        获取指数净值数据
        :param index_code: str，指数代码
        :param start_dt:  str，开始日期，格式为 'YYYY-MM-DD'
        :param end_dt:  str，结束日期，格式为 'YYYY-MM-DD'
        :return:  DataFrame，指数净值数据，为长格式，有指数代码列，交易日期列和指数净值列
        """
        if isinstance(index_code, str):
            index_code = [index_code]
        if isinstance(start_dt, pd.Timestamp):
            start_dt = start_dt.strftime('%Y-%m-%d')
        if isinstance(end_dt, pd.Timestamp):
            end_dt = end_dt.strftime('%Y-%m-%d')
        index_code_str = ','.join(index_code)
        index_nav = c.csd(index_code_str, "CLOSE", start_dt, end_dt,
                          "period=1,adjustflag=1,curtype=1,order=1,market=CNSESH,isPandas=1")
        index_nav.reset_index(inplace=True)
        index_nav.rename(columns={'CODES': '指数代码', 'DATES': '交易日期', 'CLOSE': '指数净值'}, inplace=True)
        index_nav['交易日期'] = pd.to_datetime(index_nav['交易日期'])
        return index_nav

    @staticmethod
    def get_trading_dates(start_date: str, end_date: str) -> pd.DatetimeIndex:
        """
        获取指定日期范围内的交易日期
        Args:
            start_date: 开始日期，格式为 'YYYY-MM-DD'
            end_date: 结束日期，格式为 'YYYY-MM-DD'
        Returns:
            pd.DatetimeIndex: 交易日期索引
        """
        data = c.tradedates(start_date, end_date, "period=1,order=1,market=CNSESH")
        return pd.DatetimeIndex(data.Data)

    @staticmethod
    def get_stock_close(stock_codes, start_dt, end_dt, if_adj=True, batch_size=100):
        """
        获取股票收盘价数据
        :param stock_codes: list，股票代码列表
        :param start_dt:  str，开始日期，格式为 'YYYY-MM-DD'
        :param end_dt:  str，结束日期，格式为 'YYYY-MM-DD'
        :param if_adj:  bool，是否复权，默认为 True
        :param batch_size: int，每次查询的股票数量
        :return:  DataFrame，股票收盘价数据，为长格式，有股票代码列，交易日期列和收盘价列
        """
        adj_flag = 2 if if_adj else 1
        stock_close_result = []
        for i in range(0, len(stock_codes), batch_size):
            stock_codes_batch = stock_codes[i:i + batch_size]
            stock_codes_str = ','.join(stock_codes_batch)

            stock_close = c.csd(stock_codes_str, "CLOSE", start_dt, end_dt,
                                f"period=1,adjustflag={adj_flag},curtype=1,order=1,market=CNSESH,isPandas=1")
            stock_close.reset_index(inplace=True)
            stock_close.rename(columns={'CODES': '股票代码', 'DATES': '交易日期', 'CLOSE': '收盘价'}, inplace=True)
            stock_close['交易日期'] = pd.to_datetime(stock_close['交易日期'])
            stock_close_result.append(stock_close)
            print(f"已获取{i + batch_size}/{len(stock_codes)}只股票的收盘价数据")
        stock_close = pd.concat(stock_close_result)
        return stock_close

    @staticmethod
    def get_sector_funds(sector_code: str, end_dt: str = '2024-09-30') -> pd.DataFrame:
        """
        获取某个板块的所有基金代码和信息
        Args:
            sector_code: 板块代码
            end_dt: 截止日期，格式 'YYYY-MM-DD'
        Returns:
            pd.DataFrame: 包含基金信息的数据框
        Raises:
            ValueError: 当API调用失败或数据格式错误时
        """
        # 列名映射字典
        COLUMN_MAPPING = {
            'CODES': '基金代码',
            'NAME': '基金名称',
            'STARTFUND': '是否初始基金',
            'FOUNDDATE': '成立日期',
            'MATURITYDATENEW': '终止日',
            'ISROF': '是否定开',
            'PRTNETASSET': '最新资产净值',
        }

        try:
            # 获取基金代码
            fund_codes = c.sector(sector_code, end_dt).Codes
            if not fund_codes:
                raise ValueError(f"未找到板块 {sector_code} 的基金")

            # 获取基金信息
            date_columns = ['FOUNDDATE', 'MATURITYDATENEW', 'SUSPENDDATE', 'RESUMEDATE']
            funds_info = c.css(
                ",".join(fund_codes),
                "NAME,STARTFUND,FOUNDDATE,MATURITYDATENEW,DELISTDATE,ISROF,SUSPENDDATE,RESUMEDATE,PRTNETASSET",
                f"TradeDate={end_dt},ReportDate={end_dt},isPandas=1"
            )

            funds_info.reset_index(inplace=True)

            # 一次性转换日期列
            for col in date_columns:
                if col in funds_info.columns:
                    funds_info[col] = pd.to_datetime(funds_info[col], errors='coerce')
            # 更新终止日期
            mask_rof = (funds_info['ISROF'] == '是') & (funds_info['SUSPENDDATE'].notna())
            mask_resume = (funds_info['RESUMEDATE'].notna()) & (funds_info['RESUMEDATE'] > funds_info['SUSPENDDATE'])

            funds_info.loc[mask_rof & ~mask_resume, 'MATURITYDATENEW'] = \
                funds_info.loc[mask_rof & ~mask_resume, 'SUSPENDDATE']
            # 重命名列
            funds_info.rename(columns=COLUMN_MAPPING, inplace=True)
            # 筛选初始基金
            result = funds_info[funds_info['是否初始基金'] == '是']

            return result[['基金代码', '基金名称', '成立日期', '终止日', '是否定开', '最新资产净值']].copy()

        except Exception as e:
            raise ValueError(f"获取基金信息失败: {str(e)}")

    @staticmethod
    def query_stock_holdings(fund_codes: List[str], report_date: str) -> pd.DataFrame:
        """查询多只基金单报告期的股票持仓数据
        返回的DataFrame包含以下列：
        - 基金代码
        - 报告日期
        - 股票代码
        - 持仓占比
        """
        columns_map = {
            'FUNDCODE': '基金代码',
            'REPORTDATE': '报告日期',
            'SECUCODE': '股票代码',
            'NETASSETRATIO': '持仓占比'
        }
        if report_date[-5:] in ['12-31', '06-30']:
            codes_info = 'FundIHolderStockDetailInfo'
        elif report_date[-5:] in ['03-31', '09-30']:
            codes_info = 'FundIHolderKeyStockDetailInfo'
        else:
            raise ValueError("报告期日期错误")
        query_dict = {'codes': codes_info,
                      'indicators': "FUNDCODE,REPORTDATE,SECUCODE,NETASSETRATIO"}
        # 预定义查询参数，避免重复创建
        base_options = f"ReportDate={report_date},isPandas=1"
        stock_holdings_list = []

        for fund_code in fund_codes:
            query_dict['options'] = f"FundCode={fund_code},{base_options}"
            df = fetcher.query_from_choice('ctr', query_dict)
            if not isinstance(df, pd.DataFrame):
                df = pd.DataFrame(columns=['FUNDCODE', 'REPORTDATE', 'SECUCODE', 'NETASSETRATIO'])

            # 链式处理，减少中间DataFrame
            stock_holdings = (df.rename(columns=columns_map)
                              .assign(
                报告日期=lambda x: pd.to_datetime(report_date),
                持仓占比=lambda x: pd.to_numeric(x['持仓占比'])
            ).fillna({'持仓占比': 0}))
            stock_holdings_list.append(stock_holdings)
        total_stock_holdings = pd.concat(stock_holdings_list, ignore_index=True)
        return total_stock_holdings


fetcher = ChoiceDataUtils()


if __name__ == '__main__':
    pass
    # a = fetcher.get_trading_dates('2024-12-01', '2024-12-31')
    fund_nav = fetcher.get_fund_nav(['000001.OF'], '2024-12-29', '2025-01-02', ret_flag=True)

    fund_holdings = fetcher.query_stock_holdings(["019829.OF"], "2025-12-31")

    # query_dict = {'codes': ','.join(['000001.OF', '000003.OF']),
    #               'indicators': "PRTSTOCKTONAV,PRTCONVERTIBLEBONDTONAV",
    #               'options': f'ReportDate=2024-09-30,isPandas=1'}
    # position_data = fetcher.query_from_choice('css', query_dict)

    # total_funds = fetcher.get_sector_funds('518001002002002')
    # total_funds.to_excel(r'C:\Users\Administrator\Desktop\total_funds.xlsx', index=False)
    # main_index_nav = fetcher.get_index_nav('CBA33601.CS', '2020-01-01', '2024-12-31')
    # 假如原来的命令是c.edb(codes_str, f"IsLatest=0,StartDate={start_date},EndDate={end_date},isPandas=1")
    # 现在可以改为如下方式调用
    # main_codes_str = ('EMM00087088,EMM00087089,EMM00087090,EMM00087091,EMM00087092,EMM00087093,EMM00087094,'
    #              'EMM00087095,EMM00087096,EMM00087097,EMM00087098,EMM00087099,EMM00087100,EMM00087101,')
    # main_query_dict = {'codes': main_codes_str,
    #               'options': f"IsLatest=0,StartDate=2022-01-01,EndDate=2022-01-31,isPandas=1"}
    # edb_data = fetcher.query_from_choice('edb', main_query_dict)
    # main_fund_codes = fetcher.get_sector_funds('518001002001001')
    # main_fund_info = fetcher.get_sector_funds('518001002001001')
    # fund_nav = get_target_fund_nav('000001.OF', '2020-01-01')
    # stock_close_df = get_stock_close(['600000.SH', '600004.SH'], '2020-01-01', '2020-12-31', if_adj=False)
    # main_fund_codes = ['000001', '000011']
    # main_start_date = '2020-01-01'
    # main_end_date = '2020-12-31'
    # main_fund_nav = get_fund_nav(main_fund_codes, main_start_date, main_end_date)
    # bf_codes = fetcher.get_sector_funds('518001002001002')
    # codes = ','.join(bf_codes['基金代码'].unique().tolist())
    # query_dict = {
    #     'codes': codes,
    #     'indicators': 'NAME,CORPMGRCOMPANY,FUNDMANAGER,MGRYEARS,MGRDATES,'
    #                   'FOUNDDATE,VALUESUM,EQUITYVALUE,PURCHSTATUS,MINHOLDPERIOD,ISROF,HOLDPCTN',
    #     'options': "Rank=1,ReportDate=2024-12-31,TRADEDATE=2025-03-11,isPandas=1"
    # }
    # data = fetcher.query_from_choice('css', query_dict)


