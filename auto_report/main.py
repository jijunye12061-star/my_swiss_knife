import sys
from pathlib import Path

sys_path = Path("dolphinscheduler/default/resources/jjy/clients/sub_redeem_report")
sys.path.insert(0, str(sys_path))

import os
import calendar
from datetime import datetime

import pymysql
import pandas as pd

from report_generator import generate_monthly_report
from watermark import apply_watermark_and_protection
from llm_summary import build_llm_summary_generator

# ========== Doris连接信息 ==========
DORIS_CONFIG = {
    'host': '10.189.18.47',
    'port': 10096,
    'user': 'irdev',
    'password': 'oCrxPb0osif31%TGB',
    'database': 'tytdata',
}


def query_doris(sql: str, params=None) -> pd.DataFrame:
    conn = pymysql.connect(**DORIS_CONFIG)
    try:
        df = pd.read_sql(sql, conn, params=params)
        return df
    finally:
        conn.close()


def get_max_trade_date(date_str: str = None) -> str:
    if date_str is None:
        date_str = datetime.now().strftime('%Y-%m-%d')

    sql = """
        SELECT c_max_trade_date 
        FROM tytdata.tb_trade_calendar 
        WHERE c_date = %s
    """
    df = query_doris(sql, params=[date_str])

    if df.empty:
        raise ValueError(f"未找到 {date_str} 的交易日历数据")

    return df['c_max_trade_date'].iloc[0]


def main():
    # now = datetime.now()
    now = datetime(2026, 2, 28, 16, 40, 0)
    print(f"今天是{now}")

    # 自动获取最新交易日作为today
    trade_date = get_max_trade_date('2026-02-28')  # 返回如 '2026-02-27'
    today = str(trade_date).replace('-', '')  # '20260227'
    print(f"计算{today}的结果")
    # 自动计算当月起止日期
    first_day = now.replace(day=1)
    last_day = now.replace(day=calendar.monthrange(now.year, now.month)[1])
    START_DATE = first_day.strftime('%Y-%m-%d')
    END_DATE = last_day.strftime('%Y-%m-%d')

    # 自动生成年月标题，如 "2026年2月"
    title_month = f"{now.year}年{now.month}月"

    # 配置参数
    # INPUT_FILE = f'/tmp/基金交易市场动态表-基金标签-分机构{today}.xlsx'
    # from ttjj_common.ftp import ftp_download
    # ftp_download(
    #     f'/report/excel/dongtaiD/TY_TTJJ_ZHT/基金交易市场动态表-基金标签-分机构{today}.xlsx',
    #     INPUT_FILE,
    # )
    INPUT_FILE = r'./data/基金交易市场动态表-基金标签-分机构20260226.xlsx'
    WATERMARK_TEXT = '天天基金投研专用'
    PROTECTION_PASSWORD = None

    # ── 摘要控制 ──────────────────────────────────
    ADD_LLM_SUMMARY = True

    # 输出目录 & 文件
    # output_directory = "dolphinscheduler/default/resources/jjy/clients/sub_redeem_report/output"
    output_directory = r'./output'
    os.makedirs(output_directory, exist_ok=True)

    report_file = os.path.join(output_directory, f'{title_month}基金申赎报告.xlsx')
    final_file = os.path.join(output_directory, f'{title_month}月度基金申赎报告.xlsx')

    # 选择摘要生成器
    if ADD_LLM_SUMMARY is True:
        summary_gen = build_llm_summary_generator()
        print('使用LLM生成摘要')
    else:
        summary_gen = None
        print('跳过摘要生成')

    # Step 1: 生成报告
    generate_monthly_report(
        input_file=INPUT_FILE,
        output_file=report_file,
        start_date=START_DATE,
        end_date=END_DATE,
        add_trend_charts=True,
        summary_generator=summary_gen,
    )

    # Step 2: 添加水印 + 保护
    apply_watermark_and_protection(
        input_xlsx=report_file,
        output_xlsx=final_file,
        watermark_text=WATERMARK_TEXT,
        password=PROTECTION_PASSWORD,
    )

    # # Step 3: 推送咚咚
    # if final_file:
    #     from ttjj_common.dongdong import send_file
    #     send_file('g30306', final_file)
    #
    # # Step 4: 推送邮件
    # from ttjj_common.email import connect_email_server
    # import time
    #
    # time.sleep(2)
    #
    # sender = connect_email_server(timeout=120)
    # receivers = [
    #     "jijunye@eastmoney.com"
    # ]
    # subject = f"{Path(final_file).stem}"
    # body = "请查收本期基金申赎月度报告，详见附件。"
    #
    # attachments = {}
    # final_path = Path(final_file).resolve()
    # if final_path.exists():
    #     attachments[final_path.name] = final_path
    #
    # print(f"准备发送邮件至: {receivers}")
    #
    # for receiver in receivers:
    #     sender.send(
    #         receivers=[receiver],
    #         subject=subject,
    #         text=body,
    #         attachments=attachments,
    #     )
    # print("邮件发送完成")
    #
    # print(f'\n 最终报告: {final_file}')


if __name__ == '__main__':
    main()