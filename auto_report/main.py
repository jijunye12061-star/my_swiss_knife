"""
主入口：生成基金申赎月度报告
  1. 生成汇总表 + 趋势图
  2. 插入AI摘要（可选）
  3. 添加水印 + 工作表保护
"""

from report_generator import generate_monthly_report
from watermark import apply_watermark_and_protection


# ── 示例：硬编码摘要（后续替换为LLM调用）──────────────

def example_summary_generator(df_summary):
    """
    示例摘要生成器 - 演示接口格式。
    实际使用时替换为LLM调用，例如：

        def llm_summary_generator(df_summary):
            prompt = build_prompt(df_summary)
            response = call_llm(prompt)
            return {
                'overall': response['overall'],
                'institutions': response['institutions'],
            }
    """
    overall = (
        "1、一月份保险资金在平台净申购量最大（69.99），其次为券商（59.80）、公募（19.97）；\n"
        "2、一月份银行自营大幅净赎回（-75.08），信托期货（-24.81）、理财子（-22.54）也保持净赎回状态；\n"
        "3、市场呈现明显的\"固收+化\"趋势，固收+类基金净申购最多（180.21），而纯债类基金大幅净赎回（-208.31）；\n"
        "4、权益类基金整体净申购（50.78），交易较为活跃但规模明显低于固收+。"
    )

    institutions = {
        '保险': (
            "1、保险总体净申购规模较大（69.99），配置意愿强烈；\n"
            "2、在权益类基金保持较大净申购态势（34.96），偏好大盘成长型（9.16）和中盘成长型（8.89）；\n"
            "3、在固收+类基金保持净申购（28.23），重点配置相对收益型（15.64）；"
        ),
        '券商': (
            "1、券商总体净申购规模较大（59.80），配置意愿强烈；\n"
            "2、在权益类基金保持稳定净申购态势（12.96），偏好大盘成长型（4.63）和增强指数型（2.42）；\n"
            "3、在固收+类基金加仓力度最大（75.02），重点配置相对收益型（55.56）；\n"
            "4、对纯债类基金保持净赎回态势（-29.71）；"
        ),
        '信托期货': (
            "1、信托期货总体净赎回（-24.81），主要来自纯债类基金的大幅赎回（-62.61）；\n"
            "2、在固收+类基金保持净申购（37.39），配置力度较大；\n"
            "3、权益类基金净申购规模较小（1.32）；"
        ),
        '公募': (
            "1、公募总体净申购（19.97），以固收+类为主要配置方向；\n"
            "2、固收+类基金净申购29.63，占总量比重最大；\n"
            "3、纯债类基金净赎回（-8.55），权益类小幅净申购（0.47）；"
        ),
        '理财子公司': (
            "1、理财子总体净赎回（-22.54），主要因纯债类大幅净赎回（-34.39）；\n"
            "2、在固收+类基金保持净申购（9.69）；\n"
            "3、权益类基金净申购规模较小（0.81）；"
        ),
        '银行自营': (
            "1、银行自营大幅净赎回（-75.08），为所有机构中赎回量最大；\n"
            "2、赎回集中在纯债类基金（-75.18），偏利率型赎回最多（-44.15）；\n"
            "3、其他类型基金交易量极小；"
        ),
        '私募': (
            "1、私募总体小幅净赎回（-1.89），交易量较小；\n"
            "2、纯债类小幅净赎回（-2.75），在香港互认基金净申购（1.33）；"
        ),
        '银行资管': (
            "1、银行资管总体小幅净赎回（-0.80），交易量较小；\n"
            "2、纯债类小幅净赎回（-1.41），固收+类小幅净申购（0.62）；"
        ),
        '其他': (
            "1、其他机构总体小幅净赎回（-0.27），交易量极小；"
        ),
    }

    return {
        'overall': overall,
        'institutions': institutions,
    }


# ── LLM摘要生成器模板 ────────────────────────────────

def llm_summary_generator(df_summary):
    """
    LLM摘要生成器模板 - 将df传给大模型，返回结构化摘要。

    使用方式：
        1. 实现 call_llm() 函数（调用API）
        2. 在 main() 中将 summary_generator 替换为此函数

    def call_llm(prompt):
        # 调用你的LLM API
        import openai  # 或其他SDK
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}],
        )
        return response.choices[0].message.content
    """
    # 构建prompt
    summary_text = df_summary.to_string(index=False)
    prompt = f"""请基于以下基金申赎月度汇总数据，生成分析摘要。

数据如下：
{summary_text}

请按以下JSON格式输出：
{{
    "overall": "整体情况分析，每条以序号开头，用换行符分隔",
    "institutions": {{
        "机构名称": "该机构的行为分析，每条以序号开头，用换行符分隔",
        ...
    }}
}}

要求：
1. overall部分3-5条，概括全市场趋势
2. institutions部分覆盖所有机构，每个2-4条
3. 数字保留两位小数，用括号标注
4. 语言简练专业
"""
    # response = call_llm(prompt)
    # return json.loads(response)

    # 暂时返回None，表示未实现
    return None


# ── 主流程 ────────────────────────────────────────────

def main():
    # 配置参数
    INPUT_FILE = r'./data/基金交易市场动态表-基金标签-分机构20260130.xlsx'
    START_DATE = '2026-01-01'
    END_DATE = '2026-01-31'
    WATERMARK_TEXT = '平安基金MOM专用'
    PROTECTION_PASSWORD = None  # 设为字符串即启用密码保护，如 'abc123'

    # 中间文件和最终文件
    report_file = r'./data/2026年1月基金申赎报告.xlsx'
    final_file = r'./data/2026年1月基金申赎报告-终版.xlsx'

    # Step 1: 生成报告（含汇总表 + 摘要 + 趋势图）
    generate_monthly_report(
        input_file=INPUT_FILE,
        output_file=report_file,
        start_date=START_DATE,
        end_date=END_DATE,
        add_trend_charts=True,
        summary_generator=example_summary_generator,  # 替换为 llm_summary_generator
    )

    # Step 2: 添加水印 + 保护
    apply_watermark_and_protection(
        input_xlsx=report_file,
        output_xlsx=final_file,
        watermark_text=WATERMARK_TEXT,
        password=PROTECTION_PASSWORD,
    )

    print(f'\n🎉 最终报告: {final_file}')


if __name__ == '__main__':
    main()