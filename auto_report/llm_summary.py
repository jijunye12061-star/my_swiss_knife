"""
LLM 摘要生成模块
- 调用 DeepSeek-R1 生成基金申赎数据的结构化分析摘要
- 健壮的 JSON 解析（兼容 <think> 标签、Markdown fence 等）
"""

import json
import re
from openai import OpenAI
from dotenv import load_dotenv
from pathlib import Path
import os

# ── 加载环境变量 ──────────────────────────────────────
_env_path = Path(__file__).parent.parent / '.env'
load_dotenv(_env_path)

# ── 客户端配置 ────────────────────────────────────────
_DEFAULT_API_KEY = os.getenv('LLM_API_KEY', 'VOplCUjdbDBjO1Zf4f2eE5CcBd244835Ad31D5F6Ab7699F9')
_DEFAULT_BASE_URL = os.getenv('LLM_BASE_URL', 'https://dd-ai-api.eastmoney.com/v1')
_DEFAULT_MODEL = os.getenv('LLM_MODEL', 'DeepSeek-R1')


def _get_client(api_key=None, base_url=None):
    return OpenAI(
        api_key=api_key or _DEFAULT_API_KEY,
        base_url=base_url or _DEFAULT_BASE_URL,
    )


# ── Prompt ────────────────────────────────────────────

SYSTEM_PROMPT = """\
你是一名专业的基金销售数据分析师，擅长分析机构资金申赎行为和市场趋势。
请用简练、专业的中文输出分析。

严格要求：
- 只输出合法JSON，不要有任何额外文字、注释、Markdown格式或代码块标记。
- 不要输出```json或```等标记。"""

USER_PROMPT_TEMPLATE = """\
请分析以下基金申赎月度汇总数据（单位：亿元），生成结构化摘要。

数据：
{data}

输出格式（严格JSON）：
{{
    "overall": "多行文本，用\\n分隔，每条以数字序号开头",
    "institutions": {{
        "机构名1": "多行文本",
        "机构名2": "多行文本",
        ...
    }}
}}

分析规范：
1. overall（3-5条）：
   - 指出净申购/净赎回规模最大的机构类型及金额
   - 指出最受青睐和被大幅赎回的基金大类及金额
   - 总结市场整体趋势特征
   - 金额保留两位小数，用括号标注，如（69.95）；负值表示净赎回，如（-75.00）

2. institutions（覆盖所有机构，每个2-4条）：
   - 第一条说明该机构总体净申购/净赎回方向及金额
   - 后续条目说明在各大类基金（权益、纯债、固收+、QDII、香港互认）的配置行为
   - 提及具体子类偏好时引用子类名称和金额，如"偏好大盘成长型（9.16）和中盘成长型（8.89）"
   - 每条以序号开头，用分号结尾
   - 交易量极小的机构可只写1-2条

3. institutions的键名必须与数据列名中"净申赎"前的机构名完全一致（如列名为"保险净申赎"则键名为"保险"）"""


# ── JSON 解析 ─────────────────────────────────────────

def _extract_json(raw: str) -> dict:
    """
    从 LLM 原始输出中提取 JSON，兼容：
    - DeepSeek-R1 的 <think>...</think> 前缀
    - Markdown ```json ... ``` 包裹
    - 前后多余文字
    """
    text = raw

    # 1. 去掉 <think> 块
    if '</think>' in text:
        text = text.split('</think>')[-1].strip()

    # 2. 去掉 Markdown fence
    text = re.sub(r'^```(?:json)?\s*', '', text.strip())
    text = re.sub(r'\s*```$', '', text.strip())

    # 3. 直接尝试解析
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # 4. 提取第一个 { ... } 块（贪婪匹配最外层大括号）
    match = re.search(r'\{[\s\S]*\}', text)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass

    raise ValueError(f"无法从LLM输出中解析JSON，原始输出前200字符: {raw[:200]}")


def _validate_summary(result: dict) -> dict:
    """校验并规范化摘要结构"""
    if not isinstance(result, dict):
        raise ValueError("LLM输出不是字典")

    overall = result.get('overall', '')
    institutions = result.get('institutions', {})

    if not overall:
        raise ValueError("LLM输出缺少 overall 字段")
    if not isinstance(institutions, dict) or not institutions:
        raise ValueError("LLM输出缺少 institutions 字段或为空")

    return {
        'overall': overall.strip(),
        'institutions': {k: v.strip() for k, v in institutions.items()},
    }


# ── 核心生成函数 ──────────────────────────────────────

def generate_summary(df_summary, model=None, api_key=None, base_url=None, max_retries=2):
    """
    调用 LLM 生成摘要，可直接作为 summary_generator 传入 generate_monthly_report()。

    Args:
        df_summary: pd.DataFrame - 汇总表
        model: 模型名，默认读取环境变量或 DeepSeek-R1
        max_retries: 失败重试次数

    Returns:
        dict: {'overall': str, 'institutions': {str: str}} 或 None
    """
    client = _get_client(api_key, base_url)
    model = model or _DEFAULT_MODEL
    data_str = df_summary.to_csv(index=False)
    prompt_text = USER_PROMPT_TEMPLATE.format(data=data_str)

    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            print(f'  [LLM] 第{attempt}次调用（模型: {model}）...')
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": prompt_text},
                ],
                max_tokens=2000,
            )
            raw = response.choices[0].message.content
            result = _extract_json(raw)
            summary = _validate_summary(result)
            print(f'  [LLM] 摘要生成成功，机构数: {len(summary["institutions"])}')
            return summary

        except Exception as e:
            last_error = e
            print(f'  [LLM] 第{attempt}次失败: {e}')

    print(f'  [LLM] 全部{max_retries}次尝试失败，最后错误: {last_error}')
    return None


def build_llm_summary_generator(model=None, api_key=None, base_url=None, max_retries=2):
    """
    返回符合 generate_monthly_report(summary_generator=...) 签名的函数。

    用法：
        generate_monthly_report(
            ...,
            summary_generator=build_llm_summary_generator(),
        )
    """
    def _generator(df_summary):
        return generate_summary(
            df_summary,
            model=model,
            api_key=api_key,
            base_url=base_url,
            max_retries=max_retries,
        )
    return _generator