"""
基金日内估值追踪 Flask 应用
"""
import sys
import json
import threading
from pathlib import Path
sys.path.append(str(Path(__file__).parent.parent))

from flask import Flask, render_template, request, jsonify
from data_fetcher import FundDataFetcher, get_today_str, get_last_trade_date
from calculator import FundValuationCalculator
import traceback

app = Flask(__name__)
fetcher = FundDataFetcher()

# ─── 基金列表缓存 ─────────────────────────────────────────────
_cache_lock = threading.Lock()
_fund_list_cache = None
_fund_code_to_init = {}


def load_fund_list():
    """加载基金列表，带线程安全缓存；同时构建 code→init_code 映射"""
    global _fund_list_cache, _fund_code_to_init
    if _fund_list_cache is not None:
        return _fund_list_cache
    with _cache_lock:
        if _fund_list_cache is not None:  # double-check
            return _fund_list_cache
        fund_list_path = Path(__file__).parent / 'static' / 'fund_list.json'
        data = []
        if fund_list_path.exists():
            with open(fund_list_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        _fund_code_to_init = {
            fund['code']: fund.get('init_code', fund['code'])
            for fund in data
        }
        _fund_list_cache = data
    return _fund_list_cache


def resolve_init_code(fund_code: str) -> str:
    """
    给定基金代码（可能是非主基金），返回对应的主基金代码(init_code)。
    若在列表中找不到，尝试补全 .OF 后缀再查，最终兜底返回原代码（带后缀）。
    """
    load_fund_list()
    if fund_code in _fund_code_to_init:
        return _fund_code_to_init[fund_code]
    candidate = fund_code if '.' in fund_code else fund_code + '.OF'
    return _fund_code_to_init.get(candidate, candidate)


# ─── 路由 ─────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/fund_search', methods=['GET'])
def fund_search():
    """基金搜索接口，支持代码和简称模糊匹配"""
    query = request.args.get('q', '').strip()
    if not query:
        return jsonify([])

    fund_list = load_fund_list()
    query_lower = query.lower()

    results = []
    for fund in fund_list:
        code = fund.get('code', '')
        name = fund.get('name', '')
        if query_lower in code.lower() or query_lower in name.lower():
            results.append({'code': code, 'name': name})
        if len(results) >= 10:
            break

    return jsonify(results)


@app.route('/api/valuation', methods=['POST'])
def get_valuation():
    """获取基金估值数据"""
    try:
        data = request.get_json()
        fund_code = data.get('fund_code', '').strip()

        if not fund_code:
            return jsonify({'error': '请输入基金代码'}), 400

        # ── 解析 init_code，持仓/仓位查询统一用主基金代码 ──
        init_code = resolve_init_code(fund_code)
        if init_code != fund_code:
            print(f"[init_code] {fund_code} → {init_code}")

        # 1. 获取基金持仓
        report_date = "2025-12-31"
        holdings = fetcher.get_fund_holdings(init_code, report_date)

        if holdings.empty:
            return jsonify({'error': f'未找到基金 {fund_code}（主基金 {init_code}）的持仓数据'}), 404

        # 2. 获取股票仓位
        stock_position = fetcher.get_stock_position_ratio(init_code, report_date)

        # 3. 获取股票代码列表
        stock_codes = holdings['股票代码'].unique().tolist()

        # 4. 获取昨收价
        last_trade_date = get_last_trade_date()
        prev_close_dict = fetcher.get_previous_close(stock_codes, last_trade_date)

        # 5. 获取日内K线
        today_str = get_today_str()
        intraday_klines = fetcher.batch_get_intraday_klines(stock_codes, today_str)

        # 6. 计算估值
        calculator = FundValuationCalculator(holdings, prev_close_dict, stock_position)
        valuation_data = calculator.calculate_fund_valuation(intraday_klines)

        # 7. 统计信息
        stats = calculator.get_summary_stats(valuation_data)

        # 8. 重仓股详细信息（含最新价和涨跌幅）
        holdings_detail = []
        for _, row in holdings.iterrows():
            code = row['股票代码']
            name = row['股票名称']
            ratio = row['持仓占比']
            prev_close = prev_close_dict.get(code)

            klines = intraday_klines.get(code, [])
            latest_price = klines[-1]['price'] if klines else prev_close

            change_pct = None
            if prev_close and prev_close > 0 and latest_price:
                change_pct = (latest_price - prev_close) / prev_close * 100

            holdings_detail.append({
                'code': code,
                'name': name,
                'ratio': round(ratio, 4),
                'prev_close': round(prev_close, 3) if prev_close else None,
                'latest_price': round(latest_price, 3) if latest_price else None,
                'change_pct': round(change_pct, 2) if change_pct is not None else None,
            })

        holdings_detail.sort(key=lambda x: x['ratio'], reverse=True)

        return jsonify({
            'success': True,
            'fund_code': fund_code,
            'init_code': init_code,
            'report_date': report_date,
            'stock_position': round(stock_position, 2),
            'holdings_count': len(stock_codes),
            'valuation_data': valuation_data,
            'stats': stats,
            'holdings': holdings_detail,
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f'数据获取失败: {str(e)}'}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000, use_reloader=False)