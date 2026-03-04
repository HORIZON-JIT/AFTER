"""
価格計算 Flask API
一括価格計算・A番BOM展開・掛率設定をWeb APIとして提供
"""

import os
import time
import logging
from datetime import date
from flask import Flask, render_template, request, jsonify

from pricing import (
    ORACLE_AVAILABLE,
    RateTable,
    calculate_prices,
    demo_calculate,
    price_result_to_dict,
    detect_category,
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)


# ---------------------------------------------------------------------------
# ページ
# ---------------------------------------------------------------------------
@app.route("/pricing")
def pricing_page():
    """価格計算ページ"""
    today = date.today()
    return render_template(
        "pricing.html",
        oracle_available=ORACLE_AVAILABLE,
        current_date=today.isoformat(),
    )


# ---------------------------------------------------------------------------
# API: 一括価格計算
# ---------------------------------------------------------------------------
@app.route("/api/pricing/calculate", methods=["POST"])
def api_calculate():
    """
    一括価格計算API

    Request Body (JSON)
    -------------------
    {
        "hinban_list": ["M167001-18", "4012273-00", ...],
        "calc_additional": true,  // 上代等も計算
        "rates": { ... }         // 掛率（省略時デフォルト）
    }
    """
    data = request.get_json()
    if not data or "hinban_list" not in data:
        return jsonify({"error": "hinban_listが必要です"}), 400

    hinban_list = data["hinban_list"]
    if not isinstance(hinban_list, list) or len(hinban_list) == 0:
        return jsonify({"error": "品番を1件以上指定してください"}), 400

    if len(hinban_list) > 5000:
        return jsonify({"error": "一度に5000件までです"}), 400

    calc_additional = data.get("calc_additional", False)

    # 掛率テーブル
    rates = RateTable()
    if "rates" in data and isinstance(data["rates"], dict):
        rate_data = data["rates"]
        for key, value in rate_data.items():
            if hasattr(rates, key):
                try:
                    setattr(rates, key, float(value))
                except (ValueError, TypeError):
                    pass

    # 計算実行
    start = time.time()
    try:
        if ORACLE_AVAILABLE:
            results = calculate_prices(hinban_list, rates, calc_additional)
        else:
            results = demo_calculate(hinban_list, rates, calc_additional)
    except Exception as e:
        logger.error(f"価格計算エラー: {e}")
        return jsonify({"error": f"計算エラー: {str(e)}"}), 500

    elapsed = round(time.time() - start, 3)

    # 集計
    total_genka = sum(r.genka for r in results)
    total_t = sum(r.t_sikiri for r in results)
    null_count = sum(1 for r in results if r.null_check == "有")

    return jsonify({
        "count": len(results),
        "elapsed_sec": elapsed,
        "total_genka": total_genka,
        "total_t_sikiri": total_t,
        "null_count": null_count,
        "oracle_available": ORACLE_AVAILABLE,
        "results": [price_result_to_dict(r) for r in results],
    })


# ---------------------------------------------------------------------------
# API: 単品詳細計算（M番の工程情報含む）
# ---------------------------------------------------------------------------
@app.route("/api/pricing/detail/<hinban>")
def api_detail(hinban):
    """単品の詳細価格情報を返す"""
    rates = RateTable()
    start = time.time()

    try:
        if ORACLE_AVAILABLE:
            results = calculate_prices([hinban], rates, calc_additional=True)
        else:
            results = demo_calculate([hinban], rates, calc_additional=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    if not results:
        return jsonify({"error": "計算結果なし"}), 404

    elapsed = round(time.time() - start, 3)
    result = price_result_to_dict(results[0])
    result["elapsed_sec"] = elapsed

    return jsonify(result)


# ---------------------------------------------------------------------------
# API: 掛率デフォルト値
# ---------------------------------------------------------------------------
@app.route("/api/pricing/rates")
def api_rates():
    """デフォルト掛率テーブルを返す"""
    rt = RateTable()
    fields = {}
    for f in rt.__dataclass_fields__:
        fields[f] = getattr(rt, f)
    return jsonify(fields)


# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PRICE_PORT", 5001))
    debug = os.environ.get("FLASK_DEBUG", "1") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug)
