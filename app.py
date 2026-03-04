"""
アフター部門 出荷数ランキングアプリ
年間 / 月間 / 日別の出荷数ランキング TOP500 を表示する
"""

import os
from datetime import datetime, date
from flask import Flask, render_template, request, jsonify

# Oracle接続が利用できない環境ではデモモードで動作
try:
    import oracledb
    ORACLE_AVAILABLE = True
except ImportError:
    try:
        import cx_Oracle as oracledb
        ORACLE_AVAILABLE = True
    except ImportError:
        ORACLE_AVAILABLE = False

app = Flask(__name__)

# DB接続設定（環境変数で上書き可能）
DB_CONFIG = {
    "user": os.environ.get("DB_USER", "ECOREAD"),
    "password": os.environ.get("DB_PASSWORD", "ECOread01#"),
    "dsn": os.environ.get("DB_DSN", "172.25.3.119:1521/orcl.hrz.local"),
}

RANKING_LIMIT = 500


def get_db_connection():
    """Oracle DB接続を取得する"""
    if not ORACLE_AVAILABLE:
        return None
    return oracledb.connect(
        user=DB_CONFIG["user"],
        password=DB_CONFIG["password"],
        dsn=DB_CONFIG["dsn"],
    )


def execute_ranking_query(mode, target_date_str):
    """
    出荷ランキングを取得する

    Parameters
    ----------
    mode : str
        'yearly' / 'monthly' / 'daily'
    target_date_str : str
        対象日付文字列 (yyyy, yyyy-mm, yyyy-mm-dd)

    Returns
    -------
    list[dict] : ランキング結果
    str : 集計期間の表示文字列
    """
    conn = get_db_connection()
    if conn is None:
        return _demo_data(mode, target_date_str), _period_label(mode, target_date_str)

    try:
        cursor = conn.cursor()

        if mode == "yearly":
            year = target_date_str[:4]
            date_condition = f"to_char(sm.shukka_j_date, 'YYYY') = '{year}'"
            period_label = f"{year}年"
        elif mode == "monthly":
            year = target_date_str[:4]
            month = target_date_str[5:7]
            date_condition = (
                f"to_char(sm.shukka_j_date, 'YYYY') = '{year}' "
                f"AND to_char(sm.shukka_j_date, 'MM') = '{month}'"
            )
            period_label = f"{year}年{int(month)}月"
        else:  # daily
            date_condition = (
                f"to_char(sm.shukka_j_date, 'YYYY/MM/DD') = "
                f"'{target_date_str.replace('-', '/')}'"
            )
            parts = target_date_str.split("-")
            period_label = f"{parts[0]}年{int(parts[1])}月{int(parts[2])}日"

        sql = f"""
            SELECT
                ROWNUM AS rank_no,
                t.hinban,
                t.hm_nm,
                t.total_shukka_suu,
                t.shukka_count
            FROM (
                SELECT
                    jm.hinban,
                    MAX(jm.juchuu_hm_nm) AS hm_nm,
                    SUM(sm.shukka_j_suu) AS total_shukka_suu,
                    COUNT(DISTINCT to_char(sm.shukka_j_date, 'YYYY/MM/DD')) AS shukka_count
                FROM ecouser.t_shukka_m sm
                INNER JOIN ecouser.t_lot_info li
                    ON li.lot_no = sm.lot_no
                INNER JOIN ecouser.t_juchuu_m jm
                    ON jm.juchuu_no = li.juchuu_no
                    AND jm.juchuu_line_no = li.juchuu_line_no
                INNER JOIN ecouser.t_juchuu_h jh
                    ON jh.juchuu_no = jm.juchuu_no
                    AND jh.juchuu_hansuu = jm.juchuu_hansuu
                WHERE jh.juchuu_kyoten_cd = 'A'
                    AND {date_condition}
                    AND jm.hinban NOT LIKE 'C%'
                    AND jm.hinban NOT LIKE 'U%'
                GROUP BY jm.hinban
                ORDER BY SUM(sm.shukka_j_suu) DESC
            ) t
            WHERE ROWNUM <= {RANKING_LIMIT}
        """

        cursor.execute(sql)
        columns = [col[0].lower() for col in cursor.description]
        rows = []
        for row in cursor.fetchall():
            rows.append(dict(zip(columns, row)))

        cursor.close()
        return rows, period_label

    except Exception as e:
        app.logger.error(f"DB query error: {e}")
        return [], _period_label(mode, target_date_str)
    finally:
        conn.close()


def _period_label(mode, target_date_str):
    """集計期間の表示ラベルを生成"""
    if mode == "yearly":
        return f"{target_date_str[:4]}年"
    elif mode == "monthly":
        parts = target_date_str.split("-")
        return f"{parts[0]}年{int(parts[1])}月"
    else:
        parts = target_date_str.split("-")
        return f"{parts[0]}年{int(parts[1])}月{int(parts[2])}日"


def _demo_data(mode, target_date_str):
    """Oracle未接続時のデモデータ"""
    import random

    demo_parts = [
        ("4012273-00", "ベアリング A"),
        ("4012274-01", "シャフト B"),
        ("M167001-18", "モータ C"),
        ("E505712-06", "センサー D"),
        ("A910246-01", "バルブ E"),
        ("4000064-01", "ギア F"),
        ("M206597-02", "ポンプ G"),
        ("A970831-00", "フィルタ H"),
        ("4015500-03", "カップリング I"),
        ("E300100-02", "コントローラ J"),
    ]
    rows = []
    for i in range(min(50, RANKING_LIMIT)):
        idx = i % len(demo_parts)
        hinban, name = demo_parts[idx]
        suffix = f"-{i // len(demo_parts):02d}" if i >= len(demo_parts) else ""
        qty = max(1, 500 - i * 8 + random.randint(-5, 5))
        rows.append({
            "rank_no": i + 1,
            "hinban": hinban + suffix if suffix else hinban,
            "hm_nm": f"{name} ({i + 1})",
            "total_shukka_suu": qty,
            "shukka_count": max(1, qty // 10),
        })
    return rows


@app.route("/")
def index():
    """メインページ"""
    today = date.today()
    return render_template(
        "index.html",
        current_year=today.year,
        current_month=f"{today.year}-{today.month:02d}",
        current_date=today.isoformat(),
        oracle_available=ORACLE_AVAILABLE,
    )


@app.route("/api/ranking")
def api_ranking():
    """
    ランキングAPI

    Query Parameters
    ----------------
    mode : str
        yearly / monthly / daily
    target : str
        対象期間 (yyyy / yyyy-mm / yyyy-mm-dd)
    """
    mode = request.args.get("mode", "monthly")
    target = request.args.get("target", "")

    # デフォルト値の設定
    today = date.today()
    if not target:
        if mode == "yearly":
            target = str(today.year)
        elif mode == "monthly":
            target = f"{today.year}-{today.month:02d}"
        else:
            target = today.isoformat()

    # 入力バリデーション
    if mode not in ("yearly", "monthly", "daily"):
        return jsonify({"error": "Invalid mode"}), 400

    rows, period_label = execute_ranking_query(mode, target)

    return jsonify({
        "mode": mode,
        "target": target,
        "period_label": period_label,
        "total_count": len(rows),
        "ranking": rows,
        "oracle_available": ORACLE_AVAILABLE,
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "1") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug)
