"""
価格計算エンジン
VBAマクロの処理をPythonで高速化
- バッチSQL（IN句）で一括取得
- 接続プーリング対応
- concurrent.futures による並列処理
"""

import os
import math
import logging
from dataclasses import dataclass, field
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, timedelta

logger = logging.getLogger(__name__)

try:
    import oracledb
    ORACLE_AVAILABLE = True
except ImportError:
    try:
        import cx_Oracle as oracledb
        ORACLE_AVAILABLE = True
    except ImportError:
        ORACLE_AVAILABLE = False

# ---------------------------------------------------------------------------
# DB接続
# ---------------------------------------------------------------------------
DB_CONFIG = {
    "user": os.environ.get("DB_USER", "ECOREAD"),
    "password": os.environ.get("DB_PASSWORD", "ECOread01#"),
    "dsn": os.environ.get("DB_DSN", "172.25.3.119:1521/orcl.hrz.local"),
}

# HONPS DB (標準単価用)
HONPS_DB_CONFIG = {
    "user": os.environ.get("HONPS_DB_USER", "SELECT1"),
    "password": os.environ.get("HONPS_DB_PASSWORD", "SELECT1"),
    "dsn": os.environ.get("HONPS_DB_DSN", "ORAGOLD"),
}

_pool = None
_honps_pool = None


def get_pool():
    """ECO DB接続プールを取得"""
    global _pool
    if _pool is None and ORACLE_AVAILABLE:
        _pool = oracledb.create_pool(
            user=DB_CONFIG["user"],
            password=DB_CONFIG["password"],
            dsn=DB_CONFIG["dsn"],
            min=2,
            max=10,
            increment=1,
        )
    return _pool


def get_honps_pool():
    """HONPS DB接続プールを取得"""
    global _honps_pool
    if _honps_pool is None and ORACLE_AVAILABLE:
        try:
            _honps_pool = oracledb.create_pool(
                user=HONPS_DB_CONFIG["user"],
                password=HONPS_DB_CONFIG["password"],
                dsn=HONPS_DB_CONFIG["dsn"],
                min=1,
                max=5,
                increment=1,
            )
        except Exception:
            logger.warning("HONPS DB接続プール作成失敗")
    return _honps_pool


def get_eco_connection():
    pool = get_pool()
    if pool:
        return pool.acquire()
    return None


def get_honps_connection():
    pool = get_honps_pool()
    if pool:
        return pool.acquire()
    return None


def _fetchall_dict(cursor):
    """カーソル結果をdict listで返す"""
    cols = [c[0].lower() for c in cursor.description]
    return [dict(zip(cols, row)) for row in cursor.fetchall()]


def _roundup(value, digits):
    """VBAのRoundUp相当: digits=-1 → 10の位に切り上げ"""
    if value is None or value == 0:
        return 0
    factor = 10 ** (-digits)
    return math.ceil(value / factor) * factor


# ---------------------------------------------------------------------------
# 掛率テーブル（デフォルト値、API経由でも上書き可能）
# ---------------------------------------------------------------------------
@dataclass
class RateTable:
    """テーブルシートの掛率設定"""
    # UP, チャージ
    up: float = 1.0
    charge: float = 6000.0

    # M番 T仕切り掛率
    m_rate_1: float = 1.6      # 価格帯1未満
    m_band_1: float = 5000     # 価格帯1
    m_rate_2: float = 1.5      # 価格帯1以上 価格帯2未満
    m_band_2: float = 20000    # 価格帯2
    m_rate_3: float = 1.4      # 価格帯3以上
    m_band_3: float = 20000    # 価格帯3

    # 各カテゴリ掛率
    rate_4: float = 1.3        # 4番（購入品）
    rate_e: float = 1.5        # E番
    rate_a: float = 0.9        # A番（÷掛率）
    rate_l: float = 1.4        # L番
    rate_cv: float = 1.4       # CV番
    rate_um: float = 1.3       # UM番
    rate_p: float = 1.4        # P番

    # HI仕切り
    hi_var1: float = 1.0
    hi_var2: float = 1.0
    hi_var3: float = 1.0

    # 仮上代
    ka_var1: float = 1.0
    ka_var2: float = 1.0
    ka_var3: float = 1.0   # 4番以外
    ka_var4: float = 1.0   # 4番以外

    # ディーラー仕切り
    de_var1: float = 1.0

    # 上代
    jo_rate1: float = 1.0
    jo_band1: float = 11000
    jo_rate2: float = 1.0
    jo_band2: float = 33000
    jo_rate3: float = 1.0

    # H仕切り対T仕切り比較用
    h_ratio: float = 1.0


# ---------------------------------------------------------------------------
# データモデル
# ---------------------------------------------------------------------------
@dataclass
class PriceResult:
    """1品番の価格計算結果"""
    hinban: str = ""
    category: str = ""             # M, 4, E, F, A, L, CV, UM, P
    genka: float = 0               # 原価（標準単価）
    t_sikiri: float = 0            # T仕切り
    h_sikiri: float = 0            # H仕切り（ECO）
    hi_sikiri: float = 0           # HI仕切り
    kari_joudai: float = 0         # 仮上代
    dealer_sikiri: float = 0       # ディーラー仕切り
    joudai: float = 0              # 上代
    rate_used: float = 0           # 使用した掛率
    naikote_cost: float = 0        # 内工程コスト
    gaikote_cost: float = 0        # 外工程コスト
    konyu_cost: float = 0          # 購入コスト
    first_process: str = ""        # 第1工程
    buhin_kubun: str = ""          # 部品区分
    null_check: str = ""           # NULLデータ有無
    buhin_name: str = ""           # 部品名
    # A番用
    bom_total: float = 0           # BOM部品合計
    assembly_cost: float = 0       # 組立工数コスト
    outsource_cost: float = 0      # 社外組立費
    assembly_place: str = ""       # 組立場所
    children: list = field(default_factory=list)  # 構成部品リスト


@dataclass
class ChildPart:
    """A番構成部品"""
    hinban: str = ""
    name: str = ""
    inzu: float = 0
    unit_price: float = 0
    t_sikiri: float = 0


# ---------------------------------------------------------------------------
# 一括バッチ取得（高速化の要）
# ---------------------------------------------------------------------------
def _in_clause(items, prefix="p"):
    """バインド変数付きIN句を生成"""
    binds = {f"{prefix}{i}": v for i, v in enumerate(items)}
    clause = ", ".join(f":{k}" for k in binds)
    return clause, binds


def batch_fetch_standard_prices(hinban_list):
    """HONPS標準単価を一括取得"""
    conn = get_honps_connection()
    if not conn or not hinban_list:
        return {}

    results = {}
    try:
        cursor = conn.cursor()
        # 1000件ずつ分割（OracleのIN句制限）
        for i in range(0, len(hinban_list), 900):
            chunk = hinban_list[i:i + 900]
            clause, binds = _in_clause(chunk)
            sql = f"""
                SELECT hinban,
                       CASE WHEN tan_cost_ko IS NULL
                            THEN tan_cost
                            ELSE (tan_cost + tan_cost_ko)
                       END AS genka,
                       naikote_cost, gaikote_cost, konyu_cost,
                       tan_cost_ko, kote_1
                FROM honps.hv_ma_ta_hyotanka
                WHERE hinban IN ({clause})
            """
            cursor.execute(sql, binds)
            for row in cursor.fetchall():
                results[row[0]] = {
                    "genka": float(row[1] or 0),
                    "naikote_cost": float(row[2] or 0),
                    "gaikote_cost": float(row[3] or 0),
                    "konyu_cost": float(row[4] or 0),
                    "first_process": row[6] or "",
                }
        cursor.close()
    except Exception as e:
        logger.error(f"HONPS標準単価取得エラー: {e}")
    finally:
        if conn:
            conn.close()
    return results


def batch_fetch_h_sikiri(hinban_list):
    """ECO H仕切りを一括取得"""
    conn = get_eco_connection()
    if not conn or not hinban_list:
        return {}

    results = {}
    try:
        cursor = conn.cursor()
        for i in range(0, len(hinban_list), 900):
            chunk = hinban_list[i:i + 900]
            clause, binds = _in_clause(chunk)
            sql = f"""
                SELECT shohin_buhin_cd, H_sikiri
                FROM ecouser.hv_shohin_buhin
                WHERE shohin_buhin_cd IN ({clause})
                    AND zaiko_kyoten_cd = 'A'
                    AND torisaki_cd = 'T10000'
            """
            cursor.execute(sql, binds)
            for row in cursor.fetchall():
                cd = row[0]
                val = float(row[1] or 0)
                # 最新のH仕切りを保持
                if cd not in results or val > results[cd]:
                    results[cd] = val
        cursor.close()
    except Exception as e:
        logger.error(f"H仕切り取得エラー: {e}")
    finally:
        conn.close()
    return results


def batch_fetch_h_sikiri_by_prefix(hinban_list):
    """品番先頭7桁の最大改版のH仕切りを一括取得（H仕切り追加用）"""
    conn = get_eco_connection()
    if not conn or not hinban_list:
        return {}

    results = {}
    try:
        cursor = conn.cursor()
        for hinban in hinban_list:
            prefix7 = hinban[:7]
            sql = """
                SELECT a.H_sikiri, a.shohin_buhin_cd
                FROM ecouser.hv_shohin_buhin a
                WHERE a.shohin_buhin_cd = (
                    SELECT MAX(b.shohin_buhin_cd)
                    FROM ecouser.hv_shohin_buhin b
                    WHERE b.shohin_buhin_cd LIKE :prefix
                        AND b.zaiko_kyoten_cd = 'A'
                        AND b.torisaki_cd = 'T10000'
                )
                AND a.zaiko_kyoten_cd = 'A'
                AND a.torisaki_cd = 'T10000'
                ORDER BY a.hannkou_date DESC
            """
            cursor.execute(sql, {"prefix": prefix7 + "%"})
            row = cursor.fetchone()
            if row and row[0]:
                results[hinban] = float(row[0])
        cursor.close()
    except Exception as e:
        logger.error(f"H仕切り(prefix)取得エラー: {e}")
    finally:
        conn.close()
    return results


def batch_fetch_purchase_prices(hinban_list):
    """4番購入品の単価を一括取得"""
    conn = get_eco_connection()
    if not conn or not hinban_list:
        return {}

    results = {}
    try:
        cursor = conn.cursor()
        for hinban in hinban_list:
            sql = """
                SELECT DISTINCT km.tanka, thm.tori_tuuka_tani_kbn
                FROM ecouser.t_kakakuhyou_h_mst kh
                LEFT JOIN ecouser.t_kakakuhyou_m_mst km
                    ON kh.hinban = km.hinban
                    AND kh.sgy_bumon_kbn = km.sgy_bumon_kbn
                    AND kh.koutei_cd = km.koutei_cd
                    AND kh.start_date = km.kkhh_start_date
                    AND km.seiban = '*'
                    AND km.end_date = '9999/1/1'
                    AND kh.torisaki_cd = km.torisaki_cd
                INNER JOIN ecouser.v_seizou_view sv
                    ON kh.hinban = sv.oya_hinban
                    AND kh.koutei_cd = sv.ko_hinban
                    AND sv.oya_seiban = '*'
                    AND sv.bkj_end_date = '9999/1/1'
                INNER JOIN ecouser.t_torisaki_hm_mst thm
                    ON kh.hinban = thm.hinban
                    AND kh.sgy_bumon_kbn = thm.sgy_bumon_kbn
                    AND kh.koutei_cd = thm.koutei_cd
                    AND thm.end_date = '9999/1/1'
                    AND kh.torisaki_cd = thm.torisaki_cd
                    AND thm.seiban = '*'
                WHERE kh.end_date = '9999/1/1'
                    AND kh.sgy_bumon_kbn = '*'
                    AND kh.seiban = '*'
                    AND kh.hinban = :hinban
                ORDER BY kh.hinban
            """
            cursor.execute(sql, {"hinban": hinban})
            row = cursor.fetchone()
            if row:
                tanka = float(row[0] or 0)
                currency = row[1] or "JPY"
                if currency != "JPY":
                    rate = _get_exchange_rate(cursor, currency)
                    tanka = tanka * rate
                results[hinban] = tanka
        cursor.close()
    except Exception as e:
        logger.error(f"購入品単価取得エラー: {e}")
    finally:
        conn.close()
    return results


def _get_exchange_rate(cursor, currency_from):
    """為替レート取得"""
    sql = """
        SELECT DISTINCT rm.rate
        FROM ecouser.t_rate_mst rm
        WHERE rm.tuuka_cd_from = :cur
    """
    cursor.execute(sql, {"cur": currency_from})
    row = cursor.fetchone()
    return float(row[0]) if row else 1.0


def batch_fetch_buhin_kubun(hinban_list):
    """部品区分を一括取得"""
    conn = get_eco_connection()
    if not conn or not hinban_list:
        return {}

    results = {}
    try:
        cursor = conn.cursor()
        for i in range(0, len(hinban_list), 900):
            chunk = hinban_list[i:i + 900]
            clause, binds = _in_clause(chunk)
            sql = f"""
                SELECT shohin_buhin_cd,
                       LISTAGG(DISTINCT buhinkubun, '_')
                           WITHIN GROUP (ORDER BY buhinkubun) AS kubun
                FROM ecouser.hv_shohin_buhin
                WHERE shohin_buhin_cd IN ({clause})
                    AND zaiko_kyoten_cd = 'A'
                GROUP BY shohin_buhin_cd
            """
            cursor.execute(sql, binds)
            for row in cursor.fetchall():
                results[row[0]] = row[1] or ""
        cursor.close()
    except Exception as e:
        logger.error(f"部品区分取得エラー: {e}")
    finally:
        conn.close()
    return results


# ---------------------------------------------------------------------------
# A番処理
# ---------------------------------------------------------------------------
def fetch_a_bom(hinban):
    """A番のBOM（構成部品）を取得"""
    conn = get_eco_connection()
    if not conn:
        return [], 0, "", ""

    children = []
    assembly_cost = 0
    assembly_place = ""
    outsource_cost = 0

    try:
        cursor = conn.cursor()

        # 構成部品取得
        sql = """
            SELECT a.ko_hinban AS hinban, INZUU AS inzu,
                   b.hm_nm_1 AS name
            FROM ecouser.v_seizou_view a
            LEFT JOIN ecouser.t_hm_mst b
                ON a.ko_hinban = b.hinban
                AND b.seiban = '*'
                AND b.end_date = '9999/1/1'
                AND b.sgy_bumon_kbn = '*'
                AND b.kyoten_cd = '*'
            WHERE a.oya_hinban = :hinban
                AND a.oya_seiban = '*'
                AND a.bkj_end_date = '9999/1/1'
                AND a.ko_data_kbn = 'PC001'
        """
        cursor.execute(sql, {"hinban": hinban})
        for row in cursor.fetchall():
            children.append(ChildPart(
                hinban=row[0] or "",
                inzu=float(row[1] or 0),
                name=row[2] or "",
            ))

        # 社外組立費
        sql2 = """
            SELECT a.TANKA AS cost
            FROM ecouser.t_kakakuhyou_m_mst a
            LEFT JOIN ecouser.t_kakakuhyou_h_mst b
                ON a.hinban = b.hinban
                AND a.kkhh_start_date = b.start_date
                AND b.end_date = '9999/1/1'
            WHERE a.hinban = :hinban
        """
        cursor.execute(sql2, {"hinban": hinban})
        max_cost = 0
        for row in cursor.fetchall():
            val = float(row[0] or 0)
            if val > max_cost:
                max_cost = val
        outsource_cost = max_cost

        # 組立工数・組立場所
        sql3 = """
            SELECT dandori_time, naigaisaku_kbn, line_cd, torisaki_cd
            FROM ecouser.v_seizou_view
            WHERE oya_hinban = :hinban
                AND oya_seiban = '*'
                AND bkj_end_date = '9999/1/1'
                AND (ko_hinban = 'K-S' OR ko_hinban = '@K-S')
                AND ko_data_kbn = 'PC003'
        """
        cursor.execute(sql3, {"hinban": hinban})
        row = cursor.fetchone()
        if row:
            dandori = float(row[0] or 0)
            naigai = row[1] or ""
            line_cd = row[2] or ""
            torisaki = row[3] or ""
            # KIT工程の工数追加（min 5部品×2分）
            kit_min = max(len(children), 5) * 2
            work_min = dandori / 60 + kit_min
            assembly_cost = work_min
            if naigai == "PC001":
                assembly_place = torisaki
            else:
                assembly_place = line_cd

        cursor.close()
    except Exception as e:
        logger.error(f"A番BOM取得エラー ({hinban}): {e}")
    finally:
        conn.close()

    return children, outsource_cost, assembly_place, assembly_cost


# ---------------------------------------------------------------------------
# T仕切り計算
# ---------------------------------------------------------------------------
def calc_t_sikiri(genka, category, rates, first_process="", naikote_cost=0):
    """カテゴリと掛率テーブルからT仕切りを計算"""
    rate = 0

    if category == "M":
        if genka < rates.m_band_1:
            rate = rates.m_rate_1
        elif genka < rates.m_band_2:
            rate = rates.m_rate_2
        else:
            rate = rates.m_rate_3
        # M番で第1工程が@BYまたはBYかつ内作コスト0 → 4番掛率
        if first_process in ("@BY", "BY") and naikote_cost == 0:
            rate = rates.rate_4
    elif category == "4":
        rate = rates.rate_4
    elif category in ("E", "F"):
        rate = rates.rate_e
    elif category == "A":
        rate = rates.rate_a
    elif category == "L":
        rate = rates.rate_l
    elif category == "CV":
        rate = rates.rate_cv
    elif category == "UM":
        rate = rates.rate_um
    elif category == "P":
        rate = rates.rate_p

    if rate == 0:
        return 0, 0

    if category == "A":
        # A番は ÷ 掛率
        t_sikiri = _roundup(genka / rate, -1)
    else:
        t_sikiri = _roundup(genka * rate, -1)

    return t_sikiri, rate


# ---------------------------------------------------------------------------
# 上代等計算
# ---------------------------------------------------------------------------
def calc_additional_prices(t_sikiri, category, rates):
    """HI仕切り、仮上代、ディーラー仕切り、上代を計算"""
    # HI仕切り
    hi = _roundup(t_sikiri / rates.hi_var1 * rates.hi_var2 * rates.hi_var3, 0)

    # 仮上代
    if category == "4":
        kari = _roundup(t_sikiri / rates.ka_var1 * rates.ka_var2, 0)
    else:
        kari = _roundup(
            t_sikiri / rates.ka_var1 * rates.ka_var2
            * rates.ka_var3 * rates.ka_var4, 0
        )

    # ディーラー仕切り
    dealer = _roundup(hi * rates.de_var1, 0)

    # 上代
    if kari < rates.jo_band1:
        if kari < 1000:
            joudai = _roundup(kari * rates.jo_rate1, 0)
        else:
            joudai = _roundup(kari * rates.jo_rate1, -1)
    elif kari < rates.jo_band2:
        joudai = _roundup(kari * rates.jo_rate2, -1)
    else:
        joudai = _roundup(kari * rates.jo_rate3, -1)

    return hi, kari, dealer, joudai


# ---------------------------------------------------------------------------
# カテゴリ判定
# ---------------------------------------------------------------------------
def detect_category(hinban):
    """品番先頭からカテゴリを判定"""
    if not hinban:
        return ""
    h = hinban.upper()
    if h[:2] == "CV":
        return "CV"
    if h[:2] == "UM":
        return "UM"
    first = h[0]
    if first in ("M", "4", "E", "F", "A", "L", "P"):
        return first
    return ""


# ---------------------------------------------------------------------------
# メイン: 一括価格計算
# ---------------------------------------------------------------------------
def calculate_prices(hinban_list, rates=None, calc_additional=False):
    """
    品番リストの価格を一括計算する

    Parameters
    ----------
    hinban_list : list[str]
        品番リスト
    rates : RateTable, optional
        掛率テーブル（Noneならデフォルト使用）
    calc_additional : bool
        上代等も計算するか

    Returns
    -------
    list[PriceResult]
    """
    if rates is None:
        rates = RateTable()

    if not hinban_list:
        return []

    # カテゴリ分類
    items = [(h, detect_category(h)) for h in hinban_list]
    non_a = [h for h, c in items if c != "A" and c != ""]
    a_items = [h for h, c in items if c == "A"]

    # Phase 1: バッチ取得（並列実行）
    std_prices = {}
    h_sikiris = {}
    h_sikiris_prefix = {}
    purchase_prices = {}
    buhin_kubuns = {}

    four_items = [h for h, c in items if c == "4"]

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {}
        if non_a:
            futures["std"] = executor.submit(batch_fetch_standard_prices, non_a)
        futures["h_sikiri"] = executor.submit(batch_fetch_h_sikiri, hinban_list)
        futures["h_prefix"] = executor.submit(
            batch_fetch_h_sikiri_by_prefix, hinban_list
        )
        if four_items:
            futures["purchase"] = executor.submit(
                batch_fetch_purchase_prices, four_items
            )
        futures["kubun"] = executor.submit(batch_fetch_buhin_kubun, hinban_list)

        for key, future in futures.items():
            try:
                result = future.result(timeout=60)
                if key == "std":
                    std_prices = result
                elif key == "h_sikiri":
                    h_sikiris = result
                elif key == "h_prefix":
                    h_sikiris_prefix = result
                elif key == "purchase":
                    purchase_prices = result
                elif key == "kubun":
                    buhin_kubuns = result
            except Exception as e:
                logger.error(f"バッチ取得エラー ({key}): {e}")

    # Phase 2: 品番ごとに価格計算
    results = []
    for hinban, category in items:
        pr = PriceResult(hinban=hinban, category=category)
        pr.buhin_kubun = buhin_kubuns.get(hinban, "")

        if category == "A":
            # A番: BOM展開
            pr = _calc_a_price(hinban, rates, std_prices, h_sikiris)
            pr.buhin_kubun = buhin_kubuns.get(hinban, "")
        elif category == "4":
            # 4番購入品
            pr.genka = purchase_prices.get(hinban, 0)
            if pr.genka == 0:
                sp = std_prices.get(hinban, {})
                pr.genka = sp.get("genka", 0)
            pr.t_sikiri, pr.rate_used = calc_t_sikiri(
                pr.genka, category, rates
            )
        else:
            # M, E, F, L, CV, UM, P
            sp = std_prices.get(hinban, {})
            pr.genka = sp.get("genka", 0)
            pr.naikote_cost = sp.get("naikote_cost", 0)
            pr.gaikote_cost = sp.get("gaikote_cost", 0)
            pr.konyu_cost = sp.get("konyu_cost", 0)
            pr.first_process = sp.get("first_process", "")
            pr.t_sikiri, pr.rate_used = calc_t_sikiri(
                pr.genka, category, rates,
                first_process=pr.first_process,
                naikote_cost=pr.naikote_cost,
            )

        # H仕切り
        pr.h_sikiri = h_sikiris.get(hinban, 0)
        if pr.h_sikiri == 0:
            pr.h_sikiri = h_sikiris_prefix.get(hinban, 0)

        # 上代等
        if calc_additional and pr.t_sikiri > 0:
            hi, kari, dealer, joudai = calc_additional_prices(
                pr.t_sikiri, category, rates
            )
            pr.hi_sikiri = hi
            pr.kari_joudai = kari
            pr.dealer_sikiri = dealer
            pr.joudai = joudai

        # NULLチェック
        if pr.genka == 0 and category:
            pr.null_check = "有"

        results.append(pr)

    return results


def _calc_a_price(hinban, rates, std_prices_cache, h_sikiris_cache):
    """A番の価格計算: BOM展開 → 子部品合計 + 工数 + 組立外注費"""
    pr = PriceResult(hinban=hinban, category="A")

    # BOM取得
    children, outsource_cost, assembly_place, work_min = fetch_a_bom(hinban)
    pr.outsource_cost = outsource_cost
    pr.assembly_place = assembly_place

    # 子部品の単価取得
    child_hinbans = [c.hinban for c in children]
    child_std = batch_fetch_standard_prices(
        [h for h in child_hinbans if h not in std_prices_cache]
    )
    child_std.update({k: v for k, v in std_prices_cache.items() if k in child_hinbans})

    # 子部品のT仕切り計算と合計
    bom_genka_total = 0
    bom_t_total = 0
    child_results = []

    for child in children:
        c_cat = detect_category(child.hinban)
        sp = child_std.get(child.hinban, {})
        c_genka = sp.get("genka", 0)

        # 4番子部品は購入品価格
        if c_cat == "4":
            pp = batch_fetch_purchase_prices([child.hinban])
            if child.hinban in pp:
                c_genka = pp[child.hinban]

        c_first = sp.get("first_process", "")
        c_naikote = sp.get("naikote_cost", 0)
        c_t, _ = calc_t_sikiri(
            c_genka, c_cat, rates,
            first_process=c_first,
            naikote_cost=c_naikote,
        )

        child.unit_price = c_genka
        child.t_sikiri = c_t
        child_results.append(child)

        bom_genka_total += child.inzu * c_genka
        bom_t_total += child.inzu * c_t

        if c_t == 0:
            pr.null_check = "有"

    pr.children = child_results
    pr.bom_total = bom_genka_total

    # 工数コスト
    charge_rate = rates.charge
    assembly_charge = work_min / 60 * charge_rate
    pr.assembly_cost = assembly_charge

    # 原価 = BOM合計 + 工数×チャージ + 社外組立費
    pr.genka = bom_genka_total + assembly_charge + outsource_cost

    # T仕切り合計 = BOM T仕切り合計 + 工数×チャージ + 社外組立費
    t_total = bom_t_total + assembly_charge + outsource_cost
    pr.t_sikiri = _roundup(t_total / rates.rate_a, -1)
    pr.rate_used = rates.rate_a

    return pr


# ---------------------------------------------------------------------------
# デモデータ
# ---------------------------------------------------------------------------
def demo_calculate(hinban_list, rates=None, calc_additional=False):
    """Oracle未接続時のデモ計算"""
    import random
    if rates is None:
        rates = RateTable()

    results = []
    for hinban in hinban_list:
        cat = detect_category(hinban)
        pr = PriceResult(hinban=hinban, category=cat)

        # デモ原価
        base_prices = {
            "M": lambda: random.uniform(800, 50000),
            "4": lambda: random.uniform(100, 15000),
            "E": lambda: random.uniform(500, 30000),
            "F": lambda: random.uniform(500, 30000),
            "A": lambda: random.uniform(5000, 100000),
            "L": lambda: random.uniform(1000, 20000),
            "CV": lambda: random.uniform(2000, 40000),
            "UM": lambda: random.uniform(3000, 50000),
            "P": lambda: random.uniform(500, 10000),
        }
        pr.genka = round(base_prices.get(cat, lambda: 1000)(), 2)
        pr.buhin_name = f"デモ部品 {hinban}"

        pr.t_sikiri, pr.rate_used = calc_t_sikiri(pr.genka, cat, rates)
        pr.h_sikiri = _roundup(pr.genka * random.uniform(1.1, 1.4), -1)

        if calc_additional and pr.t_sikiri > 0:
            hi, kari, dealer, joudai = calc_additional_prices(
                pr.t_sikiri, cat, rates
            )
            pr.hi_sikiri = hi
            pr.kari_joudai = kari
            pr.dealer_sikiri = dealer
            pr.joudai = joudai

        results.append(pr)

    return results


def price_result_to_dict(pr):
    """PriceResultをdict変換"""
    d = {
        "hinban": pr.hinban,
        "category": pr.category,
        "buhin_name": pr.buhin_name,
        "genka": pr.genka,
        "t_sikiri": pr.t_sikiri,
        "h_sikiri": pr.h_sikiri,
        "hi_sikiri": pr.hi_sikiri,
        "kari_joudai": pr.kari_joudai,
        "dealer_sikiri": pr.dealer_sikiri,
        "joudai": pr.joudai,
        "rate_used": pr.rate_used,
        "naikote_cost": pr.naikote_cost,
        "gaikote_cost": pr.gaikote_cost,
        "konyu_cost": pr.konyu_cost,
        "first_process": pr.first_process,
        "buhin_kubun": pr.buhin_kubun,
        "null_check": pr.null_check,
    }
    if pr.category == "A":
        d["bom_total"] = pr.bom_total
        d["assembly_cost"] = pr.assembly_cost
        d["outsource_cost"] = pr.outsource_cost
        d["assembly_place"] = pr.assembly_place
        d["children"] = [
            {
                "hinban": c.hinban,
                "name": c.name,
                "inzu": c.inzu,
                "unit_price": c.unit_price,
                "t_sikiri": c.t_sikiri,
            }
            for c in pr.children
        ]
    return d
