"""JCBレポートのテンプレートPPTX設定

テンプレートのシェイプ構造を定義。
テンプレートが更新されたらここも合わせて更新する。
"""

from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
TEMPLATE_PATH = SCRIPT_DIR / "template" / "報告資料_v3.pptx"
TASKS_DIR = SCRIPT_DIR / "tasks"

# シェイプはインデックスではなくテキスト内容・位置から自動検出する
# （generate.py の _detect_shapes() を参照）

# テーブルの列定義（ブランドCSVのカラム名に対応）
TABLE_COLUMNS = ["brand_name", "total_price", "total_count", "unique_user_count"]
TABLE_MAX_ROWS = 10  # ヘッダー除く

# CSV自動検出パターン
CSV_PATTERNS = {
    "login": ["login", "ログイン"],
    "purchase": ["purchase", "購入"],
    "brand": ["brand", "ブランド", "販売"],
    "summary": ["summary", "サマリ"],
}

# CSVカラム名マッピング（BigQuery出力に合わせる）
COLUMN_MAP = {
    "login": {
        "group_key": "client_id",
        "company_name": "company_name",
        "date": "week_start_date",
        "value": "weekly_login_users",
    },
    "purchase": {
        "group_key": "client_id",
        "company_name": "company_name",
        "date": "week_start",
        "value": "weekly_purchase_count",
    },
    "summary": {
        "group_key": "client_id",
        "company_name": "client_name",
    },
    "brand": {
        "group_key": "client_id",
        "company_name": "client_name",
    },
}

# ブランドドーナツグラフの最大表示数（超過分は「その他」に集約）
BRAND_DONUT_MAX = 8
