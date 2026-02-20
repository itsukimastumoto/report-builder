# JCBレポート運用ルール

JCB成果報告レポートの技術仕様・運用ルールです。使い方は [report-builder/README.md](../README.md) を参照してください。

---

## 入力CSV仕様

4つのCSVはすべてBigQueryダッシュボードからダウンロードし、ファイル名のキーワードで自動判別されます。

### 1. ログイン（キーワード: `ログイン`）
| カラム | 内容 |
|--------|------|
| client_id | クライアントID |
| company_name | 会社名 |
| week_start_date | 週の開始日 |
| week_end_date | 週の終了日 |
| weekly_login_users | 週次ログインユーザー数 |
| total_active_users | アクティブユーザー総数 |
| login_rate_percent | ログイン率 |

### 2. 購入数（キーワード: `購入`）
| カラム | 内容 |
|--------|------|
| client_id | クライアントID |
| company_name | 会社名 |
| week_start | 週の開始日 |
| weekly_purchase_users | 週次購入ユーザー数 |
| weekly_purchase_count | 週次購入数 |
| total_active_users | アクティブユーザー総数 |
| purchase_rate_percent | 購入率 |

> 0値の週はCSVに含まれない場合があります。ツールがログインCSVの全週を基準に自動で0埋めします。

### 3. ブランド（キーワード: `ブランド` or `販売`）
| カラム | 内容 |
|--------|------|
| client_id | クライアントID |
| client_name | 会社名 |
| brand_name | ブランド名 |
| total_count | 販売数 |
| unique_user_count | 購入ユーザー数 |
| total_price | 販売合計金額 |
| discounted_price_sum | 割引後金額合計 |
| discount_sum | 割引額合計 |

### 4. サマリ（キーワード: `サマリ`）
| カラム | 内容 |
|--------|------|
| client_id | クライアントID |
| client_name | 会社名 |
| period_start | 集計期間開始 |
| period_end | 集計期間終了 |
| first_registration_users | 初回登録ユーザー数 |
| mau | MAU（購入ユーザー数） |
| product_distribution_total | 商品代流通総額 |
| total_purchase_amount | 総購入金額 |

---

## 生成ルール

### スライド内容（1スライド = 1クライアント）
| 項目 | データソース | 単位 |
|------|------------|------|
| タイトル（会社名 ご報告資料） | サマリ | - |
| 期間 | サマリ | - |
| 初回登録ユーザー数 | サマリ | 人 |
| MAU（購入ユーザー数） | サマリ | 人 |
| 商品代流通総額・総購入金額 | サマリ | 円 |
| ログインユーザー数推移グラフ | ログイン | 人 |
| 購入数推移グラフ | 購入 | 件 |
| ブランド販売・発行割合ドーナツグラフ | ブランド | - |
| ブランドTOP10テーブル | ブランド | 円/件/人 |

### タイトル処理
- `{会社名} ご報告資料` の形式で表示
- `（申込企業：XXX）` は自動除去（例: `株式会社ポケットカード（申込企業：株式会社ジェーシービー）` → `株式会社ポケットカード ご報告資料`）

### データなしの場合
- 購入データなし → 購入チャート画像を削除
- ブランドデータなし → ドーナツ画像を削除、テーブルを空に
- ログインデータなし → ログインチャート画像を削除

### ブランドTOP10テーブル
- `total_price`（販売合計金額）の降順でTOP10を表示
- 11位以降は「その他」に集約（ドーナツグラフは最大8ブランド + その他）
- ブランド名が長い場合はフォントサイズを自動縮小（改行防止）

### グラフ
- 週次推移の棒グラフ: X軸は `M/D週` 表記（例: `1/19週`）
- 購入データの0値補完: ログインCSVの全週を基準に、欠損週を0で埋める

---

## タスクフォルダ

実行ごとに独立したフォルダが作成されます。前回の結果には影響しません。

```
jcb/tasks/
├── 20260220/          # 作業日ごとのフォルダ
│   ├── input/         # CSVのコピー（バックアップ）
│   └── output/        # JCB報告資料_YYYYMM.pptx
├── 20260220_2/        # 同日2回目は _2 サフィックス
└── 20260305/          # 別の日
```

---

## テンプレート

- ファイル: `jcb/template/報告資料_v3.pptx`
- PowerPointで直接編集可能
- シェイプはテキスト内容（「ご報告資料」「MAU」等）から自動検出するため、レイアウト変更やシェイプの移動をしても動作する
- テーブルの全行・全セルにcompact spacing（spcBef=0, spcAft=0, lnSpc=100%）を設定済み

---

## 設定ファイル（config.py）

CSVのカラム名が変わった場合は `config.py` の `COLUMN_MAP` を更新してください。

```python
TABLE_COLUMNS = ["brand_name", "total_price", "total_count", "unique_user_count"]
TABLE_MAX_ROWS = 10
BRAND_DONUT_MAX = 8

COLUMN_MAP = {
    "login":    {"group_key": "client_id", "company_name": "company_name", "date": "week_start_date", "value": "weekly_login_users"},
    "purchase": {"group_key": "client_id", "company_name": "company_name", "date": "week_start", "value": "weekly_purchase_count"},
    "summary":  {"group_key": "client_id", "company_name": "client_name"},
    "brand":    {"group_key": "client_id", "company_name": "client_name"},
}
```

---

## CLI直接実行

Claude Codeを使わずに直接実行する場合:

```bash
cd /path/to/report-builder

# CSVファイルを直接指定
python3 jcb/generate.py YYYYMM file1.csv file2.csv file3.csv file4.csv

# CSVが入ったディレクトリを指定
python3 jcb/generate.py YYYYMM --input-dir /path/to/csvs
```
