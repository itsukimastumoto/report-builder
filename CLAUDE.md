# Report Builder（レポートビルダー）

レポート生成ツールの統合フォルダ。配下にレポート種別ごとのサブツールを持つ。

## トリガーワード

| キーワード | 対象 |
|-----------|------|
| JCBレポート、JCB報告資料、成果報告レポート | `jcb/` |
| 月次レポート、月次集計、OASIS月次 | `oasis-monthly/` |

## 構成

```
report-builder/
├── shared/          # 共通ライブラリ（グラフ生成、CSV処理）
├── jcb/             # JCB成果報告レポート（PPTX）
└── oasis-monthly/   # OASIS月次レポート（Excel）
```

## JCBレポート（jcb/）

### 概要
BigQueryから出力したCSV（4種類）を元に、クライアントごとの報告スライドを**1つのPPTXファイル**にまとめて生成する。

### 実行方法
```bash
cd tools/report-builder
# CSVファイルを直接指定（タスクフォルダが自動作成される）
python3 jcb/generate.py YYYYMM file1.csv file2.csv file3.csv file4.csv

# またはCSVが入ったディレクトリを指定
python3 jcb/generate.py YYYYMM --input-dir /path/to/csvs
```

### 入力CSV（4種類、全クライアントのデータが行に入っている）
1. **ログイン**: ファイル名に `login` or `ログイン` を含む
   - カラム: `client_id, company_name, week_start_date, week_end_date, weekly_login_users, total_active_users, login_rate_percent`
   - 週次データ（全週のデータが含まれる、0値含む）
2. **購入数**: ファイル名に `purchase` or `購入` を含む
   - カラム: `client_id, company_name, week_start, weekly_purchase_users, weekly_purchase_count, total_active_users, purchase_rate_percent`
   - 週次データ（0値の週はCSVに含まれない → コードが自動で0埋め）
3. **ブランド**: ファイル名に `brand` or `ブランド` or `販売` を含む
   - カラム: `client_id, client_name, brand_name, total_count, unique_user_count, total_price, discounted_price_sum, discount_sum`
   - クライアント別のブランドデータ
4. **サマリ**: ファイル名に `summary` or `サマリ` を含む
   - カラム: `client_id, client_name, period_start, period_end, first_registration_users, mau, product_distribution_total, total_purchase_amount`
   - クライアント単位で集約済み

### 出力（タスクフォルダ方式）
毎回の作業は独立したタスクフォルダに分離される。前回の作業とは無関係。
```
jcb/tasks/
├── 20260220/          # 作業日ごとのフォルダ（自動作成）
│   ├── input/         # CSVのコピー（保存される）
│   └── output/        # JCB報告資料_YYYYMM.pptx
├── 20260220_2/        # 同日2回目は _2 サフィックス
│   ├── input/
│   └── output/
└── 20260305/          # 別の日の作業
    ├── input/
    └── output/
```

### テンプレート
- `jcb/template/報告資料_v3.pptx` — ユーザー手編集の最終版
- シェイプはテキスト内容から自動検出するため、テンプレート編集後にconfig.pyの更新は不要
- テーブルの全行・全セルにcompact spacing（spcBef=0, spcAft=0, lnSpc=100%）が設定済み

### 生成ルール
- **タイトル**: `{会社名} ご報告資料`（「（イシュア名）」は付けない）
- **会社名クリーニング**: `（申込企業：XXX）` はタイトルから自動除去（正規表現パターン）
- **データなしの場合**: テンプレートのダミー画像・ダミーテーブルデータを削除（空白にする）
  - 購入データなし → 購入チャート画像を削除
  - ブランドデータなし → ドーナツ画像を削除、テーブルを空に
  - ログインデータなし → ログインチャート画像を削除
- **週次グラフ**: X軸は `M/D週` 表記（例: `1/19週`）
- **購入データの0値補完**: ログインCSVの全週を基準に、欠損週を0で埋める
- **ブランドTOP10**: total_priceの降順でTOP10、超過分は「その他」に集約
- **テキスト置換**: run単位で行い、フォントサイズ・書式を維持する
- **スライド統合**: 個別PPTXを生成→1ファイルにマージ（画像は一時ファイル経由でadd_pictureを使い一意なPartを生成）

### config.py の COLUMN_MAP
CSVのカラム名が変わったらここを更新する。現在の設定:
```python
COLUMN_MAP = {
    "login":    {"group_key": "client_id", "company_name": "company_name", "date": "week_start_date", "value": "weekly_login_users"},
    "purchase": {"group_key": "client_id", "company_name": "company_name", "date": "week_start", "value": "weekly_purchase_count"},
    "summary":  {"group_key": "client_id", "company_name": "client_name"},
    "brand":    {"group_key": "client_id", "company_name": "client_name"},
}
```

## OASIS月次レポート（oasis-monthly/）

### 概要
giftee Benefitの月次実績集計（キャンペーン販促費・ポイント利用額）を自動生成する。

### Claudeの実行フロー

#### 1. ファイル確認（3つ揃っているか）
以下の3ファイルが揃っていることを確認：
- ポイント利用状況(月次)
- 商品×割引率ごとの発行数・購入額・割引額総計
- クライアントID単位の商品&案件割引率一覧

**3ファイル揃っていない場合は処理を実行しない。**

#### 2. 抽出条件スクショの確認
ユーザーがスクショを添付していない場合は以下を依頼：
> 「抽出条件のスクショも一緒にください。確認ポイント：
> - クライアントID = all
> - 案件ID = all
> - 集計期間 = その月の1日〜末日（例: 2026-01-01 → 2026-01-31）」

スクショで以下を目視確認：
- クライアントID: `all` になっているか
- 案件ID: `all` になっているか
- 集計期間: 対象月の初日〜末日になっているか

#### 3. input/フォルダに配置
```bash
cp ~/Downloads/*.xlsx /Users/itsuki.matsumoto/claude-code/tools/report-builder/oasis-monthly/input/
cp ~/Downloads/*.png /Users/itsuki.matsumoto/claude-code/tools/report-builder/oasis-monthly/input/  # スクショも
```

#### 4. スクリプト実行
```bash
cd /Users/itsuki.matsumoto/claude-code/tools/report-builder
python3 oasis-monthly/generate.py YYYYMM
```

ファイルを直接指定する場合:
```bash
python3 oasis-monthly/generate.py YYYYMM --campaign file1.xlsx --discount file2.xlsx --point file3.xlsx
```

#### 5. 出力確認
`oasis-monthly/output/YYYYMM/` に以下が生成される：
- `YYYYMM_【月次集計】キャンペーン販促費.xlsx`
- `YYYYMM_【月次集計】ポイント利用額.xlsx`
- `input/` - 入力ファイルのバックアップ（処理後に自動移動）

#### 6. スクショをinput/に追加
抽出条件のスクショを `output/YYYYMM/input/` に手動で追加：
```bash
cp ~/Downloads/抽出条件*.png "/Users/itsuki.matsumoto/claude-code/tools/report-builder/oasis-monthly/output/YYYYMM/input/"
```

#### 7. Dropboxへ同期（任意）
```bash
cp -r oasis-monthly/output/YYYYMM/* "/Users/itsuki.matsumoto/giftee Dropbox/Matsumoto Itsuki/giftee_Biz_template/00_biz_汎用/#sales汎用資料/g4b_Corporate Gift/02_プロダクト検討/02_gifteeBenefit/01_プロダクト概要/05_実績集計/YYYYMM/"
```

### ファイル自動検出ルール
input/フォルダのファイル名から自動判定：
- `商品×割引率` を含む → キャンペーン販売商品データ
- `割引率一覧` を含む → 基本割引率一覧
- `ポイント利用` を含む → ポイント利用状況

### ディレクトリ構成
```
oasis-monthly/
├── generate.py        # メインスクリプト
├── templates/         # Excelテンプレート
├── input/             # 入力ファイル（処理後は空になる）
└── output/
    └── YYYYMM/
        ├── YYYYMM_【月次集計】キャンペーン販促費.xlsx
        ├── YYYYMM_【月次集計】ポイント利用額.xlsx
        └── input/     # 入力ファイルのバックアップ + 抽出条件スクショ
```

### 依存パッケージ
```bash
pip3 install pandas openpyxl
```

## 共通ライブラリ（shared/）

| モジュール | 機能 |
|-----------|------|
| `chart_generator.py` | matplotlib でグラフ画像生成（棒グラフ、ドーナツグラフ） |
| `csv_utils.py` | CSV読み込み、案件グルーピング、ファイル自動検出 |

## 依存パッケージ

```bash
pip3 install python-pptx matplotlib pandas openpyxl lxml
```
