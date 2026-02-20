# レポート生成ツール（Report Builder）

giftee Benefitの各種レポートをClaude Codeで自動生成するツールです。データソースを渡して話しかけるだけでレポートが完成します。

現在対応しているレポート:

| レポート | 出力形式 | 用途 |
|---------|---------|------|
| **JCB成果報告レポート** | PPTX | クライアントごとの月次報告スライド |
| OASIS月次レポート | Excel | キャンペーン販促費・ポイント利用額の月次集計 |

---

## 初回セットアップ

### 1. Claude Codeのインストール

ターミナル（Macの「ターミナル」アプリ）を開いて以下を実行します。

```bash
npm install -g @anthropic-ai/claude-code
```

> **Node.jsが入っていない場合**: https://nodejs.org/ からLTS版をダウンロード&インストールしてから再実行してください。

### 2. report-builderフォルダを取得

Dropboxの以下のパスからフォルダをダウンロードしてデスクトップなど好きな場所に置いてください。

```
giftee_Biz_template/00_biz_汎用/#sales汎用資料/g4b_Corporate Gift/02_プロダクト/02_gifteeBenefit/01_プロダクト概要/06_便利ツール/report-builder
```

> Dropbox同期済みの場合はそのまま使ってもOKです。

### 3. Python依存パッケージのインストール

ターミナルで以下を実行します。

```bash
pip3 install python-pptx matplotlib pandas openpyxl lxml
```

> **Python3が入っていない場合**: `brew install python3` を実行するか、https://www.python.org/ からインストールしてください。

### 4. セットアップ確認

```bash
cd ~/Desktop/report-builder   # report-builderフォルダの場所に合わせて変更
claude
```

Claude Codeが起動すれば準備完了です。`Ctrl+C` で終了できます。

---

## JCB成果報告レポートの作成手順

CSチームが毎月実施する主要なレポート作成フローです。

### Step 1: CSVをダウンロード

[[oasis]JCBレポート、インプット用](https://polaris.giftee.dev/dashboards/303--oasis-jcb-) ダッシュボードを開き、以下を設定してCSVをダウンロードします。

1. **集計期間（集計月）** を対象の月に設定
2. **クライアント** をJCB案件に設定
3. 以下の **4つのCSV** をダウンロード

| # | CSV名 | 内容 |
|---|--------|------|
| 1 | ログインユーザー数推移（週次） | 週次ログインユーザー数 |
| 2 | 購入ユーザー数の推移 | 週次購入数 |
| 3 | クライアント別ブランド販売状況一覧 | ブランド別の販売実績 |
| 4 | oasisサマリ（JCB用） | クライアント別の集約データ |

### Step 2: Claude Codeで生成

ターミナルで `report-builder/` フォルダに移動してClaude Codeを起動します。

```bash
cd ~/Desktop/report-builder
claude
```

起動したら、以下のように話しかけてください。

```
JCBレポートを生成して。対象月は202603。CSVは以下:
（ここにダウンロードした4つのCSVのパスを貼り付け）
```

> **Tips**: FinderからCSVファイルをターミナルにドラッグ&ドロップするとパスが自動入力されます。4つのファイルを順番にドロップすればOKです。

あとはClaude Codeが自動で処理します:
1. CSVを読み込み&バックアップ
2. クライアントごとのスライドを生成（グラフ・テーブル含む）
3. 全スライドを1つのPPTXに統合
4. 完成ファイルのパスを案内

### Step 3: 確認

生成されたPPTXをPowerPointで開いて確認してください。

```
jcb/tasks/YYYYMMDD/output/JCB報告資料_YYYYMM.pptx
```

やり直したい場合はそのまま「もう一回生成して」と伝えればOKです。実行ごとに独立したフォルダが作られるので、前回の結果が上書きされることはありません。

---

## よくある質問

**Q: 対象月を間違えた / CSVを差し替えたい**
A: 正しいCSVを用意して「CSVを差し替えて再生成して」と伝えればOKです。

**Q: テンプレートのデザインを変えたい**
A: `jcb/template/報告資料_v3.pptx` をPowerPointで直接編集できます。レイアウトを変えても自動検出で対応します。

**Q: 特定のクライアントだけ再生成したい**
A: 現時点では全クライアント一括生成です。生成後にPowerPoint上で不要なスライドを削除してください。

---

## トラブルシューティング

| 症状 | 対処 |
|------|------|
| `command not found: claude` | `npm install -g @anthropic-ai/claude-code` を再実行 |
| `command not found: npm` | https://nodejs.org/ からNode.jsをインストール |
| `command not found: python3` | `brew install python3` を実行 |
| `ModuleNotFoundError` | `pip3 install python-pptx matplotlib pandas openpyxl lxml` を再実行 |
| CSVが見つからないエラー | ファイル名にキーワード（ログイン/購入/ブランド/サマリ）が含まれているか確認 |
| PowerPointで表示が崩れる | ファイルを一度閉じてから再度開く |

---

## 運用ルール・技術仕様

各レポートの詳細な運用ルール（CSV仕様、生成ロジック、設定項目など）は以下を参照してください。

- [JCBレポート運用ルール](jcb/README.md)
