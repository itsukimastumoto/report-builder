# JCBレポート自動生成ガイド

JCB成果報告レポート（PPTX）を自動生成するツールです。
ダッシュボードからCSVを4つダウンロードして、Claude Codeに渡すだけで全クライアント分のスライドが完成します。

---

## 初回セットアップ（1回だけ）

### 1. report-builderフォルダを取得

Dropboxの以下からダウンロードしてデスクトップ等に配置してください。

```
giftee_Biz_template/00_biz_汎用/#sales汎用資料/g4b_Corporate Gift/02_プロダクト/02_gifteeBenefit/01_プロダクト概要/06_便利ツール/report-builder
```

> Dropbox同期済みならそのまま使えます

### 2. Python依存パッケージをインストール

ターミナルで以下を実行します。

```
pip3 install python-pptx matplotlib pandas openpyxl lxml
```

> Python3未インストールの場合は `brew install python3` を先に実行

---

## 毎月のレポート作成手順

### Step 1: CSVをダウンロード

[ダッシュボード](https://polaris.giftee.dev/dashboards/303--oasis-jcb-)を開き、**集計期間**と**クライアント（JCB案件）**を選択して、以下の4つのCSVをダウンロードします。

- ログインユーザー数推移（週次）
- 購入ユーザー数の推移
- クライアント別ブランド販売状況一覧
- oasisサマリ（JCB用）

### Step 2: Claude Codeで生成

ターミナルでreport-builderフォルダに移動してClaude Codeを起動します。

```
cd ~/Desktop/report-builder
claude
```

起動したら、こう話しかけるだけです:

```
JCBレポートを生成して。対象月は202603。CSVは以下:
（FinderからCSVファイルを4つドラッグ&ドロップ）
```

> **Tips**: FinderからCSVをターミナルにドラッグ&ドロップするとパスが自動入力されます

あとはClaude Codeが自動で処理してPPTXが完成します。

### Step 3: 確認

生成されたPPTXをPowerPointで開いて確認してください。
やり直したい場合は「もう一回生成して」と伝えればOKです。

---

## 困ったときは

| 症状 | 対処 |
| --- | --- |
| `command not found: python3` | `brew install python3` を実行 |
| `ModuleNotFoundError` | `pip3 install python-pptx matplotlib pandas openpyxl lxml` を再実行 |
| CSVが見つからないエラー | ファイル名にキーワード（ログイン/購入/ブランド/サマリ）が含まれているか確認 |
| それでも解決しない | Matsumotoまで連絡してください |
