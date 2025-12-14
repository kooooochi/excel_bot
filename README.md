# Excel Processor

ローカル実行とGitHub Actionsによる自動処理に対応。
柔軟でカスタマイズ可能なExcel処理ライブラリ。
エクセルを使った定型作業をGithub Actionsで自動化するサンプルレポジトリです。

## 概要

`input/`ディレクトリのExcelファイルを処理し、タイムスタンプ付きディレクトリ（`output/YYYY-MM-DD_HHMMSS/`）に結果を出力します。

## 主な機能

- **汎用的な処理フレームワーク**: プラグイン形式で様々な処理を追加可能
- **YAML設定**: 処理内容を設定ファイルで柔軟に制御
- **動的ロード**: `excel_processor/processors/` に配置したクラスを自動で読み込み
- **カスタムプロセッサー**: 独自の処理ロジックを簡単に実装
- **タイムスタンプ管理**: 処理結果を日時別に自動整理

## クイックスタート

### 1. 基本的な使い方

```bash
# Excelファイルをinputディレクトリに配置
cp your_file.xlsx input/

# 処理を実行（デフォルト設定を使用）
python run_processor.py

# 結果を確認
ls output/
```

処理結果は `output/YYYY-MM-DD_HHMMSS/` ディレクトリに保存されます。

### 2. 設定をカスタマイズ

[config.yaml](config.yaml)を編集して処理内容を変更できます:

```yaml
processors:
  - name: "SummarySheetProcessor"
    enabled: true
    config:
      sheet_name: "Summary"
      position: 0

  - name: "FormatProcessor"
    enabled: true
    config:
      header_color: "4472C4"
```

詳細は [LIBRARY_GUIDE.md](LIBRARY_GUIDE.md) を参照してください。

## Dockerで実行

```bash
# コンテナを起動
docker-compose up -d

# Excelファイルを配置
cp your_file.xlsx input/

# 処理を実行
docker-compose exec excel_bot python run_processor.py

# サンプルデータで試す
docker-compose exec excel_bot python create_sample_data.py
docker-compose exec excel_bot python run_processor.py
```

## GitHub Actionsで自動処理

### 動作条件

以下の条件を**すべて**満たす場合に自動実行されます:

- **ブランチ名**: `process/`で始まるブランチからのPR（例: `process/update-data`、`process/feature-1`）
- **変更ファイル**: `input/`ディレクトリ内の`.xlsx`または`.xls`ファイルが変更されている

### 使用手順

```bash
# 1. process/で始まるブランチを作成
git checkout -b process/update-data

# 2. Excelファイルを追加
git add input/your_file.xlsx
git commit -m "Add Excel file for processing"
git push origin process/update-data

# 3. PRを作成
# GitHub上でPull Requestを作成
```

### 処理結果の確認

- **Artifacts**: GitHub ActionsのArtifactsタブからダウンロード（30日間保存）
- **コミット**: 処理済みファイルが`output/`ディレクトリに自動コミットされます

## カスタムプロセッサーの作成

独自の処理ロジックを実装できます。詳細は [LIBRARY_GUIDE.md](LIBRARY_GUIDE.md) を参照してください。

### 簡単な例

```python
# excel_processor/processors/my_processor.py
from openpyxl.workbook import Workbook
from excel_processor.base_processor import BaseSheetProcessor

class MyProcessor(BaseSheetProcessor):
    def process(self, workbook: Workbook, file_path: str) -> Workbook:
        # カスタム処理を実装
        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            # 処理内容...

        return workbook
```

`excel_processor/processors/` に置くだけで自動で読み込まれます。[config.yaml](config.yaml)の`processors[].name`にクラス名を記載して有効化してください。

## 技術スタック

- Python 3.11
- pandas, openpyxl, numpy, tqdm, pyyaml
- Docker / Docker Compose
- GitHub Actions

## ドキュメント

- [LIBRARY_GUIDE.md](LIBRARY_GUIDE.md) - ライブラリの詳細ガイド
- [config.yaml](config.yaml) - 設定ファイルのサンプル
- [process_excel.ipynb](process_excel.ipynb) - openpyxlのチュートリアル

## ディレクトリ構造

```text
excel_bot/
├── excel_processor/          # ライブラリ本体
│   ├── __init__.py
│   ├── core.py              # メインエンジン
│   ├── base_processor.py    # ベースクラス
│   └── processors/          # プロセッサー
│       ├── summary_sheet.py   # サンプル
│       └── format_processor.py # サンプル
├── input/                   # 入力ファイル
├── output/                  # 出力ファイル（タイムスタンプ別）
│   └── YYYY-MM-DD_HHMMSS/
├── config.yaml              # 設定ファイル
├── run_processor.py         # 実行スクリプト
└── LIBRARY_GUIDE.md         # ライブラリガイド
```
