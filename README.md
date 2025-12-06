# Excel Processing with GitHub Actions

Excelファイルを自動処理するシステム。ローカル実行とGitHub Actionsによる自動処理に対応。

## 概要

`input/`ディレクトリのExcelファイルを処理し、結果を`output/`ディレクトリに出力します。

## ローカルで実行

### 前提条件
- Docker
- Docker Compose

### 実行手順

```bash
# 1. コンテナを起動
docker-compose up -d

# 2. Excelファイルをinputディレクトリに配置
cp your_file.xlsx input/

# 3. 処理を実行
docker-compose exec python-app python process_excel.py

# 4. 結果を確認
ls output/
```

### サンプルデータで試す

```bash
docker-compose exec python-app python create_sample_data.py
docker-compose exec python-app python process_excel.py
```

## GitHub Actionsで自動処理

### 動作条件

PRで`input/`ディレクトリ内の`.xlsx`または`.xls`ファイルが変更されたとき自動実行されます。

### 使用手順

```bash
# 1. ブランチを作成
git checkout -b feature/process-data

# 2. Excelファイルを追加
git add input/your_file.xlsx
git commit -m "Add Excel file for processing"
git push origin feature/process-data

# 3. PRを作成
# GitHub上でPull Requestを作成
```

### 処理結果の確認

- **Artifacts**: GitHub ActionsのArtifactsタブからダウンロード（30日間保存）
- **コミット**: 処理済みファイルが`output/`ディレクトリに自動コミットされます

## カスタマイズ

処理ロジックを変更する場合は`process_excel.py`の`process_excel_file`関数を編集してください。

```python
def process_excel_file(input_path, output_path):
    df = pd.read_excel(input_path)

    # ここに処理を追加
    df['new_column'] = df['existing_column'] * 2

    df.to_excel(output_path, index=False)
```

## 技術スタック

- Python 3.11
- pandas, openpyxl, numpy, tqdm
- Docker / Docker Compose
- GitHub Actions
