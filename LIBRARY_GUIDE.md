# Excel Processor ライブラリガイド

Excel Processorは、Excelファイルを柔軟に処理できる汎用的なPythonライブラリです。

## 特徴

- **汎用的**: シート追加、書式設定、データ処理など、様々な処理をプラグイン形式で実装可能
- **カスタマイズ可能**: ユーザー独自の処理ロジックを簡単に追加
- **YAML設定**: 処理内容を設定ファイルで柔軟に制御
- **タイムスタンプ管理**: 処理結果を `YYYY-MM-DD_HHMMSS` 形式のディレクトリに自動保存

## 基本的な使い方

### 1. 設定ファイルの準備

[config.yaml](config.yaml)を編集して、適用したい処理を設定します。

```yaml
input_dir: "input"
output_dir: "output"

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
      font_name: "Arial"
      font_size: 11
```

### 2. 処理の実行

```bash
# 設定ファイルを使用して実行
python run_processor.py

# 設定ファイルを指定して実行
python run_processor.py -c custom_config.yaml

# 入力・出力ディレクトリを指定して実行
python run_processor.py -i input -o output
```

### 3. 処理の流れ

1. `input/`ディレクトリからExcelファイルを検出
2. 各ファイルに対して設定されたプロセッサーを順番に適用
3. 処理結果を`output/YYYY-MM-DD_HHMMSS/`ディレクトリに保存
4. 元のファイルを`input/`から移動

## 組み込みプロセッサー

### SummarySheetProcessor

サマリーシートを追加します。

**設定例:**
```yaml
- name: "SummarySheetProcessor"
  enabled: true
  config:
    sheet_name: "Summary"  # シート名
    position: 0  # 挿入位置（0=先頭）
```

**機能:**
- 処理日時
- ファイル名
- シート一覧と各シートの行数・列数

### FormatProcessor

全シートに書式を適用します。

**設定例:**
```yaml
- name: "FormatProcessor"
  enabled: true
  config:
    header_color: "4472C4"  # ヘッダー背景色
    font_name: "Arial"  # フォント名
    font_size: 11  # フォントサイズ
    apply_borders: true  # 罫線を適用
    auto_width: true  # 列幅を自動調整
    exclude_sheets: ["Summary"]  # 除外するシート
```

## カスタムプロセッサーの作成

独自の処理ロジックを実装できます。

### 1. プロセッサークラスの作成

[excel_processor/processors/](excel_processor/processors/)ディレクトリに新しいファイルを作成します。

```python
# excel_processor/processors/my_custom_processor.py

from openpyxl.workbook import Workbook
from excel_processor.base_processor import BaseSheetProcessor


class MyCustomProcessor(BaseSheetProcessor):
    """カスタムプロセッサーの例"""

    def process(self, workbook: Workbook, file_path: str) -> Workbook:
        # 設定値を取得
        my_value = self.config.get('my_setting', 'default')

        self.log(f"Processing with value: {my_value}")

        # Excelファイルの処理ロジックを実装
        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]

            # 例: 全セルに何か処理を行う
            for row in ws.iter_rows():
                for cell in row:
                    # 処理内容
                    pass

        return workbook
```

### 2. プロセッサーの登録

[run_processor.py](run_processor.py)の`BUILTIN_PROCESSORS`に追加します。

```python
from excel_processor.processors.my_custom_processor import MyCustomProcessor

BUILTIN_PROCESSORS = {
    'SummarySheetProcessor': SummarySheetProcessor,
    'FormatProcessor': FormatProcessor,
    'MyCustomProcessor': MyCustomProcessor,  # 追加
}
```

### 3. 設定ファイルで有効化

[config.yaml](config.yaml)に追加します。

```yaml
processors:
  - name: "MyCustomProcessor"
    enabled: true
    config:
      my_setting: "custom_value"
```

## カスタムプロセッサーの例

### 例1: 特定の列を削除するプロセッサー

```python
class DeleteColumnProcessor(BaseSheetProcessor):
    def process(self, workbook: Workbook, file_path: str) -> Workbook:
        target_sheet = self.config.get('target_sheet')
        column_to_delete = self.config.get('column', 1)

        if target_sheet and target_sheet in workbook.sheetnames:
            ws = workbook[target_sheet]
            ws.delete_cols(column_to_delete)
            self.log(f"Deleted column {column_to_delete} from {target_sheet}")

        return workbook
```

### 例2: データを集計するプロセッサー

```python
class AggregationProcessor(BaseSheetProcessor):
    def process(self, workbook: Workbook, file_path: str) -> Workbook:
        source_sheet = self.config.get('source_sheet', 'Sheet1')
        output_sheet = self.config.get('output_sheet', 'Aggregated')

        if source_sheet in workbook.sheetnames:
            ws_source = workbook[source_sheet]
            ws_output = self.get_or_create_sheet(workbook, output_sheet)

            # 集計ロジックを実装
            # 例: 列の合計を計算
            ws_output['A1'] = "Total"
            ws_output['B1'] = f"=SUM({source_sheet}!B:B)"

            self.log(f"Aggregated data from {source_sheet}")

        return workbook
```

## ヘルパーメソッド

`BaseSheetProcessor`が提供するヘルパーメソッド:

- `create_sheet(workbook, sheet_name, index=None)`: 新しいシートを作成
- `get_or_create_sheet(workbook, sheet_name)`: シートを取得、なければ作成
- `log(message)`: ログを出力

## ベストプラクティス

1. **エラーハンドリング**: 処理が失敗してもファイルが壊れないように注意
2. **設定の検証**: `config`から値を取得する際はデフォルト値を設定
3. **ログ出力**: `self.log()`を使って処理状況を記録
4. **不変性**: 可能な限り元のデータを保持しながら処理

## トラブルシューティング

### プロセッサーが読み込まれない

- プロセッサー名が正しいか確認
- `BUILTIN_PROCESSORS`に登録されているか確認
- インポート文が正しいか確認

### 設定が反映されない

- YAML形式が正しいか確認（インデントなど）
- `enabled: true`が設定されているか確認

### ファイルが移動されない

- `input/`ディレクトリにExcelファイルが存在するか確認
- 一時ファイル（`~$`で始まるファイル）は無視されます
