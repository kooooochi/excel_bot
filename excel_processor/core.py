"""コアエンジン - Excel処理のメインロジック"""

import os
import sys
import shutil
from pathlib import Path
from datetime import datetime
from typing import List, Type
import openpyxl
from tqdm import tqdm

from .base_processor import BaseSheetProcessor


class ExcelProcessor:
    """
    Excel処理のメインクラス

    inputディレクトリのExcelファイルを処理し、
    タイムスタンプ付きディレクトリにoutputとして保存します。
    """

    def __init__(
        self,
        input_dir: str = "input",
        output_dir: str = "output",
        processors: List[BaseSheetProcessor] = None
    ):
        """
        Args:
            input_dir: 入力ファイルのディレクトリ
            output_dir: 出力先のベースディレクトリ
            processors: 適用するプロセッサーのリスト
        """
        self.input_dir = Path(input_dir)
        self.output_base_dir = Path(output_dir)
        self.processors = processors or []
        self.timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        self.output_dir = self.output_base_dir / self.timestamp

    def run(self):
        """処理のメイン実行"""
        # 入力ファイルの検出
        excel_files = self._find_excel_files()

        if not excel_files:
            print("No Excel files found in input directory.")
            return

        print(f"Found {len(excel_files)} Excel file(s) to process.")
        print(f"Output directory: {self.output_dir}")

        # 出力ディレクトリを作成
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # 各ファイルを処理
        for input_file in tqdm(excel_files, desc="Processing files"):
            try:
                self._process_file(input_file)
            except Exception as e:
                print(f"\nError processing {input_file}: {e}")
                import traceback
                traceback.print_exc()
                sys.exit(1)

        print(f"\nAll files processed successfully!")
        print(f"Output saved to: {self.output_dir}")

    def _find_excel_files(self) -> List[Path]:
        """inputディレクトリからExcelファイルを検索"""
        if not self.input_dir.exists():
            print(f"Error: Input directory '{self.input_dir}' does not exist.")
            sys.exit(1)

        excel_files = list(self.input_dir.glob("*.xlsx")) + list(self.input_dir.glob("*.xls"))

        # 一時ファイルを除外
        excel_files = [f for f in excel_files if not f.name.startswith("~$")]

        return excel_files

    def _process_file(self, input_file: Path):
        """単一のExcelファイルを処理"""
        print(f"\nProcessing: {input_file.name}")

        # Excelファイルを読み込み
        workbook = openpyxl.load_workbook(input_file)

        # 各プロセッサーを適用
        for processor in self.processors:
            try:
                workbook = processor.process(workbook, str(input_file))
            except Exception as e:
                print(f"Error in processor {processor.__class__.__name__}: {e}")
                raise

        # 出力ファイルに保存
        output_file = self.output_dir / input_file.name
        workbook.save(output_file)
        print(f"Saved: {output_file.name}")

        # 元のファイルを出力ディレクトリに移動
        shutil.move(str(input_file), str(output_file))
        print(f"Moved: {input_file.name} -> {self.output_dir}")

    def add_processor(self, processor: BaseSheetProcessor):
        """プロセッサーを追加"""
        self.processors.append(processor)
