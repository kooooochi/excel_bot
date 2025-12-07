#!/usr/bin/env python3
"""Excel Processor メイン実行スクリプト"""

import sys
import argparse
from pathlib import Path
import yaml

from excel_processor import ExcelProcessor
from excel_processor.processors import SummarySheetProcessor, FormatProcessor


def load_config(config_path: str = "config.yaml") -> dict:
    """設定ファイルを読み込む"""
    config_file = Path(config_path)

    if not config_file.exists():
        print(f"Warning: Config file '{config_path}' not found.")
        print("Using default configuration.")
        return get_default_config()

    with open(config_file, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)

    return config


def get_default_config() -> dict:
    """デフォルト設定を返す"""
    return {
        'input_dir': 'input',
        'output_dir': 'output',
        'processors': [
            {
                'name': 'SummarySheetProcessor',
                'enabled': True,
                'config': {
                    'sheet_name': 'Summary',
                    'position': 0
                }
            },
            {
                'name': 'FormatProcessor',
                'enabled': True,
                'config': {
                    'header_color': '4472C4',
                    'font_name': 'Arial',
                    'font_size': 11,
                    'apply_borders': True,
                    'auto_width': True,
                    'exclude_sheets': ['Summary']
                }
            }
        ]
    }


def create_processor_instance(processor_config: dict):
    """設定からプロセッサーインスタンスを作成"""
    processor_name = processor_config['name']
    config = processor_config.get('config', {})

    # 組み込みプロセッサー
    BUILTIN_PROCESSORS = {
        'SummarySheetProcessor': SummarySheetProcessor,
        'FormatProcessor': FormatProcessor,
    }

    if processor_name in BUILTIN_PROCESSORS:
        return BUILTIN_PROCESSORS[processor_name](config)

    # カスタムプロセッサーの動的ロード
    # TODO: カスタムプロセッサーの動的読み込み機能を実装
    raise ValueError(f"Unknown processor: {processor_name}")


def main():
    parser = argparse.ArgumentParser(description='Excel Processor - 汎用的なExcel処理ツール')
    parser.add_argument(
        '-c', '--config',
        default='config.yaml',
        help='設定ファイルのパス（デフォルト: config.yaml）'
    )
    parser.add_argument(
        '-i', '--input-dir',
        help='入力ディレクトリ（設定ファイルの値を上書き）'
    )
    parser.add_argument(
        '-o', '--output-dir',
        help='出力ディレクトリ（設定ファイルの値を上書き）'
    )

    args = parser.parse_args()

    # 設定を読み込む
    config = load_config(args.config)

    # コマンドライン引数で上書き
    input_dir = args.input_dir or config.get('input_dir', 'input')
    output_dir = args.output_dir or config.get('output_dir', 'output')

    # プロセッサーを作成
    processors = []
    for proc_config in config.get('processors', []):
        if proc_config.get('enabled', True):
            try:
                processor = create_processor_instance(proc_config)
                processors.append(processor)
                print(f"Loaded processor: {proc_config['name']}")
            except Exception as e:
                print(f"Error loading processor {proc_config['name']}: {e}")
                sys.exit(1)

    if not processors:
        print("Warning: No processors configured. Files will be moved without processing.")

    # ExcelProcessorを実行
    print(f"\n{'='*60}")
    print("Excel Processor Starting...")
    print(f"{'='*60}\n")

    processor = ExcelProcessor(
        input_dir=input_dir,
        output_dir=output_dir,
        processors=processors
    )

    processor.run()

    print(f"\n{'='*60}")
    print("Excel Processor Completed!")
    print(f"{'='*60}")


if __name__ == '__main__':
    main()
