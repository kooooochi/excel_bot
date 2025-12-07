"""Excel Processor Library - 汎用的なExcel処理フレームワーク"""

from .core import ExcelProcessor
from .base_processor import BaseSheetProcessor
from .utils import (
    load_excel_from_input,
    get_excel_files,
    save_preview,
    print_sheet_info,
    print_sheet_preview
)

__all__ = [
    'ExcelProcessor',
    'BaseSheetProcessor',
    'load_excel_from_input',
    'get_excel_files',
    'save_preview',
    'print_sheet_info',
    'print_sheet_preview'
]
__version__ = '0.1.0'
