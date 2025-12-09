"""プロセッサーの公開エントリーポイント"""

import importlib
import inspect
import pkgutil
from pathlib import Path

from ..base_processor import BaseSheetProcessor

__all__ = []


def _load_processors_in_directory():
    """
    processors ディレクトリ直下のモジュールから BaseSheetProcessor を継承した
    クラスを動的に読み込み、モジュールレベルに公開する。
    """
    base_dir = Path(__file__).parent
    if not base_dir.exists():
        return

    for module_info in pkgutil.iter_modules([str(base_dir)]):
        # __init__ やパッケージはスキップ
        if module_info.name == "__init__" or module_info.ispkg:
            continue

        module_name = f"{__name__}.{module_info.name}"
        module = importlib.import_module(module_name)

        for name, obj in inspect.getmembers(module, inspect.isclass):
            if obj.__module__ != module_name:
                continue
            if not issubclass(obj, BaseSheetProcessor):
                continue

            globals()[name] = obj
            if name not in __all__:
                __all__.append(name)


_load_processors_in_directory()
