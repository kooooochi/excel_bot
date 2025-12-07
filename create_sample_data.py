#!/usr/bin/env python3
"""サンプルExcelデータ作成スクリプト"""

from pathlib import Path
import pandas as pd
import numpy as np
from datetime import datetime, timedelta


def create_sample_employees():
    """従業員データのサンプルを作成"""
    data = {
        '社員ID': [f'E{i:04d}' for i in range(1, 21)],
        '氏名': [
            '田中太郎', '佐藤花子', '鈴木一郎', '高橋美咲', '伊藤健太',
            '渡辺真由美', '山本大輔', '中村さくら', '小林隆', '加藤愛',
            '吉田拓也', '山田麻衣', '佐々木翔', '井上優子', '木村洋平',
            '林香織', '斎藤雄太', '松本美穂', '清水直樹', '森田恵子'
        ],
        '部署': np.random.choice(['営業', '開発', '人事', '総務', '経理'], 20),
        '役職': np.random.choice(['部長', '課長', '係長', '一般'], 20, p=[0.1, 0.2, 0.3, 0.4]),
        '年齢': np.random.randint(25, 60, 20),
        '入社日': [(datetime(2020, 1, 1) - timedelta(days=int(x))).strftime('%Y-%m-%d')
                  for x in np.random.randint(0, 3650, 20)],
        '給与': np.random.randint(250000, 800000, 20),
    }

    df = pd.DataFrame(data)
    return df


def create_sample_sales():
    """売上データのサンプルを作成"""
    dates = pd.date_range('2024-01-01', '2024-12-31', freq='D')
    data = {
        '日付': dates,
        '商品名': np.random.choice(['商品A', '商品B', '商品C', '商品D', '商品E'], len(dates)),
        '販売数': np.random.randint(1, 100, len(dates)),
        '単価': np.random.choice([1000, 1500, 2000, 2500, 3000], len(dates)),
        '地域': np.random.choice(['東京', '大阪', '名古屋', '福岡', '札幌'], len(dates)),
        '担当者': np.random.choice(['田中', '佐藤', '鈴木', '高橋', '伊藤'], len(dates)),
    }

    df = pd.DataFrame(data)
    df['売上金額'] = df['販売数'] * df['単価']
    return df


def create_sample_inventory():
    """在庫データのサンプルを作成"""
    data = {
        '商品コード': [f'P{i:03d}' for i in range(1, 51)],
        '商品名': [f'製品{chr(65 + (i-1) // 10)}{(i-1) % 10 + 1}' for i in range(1, 51)],
        'カテゴリ': np.random.choice(['電子機器', '家具', '文具', '食品', '衣類'], 50),
        '在庫数': np.random.randint(0, 500, 50),
        '単価': np.random.randint(100, 50000, 50),
        '仕入先': np.random.choice(['仕入先A', '仕入先B', '仕入先C', '仕入先D'], 50),
        '最終入荷日': [(datetime.now() - timedelta(days=int(x))).strftime('%Y-%m-%d')
                    for x in np.random.randint(0, 90, 50)],
    }

    df = pd.DataFrame(data)
    df['在庫金額'] = df['在庫数'] * df['単価']
    return df


def main():
    """サンプルデータを作成してinputディレクトリに保存"""
    input_dir = Path('input')
    input_dir.mkdir(exist_ok=True)

    print("Creating sample Excel files...")

    # 従業員データ
    df_employees = create_sample_employees()
    employees_file = input_dir / 'sample_employees.xlsx'
    df_employees.to_excel(employees_file, index=False, sheet_name='従業員一覧')
    print(f"✓ Created: {employees_file}")

    # 売上データ
    df_sales = create_sample_sales()
    sales_file = input_dir / 'sample_sales.xlsx'
    df_sales.to_excel(sales_file, index=False, sheet_name='売上データ')
    print(f"✓ Created: {sales_file}")

    # 在庫データ
    df_inventory = create_sample_inventory()
    inventory_file = input_dir / 'sample_inventory.xlsx'
    df_inventory.to_excel(inventory_file, index=False, sheet_name='在庫管理')
    print(f"✓ Created: {inventory_file}")

    # 複数シートのサンプル
    multi_sheet_file = input_dir / 'sample_multi_sheet.xlsx'
    with pd.ExcelWriter(multi_sheet_file, engine='openpyxl') as writer:
        df_employees.head(10).to_excel(writer, sheet_name='従業員', index=False)
        df_sales.head(100).to_excel(writer, sheet_name='売上', index=False)
        df_inventory.head(20).to_excel(writer, sheet_name='在庫', index=False)
    print(f"✓ Created: {multi_sheet_file}")

    print(f"\nAll sample files created in '{input_dir}' directory!")
    print("\nNext steps:")
    print("  1. Run: python run_processor.py")
    print("  2. Or open: develop_processor.ipynb")


if __name__ == '__main__':
    main()
