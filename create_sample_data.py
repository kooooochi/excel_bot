import pandas as pd
from datetime import datetime, timedelta
import random


def create_sample_data():
    """Create sample Excel data for testing."""

    # Sample 1: Sales data
    sales_data = {
        '日付': [datetime(2025, 1, i).strftime('%Y-%m-%d') for i in range(1, 11)],
        '商品名': ['商品A', '商品B', '商品C', '商品D', '商品E'] * 2,
        '販売数': [random.randint(10, 100) for _ in range(10)],
        '単価': [random.randint(500, 5000) for _ in range(10)],
        'カテゴリ': ['食品', '雑貨', '衣類', '電化製品', '書籍'] * 2,
    }

    df_sales = pd.DataFrame(sales_data)
    df_sales['売上金額'] = df_sales['販売数'] * df_sales['単価']
    df_sales.to_excel('input/sample_sales.xlsx', index=False)
    print("Created: input/sample_sales.xlsx")

    # Sample 2: Employee data
    employee_data = {
        '社員ID': [f'EMP{i:04d}' for i in range(1, 21)],
        '氏名': [f'社員{i}' for i in range(1, 21)],
        '部署': random.choices(['営業部', '技術部', '総務部', '人事部', '経理部'], k=20),
        '役職': random.choices(['一般', '主任', '課長', '部長'], weights=[10, 5, 3, 1], k=20),
        '入社年月日': [(datetime(2020, 1, 1) + timedelta(days=random.randint(0, 1825))).strftime('%Y-%m-%d') for _ in range(20)],
        '基本給': [random.randint(200, 800) * 1000 for _ in range(20)],
    }

    df_employee = pd.DataFrame(employee_data)
    df_employee.to_excel('input/sample_employees.xlsx', index=False)
    print("Created: input/sample_employees.xlsx")

    # Sample 3: Inventory data
    inventory_data = {
        '商品コード': [f'PRD{i:05d}' for i in range(1, 16)],
        '商品名': [f'製品{i}' for i in range(1, 16)],
        '在庫数': [random.randint(0, 500) for _ in range(15)],
        '発注点': [random.randint(50, 150) for _ in range(15)],
        '単価': [random.randint(100, 10000) for _ in range(15)],
        '倉庫': random.choices(['東京倉庫', '大阪倉庫', '福岡倉庫'], k=15),
    }

    df_inventory = pd.DataFrame(inventory_data)
    df_inventory['在庫金額'] = df_inventory['在庫数'] * df_inventory['単価']
    df_inventory['要発注'] = df_inventory['在庫数'] < df_inventory['発注点']
    df_inventory.to_excel('input/sample_inventory.xlsx', index=False)
    print("Created: input/sample_inventory.xlsx")

    print("\nAll sample files created successfully!")


if __name__ == "__main__":
    create_sample_data()
