import pandas as pd
import numpy as np
from faker import Faker
import os
import glob

# 创建Faker实例
fake = Faker()

# 指定包含CSV文件的文件夹路径
input_directory = "all_payout"  # 修改为你的CSV文件夹路径
output_directory = "output"

# 确保输出目录存在
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# 列出文件夹内所有CSV文件
csv_files = glob.glob(os.path.join(input_directory, "*.csv"))

# 为每个CSV文件执行处理步骤
for file_path in csv_files:
    # 读取CSV文件
    df = pd.read_csv(file_path)

    # 处理步骤
    df["Order"] = df["Order"].str.replace("#", "")
    df["Order"] = "11" + df["Order"]
    df.rename(columns={"Order": "Order ID"}, inplace=True)
    df["Transaction Date"] = pd.to_datetime(df["Transaction Date"], utc=True).dt.strftime('%m/%d/%Y')
    df.rename(columns={"Transaction Date": "Paid Date"}, inplace=True)
    df["Net"] = df["Net"].abs()
    df.rename(columns={"Net": "Order Total"}, inplace=True)
    df.rename(columns={"Currency": "Currency Code"}, inplace=True)
    df["Buyer Name/Buyer ID"] = [fake.name() for _ in range(len(df))]
    df["Shipping Address"] = [fake.address() for _ in range(len(df))]
    df["Product Title"] = [fake.catch_phrase() for _ in range(len(df))]
    df["Product Quantity"] = np.random.randint(1, 6, size=len(df))

    # 选择特定的列创建新的DataFrame
    columns_to_include = ["Order ID", "Paid Date", "Order Total", "Currency Code", "Buyer Name/Buyer ID",
                          "Shipping Address", "Product Title", "Product Quantity"]
    new_df = df[columns_to_include]

    # 构建输出文件路径，使用与输入文件相同的基本名称但扩展名为.xlsx
    base_name = os.path.basename(file_path)
    output_file_name = os.path.splitext(base_name)[0] + ".xlsx"
    xlsx_file_path = os.path.join(output_directory, output_file_name)

    # 使用openpyxl引擎保存为.xlsx格式
    new_df.to_excel(xlsx_file_path, index=False, engine='openpyxl')

    print(f"File saved at {xlsx_file_path}")
