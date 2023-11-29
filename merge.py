import os
import pandas as pd

# 创建一个空的DataFrame，用于存储所有文件的第三个工作表
combined_df = pd.DataFrame()

# 获取目录下所有xlsx文件
dir_path = 'book_dir'
file_list = [f for f in os.listdir(dir_path) if f.endswith('.xlsx')]

# 逐个读取文件的第三个工作表，并添加到combined_df中
for file in file_list:
    file_path = os.path.join(dir_path, file)
    sheets_dict = pd.read_excel(file_path, sheet_name=None)
    sheet_names = list(sheets_dict.keys())

    # 确保文件至少有三个工作表
    if len(sheet_names) >= 3:
        sheet_df = pd.read_excel(file_path, sheet_name=sheet_names[2])

        # 添加空白行
        blank_row = pd.DataFrame([None] * sheet_df.shape[1]).T
        sheet_df = pd.concat([sheet_df, blank_row], ignore_index=True)

        combined_df = pd.concat([combined_df, sheet_df])

# 将combined_df保存为新的xlsx文件
output_file = 'book/test_3.xlsx'
combined_df.to_excel(output_file, sheet_name="漏洞列表", index=False)
