import pandas as pd
import os


path = 'C:/Users/jun1yin5/vsCode/Transposed/'

files = os.listdir(path)

excel_files = [file for file in files if file.endswith('.xlsx') or file.endswith('.xls')]

transposed_data_list = []
for excel_file in excel_files:
    full_path = os.path.join(path,excel_file)
    df = pd.read_excel(full_path)
    transposed_df = df.T
    transposed_data_list.append((excel_file,transposed_df))
with pd.ExcelWriter('转置后汇总.xlsx') as writer:
    for excel_file, transposed_df in transposed_data_list:
        transposed_df.to_excel(writer,sheet_name=excel_file,index=False)
print('转置后的数据已保存到“转置汇总.xlsx”')