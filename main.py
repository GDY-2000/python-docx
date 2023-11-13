import random
from utils.xlsx_opera import get_xlsx_value
from utils.docx_opera import generate_report

match_key = ['company_name', 'name', 'phone']
table_Array = get_xlsx_value(match_key, file_url="enquiry.xlsx")

# 使用 random.shuffle() 函数来打乱数据顺序
random.shuffle(table_Array)

# 初始化一个空的列表，用于存储拆分后的子数组
split_data = []

# 使用循环将原始数据分为长度为2的子数组
for i in range(0, len(table_Array), 2):
    split_data.append(table_Array[i:i+2])

# 打印拆分后的子数组
for index, sub_array in enumerate(split_data):
    generate_report(data={'data': sub_array}, template_path='enquiry.docx', output_path=f'outdocx/report{index}.docx', open_file=False)
