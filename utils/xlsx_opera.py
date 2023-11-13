import pandas as pd

def get_xlsx_value(match_key, file_url="enquiry.xlsx"):
    """
    :param match_key: 需要匹配的列名字
    :param file_url: xlsx文件的地址路径

    return value: xlsx文件生成的字典数据
    """
    # 读取 Excel 文件
    df = pd.read_excel(file_url)
    # 创建对象列表
    people = []

    df.dropna(inplace=True)

    # 遍历 Excel 数据框的每一行，将每行的数据整理为对象
    for index, row in df.iterrows():
        person = {
            match_key[0]: row[match_key[0]],
             match_key[1]: row[ match_key[1]],
             match_key[2]: row[ match_key[2]]
        }
        people.append(person)
    return people