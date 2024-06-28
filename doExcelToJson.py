import argparse
import json
import os

import pandas as pd


def converter_excel_to_json(excel_path, json_path, sheet_name=None):
    # 获取所有工作表名称
    excel = pd.ExcelFile(excel_path)
    sheet_names = excel.sheet_names

    # 如果未指定工作表名称，则默认使用第一个工作表
    if sheet_name is None:
        sheet_name = sheet_names[0]

    df = pd.read_excel(excel_path, header=None)

    # 提取注释行（第一行）
    comment_row = df.iloc[0]

    head_row = df.columns = df.iloc[1]  # 第二行作为表头

    # 将索引和注释合并为一个字典
    index_comment_dict = dict(zip(head_row, comment_row))

    # 将注释行转换为 JSON
    json_data_head = json.dumps(index_comment_dict, ensure_ascii=False)

    df = df[2:]  # 从第三行开始的数据

    # 重置索引
    df.reset_index(drop=True, inplace=True)

    # 将整个 DataFrame 数据转换为 JSON
    json_data = df.to_json(orient='records', force_ascii=False)

    # 合并json
    combined_json = {
        "header": json.loads(json_data_head),
        "data": json.loads(json_data)
    }

    # 将 JSON 数据写入文件，确保使用 UTF-8 编码
    json_file_path = os.path.join(json_path, sheet_name+".json")
    with open(json_file_path, 'w', encoding='utf-8') as json_file:
        json.dump(combined_json, json_file, ensure_ascii=False, indent=2)

    print(f"成功将 Excel 文件 '{excel_path}' 的页签 '{sheet_name}' 转换为 JSON 文件 '{json_file_path}'")
    """
    将excel文件转换为json并存入到文件中
    """


def process_directory(excel_dir, json_dir):
    # 确保 JSON 文件目录存在
    os.makedirs(json_dir, exist_ok=True)

    # 遍历目录中的所有文件
    for filename in os.listdir(excel_dir):
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            excel_file = os.path.join(excel_dir, filename)
            json_path = os.path.join(json_dir)

            # 调用函数进行转换
            converter_excel_to_json(excel_file, json_path)


if __name__ == "__main__":
    # 设置命令行参数
    parser = argparse.ArgumentParser(description='Convert Excel file to JSON file')
    parser.add_argument('excel_file', type=str, help='Path to the Excel file')
    parser.add_argument('json_file', type=str, help='Path to save the JSON file')

    # 解析命令行参数
    args = parser.parse_args()

    # 调用函数进行转换
    process_directory(args.excel_file, args.json_file)
