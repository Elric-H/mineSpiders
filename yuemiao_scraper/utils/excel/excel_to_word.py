# main.py
import os
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
from word_processor import WordProcessor
from collections import defaultdict
from typing import List, Dict, Any

# 配置部分
EXCEL_PATH = 'output_summary.xlsx'
TEMPLATE_PATH = 'template.docx'
OUTPUT_DIR = './words'
MAX_WORKERS = 4  # 并行处理的线程数


def read_excel(file_path: str) -> pd.DataFrame:
    """读取Excel文件并保留数字格式"""
    # 读取时将所有列作为字符串处理，保留原始格式
    df = pd.read_excel(
        file_path,
        sheet_name=0,
        dtype=str,  # 所有列作为字符串读取
        na_values=['', 'NA', 'N/A'],  # 自定义NA值
        keep_default_na=False  # 不将空字符串等自动转为NaN
    )

    # 替换真正的NaN值为空字符串
    return df.fillna('')


def aggregate_data(df: pd.DataFrame) -> List[Dict[str, Any]]:
    """聚合数据，按照前四列分组"""
    grouped = defaultdict(list)

    # 确保列名正确，假设前四列名为A,B,C,D
    columns = df.columns.tolist()
    if len(columns) < 4:
        raise ValueError("Excel文件必须至少包含4列数据")

    for _, row in df.iterrows():
        # 直接使用原始字符串值，不做类型转换
        key = tuple(row.iloc[:4])
        grouped[key].append(row.iloc[4:].tolist())

    result = []
    for (a, b, c, d), rows in grouped.items():
        result.append({
            'a': a,
            'b': b,
            'c': c,
            'd': d,
            'rows': rows
        })
    return result


def generate_word_files(data: List[Dict[str, Any]], template_path: str, output_dir: str):
    """生成Word文件，使用并行处理"""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    wp = WordProcessor(template_path)
    name_counter = {}

    def process_item(item):
        # 生成文件名，处理重复情况
        base_name = f"{item['b']}-{item['a']}"
        if base_name in name_counter:
            name_counter[base_name] += 1
            file_name = f"{base_name}(重复{name_counter[base_name]}).docx"
        else:
            name_counter[base_name] = 0
            file_name = f"{base_name}.docx"

        output_path = os.path.join(output_dir, file_name)
        print(f"正在生成: {file_name}。。。")
        wp.generate_document(output_path, item)
        print(f"{file_name}创建成功！")

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        executor.map(process_item, data)


def main():
    print("开始处理...")

    # 读取Excel数据
    print("读取Excel文件...")
    df = read_excel(EXCEL_PATH)

    # 聚合数据
    aggregated_data = aggregate_data(df)

    # 生成Word文件
    print(f"生成Word文件到 {OUTPUT_DIR}...")
    generate_word_files(aggregated_data, TEMPLATE_PATH, OUTPUT_DIR)

    print(f"处理完成，共生成 {len(aggregated_data)} 个Word文件")


if __name__ == "__main__":
    main()
    print("程序执行完毕！")
    input("按 Enter 键退出...")
