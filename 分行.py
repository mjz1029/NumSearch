import pandas as pd
import re
import os


def find_phone_column(df):
    """自动查找可能的联系电话列"""
    phone_column_candidates = [col for col in df.columns if '电话' in col or '手机号' in col or '联系' in col]
    return phone_column_candidates


def select_phone_column(df):
    """让用户选择联系电话列"""
    print("请选择联系电话所在的列：")
    for i, col in enumerate(df.columns):
        print(f"{i + 1}. {col}")

    while True:
        try:
            choice = int(input("请输入列号（1-{}）: ".format(len(df.columns)))) - 1
            if 0 <= choice < len(df.columns):
                return df.columns[choice]
            else:
                print(f"请输入1到{len(df.columns)}之间的数字")
        except ValueError:
            print("请输入有效的数字")


def split_name(name, num_parts):
    """根据需要拆分姓名，适用于有多个姓名的情况"""
    # 确保name是字符串类型
    if pd.isna(name):
        name = ""
    else:
        name = str(name)

    if num_parts == 1 or not name:
        return [name]

    # 处理包含换行符的姓名
    if '\n' in name:
        parts = name.split('\n')
        # 确保拆分后的部分数量与电话号码数量一致
        if len(parts) == num_parts:
            return [f"{part}\n" if i < len(parts) - 1 else part for i, part in enumerate(parts)]
        else:
            # 如果拆分数量不匹配，尝试平均分配
            mid = len(name) // num_parts
            return [name[:mid] + '\n', name[mid:]]
    else:
        # 如果没有换行符，平均拆分
        mid = len(name) // num_parts
        return [name[:mid] + '\n', name[mid:]]


def process_phone_numbers(df, phone_column):
    """处理电话号码，将过长的号码拆分为多行"""
    processed_rows = []

    # 正则表达式匹配11位手机号码
    phone_pattern = re.compile(r'1\d{10}')

    for _, row in df.iterrows():
        phone_data = str(row[phone_column])

        # 提取所有11位电话号码
        phones = phone_pattern.findall(phone_data)

        if len(phones) == 1:
            # 只有一个电话号码，直接保留
            processed_rows.append(row)
        elif len(phones) >= 2:
            # 有多个电话号码，拆分为多行
            # 处理姓名列（如果存在）
            name = row.get('姓名', '') if '姓名' in df.columns else ''
            split_names = split_name(name, len(phones))

            for i, phone in enumerate(phones[:2]):  # 只处理前两个号码
                new_row = row.copy()
                new_row[phone_column] = phone

                # 更新姓名
                if '姓名' in df.columns:
                    new_row['姓名'] = split_names[i] if i < len(split_names) else name

                processed_rows.append(new_row)
        else:
            # 没有找到有效的电话号码，保留原数据
            processed_rows.append(row)

    return pd.DataFrame(processed_rows)


def main():
    # 支持的文件格式
    supported_formats = {'.csv', '.xlsx', '.xls'}

    # 获取输入文件路径
    while True:
        input_path = input("请输入要处理的文件路径: ").strip()
        if os.path.exists(input_path):
            file_ext = os.path.splitext(input_path)[1].lower()
            if file_ext in supported_formats:
                break
            else:
                print(f"不支持的文件格式，请提供以下格式的文件: {', '.join(supported_formats)}")
        else:
            print("文件不存在，请重新输入")

    # 读取文件
    print("正在读取文件...")
    try:
        if file_ext == '.csv':
            df = pd.read_csv(input_path)
        else:  # .xlsx 或 .xls
            df = pd.read_excel(input_path)
    except Exception as e:
        print(f"读取文件失败: {str(e)}")
        return

    print("文件读取成功，数据预览:")
    print(df.head())

    # 确定联系电话列
    phone_candidates = find_phone_column(df)
    if len(phone_candidates) == 1:
        phone_column = phone_candidates[0]
        print(f"自动识别联系电话列为: {phone_column}")
    else:
        print("未找到明确的联系电话列或找到多个候选列")
        phone_column = select_phone_column(df)

    # 处理数据
    print("正在处理数据...")
    processed_df = process_phone_numbers(df, phone_column)

    # 生成输出文件路径
    base, ext = os.path.splitext(input_path)
    output_path = f"{base}_processed{ext}"

    # 保存处理后的数据
    print(f"正在保存处理后的数据到 {output_path}...")
    try:
        if file_ext == '.csv':
            processed_df.to_csv(output_path, index=False)
        else:  # .xlsx 或 .xls
            processed_df.to_excel(output_path, index=False)
        print("数据处理完成并保存成功！")
    except Exception as e:
        print(f"保存文件失败: {str(e)}")
        return


if __name__ == "__main__":
    main()
