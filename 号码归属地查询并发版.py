import pandas as pd
import requests
import re
import time
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
import tqdm  # 修改导入方式以避免模块调用错误


def clean_phone_number(phone):
    """清洗手机号码：去除非数字字符，返回纯数字"""
    if pd.isna(phone):
        return ""
    phone_str = str(phone)
    return re.sub(r"\D", "", phone_str)  # 保留纯数字


def get_phone_info(phone_number):
    """调用API查询单个手机号的归属地和运营商"""
    if len(phone_number) != 11 or not phone_number.isdigit():
        return ("无效手机号", "无效手机号")

    api_url = f"https://cx.shouji.360.cn/phonearea.php?number={phone_number}"

    try:
        response = requests.get(api_url, timeout=8)
        response.raise_for_status()
        result = response.json()

        if result.get("code") == 0:
            data = result.get("data", {})
            province = data.get("province", "")
            city = data.get("city", "")
            sp = data.get("sp", "")
            location = f"{province}{city}" if (province or city) else "未知地区"
            operator = sp if sp else "未知运营商"
            return (location, operator)
        else:
            return ("API查询失败", "API查询失败")

    except requests.exceptions.RequestException as e:
        return (f"网络错误: {str(e)[:15]}", f"网络错误: {str(e)[:15]}")
    except Exception as e:
        return (f"解析错误: {str(e)[:15]}", f"解析错误: {str(e)[:15]}")


def process_row(row):
    """处理单行数据：清洗手机号并查询信息（供线程调用）"""
    index, data = row
    original_phone = data["联系电话"]
    cleaned_phone = clean_phone_number(original_phone)
    location, operator = get_phone_info(cleaned_phone)
    return (index, location, operator)  # 返回索引和结果，用于后续写入


def batch_query_excel(excel_path, max_workers=10):
    """多线程批量处理Excel，max_workers控制并发数，结果保存到新文件"""
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_path, engine="openpyxl")

        # 验证表头
        required_columns = ["序号", "姓名", "性别", "民族", "联系电话", "归属地", "运营商"]
        if not all(col in df.columns for col in required_columns):
            print("❌ Excel表头不符合要求！需包含指定列")
            return

        total_rows = len(df)
        print(f"✅ 成功读取 {total_rows} 行数据，启动多线程查询（并发数：{max_workers}）...")

        # 创建线程池并提交任务
        results = []
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # 提交所有行的处理任务
            futures = [executor.submit(process_row, row) for row in df.iterrows()]

            # 实时获取结果并显示进度（修复tqdm调用方式）
            for future in tqdm.tqdm(as_completed(futures), total=total_rows, desc="处理进度"):
                results.append(future.result())

        # 将结果写回DataFrame（按索引排序，确保顺序正确）
        for index, location, operator in results:
            df.at[index, "归属地"] = location
            df.at[index, "运营商"] = operator

        # 生成新文件名，避免覆盖原文件
        file_dir, file_name = os.path.split(excel_path)
        file_base, file_ext = os.path.splitext(file_name)
        new_file_name = f"{file_base}_已查询{file_ext}"
        new_file_path = os.path.join(file_dir, new_file_name)

        # 保存结果到新文件
        df.to_excel(new_file_path, index=False, engine="openpyxl")
        print(f"\n🎉 全部完成！结果已保存至新文件：{new_file_path}")

    except FileNotFoundError:
        print(f"❌ 未找到文件：{excel_path}")
    except Exception as e:
        print(f"❌ 程序出错：{str(e)}")


if __name__ == "__main__":
    # 安装进度条库（首次运行需执行）
    try:
        import tqdm
    except ImportError:
        print("正在安装进度条工具...")
        os.system("pip install tqdm")
        import tqdm

    excel_path = input("请输入Excel文件路径：").strip()
    # 可根据网络情况调整并发数（建议5-20之间）
    max_workers_input = input("请输入并发数（建议5-20）：").strip()
    max_workers = int(max_workers_input) if max_workers_input else 10
    batch_query_excel(excel_path, max_workers)
