import pandas as pd
import requests
import re
import time


def clean_phone_number(phone):
    """
    清洗手机号码：去除非数字字符（空格、横线、括号等），返回纯数字
    :param phone: 原始手机号码（字符串/数值类型）
    :return: 清洗后的纯数字手机号（字符串）
    """
    if pd.isna(phone):  # 处理空值
        return ""
    # 转为字符串后，只保留数字
    phone_str = str(phone)
    cleaned = re.sub(r"\D", "", phone_str)  # 正则匹配非数字字符并删除
    return cleaned


def get_phone_info(phone_number):
    """
    调用360手机号归属地API，获取归属地（省份+城市）和运营商
    :param phone_number: 清洗后的11位纯数字手机号
    :return: (归属地, 运营商) 元组，查询失败时返回对应提示
    """
    # 1. 验证手机号有效性（清洗后应为11位数字）
    if len(phone_number) != 11 or not phone_number.isdigit():
        return ("无效手机号", "无效手机号")

    # 2. 构造API请求URL（使用你指定的360 API）
    api_url = f"https://cx.shouji.360.cn/phonearea.php?number={phone_number}"

    try:
        # 发送GET请求（添加0.5秒延迟，避免高频请求被限制）
        time.sleep(0.5)
        response = requests.get(api_url, timeout=10)  # 超时时间10秒
        response.raise_for_status()  # 若HTTP状态码非200（如404、500），抛出异常

        # 3. 解析API返回的JSON数据
        result = response.json()

        # 4. 提取归属地和运营商（判断API返回是否正常）
        if result.get("code") == 0:  # code=0表示查询成功
            data = result.get("data", {})
            province = data.get("province", "")  # 省份（如"新疆"）
            city = data.get("city", "")  # 城市（如"阿克苏"）
            sp = data.get("sp", "")  # 运营商（如"电信"）
            location = f"{province}{city}" if (province and city) else "未知地区"
            operator = sp if sp else "未知运营商"
            return (location, operator)
        else:
            # API返回错误（如code≠0）
            return ("API查询失败", "API查询失败")

    except requests.exceptions.RequestException as e:
        # 捕获网络异常（超时、连接失败等）
        return (f"网络错误: {str(e)[:20]}", f"网络错误: {str(e)[:20]}")
    except Exception as e:
        # 捕获其他未知异常
        return (f"解析错误: {str(e)[:20]}", f"解析错误: {str(e)[:20]}")


def batch_query_excel(excel_path):
    """
    批量处理Excel：读取手机号码，查询归属地，写入F列（归属地）和G列（运营商）
    :param excel_path: Excel文件路径（如"./phone_list.xlsx"）
    """
    try:
        # 1. 读取Excel文件（使用openpyxl引擎，支持写入）
        # 假设表头为：序号、姓名、性别、民族、联系电话、归属地、运营商（对应列A-G）
        df = pd.read_excel(excel_path, engine="openpyxl")

        # 2. 验证Excel表头是否符合要求
        required_columns = ["序号", "姓名", "性别", "民族", "联系电话", "归属地", "运营商"]
        if not all(col in df.columns for col in required_columns):
            print("❌ Excel表头不符合要求！需包含：序号、姓名、性别、民族、联系电话、归属地、运营商")
            return

        # 3. 批量处理每一行的手机号码
        print(f"✅ 成功读取Excel，共{len(df)}行数据，开始查询归属地...")
        for index, row in df.iterrows():
            # 获取当前行的手机号码并清洗
            original_phone = row["联系电话"]
            cleaned_phone = clean_phone_number(original_phone)

            # 查询归属地和运营商
            location, operator = get_phone_info(cleaned_phone)

            # 写入F列（归属地）和G列（运营商）
            df.at[index, "归属地"] = location
            df.at[index, "运营商"] = operator

            # 打印进度（每10行打印一次，避免输出过多）
            if (index + 1) % 10 == 0 or (index + 1) == len(df):
                print(
                    f"进度：{index + 1}/{len(df)} 行完成 | 手机号：{cleaned_phone} → 归属地：{location}，运营商：{operator}")

        # 4. 保存处理后的Excel文件（覆盖原文件，建议先备份原文件）
        df.to_excel(excel_path, index=False, engine="openpyxl")
        print(f"\n🎉 处理完成！文件已保存至：{excel_path}")

    except FileNotFoundError:
        print(f"❌ 未找到文件：{excel_path}，请检查路径是否正确")
    except Exception as e:
        print(f"❌ 程序运行出错：{str(e)}")


# ------------------- 程序入口 -------------------
if __name__ == "__main__":
    # 提示用户输入Excel文件路径（示例：./phone_list.xlsx 或 C:/data/phone.xlsx）
    excel_path = input("请输入Excel文件的完整路径（例如：./phone_list.xlsx）：").strip()

    # 启动批量处理
    batch_query_excel(excel_path)