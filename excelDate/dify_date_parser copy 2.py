import re
import tempfile
import shutil
import os
from datetime import datetime
import pandas as pd

def parse_chinese_date(date_str):
    """
    解析多种日期格式，并返回对应的 datetime 对象。

    支持的格式包括：
    - 2025/9/11 8:11:11
    - 2025-09-11 08:11:22
    - 2025年9月11日
    - 2025/9/11
    - 2025-09-11

    :param date_str: 输入的日期字符串
    :return: datetime.datetime 或 None (如果无法匹配)
    """
    if not isinstance(date_str, str):
        return None

    date_str = date_str.strip()

    # 定义多种日期格式的正则表达式
    patterns = [
        # 格式: 2025/9/11 8:11:11
        (r'^(\d{4})/(\d{1,2})/(\d{1,2})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})$'),
        # 格式: 2025-09-11 08:11:22
        (r'^(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})$'),
        # 格式: 2025年9月11日
        (r'^(\d{4})年(\d{1,2})月(\d{1,2})日$'),
        # 格式: 2025/9/11
        (r'^(\d{4})/(\d{1,2})/(\d{1,2})$'),
        # 格式: 2025-09-11
        (r'^(\d{4})-(\d{1,2})-(\d{1,2})$')
    ]

    # 尝试每种格式
    for pattern in patterns:
        match = re.match(pattern, date_str)
        if match:
            try:
                # 提取年月日时分秒
                groups = match.groups()
                year = int(groups[0])

                if '年' in date_str:
                    # 中文格式: YYYY年MM月DD日
                    month = int(groups[1])
                    day = int(groups[2])
                    hour = 0
                    minute = 0
                    second = 0
                    if len(groups) > 3:
                        hour = int(groups[3])
                        minute = int(groups[4])
                        second = int(groups[5])
                elif '/' in date_str or '-' in date_str:
                    # 数字格式: YYYY/MM/DD[ HH:MM:SS] 或 YYYY-MM-DD[ HH:MM:SS]
                    month = int(groups[1])
                    day = int(groups[2])
                    hour = 0
                    minute = 0
                    second = 0
                    if len(groups) > 3:
                        hour = int(groups[3])
                        minute = int(groups[4])
                        second = int(groups[5])
                else:
                    continue

                dt = datetime(year=year, month=month, day=day, hour=hour, minute=minute, second=second)
                return dt
            except ValueError:
                continue

    return None

def convert_excel_dates_pandas(file_path):
    """
    使用pandas处理Excel文件中的日期格式转换

    :param file_path: Excel 文件路径
    :return: 处理结果统计
    """
    processed_count = 0
    
    try:
        # 读取Excel文件的所有工作表
        excel_file = pd.ExcelFile(file_path)
        
        # 创建一个字典来存储处理后的数据
        processed_sheets = {}
        
        for sheet_name in excel_file.sheet_names:
            # 读取工作表，保持原始数据类型
            df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
            
            # 遍历所有单元格
            for col in df.columns:
                for idx in df.index:
                    cell_value = df.at[idx, col]
                    
                    if pd.notna(cell_value) and isinstance(cell_value, str):
                        # 尝试解析日期
                        parsed_dt = parse_chinese_date(cell_value)
                        if parsed_dt:
                            # 转换为标准格式
                            formatted_date = parsed_dt.strftime("%Y-%m-%d %H:%M:%S")
                            df.at[idx, col] = formatted_date
                            processed_count += 1
            
            processed_sheets[sheet_name] = df
        
        # 保存处理后的Excel文件
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sheet_name, df in processed_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        return processed_count
        
    except Exception as e:
        # 如果pandas也不可用，尝试使用纯Python处理
        return convert_excel_dates_simple(file_path)

def convert_excel_dates_simple(file_path):
    """
    简化版本：只处理CSV格式或使用基础库
    """
    try:
        # 尝试将Excel转换为CSV进行处理
        df = pd.read_excel(file_path)
        processed_count = 0
        
        # 处理所有列
        for col in df.columns:
            for idx in df.index:
                cell_value = df.at[idx, col]
                if pd.notna(cell_value) and isinstance(cell_value, str):
                    parsed_dt = parse_chinese_date(cell_value)
                    if parsed_dt:
                        df.at[idx, col] = parsed_dt.strftime("%Y-%m-%d %H:%M:%S")
                        processed_count += 1
        
        # 保存为Excel
        df.to_excel(file_path, index=False)
        return processed_count
        
    except Exception:
        return 0

def main(files):
    """
    Dify Code Node 主函数
    处理输入的Excel文件数组，转换其中的日期格式
    
    :param files: Array[File] - 输入的文件数组
    :return: Array[File] - 处理后的文件数组
    """
    
    result = []
    
    try:
        # 处理每个输入文件
        for file_info in files:
            if not file_info.get('name', '').endswith(('.xlsx', '.xls')):
                # 跳过非Excel文件，但仍然添加到结果中
                result.append(file_info)
                continue
                
            file_name = file_info.get('name', 'unknown.xlsx')
            file_path = file_info.get('path', '')
            
            if not file_path or not os.path.exists(file_path):
                # 文件不存在，添加原文件信息到结果
                result.append(file_info)
                continue
            
            try:
                # 创建临时文件进行处理
                temp_dir = tempfile.mkdtemp()
                temp_file_path = os.path.join(temp_dir, file_name)
                
                # 复制原文件到临时位置
                shutil.copy2(file_path, temp_file_path)
                
                # 处理日期格式
                processed_count = convert_excel_dates_pandas(temp_file_path)
                
                # 创建处理后的文件信息
                processed_file = {
                    'name': f"processed_{file_name}",
                    'path': temp_file_path,
                    'type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    'size': os.path.getsize(temp_file_path)
                }
                
                result.append(processed_file)
                
            except Exception as e:
                # 处理失败，添加原文件到结果
                result.append(file_info)
                continue
    
    except Exception as e:
        # 发生错误时返回原文件列表
        return files
    
    return result