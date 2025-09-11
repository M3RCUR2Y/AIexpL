import re
import io
import os
from datetime import datetime
import zipfile
import xml.etree.ElementTree as ET
import base64

def parse_chinese_date(date_str):
    """解析多种日期格式"""
    if not isinstance(date_str, str):
        return None
        
    date_str = date_str.strip()
    patterns = [
        (r'^(\d{4})/(\d{1,2})/(\d{1,2})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})$'),
        (r'^(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})$'),
        (r'^(\d{4})年(\d{1,2})月(\d{1,2})日$'),
        (r'^(\d{4})/(\d{1,2})/(\d{1,2})$'),
        (r'^(\d{4})-(\d{1,2})-(\d{1,2})$')
    ]
    
    for pattern in patterns:
        match = re.match(pattern, date_str)
        if match:
            groups = match.groups()
            year = int(groups[0])
            
            if '年' in date_str:
                month = int(groups[1])
                day = int(groups[2])
                hour = minute = second = 0
            elif '/' in date_str or '-' in date_str:
                month = int(groups[1])
                day = int(groups[2])
                hour = minute = second = 0
                if len(groups) > 3:
                    hour = int(groups[3])
                    minute = int(groups[4])
                    second = int(groups[5])
            else:
                continue
                
            return datetime(year=year, month=month, day=day, hour=hour, minute=minute, second=second)
    return None

def process_xlsx_content_memory(file_data):
    """在内存中处理xlsx文件内容"""
    processed_count = 0
    
    with zipfile.ZipFile(io.BytesIO(file_data), 'r') as zip_ref:
        file_list = zip_ref.namelist()
        modified_files = {}
        
        # 处理共享字符串文件
        shared_strings_name = 'xl/sharedStrings.xml'
        if shared_strings_name in file_list:
            content = zip_ref.read(shared_strings_name)
            root = ET.fromstring(content)
            
            for si in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
                t_elem = si.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                if t_elem is not None and t_elem.text:
                    parsed_dt = parse_chinese_date(t_elem.text)
                    if parsed_dt:
                        t_elem.text = parsed_dt.strftime("%Y-%m-%d %H:%M:%S")
                        processed_count += 1
            
            modified_files[shared_strings_name] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
        
        # 处理工作表文件
        for file_name in file_list:
            if file_name.startswith('xl/worksheets/') and file_name.endswith('.xml'):
                content = zip_ref.read(file_name)
                root = ET.fromstring(content)
                
                for c in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                    # 处理内联字符串
                    is_elem = c.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is')
                    if is_elem is not None:
                        t_elem = is_elem.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                        if t_elem is not None and t_elem.text:
                            parsed_dt = parse_chinese_date(t_elem.text)
                            if parsed_dt:
                                t_elem.text = parsed_dt.strftime("%Y-%m-%d %H:%M:%S")
                                processed_count += 1
                    
                    # 处理值元素
                    v_elem = c.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                    if v_elem is not None and v_elem.text and c.get('t') != 's':
                        parsed_dt = parse_chinese_date(v_elem.text)
                        if parsed_dt:
                            v_elem.text = parsed_dt.strftime("%Y-%m-%d %H:%M:%S")
                            processed_count += 1
                
                modified_files[file_name] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
        
        # 重新打包为内存中的ZIP文件
        output_buffer = io.BytesIO()
        with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as new_zip:
            for file_name in file_list:
                if file_name in modified_files:
                    new_zip.writestr(file_name, modified_files[file_name])
                else:
                    new_zip.writestr(file_name, zip_ref.read(file_name))
        
        return output_buffer.getvalue(), processed_count

def process_text_file_memory(content):
    """在内存中处理文本文件"""
    processed_count = 0
    lines = content.split('\n')
    modified = False
    
    for i, line in enumerate(lines):
        parts = re.split(r'[\t,;|]', line)
        line_modified = False
        
        for j, part in enumerate(parts):
            parsed_dt = parse_chinese_date(part.strip())
            if parsed_dt:
                parts[j] = parsed_dt.strftime("%Y-%m-%d %H:%M:%S")
                processed_count += 1
                line_modified = True
        
        if line_modified:
            if '\t' in line:
                lines[i] = '\t'.join(parts)
            elif ',' in line:
                lines[i] = ','.join(parts)
            elif ';' in line:
                lines[i] = ';'.join(parts)
            elif '|' in line:
                lines[i] = '|'.join(parts)
            else:
                lines[i] = ' '.join(parts)
            modified = True
    
    return '\n'.join(lines), processed_count

def get_file_data(file_info):
    """获取文件数据，支持多种输入格式"""
    # 方式1: 直接从path读取
    file_path = file_info.get('path', '')
    if file_path and os.path.exists(file_path):
        with open(file_path, 'rb') as f:
            return f.read()
    
    # 方式2: 从content字段读取（base64编码）
    content = file_info.get('content', '')
    if content:
        return base64.b64decode(content)
    
    # 方式3: 从data字段读取
    data = file_info.get('data', b'')
    if data:
        return data if isinstance(data, bytes) else data.encode('utf-8')
    
    # 方式4: 从url下载（如果有的话）
    url = file_info.get('url', '')
    if url:
        # 这里可以添加HTTP请求下载文件的逻辑
        # 但在沙箱环境中可能不被允许
        pass
    
    # 如果都没有，返回空数据
    return b''

def main(files):
    """Dify Code Node 主函数 - 修复版本"""
    result_files = []
    
    # 检查输入
    if not files or not isinstance(files, list):
        return {"result": []}
    
    for file_info in files:
        if not isinstance(file_info, dict):
            continue
            
        file_name = file_info.get('name', 'unknown_file')
        
        # 获取文件数据
        file_data = get_file_data(file_info)
        
        if not file_data:
            # 如果无法获取文件数据，返回原文件信息
            result_files.append({
                'name': f"processed_{file_name}",
                'error': 'No file data available',
                'type': file_info.get('type', 'application/octet-stream'),
                'size': 0
            })
            continue
        
        processed_data = None
        processed_count = 0
        
        # 根据文件类型选择处理方法
        if file_name.endswith(('.xlsx', '.xls')):
            processed_data, processed_count = process_xlsx_content_memory(file_data)
        elif file_name.endswith(('.txt', '.csv', '.tsv')):
            content = file_data.decode('utf-8')
            processed_content, processed_count = process_text_file_memory(content)
            processed_data = processed_content.encode('utf-8')
        else:
            # 尝试作为文本文件处理
            try:
                content = file_data.decode('utf-8')
                processed_content, processed_count = process_text_file_memory(content)
                processed_data = processed_content.encode('utf-8')
            except UnicodeDecodeError:
                # 如果不是文本文件，直接返回原数据
                processed_data = file_data
                processed_count = 0
        
        # 将处理后的数据编码为base64（Dify常用格式）
        processed_base64 = base64.b64encode(processed_data).decode('utf-8')
        
        result_files.append({
            'name': f"processed_{file_name}",
            'content': processed_base64,  # base64编码的内容
            'type': file_info.get('type', 'application/octet-stream'),
            'size': len(processed_data),
            'processed_count': processed_count
        })
    
    return {
        "result": result_files
    }