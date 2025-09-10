import datetime
import dateutil.parser
from dateutil.tz import tzlocal

def format_date(original_text, possible_formats):
    """将日期字符串格式化为yyyy-MM-dd hh:mm:ss格式"""
    try:
        # 尝试使用dateutil自动解析
        parsed_date = dateutil.parser.parse(original_text, fuzzy=True)
        
        # 如果没有时间信息，添加默认时间00:00:00
        if parsed_date.hour == 0 and parsed_date.minute == 0 and parsed_date.second == 0:
            formatted_date = parsed_date.strftime("%Y-%m-%d 00:00:00")
        else:
            formatted_date = parsed_date.strftime("%Y-%m-%d %H:%M:%S")
            
        return {
            "original_text": original_text,
            "formatted_text": formatted_date,
            "status": "success"
        }
    except Exception as e:
        return {
            "original_text": original_text,
            "formatted_text": None,
            "status": "failed",
            "error": str(e)
        }

def main(inputs):
    # 获取提取的日期列表
    extracted_dates = inputs.get('extracted_dates', [])
    
    # 格式化每个日期
    formatted_dates = []
    for date_info in extracted_dates:
        formatted = format_date(
            date_info.get('original_text'),
            date_info.get('possible_formats', [])
        )
        # 添加位置信息
        formatted['position'] = date_info.get('position')
        formatted_dates.append(formatted)
    
    return {
        'formatted_dates': formatted_dates,
        'failed_count': sum(1 for d in formatted_dates if d['status'] == 'failed')
    }
