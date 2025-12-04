import re
from typing import Tuple

def validate_excel_range(range_str: str) -> Tuple[bool, str]:
    # 단일 셀인 경우 범위로 변환
    if ':' not in range_str:
        range_str = range_str + ":" + range_str
    
    # 패턴 검증
    pattern = r'^[A-Z]+[0-9]+:[A-Z]+[0-9]+$'
    if re.match(pattern, range_str):
        return (True, range_str)
    else:
        return (False, "")

def validate_column(column_str: str) -> bool:
    pattern = r'^[A-Z]+$'
    return bool(re.match(pattern, column_str))

def is_valid_range_format(range_str: str) -> Tuple[bool, str]:
    is_valid, converted_range = validate_excel_range(range_str)
    if not is_valid:
        return False, "올바른 엑셀 범위 형식이 아닙니다. 예: A1:C10"
    
    return True, converted_range

def is_valid_column_format(column_str: str) -> Tuple[bool, str]:
    if not validate_column(column_str):
        return False, "올바른 컬럼 형식이 아닙니다. 예: A, B, AA"
    
    return True, ""