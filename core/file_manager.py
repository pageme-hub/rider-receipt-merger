import openpyxl
import msoffcrypto
import shutil
from io import BytesIO
import os  # 추가된 import 문
from typing import List
from utils.logger import log_error
from config.settings import DEFAULT_PASSWORD
from typing import Tuple, List

def decrypt_file(file_path: str, password: str) -> openpyxl.Workbook:
    try:
        with open(file_path, 'rb') as file:
            crypto = msoffcrypto.OfficeFile(file)
            crypto.load_key(password=password)
            decrypted = BytesIO()
            crypto.decrypt(decrypted)
            decrypted.seek(0)
            return openpyxl.load_workbook(decrypted)
    except Exception as e:
        log_error(f"파일 복호화 실패: {file_path}, {str(e)}")
        raise

def save_workbook(workbook: openpyxl.Workbook, file_path: str, password: str) -> None:
    try:
        with open(file_path, 'wb') as file:
            workbook.save(file)
            # 암호화 로직이 필요하지 않으므로 제거
    except Exception as e:
        log_error(f"파일 저장 실패: {file_path}, {str(e)}")
        raise

def get_sheet_names(file_path: str, password: str) -> List[str]:
    try:
        workbook = decrypt_file(file_path, password)
        return workbook.sheetnames
    except Exception as e:
        log_error(f"시트 목록 조회 실패: {file_path}, {str(e)}")
        return []
    
def extract_date_from_filename(file_path: str) -> Tuple[str, str]:
    """
    파일명에서 날짜 범위 추출
    입력: "20250827~20250831_부릉_부산광역시동구부릉_DP2503138816.xlsx"
    출력: ("20250827", "20250831")
    """
    import re
    import os
    
    filename = os.path.basename(file_path)
    pattern = r'(\d{8})~(\d{8})_'
    match = re.search(pattern, filename)
    
    if match:
        return (match.group(1), match.group(2))
    else:
        raise ValueError(f"파일명에서 날짜를 찾을 수 없습니다: {filename}")
    
def create_merged_filename(file1_path: str, file2_path: str) -> str:
    """
    두 파일의 날짜를 병합한 새 파일명 생성
    file1: 20250827~20250831_부릉_부산광역시동구부릉_DP2503138816.xlsx
    file2: 20250901~20250902_부릉_부산광역시동구부릉_DP2503138816.xlsx
    결과: 20250827~20250902_부릉_부산광역시동구부릉_DP2503138816.xlsx
    """
    import re
    import os
    
    # 두 파일에서 날짜 추출
    start1, end1 = extract_date_from_filename(file1_path)
    start2, end2 = extract_date_from_filename(file2_path)
    
    # 최소 시작일과 최대 종료일 선택
    merged_start = min(start1, start2)
    merged_end = max(end1, end2)
    
    # file1의 파일명에서 날짜 외 부분 추출
    filename1 = os.path.basename(file1_path)
    pattern = r'\d{8}~\d{8}(_.*)'
    match = re.search(pattern, filename1)
    
    if match:
        suffix = match.group(1)  # "_부릉_부산광역시동구부릉_DP2503138816.xlsx"
    else:
        raise ValueError(f"파일명 형식이 올바르지 않습니다: {filename1}")
    
    # 새 파일명 조합
    new_filename = f"{merged_start}~{merged_end}{suffix}"
    
    # file1과 같은 디렉토리에 생성
    new_filepath = os.path.join(os.path.dirname(file1_path), new_filename)
    
    return new_filepath