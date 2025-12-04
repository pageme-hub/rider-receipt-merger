import json
import os
from typing import List, Dict
from config.settings import MappingConfig, DEFAULT_CONFIG_FILE
from utils.logger import log_error

def load_mappings(file_path: str) -> List[Dict]:
    try:
        if not os.path.exists(file_path):
            return []
        
        with open(file_path, 'r', encoding='utf-8') as file:
            mappings = json.load(file)
            return mappings
    except Exception as e:
        log_error(f"매핑 로드 실패: {str(e)}")
        return []

def save_mappings(file_path: str, mappings: List[Dict]) -> None:
    try:
        with open(file_path, 'w', encoding='utf-8') as file:
            json.dump(mappings, file, indent=4, ensure_ascii=False)
    except Exception as e:
        log_error(f"매핑 저장 실패: {str(e)}")
        raise

def mappings_to_dict_list(mappings: List[MappingConfig]) -> List[Dict]:
    result = []
    for mapping in mappings:
        result.append(mapping.to_dict())
    return result

def dict_list_to_mappings(dict_list: List[Dict]) -> List[MappingConfig]:
    result = []
    for config_dict in dict_list:
        result.append(MappingConfig.from_dict(config_dict))
    return result