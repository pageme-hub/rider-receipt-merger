from enum import Enum
from dataclasses import dataclass
from typing import Optional

DEFAULT_PASSWORD = "8016000315"
DEFAULT_CONFIG_FILE = "mapping_config.json"
LOG_FILE = "error.log"


class MergeMethod(Enum):
    CELL_SUM = "셀 합산"
    KEY_BASED = "키 기반"
    ROW_APPEND = "행 누적"
    CELL_MIN = "최소값 선택"
    CELL_MAX = "최대값 선택"

@dataclass
class MappingConfig:
    merge_method: MergeMethod
    sheet_name: str
    data_range: str
    key_column: Optional[str] = None
    index_column: Optional[str] = None
    min_cells: Optional[str] = None
    max_cells: Optional[str] = None

    def to_dict(self):
        return {
            "merge_method": self.merge_method.value,
            "sheet_name": self.sheet_name,
            "data_range": self.data_range,
            "key_column": self.key_column,
            "index_column": self.index_column,

        }

    @classmethod
    def from_dict(cls, data: dict):
        return cls(
            merge_method=MergeMethod(data["merge_method"]),
            sheet_name=data["sheet_name"],
            data_range=data["data_range"],
            key_column=data.get("key_column"),
            index_column=data.get("index_column")
        )