from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QComboBox, QLineEdit, QPushButton, QFormLayout, QHBoxLayout

class MappingDialog(QDialog):
    def __init__(self, parent=None, mode="add"):
        super().__init__(parent)
        self.mode = mode  # "add" 또는 "edit"
        if mode == "add":
            self.setWindowTitle("매핑 추가")
        else:
            self.setWindowTitle("매핑 수정")
        
        self.resize(500, 400)
        self.sheet_names = []
        self.main_layout = QVBoxLayout()
        self.setLayout(self.main_layout)
        self.setup_ui()

        # 버튼 시그널 연결
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

        # 조건부 활성화 로직
        self.method_combo.currentTextChanged.connect(self.update_fields_state)
        self.update_fields_state()

    def setup_ui(self):
        form_layout = QFormLayout()

        # 병합 방법
        self.method_combo = QComboBox()
        self.method_combo.addItems(["셀 합산", "키 기반", "행 누적", "최소값 선택", "최대값 선택"])
        form_layout.addRow("병합 방법:", self.method_combo)

        # 시트명
        self.sheet_combo = QComboBox()
        form_layout.addRow("시트명:", self.sheet_combo)

        # 데이터 범위
        self.range_input = QLineEdit()
        self.range_input.setPlaceholderText("예: C3:E10")
        form_layout.addRow("데이터 범위:", self.range_input)

        # 키 컬럼
        self.key_column_input = QLineEdit()
        self.key_column_input.setPlaceholderText("예: C")
        form_layout.addRow("키 컬럼:", self.key_column_input)

        # 인덱스 컬럼
        self.index_column_input = QLineEdit()
        self.index_column_input.setPlaceholderText("예: B")
        form_layout.addRow("행 순번 컬럼:", self.index_column_input)

        self.main_layout.addLayout(form_layout)

        # 버튼
        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("확인")
        self.cancel_button = QPushButton("취소")
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        self.main_layout.addLayout(button_layout)

    def update_fields_state(self):
        current_method = self.method_combo.currentText()
        is_key_based = current_method == "키 기반"
        self.key_column_input.setEnabled(is_key_based)
        self.index_column_input.setEnabled(is_key_based)

    def get_mapping(self) -> dict:
        return {
            "merge_method": self.method_combo.currentText(),
            "sheet_name": self.sheet_combo.currentText(),
            "data_range": self.range_input.text(),
            "key_column": self.key_column_input.text(),
            "index_column": self.index_column_input.text()
        }

    def set_mapping(self, mapping: dict) -> None:
        self.method_combo.setCurrentText(mapping.get("merge_method", "셀 합산"))
        self.sheet_combo.setCurrentText(mapping.get("sheet_name", ""))
        self.range_input.setText(mapping.get("data_range", ""))
        self.key_column_input.setText(mapping.get("key_column", ""))
        self.index_column_input.setText(mapping.get("index_column", ""))

    def set_sheet_names(self, sheet_names: list) -> None:
        self.sheet_names = sheet_names
        self.sheet_combo.clear()
        self.sheet_combo.addItems(sheet_names)