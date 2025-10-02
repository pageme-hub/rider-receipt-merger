from PySide6.QtWidgets import QWidget, QHBoxLayout, QLabel, QPushButton, QFileDialog, QVBoxLayout, QTextEdit
from PySide6.QtCore import Signal
from datetime import datetime

class FileSelectWidget(QWidget):
    file_selected = Signal(str)
    def __init__(self, label_text: str):
        super().__init__()
        
        self.file_path = ""
        
        # 위젯 생성
        label = QLabel(label_text)
        select_button = QPushButton("파일 선택")
        self.path_label = QLabel("선택된 파일 없음")
        
        # 레이아웃 구성
        layout = QHBoxLayout()
        layout.addWidget(label)
        layout.addWidget(select_button)
        layout.addWidget(self.path_label)
        
        self.setLayout(layout)
        
        # 시그널 연결
        select_button.clicked.connect(self.on_select_file)

    def on_select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "파일 선택", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.file_path = file_path
            self.path_label.setText(file_path)
            self.file_selected.emit(file_path)
        else:
            self.file_path = ""
            self.path_label.setText("선택된 파일 없음")

    def get_file_path(self) -> str:
        return self.file_path

class StatusDisplay(QWidget):
    def __init__(self):
        super().__init__()
        
        # QTextEdit 생성
        self.text_edit = QTextEdit()
        
        # 읽기 전용 설정
        self.text_edit.setReadOnly(True)
        
        # 최소 높이 설정
        self.text_edit.setMinimumHeight(150)
        
        # 레이아웃 구성
        layout = QVBoxLayout()
        layout.addWidget(self.text_edit)
        
        self.setLayout(layout)

    def add_message(self, message: str):
        # 타임스탬프 생성
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # 메시지 추가
        self.text_edit.append(f"[{timestamp}] {message}")

    def clear(self):
        # 텍스트 비우기
        self.text_edit.clear()

