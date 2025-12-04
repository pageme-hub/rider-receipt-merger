from PySide6.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QPushButton, QHBoxLayout, QListWidget, QLabel, QMessageBox
from ui.ui_components import FileSelectWidget, StatusDisplay
from core.file_manager import get_sheet_names
from config.settings import DEFAULT_PASSWORD
from ui.mapping_dialog import MappingDialog
from core.mapping_manager import load_mappings, save_mappings, mappings_to_dict_list, dict_list_to_mappings
from config.settings import DEFAULT_CONFIG_FILE, MappingConfig
import os

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("엑셀 파일 병합 프로그램")
        self.resize(800, 600)
        self.file1_path = ""
        self.file2_path = ""
        self.mappings = []
        self.current_mapping_index = 0
        self.setup_ui()

        # 버튼 시그널 연결
        self.file1_widget.file_selected.connect(self.on_file1_selected)
        self.file2_widget.file_selected.connect(self.on_file2_selected)

        self.start_button.clicked.connect(self.on_start_button_clicked)
        self.add_mapping_button.clicked.connect(self.on_add_mapping_clicked)
        self.edit_mapping_button.clicked.connect(self.on_edit_mapping_clicked)
        self.delete_mapping_button.clicked.connect(self.on_delete_mapping_clicked)
        
        # 초기 매핑 목록 로드
        self.refresh_mapping_list()

    def setup_ui(self):
        central_widget = QWidget()
        self.main_layout = QVBoxLayout()
        central_widget.setLayout(self.main_layout)
        self.setCentralWidget(central_widget)

        # 파일 선택 위젯
        self.file1_widget = FileSelectWidget("File 1:")
        self.file2_widget = FileSelectWidget("File 2:")
        self.main_layout.addWidget(self.file1_widget)
        self.main_layout.addWidget(self.file2_widget)

        # 매핑 목록 영역
        mapping_section_layout = QVBoxLayout()

        mapping_label = QLabel("저장된 매핑 목록:")
        mapping_section_layout.addWidget(mapping_label)

        self.mapping_list_widget = QListWidget()
        self.mapping_list_widget.setMaximumHeight(150)
        mapping_section_layout.addWidget(self.mapping_list_widget)

        # 매핑 관리 버튼
        mapping_button_layout = QHBoxLayout()
        self.add_mapping_button = QPushButton("추가")
        self.edit_mapping_button = QPushButton("수정")
        self.delete_mapping_button = QPushButton("삭제")
        mapping_button_layout.addWidget(self.add_mapping_button)
        mapping_button_layout.addWidget(self.edit_mapping_button)
        mapping_button_layout.addWidget(self.delete_mapping_button)
        mapping_section_layout.addLayout(mapping_button_layout)

        self.main_layout.addLayout(mapping_section_layout)

        # 작업 시작 버튼
        button_layout = QHBoxLayout()
        self.start_button = QPushButton("작업 시작")
        button_layout.addWidget(self.start_button)
        self.main_layout.addLayout(button_layout)

        # 상태 표시
        self.status_display = StatusDisplay()
        self.main_layout.addWidget(self.status_display)

    def on_file1_selected(self):
        path = self.file1_widget.get_file_path()
        if path:
            self.file1_path = path
            self.update_sheet_names()
            self.status_display.add_message("File 1 선택됨")

    def on_file2_selected(self):
        path = self.file2_widget.get_file_path()
        if path:
            self.file2_path = path
            self.update_sheet_names()
            self.status_display.add_message("File 2 선택됨")

    def update_sheet_names(self):
        if not self.file1_path or not self.file2_path:
            return
        try:
            sheet_names = get_sheet_names(self.file1_path, DEFAULT_PASSWORD)
            self.sheet_names = sheet_names
            self.status_display.add_message(f"시트 목록 로드 완료: {len(sheet_names)}개")
        except Exception as e:
            self.status_display.add_message(f"시트 목록 로드 실패: {str(e)}")
    
    def on_mapping_button_clicked(self):
        if not self.file1_path or not self.file2_path:
            self.status_display.add_message("먼저 두 파일을 모두 선택하세요")
            return

        # 기존 매핑 로드
        mapping_dicts = load_mappings(DEFAULT_CONFIG_FILE)

        # 다이얼로그 생성
        dialog = MappingDialog(self)
        dialog.set_sheet_names(self.sheet_names)

        # 첫 번째 매핑이 있으면 표시
        if mapping_dicts:
            dialog.set_mapping(mapping_dicts[0])

        # 다이얼로그 실행
        if dialog.exec():
            mapping_dict = dialog.get_mapping()
            # MappingConfig로 변환
            mapping_config = MappingConfig.from_dict(mapping_dict)
            # 리스트에 추가/교체
            if mapping_dicts:
                mapping_dicts[0] = mapping_dict
            else:
                mapping_dicts = [mapping_dict]
            # 저장
            save_mappings(DEFAULT_CONFIG_FILE, mapping_dicts)
            self.mappings = dict_list_to_mappings(mapping_dicts)
            self.status_display.add_message("매핑 설정 저장 완료")
    
    def on_add_mapping_clicked(self):
        """새 매핑 추가"""
        if not self.file1_path or not self.file2_path:
            self.status_display.add_message("먼저 두 파일을 모두 선택하세요")
            return
        
        # 빈 다이얼로그 열기
        dialog = MappingDialog(self)
        dialog.setWindowTitle("매핑 추가")
        dialog.set_sheet_names(self.sheet_names)
        
        # 다이얼로그 실행
        if dialog.exec():
            mapping_dict = dialog.get_mapping()
            
            # 기존 매핑 로드
            mapping_dicts = load_mappings(DEFAULT_CONFIG_FILE)
            
            # 새 매핑 추가
            mapping_dicts.append(mapping_dict)
            
            # 저장
            save_mappings(DEFAULT_CONFIG_FILE, mapping_dicts)
            
            # 목록 갱신
            self.refresh_mapping_list()
            
            self.status_display.add_message("매핑 추가 완료")
    
    def on_edit_mapping_clicked(self):
        """선택된 매핑 수정"""
        current_row = self.mapping_list_widget.currentRow()
        
        if current_row == -1:
            self.status_display.add_message("수정할 매핑을 선택하세요")
            return
        
        if not self.file1_path or not self.file2_path:
            self.status_display.add_message("먼저 두 파일을 모두 선택하세요")
            return
        
        # 기존 매핑 로드
        mapping_dicts = load_mappings(DEFAULT_CONFIG_FILE)
        
        if current_row >= len(mapping_dicts):
            self.status_display.add_message("잘못된 선택입니다")
            return
        
        # 다이얼로그 열기
        dialog = MappingDialog(self)
        dialog.setWindowTitle("매핑 수정")
        dialog.set_sheet_names(self.sheet_names)
        
        # 선택된 매핑 데이터 로드
        dialog.set_mapping(mapping_dicts[current_row])
        
        # 다이얼로그 실행
        if dialog.exec():
            mapping_dict = dialog.get_mapping()
            
            # 해당 인덱스 수정
            mapping_dicts[current_row] = mapping_dict
            
            # 저장
            save_mappings(DEFAULT_CONFIG_FILE, mapping_dicts)
            
            # 목록 갱신
            self.refresh_mapping_list()
            
            self.status_display.add_message(f"매핑 수정 완료 (인덱스: {current_row + 1})")
    
    def on_delete_mapping_clicked(self):
        """선택된 매핑 삭제"""
        current_row = self.mapping_list_widget.currentRow()
        
        if current_row == -1:
            self.status_display.add_message("삭제할 매핑을 선택하세요")
            return
        
        # 기존 매핑 로드
        mapping_dicts = load_mappings(DEFAULT_CONFIG_FILE)
        
        if current_row >= len(mapping_dicts):
            self.status_display.add_message("잘못된 선택입니다")
            return
        
        # 확인 메시지 (선택사항)
        from PySide6.QtWidgets import QMessageBox
        reply = QMessageBox.question(
            self, 
            "삭제 확인", 
            f"선택한 매핑을 삭제하시겠습니까?\n\n{mapping_dicts[current_row].get('sheet_name', '')} - {mapping_dicts[current_row].get('data_range', '')}",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.No:
            return
        
        # 해당 인덱스 삭제
        del mapping_dicts[current_row]
        
        # 저장
        save_mappings(DEFAULT_CONFIG_FILE, mapping_dicts)
        
        # 목록 갱신
        self.refresh_mapping_list()
        
        self.status_display.add_message(f"매핑 삭제 완료 (인덱스: {current_row + 1})")

    def on_start_button_clicked(self):
        if not self.file1_path or not self.file2_path:
            self.status_display.add_message("먼저 두 파일을 모두 선택하세요")
            return
        
        from core.mapping_manager import load_mappings, dict_list_to_mappings
        from core.file_manager import decrypt_file, save_workbook
        from core.data_processor import merge_cell_sum, merge_key_based, merge_row_append, merge_cell_max, merge_cell_min
        from config.settings import DEFAULT_CONFIG_FILE, DEFAULT_PASSWORD, MergeMethod
        from utils.validators import is_valid_range_format, is_valid_column_format
        
        # 매핑 로드
        mapping_dicts = load_mappings(DEFAULT_CONFIG_FILE)
        if not mapping_dicts:
            self.status_display.add_message("매핑 설정이 없습니다. 먼저 매핑을 설정하세요")
            return
        
        self.mappings = dict_list_to_mappings(mapping_dicts)
        self.status_display.add_message(f"병합 작업 시작: {len(self.mappings)}개 매핑")
        
        try:
            # 병합 파일명 생성
            from core.file_manager import create_merged_filename
            merged_path = create_merged_filename(self.file1_path, self.file2_path)
            # file1을 새 파일명으로 복사
            import shutil
            shutil.copy2(self.file1_path, merged_path)
            self.status_display.add_message(f"병합 파일 생성: {os.path.basename(merged_path)}")
            
            # 복사본과 file2 로드
            wb1 = decrypt_file(merged_path, DEFAULT_PASSWORD)
            wb2 = decrypt_file(self.file2_path, DEFAULT_PASSWORD)
            self.status_display.add_message("파일 로드 완료")
            
            # 각 매핑 처리
            for idx, mapping in enumerate(self.mappings, 1):
                self.status_display.add_message(f"[{idx}/{len(self.mappings)}] {mapping.sheet_name} 처리 중...")
                
                # 유효성 검사
                valid, msg = is_valid_range_format(mapping.data_range)
                if not valid:
                    self.status_display.add_message(f"오류: {msg}")
                    continue
                
                # 병합 방식에 따라 처리
                if mapping.merge_method.value == "셀 합산":
                    merge_cell_sum(wb1, wb2, mapping.sheet_name, mapping.data_range)
                elif mapping.merge_method.value == "키 기반":
                    if not mapping.key_column or not mapping.index_column:
                        self.status_display.add_message("오류: 키 컬럼과 인덱스 컬럼이 필요합니다")
                        continue
                    merge_key_based(wb1, wb2, mapping.sheet_name, mapping.data_range, 
                                  mapping.key_column, mapping.index_column)
                elif mapping.merge_method.value == "행 누적":
                    merge_row_append(wb1, wb2, mapping.sheet_name, mapping.data_range)
                elif mapping.merge_method.value == "최소값 선택":
                    merge_cell_min(wb1, wb2, mapping.sheet_name, mapping.data_range)
                elif mapping.merge_method.value == "최대값 선택":
                    merge_cell_max(wb1, wb2, mapping.sheet_name, mapping.data_range)
                
                self.status_display.add_message(f"[{idx}/{len(self.mappings)}] 완료")
            
            # 저장
            save_workbook(wb1, merged_path, DEFAULT_PASSWORD)
            self.status_display.add_message("===== 병합 작업 완료 =====")
            
        except Exception as e:
            self.status_display.add_message(f"오류 발생: {str(e)}")

    def refresh_mapping_list(self):
        """매핑 목록 갱신"""
        self.mapping_list_widget.clear()
        
        mapping_dicts = load_mappings(DEFAULT_CONFIG_FILE)
        
        for idx, mapping in enumerate(mapping_dicts):
            method = mapping.get("merge_method", "")
            sheet = mapping.get("sheet_name", "")
            range_str = mapping.get("data_range", "")
            
            display_text = f"[{idx+1}] {method} - {sheet} ({range_str})"
            self.mapping_list_widget.addItem(display_text)