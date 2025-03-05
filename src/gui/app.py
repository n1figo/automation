from typing import List, Dict, Optional, Tuple
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import re
import logging
import os
from datetime import datetime
from pathlib import Path
import pandas as pd
import sys
from PIL import Image, ImageTk
# 실제 애플리케이션에 필요한 모듈들을 import
# from src.utils.excel_writer import ExcelWriter
# from src.processors.xml_processor import XMLProcessor
# from src.analyzers.pdf_analyzer import PDFAnalyzer
# from src.analyzers.example_analyzer import ExamplePDFAnalyzer
# from src.extractors.html_extractor import HTMLFileExtractor

# customtkinter 설정
ctk.set_appearance_mode("light")  # 테마 설정: "light" 또는 "dark"
ctk.set_default_color_theme("blue")  # 기본 색상 테마

class PDFAnalyzerGUI:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("KB손해보험 상품개정 자동화 서비스")
        self.root.geometry("1000x700")
        
        # KB 브랜드 색상 설정
        self.kb_yellow = "#FFC423"     # KB 노란색
        self.kb_light_yellow = "#FFFBEE"  # 은은한 노란색 배경
        self.kb_dark = "#333333"    # 어두운 회색
        self.kb_light = "#F8F8F8"   # 밝은 회색
        self.kb_border = "#E0E0E0"  # 테두리 색상
        
        # 폰트 설정
        self.title_font = ("맑은 고딕", 22)
        self.subtitle_font = ("맑은 고딕", 16)
        self.normal_font = ("맑은 고딕", 14)
        self.small_font = ("맑은 고딕", 12)
        
        # 로깅 설정
        self.setup_logging()
        
        # 전체 배경색 설정
        self.root.configure(fg_color=self.kb_light_yellow)
        
        # UI 구성
        self.setup_gui()
        
    def setup_logging(self):
        """로깅 설정"""
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)
        
        log_filename = log_dir / f'pdf_analyzer_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_filename, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)

    def setup_gui(self):
        """GUI 설정"""
        # 메인 프레임
        self.main_frame = ctk.CTkFrame(self.root, fg_color=self.kb_light_yellow)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 헤더 - KB손해보험 로고 및 제목
        self.create_header()
        
        # 제목 섹션
        self.create_title_section()
        
        # 파일 선택 영역 생성
        self.create_file_selection_area()
        
        # 처리 옵션 영역 생성
        self.create_options_area()
        
        # 진행 상태 영역 생성
        self.create_progress_area()
        
        # 로그 영역 생성
        self.create_log_area()
        
        # 푸터
        self.create_footer()
    
    def create_header(self):
        """헤더 영역 생성"""
        header_frame = ctk.CTkFrame(self.main_frame, fg_color="white", height=50, corner_radius=0)
        header_frame.pack(fill="x", pady=(0, 15))
        
        # KB 로고 이미지 로드 및 표시
        try:
            # 직접 파일 경로 지정
            logo_path = "D:/github/pdf_local_11/mindful-revision-helper/assets/kb_logo.png"
            
            if os.path.exists(logo_path):
                # customtkinter의 CTkImage 사용
                logo_img = ctk.CTkImage(
                    light_image=Image.open(logo_path),
                    size=(100, 30)
                )
                
                # 로고 이미지 라벨
                logo_label = ctk.CTkLabel(
                    header_frame,
                    image=logo_img,
                    text=""
                )
                logo_label.pack(side="left", padx=10, pady=10)
            else:
                raise FileNotFoundError(f"로고 이미지 파일을 찾을 수 없습니다: {logo_path}")
            
        except Exception as e:
            # 이미지 로드 실패시 텍스트로 대체
            self.logger.warning(f"로고 이미지 로드 실패: {str(e)}")
            logo_label = ctk.CTkLabel(
                header_frame, 
                text="KB손해보험",
                font=ctk.CTkFont(family="맑은 고딕", size=14, weight="bold"),
                text_color=self.kb_dark
            )
            logo_label.pack(side="left", padx=10, pady=10)
    
    def create_title_section(self):
        """제목 섹션 생성"""
        title_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        title_frame.pack(fill="x", pady=(0, 15))
        
        # 메인 제목
        title_label = ctk.CTkLabel(
            title_frame, 
            text="상품개정 자동화 서비스",
            font=ctk.CTkFont(family="맑은 고딕", size=26, weight="bold"),
            text_color=self.kb_dark
        )
        title_label.pack(pady=(0, 5))
        
        # 부제목
        subtitle_label = ctk.CTkLabel(
            title_frame, 
            text="상품 개정 관련 파일을 업로드하여 자동으로 분석하실 수 있습니다.",
            font=ctk.CTkFont(family="맑은 고딕", size=16),
            text_color=self.kb_dark
        )
        subtitle_label.pack()
    
    def create_file_selection_area(self):
        """파일 선택 영역 생성"""
        # 파일 선택 프레임
        file_frame = ctk.CTkFrame(self.main_frame, fg_color="white", corner_radius=10)
        file_frame.pack(fill="x", pady=10, padx=5)
        
        # 보장내용 PDF 선택
        self.coverage_path_var = ctk.StringVar()
        self.create_file_input(
            file_frame, 
            "보장내용 PDF 선택", 
            self.coverage_path_var,
            [("PDF files", "*.pdf")]
        )
        
        # 가입예시 PDF 선택
        self.example_pdf_path_var = ctk.StringVar()
        self.create_file_input(
            file_frame, 
            "가입예시 PDF 선택", 
            self.example_pdf_path_var,
            [("PDF files", "*.pdf")]
        )
        
        # MHTML 파일 선택
        self.mhtml_path_var = ctk.StringVar()
        self.create_file_input(
            file_frame, 
            "MHTML 파일 선택", 
            self.mhtml_path_var,
            [("MHTML files", "*.mhtml")]
        )
        
        # XML 파일 선택
        self.xml_path_var = ctk.StringVar()
        self.create_file_input(
            file_frame, 
            "XML 파일 선택", 
            self.xml_path_var,
            [("XML files", "*.xml")]
        )
        
        # 하단 패딩
        bottom_padding = ctk.CTkFrame(file_frame, fg_color="transparent", height=5)
        bottom_padding.pack(fill="x")
    
    def create_file_input(self, parent, label_text, string_var, filetypes):
        """파일 입력 필드 생성"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=15, pady=5)
        
        # 라벨
        label = ctk.CTkLabel(
            frame,
            text=label_text,
            font=ctk.CTkFont(family="맑은 고딕", size=14),
            width=140,
            anchor="w",
            text_color=self.kb_dark
        )
        label.pack(side="left")
        
        # 입력 필드
        entry = ctk.CTkEntry(
            frame,
            textvariable=string_var,
            font=ctk.CTkFont(family="맑은 고딕", size=11),
            border_width=1,
            height=36,
            fg_color="white",
            border_color=self.kb_border,
            placeholder_text="선택된 파일 없음"
        )
        entry.pack(side="left", fill="x", expand=True, padx=10)
        
        # 찾아보기 버튼
        browse_button = ctk.CTkButton(
            frame,
            text="찾아보기...",
            font=ctk.CTkFont(family="맑은 고딕", size=14),
            fg_color=self.kb_yellow,
            text_color=self.kb_dark,
            hover_color="#FFD54F",
            height=36,
            width=90,
            corner_radius=4,
            cursor="hand2",
            command=lambda: self.browse_file(string_var, filetypes)
        )
        browse_button.pack(side="right")
    
    def create_options_area(self):
        """처리 옵션 영역 생성"""
        options_frame = ctk.CTkFrame(self.main_frame, fg_color="white", corner_radius=10)
        options_frame.pack(fill="x", pady=10, padx=5)
        
        # 제목
        title_label = ctk.CTkLabel(
            options_frame,
            text="처리 옵션",
            font=ctk.CTkFont(family="맑은 고딕", size=14, weight="bold"),
            text_color=self.kb_dark
        )
        title_label.pack(anchor="w", padx=15, pady=(10, 5))
        
        # 체크박스 프레임
        checkbox_frame = ctk.CTkFrame(options_frame, fg_color="transparent")
        checkbox_frame.pack(fill="x", padx=15, pady=(5, 10))
        
        # 체크박스 변수들
        self.coverage_analysis_var = ctk.BooleanVar(value=True)
        self.example_analysis_var = ctk.BooleanVar(value=True)
        
        # 보장내용 분석 체크박스
        coverage_check = ctk.CTkCheckBox(
            checkbox_frame,
            text="보장내용 분석",
            variable=self.coverage_analysis_var,
            font=ctk.CTkFont(family="맑은 고딕", size=11),
            checkbox_width=20,
            checkbox_height=20,
            border_width=2,
            fg_color=self.kb_yellow,
            hover_color="#FFD54F",
            text_color=self.kb_dark,
            cursor="hand2"
        )
        coverage_check.pack(side="left", padx=(0, 30))
        
        # 가입예시 분석 체크박스
        example_check = ctk.CTkCheckBox(
            checkbox_frame,
            text="가입예시 분석",
            variable=self.example_analysis_var,
            font=ctk.CTkFont(family="맑은 고딕", size=11),
            checkbox_width=20,
            checkbox_height=20,
            border_width=2,
            fg_color=self.kb_yellow,
            hover_color="#FFD54F",
            text_color=self.kb_dark,
            cursor="hand2"
        )
        example_check.pack(side="left")
    
    def create_progress_area(self):
        """진행 상태 영역 생성"""
        progress_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        progress_frame.pack(fill="x", pady=10)
        
        # 진행 상태 라벨
        self.progress_var = ctk.StringVar(value="대기 중...")
        progress_label = ctk.CTkLabel(
            progress_frame,
            textvariable=self.progress_var,
            font=ctk.CTkFont(family="맑은 고딕", size=11),
            text_color=self.kb_dark
        )
        progress_label.pack(pady=(0, 5))
        
        # 진행 상태 바
        self.progress_bar = ctk.CTkProgressBar(
            progress_frame,
            height=10,
            corner_radius=2,
            fg_color="#e0e0e0",
            progress_color=self.kb_yellow,
            border_width=0
        )
        self.progress_bar.pack(fill="x", pady=(0, 15))
        self.progress_bar.set(0)  # 초기값 설정
        
        # 분석 시작 버튼 프레임 (오른쪽 정렬용)
        button_frame = ctk.CTkFrame(progress_frame, fg_color="transparent")
        button_frame.pack(anchor="e")
        
        # 분석 시작 버튼
        self.process_button = ctk.CTkButton(
            button_frame,
            text="처리 시작",
            font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"),
            fg_color=self.kb_yellow,
            text_color=self.kb_dark,
            hover_color="#FFD54F",
            width=120,
            height=32,
            corner_radius=4,
            cursor="hand2",
            command=self.process_start,
            state="disabled"
        )
        self.process_button.pack()
    
    def create_log_area(self):
        """로그 영역 생성"""
        log_frame = ctk.CTkFrame(self.main_frame, fg_color="white", corner_radius=10)
        log_frame.pack(fill="both", expand=True, pady=10, padx=5)
        
        # 제목
        title_label = ctk.CTkLabel(
            log_frame,
            text="처리 로그",
            font=ctk.CTkFont(family="맑은 고딕", size=14, weight="bold"),
            text_color=self.kb_dark
        )
        title_label.pack(anchor="w", padx=15, pady=(10, 5))
        
        # 로그 텍스트 영역 프레임
        text_frame = ctk.CTkFrame(log_frame, fg_color="transparent")
        text_frame.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        
        # 로그 텍스트 영역
        self.log_text = ctk.CTkTextbox(
            text_frame,
            font=ctk.CTkFont(family="맑은 고딕", size=11),
            fg_color="#f8f8f8",
            border_width=1,
            border_color=self.kb_border,
            corner_radius=4,
            wrap="word"
        )
        self.log_text.pack(fill="both", expand=True)
        
        # 초기 로그 메시지
        self.log_message("시스템이 준비되었습니다. 파일을 선택하고 분석을 시작하세요.")
    
    def create_footer(self):
        """푸터 영역 생성"""
        footer_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent", height=20)
        footer_frame.pack(fill="x", pady=(5, 0))
        
        # 저작권 정보
        copyright_label = ctk.CTkLabel(
            footer_frame,
            text="© KB손해보험. All rights reserved.",
            font=ctk.CTkFont(family="맑은 고딕", size=9),
            text_color="#999999"
        )
        copyright_label.pack(side="right")

    def check_process_button_state(self):
        """분석 시작 버튼 활성화 상태 확인"""
        if (self.coverage_path_var.get() or 
            self.example_pdf_path_var.get() or
            self.mhtml_path_var.get() or
            self.xml_path_var.get()):
            self.process_button.configure(state="normal")
        else:
            self.process_button.configure(state="disabled")

    def browse_file(self, string_var, filetypes):
        """파일 탐색기 실행"""
        file_path = filedialog.askopenfilename(title="파일 선택", filetypes=filetypes)
        if file_path:
            string_var.set(file_path)
            self.check_process_button_state()

    def process_start(self):
        """분석 작업 시작"""
        try:
            self.process_button.configure(state="disabled")
            self.progress_bar.set(0)
            
            coverage_path = self.coverage_path_var.get()
            example_pdf_path = self.example_pdf_path_var.get()
            mhtml_path = self.mhtml_path_var.get()
            xml_path = self.xml_path_var.get()
            
            # 최소한 하나의 파일이 선택되었는지 확인
            if not any([coverage_path, example_pdf_path, mhtml_path, xml_path]):
                messagebox.showerror("오류", "최소한 하나의 파일을 선택해주세요.")
                self.process_button.configure(state="normal")
                return

            # 분석 작업을 별도 스레드에서 실행
            thread = threading.Thread(
                target=self.process_files, 
                args=(coverage_path, mhtml_path, example_pdf_path, xml_path)
            )
            thread.daemon = True
            thread.start()

        except Exception as e:
            self.log_message(f"처리 시작 중 오류 발생: {str(e)}", "ERROR")
            self.process_button.configure(state="normal")

    def process_files(self, coverage_path: Optional[str], mhtml_path: Optional[str], 
                  example_pdf_path: Optional[str], xml_path: Optional[str] = None) -> None:
        """파일 처리 작업 수행 (멀티스레딩 버전)"""
        try:
            self.update_progress(0, "파일 분석 준비 중...")
            
            # 실제 구현 시 여기에 파일 처리 로직을 추가
            # 예시용 진행 상태 시뮬레이션
            import time
            steps = 5
            for i in range(steps):
                progress = (i + 1) / steps * 100
                
                # 진행 상태 메시지
                messages = [
                    "파일 분석 준비 중...",
                    "데이터 추출 중...",
                    "보장 내용 분석 중...",
                    "결과 정리 중...",
                    "완료되었습니다!"
                ]
                
                self.update_progress(progress, messages[i])
                self.log_message(f"단계 {i+1}/{steps}: {messages[i]}")
                time.sleep(1)  # 실제 작업에서는 제거
            
            # 완료 메시지
            messagebox.showinfo("완료", "분석이 완료되었습니다!")

        except Exception as e:
            self.log_message(f"처리 중 오류 발생: {str(e)}", "ERROR")
            messagebox.showerror("오류", f"처리 중 오류가 발생했습니다:\n{str(e)}")
        finally:
            self.process_button.configure(state="normal")
            self.progress_var.set("대기 중...")
            self.progress_bar.set(0)

    def update_progress(self, value, message):
        """진행 상태 업데이트"""
        self.progress_bar.set(value / 100)  # 0-1 사이 값으로 변환
        self.progress_var.set(message)
        self.root.update_idletasks()

    def log_message(self, message, level="INFO"):
        """로그 메시지 출력"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # 로그 메시지 서식 지정
        log_message = f"[{timestamp}] {level}: {message}\n"
        
        # 텍스트 삽입
        self.log_text.insert("end", log_message)
        self.log_text.see("end")
        
        # 로깅
        if level == "INFO":
            self.logger.info(message)
        elif level == "ERROR":
            self.logger.error(message)
        elif level == "WARNING":
            self.logger.warning(message)

    def run(self):
        """애플리케이션 실행"""
        self.root.mainloop()


if __name__ == "__main__":
    app = PDFAnalyzerGUI()
    app.run()