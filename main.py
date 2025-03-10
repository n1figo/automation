import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import pandas as pd
from datetime import datetime
import tempfile
import sys
from analyzer.pdf_parser import PDFParser
from analyzer.highlight_analyzer import HighlightAnalyzer
from utils.excel_writer import create_business_definition_excel

class PDFAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 보험약관 분석기")
        self.root.geometry("900x700")
        self.parser = None
        self.analysis_results = {}
        self.all_tables = []
        self.setup_ui()
        
    def setup_ui(self):
        # 상단 프레임
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.pack(fill=tk.X)
        
        ttk.Label(top_frame, text="PDF 보험약관 분석기", font=("Arial", 16, "bold")).pack()
        
        # 파일 선택 프레임
        file_frame = ttk.LabelFrame(self.root, text="파일 선택", padding="10")
        file_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=70).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        ttk.Button(file_frame, text="파일 선택", command=self.browse_file).pack(side=tk.LEFT, padx=5)
        
        # 옵션 프레임
        options_frame = ttk.LabelFrame(self.root, text="분석 옵션", padding="10")
        options_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.analyze_content_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="보장내용 분석", variable=self.analyze_content_var).pack(side=tk.LEFT, padx=20)
        
        self.find_highlights_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="강조표시 검출", variable=self.find_highlights_var).pack(side=tk.LEFT, padx=20)
        
        self.extract_tables_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="테이블 추출", variable=self.extract_tables_var).pack(side=tk.LEFT, padx=20)
        
        # 작업 버튼
        button_frame = ttk.Frame(self.root, padding="10")
        button_frame.pack(fill=tk.X, padx=10)
        
        self.progress_var = tk.StringVar(value="대기 중...")
        ttk.Label(button_frame, textvariable=self.progress_var).pack(side=tk.LEFT)
        
        self.progress = ttk.Progressbar(button_frame, length=400, mode='determinate')
        self.progress.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        self.analyze_button = ttk.Button(button_frame, text="분석 시작", command=self.start_analysis)
        self.analyze_button.pack(side=tk.LEFT, padx=5)
        
        # 결과 표시 영역
        result_frame = ttk.LabelFrame(self.root, text="분석 결과", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.result_text = tk.Text(result_frame, wrap=tk.WORD, width=80, height=20)
        scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 하단 버튼들
        bottom_frame = ttk.Frame(self.root, padding="10")
        bottom_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.save_button = ttk.Button(bottom_frame, text="결과 저장", command=self.save_results, state=tk.DISABLED)
        self.save_button.pack(side=tk.RIGHT, padx=5)
        
        self.export_button = ttk.Button(bottom_frame, text="Excel 내보내기", command=self.export_to_excel, state=tk.DISABLED)
        self.export_button.pack(side=tk.RIGHT, padx=5)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF 파일", "*.pdf")])
        if file_path:
            self.file_path_var.set(file_path)
            self.log_message(f"파일 선택됨: {file_path}")
    
    def start_analysis(self):
        pdf_path = self.file_path_var.get()
        if not pdf_path:
            messagebox.showerror("오류", "PDF 파일을 선택해주세요.")
            return
        
        # 분석 시작 전 UI 상태 변경
        self.analyze_button.config(state=tk.DISABLED)
        self.save_button.config(state=tk.DISABLED)
        self.export_button.config(state=tk.DISABLED)
        self.progress_var.set("분석 중...")
        self.progress["value"] = 0
        self.result_text.delete(1.0, tk.END)
        
        # 백그라운드 스레드에서 분석 실행
        threading.Thread(target=self.run_analysis, args=(pdf_path,), daemon=True).start()
    
    def run_analysis(self, pdf_path):
        try:
            self.log_message("PDF 분석을 시작합니다...")
            self.update_progress(5, "PDF 파일 로딩 중...")
            
            # PDF 파서 초기화
            self.parser = PDFParser(pdf_path)
            file_name = os.path.basename(pdf_path)
            
            # 결과 초기화
            self.analysis_results = {
                "파일명": file_name,
                "나. 보험금 페이지": [],
                "종별_모든발견": {},
                "종별_시작": {},
                "종별_범위": {},
                "상해관련특별약관 페이지": [],
                "질병관련특별약관 페이지": [],
                "상해및질병관련특별약관 페이지": [],
                "상해및질병관련특별약관 종료 페이지": None,
                "강조색 있는 페이지": set(),
                "취소선 있는 페이지": set(),
                "처리 상태": "완료",
                "상세 로그": []
            }
            
            # PDF 로드 및 기본 정보 확인
            total_pages = self.parser.get_total_pages()
            self.log_message(f"총 {total_pages}페이지 로드됨")
            self.update_progress(10, "문서 구조 분석 중...")
            
            # 1. "나. 보험금" 검색
            self.log_message("'나. 보험금' 키워드 검색 중...")
            parsing_start_page = self.parser.find_insurance_payment_page()
            if parsing_start_page is not None:
                self.analysis_results["나. 보험금 페이지"] = [parsing_start_page + 1]
                self.log_message(f"'나. 보험금' 문구 발견: 페이지 {parsing_start_page + 1}")
            else:
                self.log_message("'나. 보험금' 문구를 찾을 수 없어 첫 페이지부터 검색합니다.", "WARNING")
                parsing_start_page = 0
            
            # 2. 종별 정보 검색
            self.update_progress(20, "종별 정보 검색 중...")
            self.log_message("종별 패턴 검색 중...")
            
            self.analysis_results["종별_모든발견"] = self.parser.find_all_insurance_types()
            if self.analysis_results["종별_모든발견"]:
                types_found = ", ".join([f"{key}: {', '.join(map(str, pages))}" 
                                        for key, pages in self.analysis_results["종별_모든발견"].items()])
                self.log_message(f"종별 발견: {types_found}")
                
                # 종별 시작 페이지 식별
                self.update_progress(30, "종별 시작 페이지 식별 중...")
                type_starts = self.parser.identify_type_start_pages(parsing_start_page)
                
                if type_starts:
                    self.analysis_results["종별_시작"] = type_starts
                    for type_key, info in type_starts.items():
                        self.log_message(f"'{type_key}' 시작 페이지: {info['page']} (신뢰도: {info['confidence']})")
                    
                    # 종별 범위 계산
                    self.update_progress(40, "종별 범위 계산 중...")
                    type_ranges = self.parser.calculate_type_ranges(type_starts)
                    self.analysis_results["종별_범위"] = type_ranges
                    
                    for type_key, range_info in type_ranges.items():
                        self.log_message(f"'{type_key}' 범위: {range_info['start_page']}~{range_info['end_page']}페이지")
                else:
                    self.log_message("종별 시작 페이지를 식별할 수 없습니다.", "WARNING")
            else:
                self.log_message("종별 정보를 찾을 수 없습니다.", "WARNING")
            
            # 3. 특별약관 검색
            self.update_progress(50, "특별약관 섹션 검색 중...")
            self.analysis_results["상해관련특별약관 페이지"] = self.parser.find_injury_special_sections()
            self.analysis_results["질병관련특별약관 페이지"] = self.parser.find_disease_special_sections()
            self.analysis_results["상해및질병관련특별약관 페이지"] = self.parser.find_combined_special_sections()
            
            if self.analysis_results["상해관련특별약관 페이지"]:
                self.log_message(f"상해관련특별약관 발견: {', '.join(map(str, self.analysis_results['상해관련특별약관 페이지']))}")
            else:
                self.log_message("상해관련특별약관을 찾을 수 없습니다.", "WARNING")
                
            if self.analysis_results["질병관련특별약관 페이지"]:
                self.log_message(f"질병관련특별약관 발견: {', '.join(map(str, self.analysis_results['질병관련특별약관 페이지']))}")
            else:
                self.log_message("질병관련특별약관을 찾을 수 없습니다.", "WARNING")
                
            if self.analysis_results["상해및질병관련특별약관 페이지"]:
                self.log_message(f"상해및질병관련특별약관 발견: {', '.join(map(str, self.analysis_results['상해및질병관련특별약관 페이지']))}")
                
                # 종료 페이지 찾기
                end_page = self.parser.find_combined_section_end_page()
                if end_page:
                    self.analysis_results["상해및질병관련특별약관 종료 페이지"] = end_page
                    self.log_message(f"상해및질병관련특별약관 종료 페이지: {end_page}")
            
            # 4. 강조색 및 취소선 검출
            if self.find_highlights_var.get():
                self.update_progress(70, "강조색 및 취소선 검출 중...")
                
                # 검색 범위 설정
                search_start_page = parsing_start_page
                search_end_page = total_pages - 1
                
                if self.analysis_results["상해및질병관련특별약관 종료 페이지"]:
                    search_end_page = self.analysis_results["상해및질병관련특별약관 종료 페이지"] - 1
                
                self.log_message(f"강조색/취소선 검색 범위: 페이지 {search_start_page + 1}~{search_end_page + 1}")
                
                highlight_analyzer = HighlightAnalyzer()
                highlights, strikethroughs = self.parser.detect_all_highlights(
                    search_start_page, search_end_page, highlight_analyzer
                )
                
                self.analysis_results["강조색 있는 페이지"] = set(highlights)
                self.analysis_results["취소선 있는 페이지"] = set(strikethroughs)
                
                if highlights:
                    self.log_message(f"강조색 있는 페이지: {', '.join(map(str, sorted(highlights)))}")
                else:
                    self.log_message("강조색이 발견되지 않았습니다.", "WARNING")
                
                if strikethroughs:
                    self.log_message(f"취소선 있는 페이지: {', '.join(map(str, sorted(strikethroughs)))}")
            
            # 5. 테이블 추출
            if self.extract_tables_var.get():
                self.update_progress(80, "테이블 추출 중...")
                
                # 추출 범위 설정 
                extract_start_page = parsing_start_page
                extract_end_page = search_end_page if 'search_end_page' in locals() else total_pages - 1
                
                self.log_message(f"테이블 추출 범위: 페이지 {extract_start_page + 1}~{extract_end_page + 1}")
                
                self.all_tables = self.parser.extract_tables_from_range(extract_start_page, extract_end_page)
                
                if self.all_tables:
                    self.log_message(f"총 {len(self.all_tables)}개 테이블 추출 완료")
                else:
                    self.log_message("추출된 테이블이 없습니다.", "WARNING")
            
            # 분석 완료
            self.update_progress(100, "분석 완료")
            self.log_message("PDF 분석이 완료되었습니다.")
            
            # 결과 버튼 활성화
            self.root.after(0, lambda: self.save_button.config(state=tk.NORMAL))
            if self.all_tables:
                self.root.after(0, lambda: self.export_button.config(state=tk.NORMAL))
            
        except Exception as e:
            self.update_progress(0, "오류 발생")
            self.log_message(f"분석 중 오류 발생: {str(e)}", "ERROR")
            import traceback
            self.log_message(traceback.format_exc(), "ERROR")
            self.analysis_results["처리 상태"] = "오류"
        finally:
            # UI 상태 복원
            self.root.after(0, lambda: self.analyze_button.config(state=tk.NORMAL))
    
    def update_progress(self, value, message):
        self.root.after(0, lambda: self.progress.config(value=value))
        self.root.after(0, lambda: self.progress_var.set(message))
    
    def log_message(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_text = f"[{timestamp}] {level}: {message}\n"
        self.root.after(0, lambda: self.result_text.insert(tk.END, log_text))
        self.root.after(0, lambda: self.result_text.see(tk.END))
        
        # 로그 저장
        if hasattr(self, 'analysis_results') and isinstance(self.analysis_results, dict):
            if "상세 로그" in self.analysis_results:
                self.analysis_results["상세 로그"].append(f"{level}: {message}")
    
    def save_results(self):
        # 결과 저장 로직
        save_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON 파일", "*.json"), ("텍스트 파일", "*.txt")]
        )
        if not save_path:
            return
            
        import json
        
        # set 객체를 리스트로 변환
        json_safe_results = self.analysis_results.copy()
        for key, value in json_safe_results.items():
            if isinstance(value, set):
                json_safe_results[key] = sorted(list(value))
        
        try:
            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(json_safe_results, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("저장 완료", f"분석 결과가 저장되었습니다: {save_path}")
        except Exception as e:
            messagebox.showerror("저장 오류", f"결과 저장 중 오류 발생: {str(e)}")
    
    def export_to_excel(self):
        if not self.all_tables:
            messagebox.showwarning("내보내기 오류", "내보낼 테이블이 없습니다.")
            return
            
        # Excel 내보내기 로직
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")]
        )
        if not save_path:
            return
            
        try:
            file_name = self.analysis_results.get("파일명", "테이블_결과")
            excel_output = create_business_definition_excel(self.all_tables, file_name)
            
            if excel_output:
                with open(save_path, "wb") as f:
                    f.write(excel_output.getvalue())
                messagebox.showinfo("내보내기 완료", f"Excel 파일이 생성되었습니다: {save_path}")
            else:
                messagebox.showerror("내보내기 오류", "Excel 생성에 실패했습니다.")
        except Exception as e:
            messagebox.showerror("내보내기 오류", f"Excel 생성 중 오류 발생: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFAnalyzerApp(root)
    root.mainloop()