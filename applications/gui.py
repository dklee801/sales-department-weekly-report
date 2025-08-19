#!/usr/bin/env python3
"""
주간보고서 자동화 GUI - 리팩토링된 버전
새로운 모듈 구조에 맞춘 import 경로 및 기능 개선
"""

import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
from pathlib import Path
import shutil
import logging
import re
from typing import Dict, List, Optional
import threading
import queue

# 프로젝트 루트를 Python 경로에 추가
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# 날짜 선택 기능을 위한 추가 import
try:
    from tkcalendar import DateEntry
    TKCALENDAR_AVAILABLE = True
except ImportError:
    TKCALENDAR_AVAILABLE = False
    print("⚠️ tkcalendar not available - using basic date selection")

import pandas as pd

# 리팩토링된 모듈들 import
try:
    print("통합 테스트: 리팩토링된 모듈 import 시도...")
    
    # 1. 유틸리티 모듈들
    from modules.utils.config_manager import get_config
    print("   ✓ config_manager import 성공")
    
    from modules.gui.login_dialog import get_erp_accounts
    print("   ✓ login_dialog import 성공")
    
    # 2. 핵심 분석 모듈들
    from modules.core.sales_calculator import main as analyze_sales
    print("   ✓ sales_calculator import 성공")
    
    from modules.core.accounts_receivable_analyzer import main as analyze_receivables
    print("   ✓ accounts_receivable_analyzer import 성공")
    
    # 3. 데이터 처리 모듈들
    from modules.data.unified_data_collector import UnifiedDataCollector
    print("   ✓ unified_data_collector import 성공")
    
    # 4. 보고서 생성 모듈들
    try:
        from modules.reports.xml_safe_report_generator import StandardFormatReportGenerator
        WeeklyReportGenerator = StandardFormatReportGenerator
        print("✅ StandardFormatReportGenerator 로드 성공")
    except ImportError:
        try:
            from modules.reports.xml_safe_report_generator import XMLSafeReportGenerator
            WeeklyReportGenerator = XMLSafeReportGenerator
            print("✅ XML 안전 보고서 생성기 import 성공")
        except ImportError:
            WeeklyReportGenerator = None
            print("⚠️ 보고서 생성기를 찾을 수 없습니다")
        
except ImportError as e:
    print(f"필수 모듈 import 실패: {e}")
    print("기존 모듈들로 fallback 시도...")
    
    # 기존 모듈들 fallback import
    # 원본 프로젝트에서 모듈 가져오기 시도
    original_project = Path(__file__).parent.parent.parent / "Sales_department" / "modules"
    if original_project.exists():
        sys.path.append(str(original_project))
        sys.path.append(str(original_project.parent))
    
    try:
        from config_manager import get_config
        from sales_calculator_v3 import main as analyze_sales
        from unified_data_collector import UnifiedDataCollector
        from login_dialog import get_erp_accounts
        
        # 보고서 생성기 fallback
        WeeklyReportGenerator = None
        try:
            from xml_safe_report_generator import XMLSafeReportGenerator
            WeeklyReportGenerator = XMLSafeReportGenerator
            print("✅ Fallback 보고서 생성기 import 성공")
        except ImportError:
            print("⚠️ Fallback 보고서 생성기도 찾을 수 없습니다")
            
    except ImportError as fallback_error:
        print(f"Fallback 모듈 import도 실패: {fallback_error}")
        messagebox.showerror("오류", "필수 모듈을 찾을 수 없습니다. 프로그램을 종료합니다.")
        sys.exit(1)

# 매출채권 분석기들 import
try:
    from modules.receivables.analyzers.processed_receivables_analyzer import ProcessedReceivablesAnalyzer
    from modules.receivables.managers.nas_manager import NASReceivablesManager
    from modules.receivables.managers.file_manager import WeeklyReportDateSelector
    from modules.receivables.processors.report_integrator import ReceivablesReportIntegrator
    from modules.gui.receivables_components import ReceivablesGUIComponent, ReceivablesSourceDialog
    RECEIVABLES_AVAILABLE = True
    print("✅ 리팩토링된 매출채권 모듈들 import 성공")
except ImportError as e:
    print(f"⚠️ 리팩토링된 매출채권 모듈 import 실패: {e}")
    try:
        from processed_receivables_analyzer import ProcessedReceivablesAnalyzer
        RECEIVABLES_AVAILABLE = True
        print("✅ 기존 매출채권 분석기 import 성공")
    except ImportError:
        RECEIVABLES_AVAILABLE = False
        print("⚠️ 매출채권 분석기를 찾을 수 없습니다")


class ReportAutomationGUI:
    """주간보고서 자동화 GUI 메인 클래스 - 리팩토링된 버전"""
    
    def __init__(self):
        try:
            # GUI 기본 설정
            self.root = tk.Tk()
            self.root.title("주간보고서 자동화 프로그램 v4.0 (리팩토링 완료)")
            self.root.geometry("900x800")
            self.root.minsize(800, 700)
            
            # ERP 계정 정보 입력
            self.erp_accounts = get_erp_accounts(self.root)
            if not self.erp_accounts:
                messagebox.showinfo("취소", "ERP 계정 정보 입력이 취소되었습니다.\n프로그램을 종료합니다.")
                self.root.destroy()
                return
            
            self.config = get_config()
            self.config.set_runtime_accounts(self.erp_accounts)
            
            # 쓰레드 통신용 큐
            self.progress_queue = queue.Queue()
            
            # 진행상황 추가 변수들
            self.current_task_total = 0
            self.current_task_step = 0
            
            # 매출채권 관련 컴포넌트 초기화 - 주간 선택 문제로 인해 임시 비활성화
            self.receivables_component = None
            print("⚠️ 임시 조치: ReceivablesGUIComponent 비활성화 (선택된 주간 전달 문제)")
                
            self.setup_ui()
            self.setup_logging()
            
        except Exception as e:
            print(f"❌ GUI 초기화 실패: {e}")
            if hasattr(self, 'root'):
                try:
                    self.root.destroy()
                except:
                    pass
            raise
    
    def setup_logging(self):
        """로깅 설정"""
        log_format = '%(asctime)s - %(levelname)s - %(message)s'
        logging.basicConfig(level=logging.INFO, format=log_format)
        self.logger = logging.getLogger(__name__)
    
    def setup_ui(self):
        """UI 구성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 제목
        title_label = ttk.Label(main_frame, text="주간보고서 자동화 프로그램 v4.0", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 5))
        
        # 부제목
        subtitle_label = ttk.Label(main_frame, text="🆕 리팩토링 완료 • 모듈화 구조 • 향상된 안정성", 
                                  font=('Arial', 10), foreground="gray")
        subtitle_label.grid(row=1, column=0, columnspan=3, pady=(0, 15))
        
        # 1. 데이터 현황 표시
        self.setup_status_section(main_frame, row=2)
        
        # 2. 데이터 갱신 섹션
        self.setup_data_section(main_frame, row=3)
        
        # 3. 보고서 생성 섹션
        self.setup_report_section(main_frame, row=4)
        
        # 4. 진행상황 표시
        self.setup_progress_section(main_frame, row=5)
        
        # Grid 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
    
    def setup_status_section(self, parent, row):
        """데이터 현황 섹션"""
        frame = ttk.LabelFrame(parent, text="1. 데이터 현황", padding="10")
        frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 상태 표시 텍스트
        self.status_text = tk.Text(frame, height=6, width=80)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.status_text.yview)
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 상태 확인 버튼
        ttk.Button(frame, text="📊 데이터 현황 확인", 
                  command=self.check_data_status).grid(row=1, column=0, pady=(10, 0))
        
        frame.columnconfigure(0, weight=1)
    
    def setup_data_section(self, parent, row):
        """데이터 갱신 섹션"""
        frame = ttk.LabelFrame(parent, text="2. 데이터 갱신 (리팩토링된 모듈)", padding="10")
        frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 매출 수집 기간 선택 섹션
        period_frame = ttk.Frame(frame)
        period_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(period_frame, text="매출 수집 기간:").grid(row=0, column=0, sticky=tk.W)
        
        # 매출 수집 기간 드롭다운 (1-24개월로 확장)
        self.sales_period_var = tk.StringVar(value="3개월")
        self.sales_period_combo = ttk.Combobox(period_frame, textvariable=self.sales_period_var,
                                              values=[f"{i}개월" for i in range(1, 25)],
                                              state="readonly", width=8)
        self.sales_period_combo.grid(row=0, column=1, padx=(10, 10), sticky=tk.W)
        
        # 설명 라벨
        ttk.Label(period_frame, text="💡 최신 데이터부터 선택한 기간만큼 수집 (최대 24개월)", 
                 foreground="gray", font=('Arial', 8)).grid(row=0, column=2, sticky=tk.W, padx=(10, 0))
        
        # 버튼들
        buttons_frame = ttk.Frame(frame)
        buttons_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        self.sales_button = ttk.Button(buttons_frame, text="📈 매출 데이터 갱신", 
                                      command=self.start_sales_update)
        self.sales_button.grid(row=0, column=0, padx=(0, 10))
        
        self.sales_process_button = ttk.Button(buttons_frame, text="🔄 매출집계 처리", 
                                             command=self.start_sales_processing)
        self.sales_process_button.grid(row=0, column=1, padx=(0, 10))
        
        self.receivables_button = ttk.Button(buttons_frame, text="💰 매출채권 분석", 
                                           command=self.start_receivables_analysis)
        self.receivables_button.grid(row=0, column=2)
        
        # 매출채권 데이터 동기화 버튼 추가
        self.receivables_sync_button = ttk.Button(buttons_frame, text="🔄 매출채권 데이터 동기화", 
                                                 command=self.start_receivables_sync)
        self.receivables_sync_button.grid(row=1, column=0, columnspan=3, pady=(10, 0))
        
        buttons_frame.columnconfigure(0, weight=1)
        buttons_frame.columnconfigure(1, weight=1)
        buttons_frame.columnconfigure(2, weight=1)
    
    def setup_report_section(self, parent, row):
        """보고서 생성 섹션"""
        frame = ttk.LabelFrame(parent, text="3. 보고서 생성 (리팩토링된 모듈)", padding="10")
        frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 주간 선택 섹션
        week_selection_frame = ttk.Frame(frame)
        week_selection_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(week_selection_frame, text="보고서 주간:", font=('Arial', 9)).grid(row=0, column=0, sticky=tk.W)
        
        self.friday_selection_var = tk.StringVar()
        self.friday_combobox = ttk.Combobox(week_selection_frame, textvariable=self.friday_selection_var,
                                           width=25, state="readonly")
        self.friday_combobox.grid(row=0, column=1, padx=(10, 10), sticky=tk.W)
        
        # 주간 목록 로드 버튼
        ttk.Button(week_selection_frame, text="🔄 새로고침", 
                  command=self.load_available_weeks).grid(row=0, column=2)
        
        # 보고서 설정 섹션
        settings_frame = ttk.Frame(frame)
        settings_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 기준 월 선택
        ttk.Label(settings_frame, text="기준 월:").grid(row=0, column=0, sticky=tk.W)
        self.base_month_var = tk.StringVar(value="8월")
        self.base_month_combo = ttk.Combobox(settings_frame, textvariable=self.base_month_var,
                                            values=[f"{i}월" for i in range(1, 13)],
                                            state="readonly", width=8)
        self.base_month_combo.grid(row=0, column=1, padx=(10, 10), sticky=tk.W)
        
        # 월시작일 선택 (금요일만)
        ttk.Label(settings_frame, text="월시작일:").grid(row=0, column=2, sticky=tk.W, padx=(20, 0))
        
        if TKCALENDAR_AVAILABLE:
            self.start_date_picker = DateEntry(settings_frame, width=12, background='darkblue',
                                              foreground='white', borderwidth=2,
                                              date_pattern='yyyy-mm-dd')
            self.start_date_picker.grid(row=0, column=3, padx=(10, 0), sticky=tk.W)
            self._configure_friday_only_selection()
        else:
            initial_friday = self._get_nearest_friday()
            self.start_date_var = tk.StringVar(value=initial_friday.strftime('%Y-%m-%d'))
            self.start_date_entry = ttk.Entry(settings_frame, textvariable=self.start_date_var, width=12)
            self.start_date_entry.grid(row=0, column=3, padx=(10, 0), sticky=tk.W)
            self.start_date_entry.bind('<FocusOut>', self._validate_friday_entry)
            self.start_date_entry.bind('<Return>', self._validate_friday_entry)
        
        # 실행 버튼들
        buttons_frame = ttk.Frame(frame)
        buttons_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        self.full_process_button = ttk.Button(buttons_frame, text="🚀 전체 프로세스 실행", 
                                            command=self.start_full_process_with_selected_week,
                                            style="Accent.TButton")
        self.full_process_button.grid(row=0, column=0, padx=(0, 10))
        
        self.report_only_button = ttk.Button(buttons_frame, text="📄 보고서만 생성", 
                                           command=self.start_report_generation_with_selected_week)
        self.report_only_button.grid(row=0, column=1)
        
        # 초기 데이터 로드
        self.load_available_weeks()
        
        frame.columnconfigure(0, weight=1)
    
    def setup_progress_section(self, parent, row):
        """진행상황 섹션"""
        frame = ttk.LabelFrame(parent, text="진행상황", padding="10")
        frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 현재 작업 표시
        self.current_task_var = tk.StringVar(value="대기 중...")
        self.current_task_label = ttk.Label(frame, textvariable=self.current_task_var, 
                                           font=('Arial', 10, 'bold'))
        self.current_task_label.grid(row=0, column=0, sticky=tk.W)
        
        # 상세 진행 메시지
        self.progress_var = tk.StringVar(value="작업을 시작하려면 버튼을 클릭하세요.")
        self.progress_label = ttk.Label(frame, textvariable=self.progress_var, 
                                       foreground="gray")
        self.progress_label.grid(row=1, column=0, sticky=tk.W, pady=(2, 0))
        
        # 진행바
        self.progress_bar = ttk.Progressbar(frame, mode='determinate', length=400)
        self.progress_bar.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        frame.columnconfigure(0, weight=1)
    
    def update_status(self, message: str):
        """상태 텍스트 업데이트"""
        if hasattr(self, 'status_text') and self.status_text:
            try:
                self.status_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
                self.status_text.see(tk.END)
                self.root.update_idletasks()
            except Exception as e:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")
                print(f"   ⚠️ GUI 상태 표시 오류: {e}")
        else:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")
    
    def update_progress(self, message: str):
        """진행상황 업데이트"""
        self.progress_var.set(message)
        self.root.update_idletasks()
    
    def check_data_status(self):
        """데이터 현황 확인"""
        self.status_text.delete(1.0, tk.END)
        self.update_status("🔍 리팩토링된 모듈 기반 데이터 현황 확인...")
        
        try:
            # 리팩토링된 구조 정보 표시
            self.update_status("✅ 리팩토링 완료 상태:")
            self.update_status("   📁 modules/core/ - 핵심 분석 로직")
            self.update_status("   📁 modules/data/ - 데이터 처리")
            self.update_status("   📁 modules/gui/ - GUI 컴포넌트")
            self.update_status("   📁 modules/utils/ - 유틸리티")
            self.update_status("   📁 modules/reports/ - 보고서 생성")
            self.update_status("")
            
            # 모듈 가용성 확인
            self.update_status("🔧 모듈 가용성:")
            if WeeklyReportGenerator:
                self.update_status("   ✅ 보고서 생성기: 사용 가능")
            else:
                self.update_status("   ❌ 보고서 생성기: 사용 불가")
            
            if RECEIVABLES_AVAILABLE:
                self.update_status("   ✅ 매출채권 분석기: 사용 가능")
            else:
                self.update_status("   ❌ 매출채권 분석기: 사용 불가")
            
            self.update_status("")
            
            # 파일 존재 확인
            base_dir = Path(__file__).parent.parent
            template_file = base_dir / "2025년도 주간보고 양식_2.xlsx"
            processed_dir = base_dir / "data/processed"
            
            self.update_status("📂 파일 현황:")
            if template_file.exists():
                self.update_status("   ✅ 보고서 템플릿: 존재")
            else:
                self.update_status("   ❌ 보고서 템플릿: 없음")
            
            if processed_dir.exists():
                excel_files = list(processed_dir.glob("*.xlsx"))
                self.update_status(f"   📊 처리된 데이터: {len(excel_files)}개 파일")
            else:
                self.update_status("   📊 처리된 데이터: 디렉토리 없음")
            
        except Exception as e:
            self.update_status(f"❌ 데이터 현황 확인 중 오류: {e}")
    
    def get_selected_sales_period_months(self):
        """선택된 매출 수집 기간을 숫자로 변환"""
        try:
            period_text = self.sales_period_var.get()
            return int(period_text.replace('개월', ''))
        except:
            return 3  # 기본값
    
    def load_available_weeks(self):
        """사용 가능한 주간 목록 로드"""
        try:
            current_date = datetime.now()
            
            # 현재 날짜에서 가장 가까운 금요일 찾기
            days_until_friday = (4 - current_date.weekday()) % 7
            if days_until_friday == 0 and current_date.weekday() != 4:
                days_until_friday = 7
            
            next_friday = current_date + timedelta(days=days_until_friday)
            
            # 최근 8주간의 금요일 목록 생성
            friday_options = []
            for i in range(8):
                friday = next_friday - timedelta(weeks=i)
                next_thursday = friday + timedelta(days=6)
                display_text = f"{friday.strftime('%Y-%m-%d')} (금) ~ {next_thursday.strftime('%m-%d')} (목)"
                friday_options.append(display_text)
            
            self.friday_combobox['values'] = friday_options
            if friday_options:
                self.friday_combobox.set(friday_options[0])  # 최신 주간 선택
        except Exception as e:
            self.update_status(f"⚠️ 주간 목록 로드 오류: {e}")
    
    def get_selected_base_month(self):
        """선택된 기준 월 반환"""
        return self.base_month_var.get()
    
    def get_selected_start_date_range(self):
        """선택된 시작일을 기반으로 기간 범위 생성"""
        try:
            if TKCALENDAR_AVAILABLE and hasattr(self, 'start_date_picker'):
                start_date = self.start_date_picker.get_date()
            else:
                start_date_str = self.start_date_var.get()
                start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
            
            end_date = start_date + timedelta(days=6)
            return f"{start_date.strftime('%Y-%m-%d')} - {end_date.strftime('%Y-%m-%d')}"
            
        except Exception as e:
            self.update_status(f"⚠️ 시작일 범위 생성 오류: {e}")
            return None

    # 이하 메서드들은 실제 구현이 필요한 스텁들입니다
    def start_sales_update(self):
        """매출 데이터 갱신 시작 - 구현 필요"""
        self.update_status("매출 데이터 갱신 기능 - 구현 필요")
        
    def start_sales_processing(self):
        """매출집계 처리 시작 - 구현 필요"""
        self.update_status("매출집계 처리 기능 - 구현 필요")
        
    def start_receivables_analysis(self):
        """매출채권 분석 실행 - 구현 필요"""
        self.update_status("매출채권 분석 기능 - 구현 필요")
        
    def start_receivables_sync(self):
        """매출채권 데이터 동기화 - 구현 필요"""
        self.update_status("매출채권 데이터 동기화 기능 - 구현 필요")
        
    def start_report_generation_with_selected_week(self):
        """선택된 주간으로 보고서 생성 - 구현 필요"""
        self.update_status("보고서 생성 기능 - 구현 필요")
        
    def start_full_process_with_selected_week(self):
        """전체 프로세스 실행 - 구현 필요"""
        self.update_status("전체 프로세스 기능 - 구현 필요")
    
    def _configure_friday_only_selection(self):
        """금요일만 선택 가능하도록 설정"""
        if not hasattr(self, 'start_date_picker'):
            return
            
        def validate_date_selection(event=None):
            try:
                selected_date = self.start_date_picker.get_date()
                if selected_date.weekday() != 4:  # 금요일이 아니면
                    days_to_friday = (4 - selected_date.weekday()) % 7
                    if days_to_friday == 0:
                        days_to_friday = 7
                    friday_date = selected_date + timedelta(days=days_to_friday)
                    self.start_date_picker.set_date(friday_date)
                    messagebox.showinfo("알림", f"금요일만 선택 가능합니다.\n{friday_date.strftime('%Y-%m-%d')} (금)로 변경했습니다.")
            except Exception as e:
                print(f"날짜 유효성 검사 오류: {e}")
        
        self.start_date_picker.bind('<<DateEntrySelected>>', validate_date_selection)
        self.start_date_picker.bind('<FocusOut>', validate_date_selection)
        self._set_initial_friday()
    
    def _get_nearest_friday(self):
        """가장 가까운 이전 금요일 날짜 반환"""
        today = datetime.now().date()
        days_since_friday = (today.weekday() - 4) % 7
        if days_since_friday == 0 and today.weekday() == 4:
            return today
        else:
            return today - timedelta(days=days_since_friday)
    
    def _set_initial_friday(self):
        """초기 날짜를 가장 가까운 금요일로 설정"""
        try:
            last_friday = self._get_nearest_friday()
            self.start_date_picker.set_date(last_friday)
        except Exception as e:
            print(f"초기 금요일 설정 오류: {e}")
    
    def _validate_friday_entry(self, event=None):
        """기본 Entry 위젯에서 금요일 유효성 검사"""
        try:
            date_str = self.start_date_var.get()
            selected_date = datetime.strptime(date_str, '%Y-%m-%d').date()
            
            if selected_date.weekday() != 4:  # 금요일이 아니면
                days_to_friday = (4 - selected_date.weekday()) % 7
                if days_to_friday == 0:
                    days_to_friday = 7
                friday_date = selected_date + timedelta(days=days_to_friday)
                self.start_date_var.set(friday_date.strftime('%Y-%m-%d'))
                messagebox.showinfo("알림", f"금요일만 선택 가능합니다.\n{friday_date.strftime('%Y-%m-%d')} (금)로 변경했습니다.")
                
        except ValueError:
            friday = self._get_nearest_friday()
            self.start_date_var.set(friday.strftime('%Y-%m-%d'))
            messagebox.showerror("오류", f"잘못된 날짜 형식입니다.\nYYYY-MM-DD 형식으로 입력해주세요.\n{friday.strftime('%Y-%m-%d')} (금)로 재설정했습니다.")
        except Exception as e:
            print(f"Entry 날짜 유효성 검사 오류: {e}")
    
    def run(self):
        """GUI 실행"""
        self.root.mainloop()


def main():
    """메인 실행 함수"""
    try:
        app = ReportAutomationGUI()
        app.run()
    except Exception as e:
        print(f"프로그램 실행 중 오류 발생: {e}")
        messagebox.showerror("오류", f"프로그램 실행 중 오류가 발생했습니다:\n{e}")


if __name__ == "__main__":
    main()
