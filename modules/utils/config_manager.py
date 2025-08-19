import json
import os
from pathlib import Path
from typing import Dict, Any, List, Tuple
from datetime import datetime, timedelta

class ConfigManager:
    """설정 파일 관리 클래스 - 리팩토링된 버전"""
    
    def __init__(self, config_dir: str = None):
        if config_dir is None:
            current_file = Path(__file__).resolve()
            # 리팩토링된 구조: modules/utils/config_manager.py에서 2단계 위로
            self.base_dir = current_file.parent.parent.parent
            
            # 환경변수 확인 (우선순위 1)
            if os.environ.get('SALES_REPORT_HOME'):
                self.base_dir = Path(os.environ.get('SALES_REPORT_HOME'))
                print(f"🏠 환경변수 SALES_REPORT_HOME 사용: {self.base_dir}")
            
            self.config_dir = self.base_dir / "config"
        else:
            self.config_dir = Path(config_dir)
            self.base_dir = self.config_dir.parent
        
        self.accounts_path = self.config_dir / "accounts.json"
        self.report_config_path = self.config_dir / "report_config.json"
        
        # 설정 로드
        self.accounts = self._load_accounts()
        self.report_config = self._load_report_config()
        
        # 런타임 계정 정보 (메모리 저장용)
        self.runtime_accounts = None
        
        # 프로젝트 루트 설정
        self.project_root = self.base_dir
    
    def set_runtime_accounts(self, accounts: list):
        """런타임에 계정 정보 설정 (메모리에만 저장)"""
        self.runtime_accounts = {"accounts": accounts}
        print(f"✅ 런타임 계정 설정 완료: {len(accounts)}개 회사")
    
    def get_accounts(self) -> list:
        """계정 정보 반환 (런타임 우선, 파일 없어도 오류 없이)"""
        # 1순위: 런타임 계정
        if self.runtime_accounts and self.runtime_accounts.get("accounts"):
            print(f"🚀 런타임 계정 사용: {len(self.runtime_accounts['accounts'])}개 회사")
            return self.runtime_accounts["accounts"]
        
        # 2순위: 파일 계정 (있는 경우)
        if self.accounts and self.accounts.get("accounts"):
            print(f"📁 파일 계정 사용: {len(self.accounts['accounts'])}개 회사")
            return self.accounts.get("accounts", [])
        
        # 3순위: 빈 배열 (오류 예방)
        print("⚠️ 사용 가능한 계정 정보가 없습니다")
        return []

    def _load_accounts(self) -> Dict[str, Any]:
        """계정 정보 로드"""
        try:
            with open(self.accounts_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"❌ accounts.json 파일을 찾을 수 없습니다: {self.accounts_path}")
            return {"accounts": []}
        except json.JSONDecodeError as e:
            print(f"❌ accounts.json 파일 형식 오류: {e}")
            return {"accounts": []}
    
    def _load_report_config(self) -> Dict[str, Any]:
        """보고서 설정 로드"""
        try:
            with open(self.report_config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                if config.get("paths", {}).get("base_dir") == "auto_detect":
                    config["paths"]["base_dir"] = str(self.base_dir)
                return config
        except FileNotFoundError:
            print(f"❌ report_config.json 파일을 찾을 수 없습니다: {self.report_config_path}")
            return self._get_default_config()
        except json.JSONDecodeError as e:
            print(f"❌ report_config.json 파일 형식 오류: {e}")
            return self._get_default_config()
    
    def _get_default_config(self) -> Dict[str, Any]:
        """기본 설정 반환"""
        return {
            "paths": {
                "base_dir": str(self.base_dir),
                "staff_file": "data/판매_담당자목록.xlsx",
                "staff_sheet": "담당자목록",
                "sales_raw_data": "data/sales_raw_data",
                "receivable_raw_data": "data/receivable_calculator_raw_data",
                "receivables_raw_data": "data/receivables",
                "processed_data": "data/processed",
                "report_output": "data/report",
                "downloads": "data/downloads",
                "weekly_report_save": "data/report"
            },
            "date_settings": {
                "week_start_day": "friday",
                "week_end_day": "thursday",
                "report_period_days": 5,
                "auto_calculate_period": True
            },
            "sales": {
                "exclude_codes": ["1078711207", "02644", "10219"],
                "exclude_products": [],
                "default_num_months": 3,
                "retry_count": 3,
                "download_timeout": 60,
                "category_mappings": {"수출": "무역"},
                "category_order": ["구동기", "일반부품", "무역", "티케이"]
            },
            "receivables": {
                "target_companies": ["디앤드디", "디앤아이"],
                "collection_day": "friday",
                "retry_count": 5,
                "download_timeout": 120,
                "collection_delay": 10,
                "batch_size": 6,
                "batch_delay": 30,
                "file_naming_format": "{company}_채권잔액분석_{date}.xlsx"
            },
            "report": {
                "kpi_target": {"장기미수채권_비율": 18.25},
                "targets_2025": {
                    "구동기": 922, "일반부품": 1015, "티케이": 42, "무역": 143
                }
            },
            "selenium": {
                "headless": False,
                "implicit_wait": 15,
                "page_load_timeout": 60,
                "script_timeout": 30,
                "detach_browser": True,
                "disable_automation_flags": True,
                "window_size": "1920,1080",
                "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "disable_images": False,
                "disable_javascript": False
            },
            "companies": {
                "디앤드디": {
                    "default_category": None,
                    "month_xpath_format": "//a[text()='{month}']",
                    "collect_receivables": True
                },
                "디앤아이": {
                    "default_category": "구동기",
                    "month_xpath_format": "//a[text()='{month}']",
                    "collect_receivables": True
                },
                "후지리프트코리아": {
                    "default_category": "무역",
                    "month_xpath_format": "//a[text()='{int_month}월']",
                    "collect_receivables": False
                }
            }
        }
    
    def get_base_dir(self) -> Path:
        """기본 디렉토리 반환"""
        return Path(self.report_config["paths"]["base_dir"])
    
    def get_processed_data_dir(self) -> Path:
        """처리된 데이터 디렉토리 반환"""
        return self.get_base_dir() / "data/processed"
    
    def get_report_output_dir(self) -> Path:
        """보고서 출력 디렉토리 반환"""
        return self.get_base_dir() / "data/report"
    
    def get_sales_config(self) -> Dict[str, Any]:
        """매출 관련 설정 반환"""
        return self.report_config.get("sales", {})
    
    def get_receivables_config(self) -> Dict[str, Any]:
        """매출채권 관련 설정 반환"""
        return self.report_config.get("receivables", {})
    
    # 추가 필요한 메서드들
    def get_selenium_config(self) -> Dict[str, Any]:
        """Selenium 관련 설정 반환"""
        return self.report_config.get("selenium", {
            "headless": False,
            "implicit_wait": 15,
            "page_load_timeout": 60,
            "script_timeout": 30,
            "detach_browser": True,
            "disable_automation_flags": True
        })
    
    def get_paths(self) -> Dict[str, str]:
        """경로 설정 반환"""
        paths = self.report_config.get("paths", {})
        base_dir = Path(paths.get("base_dir", self.base_dir))
        
        return {
            "base_dir": str(base_dir),
            "downloads": str(base_dir / paths.get("downloads", "data/downloads")),
            "sales_raw_data": str(base_dir / paths.get("sales_raw_data", "data/sales_raw_data")),
            "receivable_raw_data": str(base_dir / paths.get("receivable_raw_data", "data/receivable_calculator_raw_data")),
            "processed_data": str(base_dir / paths.get("processed_data", "data/processed")),
            "report_output": str(base_dir / paths.get("report_output", "data/report"))
        }
    
    def get_downloads_dir(self) -> Path:
        """다운로드 디렉토리 반환"""
        paths = self.get_paths()
        return Path(paths["downloads"])
    
    def get_sales_raw_data_dir(self) -> Path:
        """매출 원시 데이터 디렉토리 반환"""
        paths = self.get_paths()
        return Path(paths["sales_raw_data"])
    
    def get_receivable_raw_data_dir(self) -> Path:
        """매출채권 원시 데이터 디렉토리 반환"""
        paths = self.get_paths()
        return Path(paths["receivable_raw_data"])
    
    def get_receivables_raw_data_dir(self) -> Path:
        """매출채권 데이터 디렉토리 반환 (별칭)"""
        return self.get_receivable_raw_data_dir()
    
    def get_default_num_months(self) -> int:
        """기본 수집 개월 수 반환"""
        return self.report_config.get("sales", {}).get("default_num_months", 3)
    
    def get_download_timeout(self) -> int:
        """다운로드 타임아웃 반환"""
        return self.report_config.get("sales", {}).get("download_timeout", 60)
    
    def get_exclude_codes(self) -> List[str]:
        """제외 코드 목록 반환"""
        return self.report_config.get("sales", {}).get("exclude_codes", [])
    
    def get_exclude_products(self) -> List[str]:
        """제외 제품 목록 반환"""
        return self.report_config.get("sales", {}).get("exclude_products", [])
    
    def get_category_mappings(self) -> Dict[str, str]:
        """카테고리 매핑 반환"""
        return self.report_config.get("sales", {}).get("category_mappings", {})
    
    def get_category_order(self) -> List[str]:
        """카테고리 순서 반환"""
        return self.report_config.get("sales", {}).get("category_order", ["구동기", "일반부품", "무역", "티케이"])
    
    def get_company_config(self, company_name: str) -> Dict[str, Any]:
        """회사별 설정 반환"""
        return self.report_config.get("companies", {}).get(company_name, {})
    
    def get_staff_file_path(self) -> Path:
        """담당자 파일 경로 반환"""
        paths = self.report_config.get("paths", {})
        staff_file = paths.get("staff_file", "data/판매_담당자목록.xlsx")
        return self.get_base_dir() / staff_file
    
    def get_staff_sheet_name(self) -> str:
        """담당자 시트명 반환"""
        paths = self.report_config.get("paths", {})
        return paths.get("staff_sheet", "담당자목록")
    
    def get_kpi_target(self, kpi_name: str) -> float:
        """KPI 목표값 반환"""
        kpi_targets = self.report_config.get("report", {}).get("kpi_target", {})
        return kpi_targets.get(kpi_name, 0.0)


# 전역 설정 인스턴스
_config_instance = None

def get_config() -> ConfigManager:
    """전역 설정 인스턴스 반환"""
    global _config_instance
    if _config_instance is None:
        _config_instance = ConfigManager()
    return _config_instance
