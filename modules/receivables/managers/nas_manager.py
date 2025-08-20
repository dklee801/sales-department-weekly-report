"""
NAS 매출채권 파일 관리자 - 리팩토링된 버전
NAS 연동 및 파일 동기화 기능 제공
"""

import pandas as pd
import os
from pathlib import Path
from datetime import datetime, timedelta
import re
import logging
from typing import Dict, List, Optional, Tuple, Union, Callable
import sys
import shutil

# 프로젝트 루트 경로 설정
project_root = Path(__file__).parent.parent.parent.parent
sys.path.insert(0, str(project_root))

try:
    from modules.utils.config_manager import get_config
    from modules.receivables.managers.file_manager import ReceivablesFileManager
except ImportError:
    # Fallback import
    sys.path.append(str(project_root / "modules"))
    from config_manager import get_config
    from receivables_modules.file_manager import ReceivablesFileManager

class NASReceivablesManager:
    """NAS 매출채권 파일 관리자 - UI 경로 선택 방식"""
    
    def __init__(self, config=None):
        self.config = config or get_config()
        self.logger = logging.getLogger('NASReceivablesManager')
        self.nas_selected_path = None
        self.local_receivables_dir = self.config.get_receivable_raw_data_dir()
        self.file_manager = ReceivablesFileManager(config)
        
    def set_nas_path(self, nas_path: Union[str, Path]):
        """NAS 경로 설정"""
        self.nas_selected_path = Path(nas_path) if nas_path else None
        
    def scan_nas_files_recursive(self, nas_path: Path = None) -> List[Path]:
        """NAS 폴더에서 매출채권 파일들 재귀적 스캔"""
        if nas_path is None:
            nas_path = self.nas_selected_path
            
        if not nas_path or not nas_path.exists():
            raise FileNotFoundError(f"NAS 경로에 접근할 수 없습니다: {nas_path}")
        
        excel_files = []
        
        try:
            # 루트 폴더에서 Excel 파일 찾기
            excel_files.extend(list(nas_path.glob("*.xlsx")))
            excel_files.extend(list(nas_path.glob("*.xls")))
            
            # 하위 폴더에서도 재귀적으로 찾기
            for subfolder in nas_path.iterdir():
                if subfolder.is_dir():
                    try:
                        excel_files.extend(list(subfolder.glob("**/*.xlsx")))
                        excel_files.extend(list(subfolder.glob("**/*.xls")))
                    except PermissionError:
                        self.logger.warning(f"권한 없음: {subfolder}")
                        continue
            
            # 매출채권 관련 파일만 필터링
            receivables_files = []
            for file_path in excel_files:
                if file_path.name.startswith("~$"):  # 임시 파일 제외
                    continue
                    
                filename = file_path.name.lower()
                if any(keyword in filename for keyword in ["매출채권", "채권잔액", "채권계산", "receivable"]):
                    receivables_files.append(file_path)
            
            self.logger.info(f"NAS에서 매출채권 파일 {len(receivables_files)}개 발견")
            return receivables_files
            
        except Exception as e:
            raise Exception(f"NAS 파일 스캔 실패: {e}")
    
    def sync_files_simple(self, overwrite_duplicates=True, progress_callback: Optional[Callable] = None) -> Dict[str, any]:
        """단순한 파일 동기화 - 모든 파일을 한 폴더에"""
        result = {
            "success": False,
            "copied_files": [],
            "skipped_files": [],
            "failed_files": [],
            "total_scanned": 0,
            "error": None
        }
        
        try:
            if not self.nas_selected_path:
                raise Exception("NAS 경로가 선택되지 않았습니다.")
            
            if progress_callback:
                progress_callback(f"NAS 경로 스캔 중: {self.nas_selected_path}")
            
            # 1단계: NAS에서 매출채권 파일들 스캔
            nas_files = self.scan_nas_files_recursive()
            result["total_scanned"] = len(nas_files)
            
            if progress_callback:
                progress_callback(f"발견된 매출채권 파일: {len(nas_files)}개")
            
            if not nas_files:
                if progress_callback:
                    progress_callback("매출채권 파일을 찾을 수 없습니다")
                result["success"] = True  # 파일이 없어도 성공으로 처리
                return result
            
            # 2단계: 로컬 폴더 준비
            self.local_receivables_dir.mkdir(parents=True, exist_ok=True)
            
            if progress_callback:
                progress_callback(f"파일 복사 시작: {len(nas_files)}개 파일")
            
            # 3단계: 각 파일 처리
            for i, file_path in enumerate(nas_files, 1):
                try:
                    if progress_callback:
                        progress_callback(f"처리 중 [{i}/{len(nas_files)}]: {file_path.name}")
                    
                    # 파일 형식 검증
                    if not self.validate_receivables_file(file_path):
                        result["failed_files"].append(f"{file_path.name} (형식 오류)")
                        continue
                    
                    # 로컬 파일 경로
                    local_file = self.local_receivables_dir / file_path.name
                    
                    # 중복 파일 처리 - 기본 덮어쓰기
                    if local_file.exists():
                        if overwrite_duplicates:
                            # 덮어쓰기 모드: 기존 파일을 그대로 교체
                            pass  # 파일 경로 그대로 사용
                        else:
                            # 파일 크기 비교로 동일성 확인
                            if local_file.stat().st_size == file_path.stat().st_size:
                                result["skipped_files"].append(f"{file_path.name} (동일 파일)")
                                continue
                            else:
                                # 크기가 다르면 새로운 이름으로 저장
                                name_without_ext = file_path.stem
                                extension = file_path.suffix
                                counter = 1
                                while True:
                                    new_name = f"{name_without_ext}_{counter}{extension}"
                                    new_local_file = self.local_receivables_dir / new_name
                                    if not new_local_file.exists():
                                        local_file = new_local_file
                                        break
                                    counter += 1
                    
                    # 파일 복사
                    shutil.copy2(file_path, local_file)
                    result["copied_files"].append(file_path.name)
                    self.logger.info(f"파일 복사 완료: {file_path.name}")
                    
                except Exception as e:
                    result["failed_files"].append(f"{file_path.name} ({str(e)})")
                    self.logger.error(f"파일 복사 실패 {file_path.name}: {e}")
            
            result["success"] = True
            if progress_callback:
                progress_callback("동기화 완료")
                
        except Exception as e:
            result["error"] = str(e)
            result["success"] = False
            self.logger.error(f"동기화 실패: {e}")
            if progress_callback:
                progress_callback(f"동기화 실패: {e}")
        
        return result
    
    def sync_files_to_local_organized(self, progress_callback: Optional[Callable] = None) -> Dict[str, any]:
        """레거시 호환성 메서드 - 단순 동기화로 리다이렉트 (기본 덮어쓰기)"""
        return self.sync_files_simple(overwrite_duplicates=True, progress_callback=progress_callback)
    
    def validate_receivables_file(self, file_path: Path) -> bool:
        """매출채권 파일 형식 검증"""
        try:
            # 파일 크기 체크
            if file_path.stat().st_size < 1000:  # 1KB 미만
                return False
            
            # Excel 파일인지 확인
            if file_path.suffix.lower() not in ['.xlsx', '.xls']:
                return False
            
            # 기본 구조 검증
            xl = pd.ExcelFile(file_path)
            
            if len(xl.sheet_names) == 0:
                return False
            
            # 회사 시트 찾기
            company_sheets = []
            for sheet_name in xl.sheet_names:
                if any(company in sheet_name for company in ["디앤드디", "디앤아이", "DND", "DNI"]):
                    company_sheets.append(sheet_name)
            
            # 회사 시트가 있으면 해당 시트 검증, 없으면 첫 번째 시트 검증
            test_sheet = company_sheets[0] if company_sheets else xl.sheet_names[0]
            df = xl.parse(test_sheet, nrows=2)
            
            if df.empty:
                return False
            
            # 필수 컬럼 존재 여부 확인
            required_columns = ["거래처코드", "거래처명", "총채권", "기간초과 매출채권", "90일초과 매출채권"]
            columns_str = str(df.columns).lower()
            
            found_columns = 0
            for col in required_columns:
                if col in columns_str:
                    found_columns += 1
            
            return found_columns >= len(required_columns) * 0.6  # 60% 이상 매칭
                    
        except Exception as e:
            self.logger.error(f"파일 검증 실패 {file_path}: {e}")
            return False
    
    def get_sync_summary(self) -> Dict[str, any]:
        """동기화 요약 정보"""
        if not self.nas_selected_path:
            return {"error": "NAS 경로가 설정되지 않았습니다"}
        
        try:
            nas_files = self.scan_nas_files_recursive()
            local_files = self.file_manager.find_all_receivables_files()
            
            # 파일명 기준으로 비교
            nas_file_names = set(f.name for f in nas_files)
            local_file_names = set(f[0].name for f in local_files if f[1] is not None)
            
            missing_in_local = nas_file_names - local_file_names
            extra_in_local = local_file_names - nas_file_names
            common_files = nas_file_names & local_file_names
            
            return {
                "nas_path": str(self.nas_selected_path),
                "nas_files_count": len(nas_files),
                "local_files_count": len(local_files),
                "common_files_count": len(common_files),
                "missing_in_local": list(missing_in_local),
                "extra_in_local": list(extra_in_local),
                "sync_needed": len(missing_in_local) > 0
            }
        except Exception as e:
            return {"error": str(e)}
    
    def check_nas_connectivity(self) -> Dict[str, any]:
        """NAS 연결 상태 확인"""
        result = {
            "connected": False,
            "path": str(self.nas_selected_path) if self.nas_selected_path else None,
            "accessible": False,
            "files_found": 0,
            "error": None
        }
        
        try:
            if not self.nas_selected_path:
                result["error"] = "NAS 경로가 설정되지 않았습니다"
                return result
            
            result["connected"] = self.nas_selected_path.exists()
            
            if result["connected"]:
                result["accessible"] = True
                try:
                    files = self.scan_nas_files_recursive()
                    result["files_found"] = len(files)
                except Exception as e:
                    result["accessible"] = False
                    result["error"] = f"파일 스캔 실패: {e}"
            else:
                result["error"] = "NAS 경로에 접근할 수 없습니다"
                
        except Exception as e:
            result["error"] = str(e)
        
        return result
    
    def create_sync_report(self) -> pd.DataFrame:
        """동기화 보고서 생성"""
        try:
            summary = self.get_sync_summary()
            
            if "error" in summary:
                return pd.DataFrame([{"상태": "오류", "메시지": summary["error"]}])
            
            report_data = [
                {"항목": "NAS 경로", "값": summary["nas_path"]},
                {"항목": "NAS 파일 수", "값": summary["nas_files_count"]},
                {"항목": "로컬 파일 수", "값": summary["local_files_count"]},
                {"항목": "공통 파일 수", "값": summary["common_files_count"]},
                {"항목": "동기화 필요", "값": "예" if summary["sync_needed"] else "아니오"}
            ]
            
            # 누락된 파일들 추가
            if summary["missing_in_local"]:
                for i, filename in enumerate(summary["missing_in_local"][:5]):  # 최대 5개만
                    report_data.append({"항목": f"누락 파일 {i+1}", "값": filename})
            
            return pd.DataFrame(report_data)
            
        except Exception as e:
            return pd.DataFrame([{"상태": "오류", "메시지": str(e)}])
    
    def schedule_auto_sync(self, interval_hours: int = 24):
        """자동 동기화 스케줄링 (미구현 - 향후 확장용)"""
        # 향후 스케줄러 기능 추가 시 구현
        self.logger.info(f"자동 동기화 스케줄링 요청: {interval_hours}시간 간격")
        return {"scheduled": False, "message": "자동 동기화 기능은 향후 구현 예정입니다"}


class DataValidator:
    """개선된 데이터 검증 클래스 - 매출채권 전용"""
    
    def __init__(self, config=None):
        self.config = config or get_config()
        self.logger = logging.getLogger('DataValidator')
        
        self.receivables_patterns = [
            r'.*매출채권.*\.xlsx?$',
            r'.*채권잔액.*\.xlsx?$',
            r'.*채권계산.*\.xlsx?$',
            r'.*receivable.*\.xlsx?$'
        ]
    
    def validate_date_range(self, start_date: str, end_date: str) -> bool:
        """날짜 범위 유효성 검증"""
        try:
            start = datetime.strptime(start_date, '%Y-%m-%d')
            end = datetime.strptime(end_date, '%Y-%m-%d')
            
            # 시작일이 종료일보다 이전이어야 함
            if start >= end:
                return False
            
            # 범위가 너무 크면 안됨 (예: 2년 이상)
            if (end - start).days > 730:
                return False
            
            return True
            
        except ValueError:
            return False
    
    def validate_receivables_data_structure(self, file_path: Path) -> Dict[str, any]:
        """매출채권 데이터 구조 검증"""
        result = {
            "valid": False,
            "errors": [],
            "warnings": [],
            "structure_info": {},
            "data_quality": {}
        }
        
        try:
            xl = pd.ExcelFile(file_path)
            result["structure_info"]["total_sheets"] = len(xl.sheet_names)
            result["structure_info"]["sheet_names"] = xl.sheet_names
            
            # 회사 시트 찾기
            company_sheets = []
            for sheet_name in xl.sheet_names:
                if any(company in sheet_name for company in ["디앤드디", "디앤아이", "DND", "DNI"]):
                    company_sheets.append(sheet_name)
            
            result["structure_info"]["company_sheets"] = company_sheets
            
            if not company_sheets:
                result["warnings"].append("회사별 시트를 찾을 수 없습니다")
                test_sheet = xl.sheet_names[0]
            else:
                test_sheet = company_sheets[0]
            
            # 데이터 구조 검증
            df = xl.parse(test_sheet)
            result["structure_info"]["columns"] = list(df.columns)
            result["structure_info"]["row_count"] = len(df)
            
            # 필수 컬럼 확인
            required_columns = ["거래처코드", "거래처명", "총채권"]
            optional_columns = ["기간초과 매출채권", "90일초과 매출채권"]
            
            found_required = 0
            found_optional = 0
            
            columns_str = str(df.columns).lower()
            
            for col in required_columns:
                if col in columns_str:
                    found_required += 1
                else:
                    result["errors"].append(f"필수 컬럼 누락: {col}")
            
            for col in optional_columns:
                if col in columns_str:
                    found_optional += 1
            
            # 데이터 품질 검증
            if not df.empty:
                # 빈 값 비율
                total_cells = len(df) * len(df.columns)
                empty_cells = df.isnull().sum().sum()
                empty_ratio = empty_cells / total_cells if total_cells > 0 else 0
                
                result["data_quality"]["empty_ratio"] = round(empty_ratio, 3)
                result["data_quality"]["total_rows"] = len(df)
                result["data_quality"]["total_columns"] = len(df.columns)
                
                if empty_ratio > 0.5:
                    result["warnings"].append(f"빈 데이터 비율이 높습니다: {empty_ratio:.1%}")
            
            # 유효성 판정
            result["valid"] = (found_required >= len(required_columns) * 0.8) and len(df) > 0
            
            if result["valid"]:
                result["structure_info"]["quality_score"] = min(100, int((found_required + found_optional) / (len(required_columns) + len(optional_columns)) * 100))
            
        except Exception as e:
            result["errors"].append(f"파일 분석 실패: {str(e)}")
        
        return result
    
    def batch_validate_files(self, file_paths: List[Path]) -> pd.DataFrame:
        """여러 파일 일괄 검증"""
        validation_results = []
        
        for file_path in file_paths:
            result = self.validate_receivables_data_structure(file_path)
            
            validation_results.append({
                "파일명": file_path.name,
                "유효성": "유효" if result["valid"] else "무효",
                "품질점수": result["structure_info"].get("quality_score", 0),
                "시트수": result["structure_info"].get("total_sheets", 0),
                "행수": result["structure_info"].get("row_count", 0),
                "오류수": len(result["errors"]),
                "경고수": len(result["warnings"]),
                "주요오류": result["errors"][0] if result["errors"] else ""
            })
        
        return pd.DataFrame(validation_results)