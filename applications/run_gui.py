#!/usr/bin/env python3
"""
주간보고서 자동화 GUI 실행 런처 - 새 구조용 (수정됨)
"""

import sys
import os
from pathlib import Path

def setup_environment():
    """환경 설정"""
    # 프로젝트 루트 디렉토리를 Python 경로에 추가
    project_root = Path(__file__).parent.parent
    sys.path.insert(0, str(project_root))
    
    # modules 디렉토리를 Python 경로에 추가
    modules_dir = project_root / "modules"
    if modules_dir.exists():
        sys.path.insert(0, str(modules_dir))
    
    print(f"프로젝트 루트: {project_root}")
    print(f"모듈 디렉토리: {modules_dir}")

def check_dependencies():
    """필수 모듈 확인"""
    required_modules = [
        'pandas',
        'tkinter', 
        'openpyxl',
        'xlsxwriter',
        'tkcalendar'
    ]
    
    missing_modules = []
    
    for module in required_modules:
        try:
            if module == 'tkinter':
                import tkinter
            elif module == 'tkcalendar':
                from tkcalendar import DateEntry
            else:
                __import__(module)
            print(f"✅ {module}")
        except ImportError:
            missing_modules.append(module)
            print(f"❌ {module}")
    
    if missing_modules:
        print(f"\n다음 모듈을 설치해주세요:")
        for module in missing_modules:
            if module == 'tkcalendar':
                print(f"pip install {module}>=1.6.1")
            else:
                print(f"pip install {module}")
        return False
    
    return True

def check_project_files():
    """프로젝트 파일 확인"""
    project_root = Path(__file__).parent.parent
    modules_dir = project_root / "modules"
    
    required_files = [
        "modules/utils/config_manager.py",
        "modules/core/sales_calculator.py",
        "modules/data/unified_data_collector.py",
        "modules/gui/components/weekly_date_selector.py",
        "modules/gui/login_dialog.py"
    ]
    
    missing_files = []
    
    for file_path in required_files:
        full_path = project_root / file_path
        if full_path.exists():
            print(f"✅ {file_path}")
        else:
            missing_files.append(file_path)
            print(f"❌ {file_path}")
    
    if missing_files:
        print(f"\n다음 파일들이 누락되었습니다:")
        for file in missing_files:
            print(f"  - {file}")
        return False
    
    return True

def run_gui():
    """GUI 실행"""
    try:
        print("\n🔍 개선 기능 테스트 중...")
        
        # 날짜 선택기 기능 테스트
        try:
            from modules.gui.components.weekly_date_selector import WeeklyDateSelector
            print("✅ 주간 날짜 선택기 정상")
        except ImportError as e:
            print(f"⚠️ 주간 날짜 선택기 비활성화: {e}")
        
        # tkcalendar 기능 테스트
        try:
            from tkcalendar import DateEntry
            print("✅ 달력 위젯 정상")
        except ImportError:
            print("⚠️ 달력 위젯 비활성화 (기본 기능만 사용)")
        
        print("\n🚀 주간보고서 자동화 GUI 시작...")
        print("🆕 새로운 모듈 구조 적용")
        
        # 실제 GUI 모듈 실행
        try:
            # GUI 모듈 import
            gui_file = Path(__file__).parent / "gui.py"
            if gui_file.exists():
                print("✅ GUI 모듈 파일 발견")
                
                # GUI 클래스 import 시도
                from applications.gui import ReportAutomationGUI
                print("✅ GUI 클래스 import 성공")
                
                # GUI 실행
                print("🎯 GUI 인스턴스 생성 중...")
                app = ReportAutomationGUI()
                print("🎯 GUI 실행 시작...")
                app.run()
                
            else:
                print("❌ applications/gui.py 파일이 없습니다.")
                print("   GUI 파일을 확인해주세요.")
                
        except ImportError as e:
            print(f"❌ GUI 모듈 import 실패: {e}")
            print("GUI 모듈의 클래스 구조를 확인해주세요.")
            
        except AttributeError as e:
            print(f"❌ GUI 클래스 실행 실패: {e}")
            print("ReportAutomationGUI 클래스와 run() 메서드를 확인해주세요.")
            
        except Exception as e:
            print(f"❌ GUI 실행 중 오류: {e}")
            import traceback
            traceback.print_exc()
            
    except Exception as e:
        print(f"❌ GUI 준비 실패: {e}")
        import traceback
        traceback.print_exc()

def main():
    """메인 실행 함수"""
    print("="*60)
    print("주간보고서 자동화 프로그램 GUI 런처 (새 구조)")
    print("="*60)
    
    # 1. 환경 설정
    print("\n1. 환경 설정 중...")
    setup_environment()
    
    # 2. 의존성 확인
    print("\n2. 필수 모듈 확인 중...")
    if not check_dependencies():
        input("\nEnter 키를 눌러 종료...")
        return
    
    # 3. 프로젝트 파일 확인
    print("\n3. 프로젝트 파일 확인 중...")
    if not check_project_files():
        input("\nEnter 키를 눌러 종료...")
        return
    
    # 4. GUI 실행
    print("\n4. GUI 실행 중...")
    run_gui()

if __name__ == "__main__":
    main()
