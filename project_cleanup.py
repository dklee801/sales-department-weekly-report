#!/usr/bin/env python3
"""
프로젝트 정리 스크립트
불필요한 개발/임시 파일들을 정리합니다.
"""

import os
import shutil
from pathlib import Path

def clean_project():
    """프로젝트 폴더 정리"""
    base_path = Path("C:/Users/dklee/Python/weelky_report/Sales_department_refactored")
    
    # 1. 개발/임시 파일들 삭제
    temp_files = [
        "debug_nas_files.py",
        "final_guide_report_week_current_week.py",
        "final_improvement_summary.py", 
        "fix_weekly_selection_guide.py",
        "gui_message_enhancement.py",
        "gui_week_selection_fix.py",
        "verify_fix.py",
        "verify_receivables_fix.py",
        "git_sync_en.bat",
        "sync_git.bat",
        "README_receivables_sync.md"
    ]
    
    deleted_files = []
    for file_name in temp_files:
        file_path = base_path / file_name
        if file_path.exists():
            file_path.unlink()
            deleted_files.append(file_name)
            print(f"✅ 삭제완료: {file_name}")
    
    # 2. __pycache__ 폴더들 삭제
    pycache_folders = []
    for pycache in base_path.rglob("__pycache__"):
        shutil.rmtree(pycache)
        pycache_folders.append(str(pycache.relative_to(base_path)))
        print(f"✅ 캐시삭제: {pycache.relative_to(base_path)}")
    
    # 3. 백업 폴더 정리
    backup_path = base_path / "data" / "processed" / "backup"
    if backup_path.exists():
        backup_files = list(backup_path.glob("*.xlsx"))
        if len(backup_files) > 3:  # 최신 3개만 유지
            backup_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            for old_backup in backup_files[3:]:
                old_backup.unlink()
                print(f"✅ 백업삭제: {old_backup.name}")
    
    # 4. 보고서 폴더 정리 (8월 19일 이후만 유지)
    report_path = base_path / "data" / "report"
    if report_path.exists():
        report_files = list(report_path.glob("주간보고서_202508*.xlsx"))
        old_reports = [f for f in report_files if "20250814" in f.name or "20250818" in f.name and f.name < "주간보고서_20250819"]
        for old_report in old_reports[:10]:  # 안전하게 최대 10개만
            old_report.unlink()
            print(f"✅ 보고서삭제: {old_report.name}")
    
    print(f"\n🎉 정리 완료!")
    print(f"📁 삭제된 파일: {len(deleted_files)}개")
    print(f"🗂️ 삭제된 캐시: {len(pycache_folders)}개")
    print(f"💾 프로젝트 크기 최적화 완료")

if __name__ == "__main__":
    clean_project()
