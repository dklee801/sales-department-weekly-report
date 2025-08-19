#!/usr/bin/env python3
"""
프로젝트 초기 설정 스크립트
새로운 환경에서 필요한 폴더들을 생성합니다.
"""
import os
from pathlib import Path

def create_data_structure():
    """데이터 폴더 구조 생성"""
    folders = [
        "data/sales_raw_data/2023",
        "data/sales_raw_data/2024", 
        "data/sales_raw_data/2025",
        "data/receivable_calculator_raw_data",
        "data/processed/backup",
        "data/downloads",
        "data/receivables", 
        "data/report"
    ]
    
    for folder in folders:
        Path(folder).mkdir(parents=True, exist_ok=True)
        print(f"✅ 폴더 생성: {folder}")

def create_config_template():
    """설정 파일 템플릿 생성"""
    accounts_template = {
        "erp_accounts": {
            "username": "YOUR_USERNAME_HERE",
            "password": "YOUR_PASSWORD_HERE"
        }
    }
    
    template_path = Path("config/accounts.json.example")
    template_path.parent.mkdir(parents=True, exist_ok=True)
    
    import json
    with open(template_path, 'w', encoding='utf-8') as f:
        json.dump(accounts_template, f, indent=2, ensure_ascii=False)
    
    print(f"✅ 설정 템플릿 생성: {template_path}")
    print("📝 실제 사용 시: cp config/accounts.json.example config/accounts.json")

if __name__ == "__main__":
    print("🚀 Sales Department Weekly Report 프로젝트 초기 설정")
    print("="*60)
    
    create_data_structure()
    create_config_template()
    
    print("\n🎉 프로젝트 초기 설정 완료!")
    print("\n📋 다음 단계:")
    print("1. pip install -r requirements.txt")
    print("2. config/accounts.json.example을 복사하여 실제 계정 정보 입력")
    print("3. python applications/gui.py 실행")
