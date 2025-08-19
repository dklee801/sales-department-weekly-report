# Sales Department Weekly Report System

**주간보고서 자동화 시스템 v4.0 (리팩토링 완료)**

매출부서 주간보고서 생성을 위한 완전 자동화 시스템입니다. 매출 데이터 수집부터 매출채권 분석, 최종 보고서 생성까지 전 과정을 자동화합니다.

## 🚀 **주요 기능**

### 📊 **매출 데이터 처리**
- **자동 수집**: ERP 시스템에서 매출 데이터 자동 추출 (1-24개월)
- **다중 회사 지원**: 디앤드디, 디앤아이, 후지리프트코리아
- **지능형 집계**: 월별/주차별 자동 집계 및 비교 분석
- **데이터 검증**: 자동 오류 검출 및 수정

### 💰 **매출채권 분석**
- **주간 기준 분석**: 선택된 목요일 기준 주간 분석
- **전주 대비 변화**: 90일 초과 채권, 결제기간 초과 채권 증감률
- **TOP20 분석**: 결제기간 초과 채권 상위 거래처 분석
- **자동 동기화**: 분석 결과를 보고서 양식에 자동 반영

### 📋 **보고서 생성**
- **표준 양식 준수**: 2025년도 주간보고 양식_2.xlsx 호환
- **병합 셀 처리**: Excel 병합 셀 안전 처리
- **자동 데이터 통합**: 매출 + 매출채권 데이터 통합
- **XML 안전성**: 특수문자 및 대용량 데이터 안전 처리

## 🏗️ **시스템 아키텍처 (리팩토링 완료)**

```
📁 sales-department-weekly-report/
├── 📁 applications/           # 실행 파일들
│   ├── main.py               # CLI 메인 실행
│   ├── gui.py                # GUI 실행
│   └── run_gui.py            # GUI 런처
├── 📁 modules/               # 핵심 모듈들
│   ├── 📁 core/              # 핵심 비즈니스 로직
│   │   ├── sales_calculator.py         # 매출 계산
│   │   ├── accounts_receivable_analyzer.py  # 매출채권 분석
│   │   └── processed_receivables_analyzer.py # 매출채권 처리
│   ├── 📁 data/              # 데이터 처리
│   │   ├── 📁 processors/    # 데이터 가공
│   │   └── unified_data_collector.py   # 통합 데이터 수집
│   ├── 📁 reports/           # 보고서 생성
│   │   └── xml_safe_report_generator.py # 안전 보고서 생성
│   ├── 📁 gui/               # GUI 컴포넌트
│   │   └── login_dialog.py   # 로그인 처리
│   └── 📁 utils/             # 유틸리티
│       └── config_manager.py # 설정 관리
├── 📁 data/                  # 데이터 저장소
│   ├── 📁 raw/               # 원시 데이터
│   ├── 📁 processed/         # 처리된 데이터
│   └── 📁 report/            # 최종 보고서
└── 📄 2025년도 주간보고 양식_2.xlsx  # 보고서 템플릿
```

## 🚀 **빠른 시작**

### 1. **설치**
```bash
git clone https://github.com/dklee801/sales-department-weekly-report.git
cd sales-department-weekly-report
pip install -r requirements.txt
```

### 2. **GUI 실행** (권장)
```bash
python applications/gui.py
```

### 3. **CLI 실행**
```bash
# 전체 프로세스 (매출 + 매출채권 + 보고서)
python applications/main.py

# 매출 데이터만 분석
python applications/main.py --process

# 보고서만 생성
python applications/main.py --report

# 6개월 데이터로 전체 프로세스
python applications/main.py --months 6
```

## 🎯 **사용 시나리오**

### 📈 **시나리오 1: 주간 정기 보고서**
1. **GUI 실행** → ERP 계정 입력
2. **보고서 주간 선택** → 목요일 기준 선택
3. **🚀 전체 프로세스 실행** → 자동 완성 (5-10분)
4. **결과 확인** → `data/report/주간보고서_YYYYMMDD_HHMM.xlsx`

### 💰 **시나리오 2: 매출채권 분석만**
1. **GUI 실행** → 보고서 주간 선택
2. **💰 매출채권 분석** → 선택된 주간 기준 분석
3. **결과 확인** → `data/processed/채권_분석_결과.xlsx`

### 📊 **시나리오 3: 매출 데이터 수집**
1. **매출 수집 기간 선택** (1-24개월)
2. **📈 매출 데이터 갱신** → 원시 데이터 수집
3. **🔄 매출집계 처리** → 분석용 데이터 생성

## ⚙️ **설정 및 환경**

### 🔧 **필수 요구사항**
- **Python**: 3.8 이상
- **Chrome/Edge**: 웹드라이버 자동 설치
- **Windows**: 10/11 (테스트 완료)
- **네트워크**: ERP 시스템 접근 권한

### 📂 **디렉토리 구조**
```bash
# 데이터 디렉토리 자동 생성
data/
├── raw/                    # 원시 데이터 (수집된 파일들)
│   ├── sales/             # 매출 원시 데이터
│   └── receivables/       # 매출채권 원시 데이터
├── processed/             # 처리된 중간 결과
│   ├── 매출집계_결과.xlsx
│   └── 채권_분석_결과.xlsx
└── report/                # 최종 보고서
    └── 주간보고서_YYYYMMDD_HHMM.xlsx
```

## 🛠️ **고급 사용법**

### 🎛️ **CLI 옵션**
```bash
# 조용한 모드 (최소 출력)
python applications/main.py --quiet

# 브라우저 창 표시 (디버깅)
python applications/main.py --show-browser

# 매출채권만 분석
python applications/main.py --collect-receivables

# 도움말
python applications/main.py --help
```

### 🔧 **설정 파일**
- **자동 생성**: 첫 실행 시 자동으로 설정 파일 생성
- **경로 설정**: 데이터 저장 경로 사용자 정의 가능
- **계정 관리**: ERP 계정 정보 안전 저장

## 📋 **주요 개선사항 (v4.0)**

### 🏗️ **아키텍처 리팩토링**
- ✅ **모듈화 완료**: 기능별 독립 모듈 분리
- ✅ **설정 중앙화**: config_manager를 통한 통합 관리
- ✅ **표준화된 import**: 일관된 모듈 참조 방식
- ✅ **백워드 호환성**: 기존 기능 100% 보장

### 🛡️ **안정성 강화**
- ✅ **오류 격리**: 모듈 간 오류 전파 방지
- ✅ **자동 복구**: 데이터 오류 자동 감지 및 수정
- ✅ **병합 셀 처리**: Excel 복잡 구조 안전 처리
- ✅ **메모리 최적화**: 대용량 데이터 효율 처리

### 🚀 **성능 향상**
- ✅ **병렬 처리**: 다중 회사 데이터 동시 수집
- ✅ **지능형 캐싱**: 중복 데이터 처리 최적화
- ✅ **점진적 로딩**: 필요한 모듈만 동적 로드
- ✅ **자동 정리**: 임시 파일 자동 정리

## 🐛 **문제해결**

### ❓ **자주 묻는 질문**

**Q: "매출채권 분석 실패" 오류가 발생해요**
```
A: 1. 매출채권 파일이 data/raw/receivables/에 있는지 확인
   2. 파일명이 "매출채권계산결과YYYYMMDD.xlsx" 형식인지 확인
   3. GUI에서 올바른 보고서 주간을 선택했는지 확인
```

**Q: GUI가 실행되지 않아요**
```
A: 1. pip install -r requirements.txt 실행
   2. tkinter 설치: sudo apt-get install python3-tk (Linux)
   3. tkcalendar 설치: pip install tkcalendar
```

**Q: 브라우저가 자동으로 닫혀요**
```
A: 1. Chrome/Edge 최신 버전 확인
   2. --show-browser 옵션으로 디버깅 모드 실행
   3. 방화벽/보안 프로그램 확인
```

### 🔍 **로그 확인**
```bash
# 상세 로그 확인
python applications/main.py --show-browser

# GUI 오류 확인
python applications/gui.py  # 콘솔에서 실행
```

## 📞 **지원 및 문의**

- **GitHub Issues**: [이슈 등록](https://github.com/dklee801/sales-department-weekly-report/issues)
- **버그 리포트**: 재현 단계와 오류 로그 첨부
- **기능 요청**: 구체적인 사용 사례와 함께 요청

## 📄 **라이선스**

이 프로젝트는 내부 사용을 위한 프로젝트입니다.

---

## 🎉 **성공적인 리팩토링 완료!**

**v4.0에서는 완전히 새로운 아키텍처로 재구성되어 더욱 안정적이고 확장 가능한 시스템이 되었습니다.**

✨ **모듈화 → 유지보수성 향상**  
🛡️ **안정성 → 오류 복구 능력 강화**  
🚀 **성능 → 처리 속도 대폭 개선**  
🎯 **사용성 → 더욱 직관적인 인터페이스**
