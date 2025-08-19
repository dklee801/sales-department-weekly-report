# Sales Department 주간보고서 자동화 시스템 v4.0

## 🎉 리팩토링 완료!

**sales-department-weekly-report**는 기존 Sales_department 프로젝트를 완전히 리팩토링한 버전입니다.

### 📋 주요 개선사항

#### 🏗️ 모듈화된 구조
- **명확한 책임 분리**: 각 모듈이 고유한 역할을 담당
- **유지보수성 향상**: 코드 수정 시 영향 범위가 제한됨
- **재사용성 증대**: 모듈 간 독립성으로 다른 프로젝트에서도 활용 가능

#### 📁 새로운 디렉토리 구조
```
sales-department-weekly-report/
├── applications/              # 실행 가능한 애플리케이션들
│   ├── main.py               # CLI 인터페이스
│   ├── gui.py                # GUI 애플리케이션
│   └── run_gui.py            # GUI 런처
├── modules/                   # 핵심 모듈들
│   ├── core/                 # 핵심 분석 로직
│   │   ├── sales_calculator.py
│   │   └── accounts_receivable_analyzer.py
│   ├── data/                 # 데이터 처리
│   │   ├── collectors/       # 데이터 수집기들
│   │   ├── processors/       # 데이터 가공기들
│   │   ├── validators/       # 데이터 검증기들
│   │   └── unified_data_collector.py
│   ├── gui/                  # GUI 컴포넌트들
│   │   ├── components/       # 재사용 가능한 UI 컴포넌트
│   │   └── login_dialog.py
│   ├── utils/                # 유틸리티 모듈들
│   │   ├── config_manager.py
│   │   ├── backup_manager.py
│   │   └── excel_safety_helper.py
│   └── reports/              # 보고서 생성 모듈들
│       └── xml_safe_report_generator.py
└── test_integration.py       # 통합 테스트 스크립트
```

## 🚀 사용 방법

### 1. GUI 모드 (권장)
```bash
cd sales-department-weekly-report
python applications/gui.py
```
또는
```bash
python applications/run_gui.py
```

### 2. CLI 모드
```bash
cd sales-department-weekly-report
python applications/main.py
```

### 3. 통합 테스트
```bash
python test_integration.py
```

## 🔧 기능 소개

### 📈 매출 데이터 관리
- **자동 수집**: ERP 시스템에서 매출 데이터 자동 다운로드
- **데이터 검증**: 수집된 데이터의 무결성 검사
- **집계 처리**: 월별/주차별 매출 집계 생성

### 💰 매출채권 분석
- **채권 현황 분석**: 회사별 매출채권 상태 분석
- **기간별 분류**: 90일 초과 채권, 결제기간 초과 채권 분류
- **자동 보고서**: Excel 형태의 분석 보고서 생성

### 📄 주간보고서 생성
- **표준 양식**: 2025년도 주간보고 양식에 맞춘 자동 생성
- **데이터 통합**: 매출 및 매출채권 데이터를 하나의 보고서로 통합
- **XML 안전**: 특수문자 처리로 Excel 호환성 보장

## 🎯 리팩토링 전후 비교

### ❌ 리팩토링 전 (기존 구조)
- 단일 파일에 모든 기능 집중
- 복잡한 의존성 관계
- 코드 재사용 어려움
- 유지보수 시 전체 시스템 영향

### ✅ 리팩토링 후 (새 구조)
- **모듈별 독립성**: 각 모듈이 독립적으로 작동
- **명확한 인터페이스**: 모듈 간 표준화된 통신 방식
- **확장 가능성**: 새로운 기능 추가 시 기존 코드에 영향 없음
- **테스트 용이성**: 개별 모듈 단위 테스트 가능

## 📦 필요한 패키지

```bash
pip install pandas openpyxl selenium tkinter pathlib logging xlsxwriter
```

선택적 패키지:
```bash
pip install tkcalendar  # 달력 위젯 (GUI 개선)
```

## 🏃‍♂️ 빠른 시작

1. **프로젝트 복제**
   ```bash
   git clone https://github.com/dklee801/sales-department-weekly-report.git
   cd sales-department-weekly-report
   ```

2. **의존성 설치**
   ```bash
   pip install pandas openpyxl selenium tkinter pathlib logging xlsxwriter
   pip install tkcalendar  # 선택적
   ```

3. **GUI 애플리케이션 실행**
   ```bash
   python applications/gui.py
   ```

## 🔍 문제 해결

### 자주 발생하는 문제

1. **ModuleNotFoundError**
   ```
   해결: 프로젝트 루트 디렉토리에서 실행하는지 확인
   cd sales-department-weekly-report
   python applications/gui.py
   ```

2. **Import 오류**
   ```
   해결: __init__.py 파일들이 모든 디렉토리에 있는지 확인
   python test_integration.py로 모듈 상태 점검
   ```

3. **GUI 실행 오류**
   ```
   해결: tkinter가 설치되어 있는지 확인
   python -m tkinter  # 테스트 실행
   ```

## 📈 성능 개선

### 메모리 사용량
- **모듈별 로딩**: 필요한 모듈만 선택적 로딩
- **의존성 최적화**: 불필요한 import 제거

### 실행 속도
- **지연 로딩**: 사용 시점에 모듈 로딩
- **캐시 활용**: 중복 계산 방지

## 🎯 향후 계획

### Phase 1: 안정화 (완료)
- ✅ 모듈 구조 리팩토링
- ✅ Import 경로 표준화
- ✅ 기본 기능 검증

### Phase 2: 확장 (예정)
- 📋 플러그인 시스템 도입
- 📋 웹 인터페이스 추가
- 📋 API 서버 구축

### Phase 3: 최적화 (예정)
- 📋 성능 튜닝
- 📋 대용량 데이터 처리
- 📋 병렬 처리 지원

## 🤝 기여 방법

1. 새로운 기능은 별도 모듈로 개발
2. 기존 인터페이스 유지
3. 단위 테스트 포함
4. 문서화 완료

## 📞 지원

- **기술 문의**: GitHub Issues
- **버그 신고**: GitHub Issues

---

**Sales Department Weekly Report v4.0** - 더 나은 코드, 더 나은 경험을 위해 🚀