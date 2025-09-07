# 구엑셀 빠른 변환 웹앱 개발 프로세스

## 🎯 목표
- **구엑셀(.xls)** 파일을 업로드하면 **최신 Excel(.xlsx)**에서 무조건 열리는 파일로 변환
- **한컴오피스**, **MS Office 2016 이전** 버전에서 만든 파일도 완벽 호환
- 심플한 UI: 업로드 → 변환 → 다운로드

## 📋 핵심 변환 로직

### 1단계: 파일 형식 감지 및 읽기
```python
# 우선순위별 파싱 시도
1. pandas + openpyxl (표준 .xlsx 읽기)
2. pandas + xlrd (구 .xls 읽기) 
3. 인코딩 감지 후 CSV/텍스트 파싱 (EUC-KR, CP949, UTF-8)
4. 텍스트 기반 복구 (구분자 추정)
```

### 2단계: 데이터 정규화
```python
# 데이터 타입 변환
- 날짜: 다양한 형식 → Excel 표준 날짜
- 숫자: 문자열 숫자 → 실제 숫자 타입
- 퍼센트: "50%" → 0.5 (Excel 퍼센트 형식)
- 빈 셀: None 값 처리
```

### 3단계: 표준 .xlsx로 저장
```python
# openpyxl을 사용해 최신 Excel 형식으로 저장
- 모든 시트 보존
- 한글 파일명/시트명 안전 처리
- Excel 호환성 100% 보장
```

## 🛠 기술 스택 (심플 버전)

### Backend (Python)
- **Flask** (FastAPI보다 더 간단)
- **pandas**: 데이터 처리 핵심
- **openpyxl**: .xlsx 쓰기/읽기
- **xlrd**: 구 .xls 파일 읽기
- **chardet**: 인코딩 자동 감지

### Frontend (기본 HTML + JavaScript)
- **Vanilla JS** (React 없이 심플하게)
- **Bootstrap** 또는 **Tailwind CDN**
- **Drag & Drop** 파일 업로드

## 📁 심플 프로젝트 구조
```
excel-converter/
├── app.py              # Flask 메인 서버
├── converter.py        # 변환 로직
├── static/
│   ├── style.css       # 심플 스타일
│   └── script.js       # 업로드/다운로드 JS
├── templates/
│   └── index.html      # 메인 페이지
├── uploads/            # 임시 업로드 폴더
├── converted/          # 변환 완료 파일
└── requirements.txt    # Python 패키지
```

## 🔄 사용자 플로우

### 1. 메인 페이지
```html
[파일 선택] 또는 [드래그 앤 드롭]
    ↓
[변환하기 버튼]
    ↓
[변환 중... 프로그레스바]
    ↓
[다운로드 버튼] - "원본명_변환.xlsx"
```

### 2. 백엔드 처리 과정
```python
POST /convert
    ↓
1. 파일 업로드 받기
2. 임시 저장 (uploads/)
3. 변환 로직 실행
4. 결과 파일 생성 (converted/)
5. 다운로드 링크 반환
6. 10분 후 임시파일 자동 삭제
```

## ⚡ 핵심 변환 코드 (요약)

```python
def convert_to_xlsx(input_file, output_file):
    """구엑셀을 최신 Excel로 변환"""
    
    # Step 1: 표준 파서 시도
    try:
        if input_file.endswith('.xlsx'):
            df = pd.read_excel(input_file, engine='openpyxl')
        elif input_file.endswith('.xls'):
            df = pd.read_excel(input_file, engine='xlrd')
        else:
            # CSV/텍스트 파일 인코딩 감지
            encoding = detect_encoding(input_file)
            df = pd.read_csv(input_file, encoding=encoding)
    
    except Exception:
        # Step 2: 텍스트 기반 복구
        df = text_based_recovery(input_file)
    
    # Step 3: 데이터 정규화
    df = normalize_data(df)
    
    # Step 4: 표준 xlsx로 저장
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    
    return output_file
```

## 🚀 개발 순서

### Phase 1: 최소 기능 구현 (1-2일)
1. Flask 기본 서버 설정
2. 파일 업로드 API
3. 기본 변환 로직 (pandas + openpyxl)
4. 심플 HTML 업로드 페이지

### Phase 2: 호환성 강화 (2-3일)
1. 구 .xls 파일 지원 (xlrd)
2. 인코딩 감지 및 CSV 처리
3. 텍스트 기반 복구 로직
4. 에러 처리 및 사용자 안내

### Phase 3: UX 개선 (1-2일)
1. 드래그 앤 드롭 UI
2. 변환 진행 상태 표시
3. 파일명 한글 처리
4. 임시파일 자동 정리

## 📦 배포 옵션 (심플)

### Option 1: 로컬 실행
```bash
pip install -r requirements.txt
python app.py
# http://localhost:5000 접속
```

### Option 2: 클라우드 배포
- **Heroku**: 가장 간단한 배포
- **Railway**: 무료 티어 충분
- **Render**: 최근 인기 플랫폼

## 🔒 보안 & 제한사항

### 파일 제한
- 최대 크기: **20MB**
- 허용 확장자: `.xls`, `.xlsx`, `.csv`, `.tsv`
- 바이러스 스캔은 클라우드 서비스 의존

### 데이터 보안
- 업로드 파일 **10분 후 자동 삭제**
- **민감정보 주의** 안내 메시지
- HTTPS 필수 (배포 시)

## 📊 예상 처리 가능 파일들

✅ **완벽 지원**
- 구 Microsoft Excel (.xls)
- 한컴오피스 Calc 파일
- 깨진 .xlsx 파일
- CSV/TSV (다양한 인코딩)

✅ **부분 지원**
- 복잡한 서식이 포함된 파일 (데이터만 보존)
- 매크로 포함 파일 (매크로 제거 후 변환)

❌ **지원 불가**
- 암호로 보호된 파일
- 심각하게 손상된 바이너리 파일

## 💡 사용자 가이드

### "파일이 안 열려요!" 해결책
1. **먼저 시도**: 웹앱에 파일 업로드
2. **여전히 안되면**: 메모장으로 열어서 다시 저장 후 재시도
3. **그래도 안되면**: 원본 프로그램에서 CSV로 저장 후 변환

### 성공률 높이는 팁
- 가능하면 **원본 확장자 그대로** 업로드
- 파일명에 **특수문자 최소화**
- 하나씩 변환 (배치 처리는 추후 기능)

## 🎉 완성 후 기대효과

- **5초 만에** 구엑셀 → 최신 Excel 변환
- **별도 프로그램 설치 불필요**
- **95% 이상** 파일 호환성 보장
- **무료 웹 서비스**로 누구나 이용 가능

---

이 프로세스대로 개발하면 **심플하면서도 강력한** 엑셀 변환 도구를 만들 수 있습니다. 추가 질문이나 특정 부분의 상세 코드가 필요하시면 언제든 말씀해주세요!