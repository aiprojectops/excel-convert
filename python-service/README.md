# Excel Converter Python Service

구 Excel 파일을 최신 .xlsx 형식으로 변환하는 FastAPI 마이크로서비스입니다.

## 🚀 기능

- ✅ **다양한 Excel 형식 지원**: .xls, .xlsx, .xlsm
- ✅ **한글 인코딩 자동 감지**: CP949, EUC-KR, UTF-8
- ✅ **CSV 파일 변환**: 다양한 구분자 자동 감지
- ✅ **강력한 오류 처리**: 손상된 파일도 최대한 복구
- ✅ **RESTful API**: 간단한 POST 요청으로 변환

## 📦 설치 및 실행

### 로컬 개발
```bash
cd python-service
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```

### Docker 실행
```bash
docker build -t excel-converter .
docker run -p 8000:8000 excel-converter
```

## 🔧 API 사용법

### 파일 변환
```bash
curl -X POST "http://localhost:8000/convert" \
  -H "accept: application/json" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@example.xls"
```

### 헬스 체크
```bash
curl http://localhost:8000/health
```

## 🌐 배포 (Render)

1. GitHub에 코드 푸시
2. Render 대시보드에서 새 서비스 생성
3. `python-service` 폴더를 루트로 설정
4. 자동 배포 완료!

## 📊 지원 형식

### 입력 형식
- `.xls` (Excel 97-2003)
- `.xlsx` (Excel 2007+)
- `.xlsm` (Excel 매크로 포함)
- `.csv` (다양한 인코딩)
- `.tsv` (탭 구분)

### 출력 형식
- `.xlsx` (Excel 2007+ 호환)

## 🔍 변환 과정

1. **openpyxl 엔진**: 최신 Excel 파일 우선 시도
2. **xlrd 엔진**: 구 Excel 파일 (.xls) 처리
3. **인코딩 감지**: chardet으로 자동 감지
4. **CSV 변환**: 다양한 구분자와 인코딩 시도
5. **스타일링**: 헤더 볼드 처리 및 색상 적용

## 🚨 제한사항

- 최대 파일 크기: 50MB
- 매크로는 제거됨 (보안상 이유)
- 복잡한 차트/그래프는 단순화될 수 있음
