# 📊 Excel Quick Convert

구엑셀(.xls), 한컴오피스, 깨진 Excel 파일을 **최신 .xlsx 형식**으로 빠르게 변환하는 무료 웹 서비스입니다.

![Excel Quick Convert](https://img.shields.io/badge/Excel-Quick%20Convert-blue?style=for-the-badge&logo=microsoft-excel)
![Next.js](https://img.shields.io/badge/Next.js-14-black?style=for-the-badge&logo=next.js)
![TypeScript](https://img.shields.io/badge/TypeScript-5.3-blue?style=for-the-badge&logo=typescript)
![Vercel](https://img.shields.io/badge/Vercel-Deploy-black?style=for-the-badge&logo=vercel)

## 🎯 주요 기능

- **🚀 빠른 변환**: 업로드부터 다운로드까지 단 몇 초
- **📁 다양한 형식 지원**: .xls, .xlsx, .csv, .tsv, .txt
- **🇰🇷 한국어 최적화**: EUC-KR, CP949 인코딩 자동 감지
- **🛡️ 안전한 처리**: 파일은 변환 후 자동 삭제
- **📱 반응형 디자인**: 모든 기기에서 완벽 동작
- **🔄 텍스트 복구**: 표준 파서 실패 시 텍스트 기반 복구

## 🛠️ 기술 스택

- **Frontend**: Next.js 14, React 18, TypeScript
- **Styling**: Tailwind CSS
- **Excel Processing**: SheetJS (xlsx), iconv-lite, jschardet
- **Deployment**: Vercel
- **Icons**: Lucide React

## 🚀 빠른 시작

### 로컬 개발 환경 설정

1. **저장소 클론**
```bash
git clone https://github.com/your-username/excel-quick-convert.git
cd excel-quick-convert
```

2. **의존성 설치**
```bash
npm install
# 또는
yarn install
```

3. **환경 변수 설정**
```bash
cp env.example .env.local
```

4. **개발 서버 실행**
```bash
npm run dev
# 또는
yarn dev
```

5. **브라우저에서 확인**
```
http://localhost:3000
```

### 프로덕션 빌드

```bash
npm run build
npm run start
```

## 📦 Vercel 배포

### 1. GitHub 저장소 연결

1. GitHub에 저장소 생성 및 푸시
```bash
git add .
git commit -m "Initial commit: Excel Quick Convert"
git remote add origin https://github.com/your-username/excel-quick-convert.git
git push -u origin main
```

2. [Vercel 대시보드](https://vercel.com/dashboard)에서 "New Project" 클릭
3. GitHub 저장소 선택
4. 프로젝트 설정:
   - **Framework Preset**: Next.js
   - **Root Directory**: ./
   - **Build Command**: `npm run build`
   - **Output Directory**: `.next`

### 2. 환경 변수 설정 (선택사항)

Vercel 대시보드 → Settings → Environment Variables에서 설정:

```env
NEXT_PUBLIC_MAX_FILE_SIZE=20971520
NEXT_PUBLIC_ALLOWED_EXTENSIONS=.xls,.xlsx,.csv,.tsv,.txt
```

### 3. 배포 완료

- 자동 배포가 시작되며 몇 분 후 완료
- `https://your-project-name.vercel.app` 형태의 URL 제공
- 커스텀 도메인 연결 가능

## 📋 API 문서

### POST /api/convert

파일을 .xlsx 형식으로 변환합니다.

**요청**
- Content-Type: `multipart/form-data`
- Body:
  - `file`: 변환할 파일 (필수)
  - `forceTextRecovery`: 텍스트 복구 강제 실행 (선택, boolean)

**응답**
- 성공: `.xlsx` 파일 다운로드
- 실패: JSON 에러 메시지

**예시**
```javascript
const formData = new FormData();
formData.append('file', selectedFile);
formData.append('forceTextRecovery', 'false');

const response = await fetch('/api/convert', {
  method: 'POST',
  body: formData,
});

if (response.ok) {
  const blob = await response.blob();
  // 파일 다운로드 처리
}
```

### GET /api/health

서비스 상태를 확인합니다.

**응답**
```json
{
  "status": "ok",
  "timestamp": "2024-01-01T00:00:00.000Z",
  "service": "Excel Quick Convert API",
  "version": "1.0.0"
}
```

## 🔧 설정

### 파일 크기 제한

`next.config.js`에서 수정:
```javascript
api: {
  bodyParser: {
    sizeLimit: '20mb', // 원하는 크기로 변경
  },
}
```

### 지원 파일 형식

`lib/converter.ts`의 `SUPPORTED_EXTENSIONS` 배열에서 수정:
```typescript
export const SUPPORTED_EXTENSIONS = ['.xls', '.xlsx', '.csv', '.tsv', '.txt'];
```

## 🐛 문제 해결

### 변환 실패 시

1. **표준 변환 실패**: "텍스트 기반 복구" 옵션 체크
2. **인코딩 문제**: 파일을 메모장에서 UTF-8로 저장 후 재시도
3. **파일 손상**: 원본 프로그램에서 CSV로 저장 후 변환

### 성능 최적화

- 파일 크기 20MB 이하 권장
- 복잡한 서식이 포함된 파일은 데이터만 보존됨
- 매크로 포함 파일은 매크로 제거 후 변환

## 📊 지원 현황

### ✅ 완벽 지원
- 구 Microsoft Excel (.xls)
- 한컴오피스 Calc 파일
- 깨진 .xlsx 파일
- CSV/TSV (다양한 인코딩)

### ⚠️ 부분 지원
- 복잡한 서식 포함 파일 (데이터만 보존)
- 매크로 포함 파일 (매크로 제거 후 변환)

### ❌ 지원 불가
- 암호로 보호된 파일
- 심각하게 손상된 바이너리 파일

## 🤝 기여하기

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다. 자세한 내용은 `LICENSE` 파일을 참조하세요.

## 📞 문의

- 이슈 리포트: [GitHub Issues](https://github.com/your-username/excel-quick-convert/issues)
- 이메일: your-email@example.com

## 🙏 감사의 말

- [SheetJS](https://sheetjs.com/) - 강력한 Excel 처리 라이브러리
- [Next.js](https://nextjs.org/) - 훌륭한 React 프레임워크
- [Vercel](https://vercel.com/) - 간편한 배포 플랫폼
- [Tailwind CSS](https://tailwindcss.com/) - 유틸리티 우선 CSS 프레임워크

---

**⭐ 이 프로젝트가 도움이 되었다면 스타를 눌러주세요!**
