아래는 구엑셀 파일을 엑셀에서 열리도록 변환하는 “빠른 변환(quick convert) 전용 웹앱”의 **개발 계획 + 저장소 구조 + 필수 설정/패키지 + README 초안**입니다. 서버는 Python(FastAPI), 프런트는 Next.js(React + Tailwind) 기준으로 설계했습니다. 한글/영문 파일명, EUC-KR/CP949 인코딩도 안전하게 처리합니다.

---

# 1) 제품 개요

- 목적: 업로드한 구엑셀(.xls)·한컴/비표준 계열·깨진 .xlsx·csv/tsv를 **가급적 표준 파서로** 읽고, 실패 시 **텍스트 기반 파싱**으로 **깨끗한 .xlsx**로 즉시 재생성하여 다운로드 제공
- 범위: “빠른 변환”만 (정밀 변환·LibreOffice 경로 제외)
- 목표 크기: 단일 파일 ≤ 20–30MB(기본), 확장 가능
- 비기능:
    - 안전한 파일 처리(격리 디렉터리·자동 삭제)
    - 확실한 인코딩 감지(fallback) 및 날짜/숫자/퍼센트 캐스팅
    - 한글 파일명/시트명 안전화(영문/숫자 대체 규칙 포함)

---

# 2) 기술 스택

- **Frontend**: Next.js 14 (App Router), React 18, Tailwind CSS
- **Backend**: FastAPI + Uvicorn, Python 3.11
- **Parsing**: pandas, openpyxl, xlrd(=1.2.0), chardet, python-dateutil, (선택) pyxlsb
- **빌드/배포**: Docker, docker-compose (또는 Render/Heroku/Fly.io)
- **테스트/품질**: pytest, ruff(린터), black(포매터)
- **기타**: dotenv, Sentry(선택), rate limiting(Starlette middleware)

---

# 3) 아키텍처

- Next.js 정적/서버액션 → `/api/convert`(FastAPI)로 **multipart/form-data 업로드**
- FastAPI가 업로드 파일을 임시 폴더에 저장 → “표준 파서 → 텍스트 복구” 순으로 변환 → 결과 `.xlsx` 파일 경로 반환 → 프런트가 다운로드 트리거
- 변환 실패/경고 로그는 JSON으로 함께 반환(프런트에 토스트/배너)

---

# 4) API 설계

`POST /api/convert`

- 요청: multipart/form-data
    - `file`: 업로드 파일(필수)
    - `force_text`(optional, bool): 표준 파서 무시하고 텍스트 복구 우선
- 응답(성공): `application/octet-stream` (바이너리 스트림 다운로드)
    - 헤더: `Content-Disposition: attachment; filename="<원본명_복원.xlsx>"`
    - 추가 메타는 쿼리 `?meta=1`일 때 JSON 프리플라이트로 제공 가능
- 응답(에러): `4xx/5xx` JSON `{code, message, details?}`


제한/보안:

- 파일 크기 제한(기본 30MB), 확장자 화이트리스트
- 업로드 후 10분 내 임시파일 삭제(백그라운드 GC)
- CORS: 프런트 도메인만 허용
- Rate limit: IP 단위 20 req/min

# 5) 저장소 구조 (monorepo)

excel-quickconvert-web/
├─ apps/
│  ├─ web/                      # Next.js 프런트엔드
│  │  ├─ app/
│  │  │  ├─ page.tsx           # 업로드/결과 페이지
│  │  │  ├─ api/health/route.ts
│  │  │  └─ (components ...)   # UI 컴포넌트
│  │  ├─ public/
│  │  ├─ styles/
│  │  ├─ package.json
│  │  └─ tailwind.config.ts
│  └─ api/                      # FastAPI 백엔드
│     ├─ src/
│     │  ├─ main.py             # FastAPI 엔트리
│     │  ├─ config.py
│     │  ├─ services/
│     │  │  ├─ quick_convert.py # 변환 핵심 로직(빠른 변환)
│     │  │  └─ io_utils.py      # 저장/삭제/이름정리
│     │  ├─ middleware/
│     │  │  ├─ rate_limit.py
│     │  │  └─ cleanup.py
│     │  └─ schemas.py
│     ├─ pyproject.toml
│     └─ README.md
├─ docker/
│  ├─ web.Dockerfile
│  └─ api.Dockerfile
├─ compose.yaml
├─ .env.example
├─ README.md                    # 루트 설명서
└─ LICENSE


---

# 6) 패키지 정의

## 6.1 Backend (FastAPI) – `apps/api/pyproject.toml`
[project]
name = "excel-quickconvert-api"
version = "0.1.0"
requires-python = ">=3.11"
dependencies = [
  "fastapi==0.112.2",
  "uvicorn[standard]==0.30.6",
  "python-multipart==0.0.9",
  "pandas==2.2.2",
  "openpyxl==3.1.5",
  "xlrd==1.2.0",
  "chardet==5.2.0",
  "python-dateutil==2.9.0.post0",
  "pyxlsb==1.0.10",
  "pydantic==2.8.2",
  "python-dotenv==1.0.1",
]

[tool.black]
line-length = 100

[tool.ruff]
line-length = 100
select = ["E", "F", "I"]

## 6.2 Frontend (Next.js) – `apps/web/package.json`
{
  "name": "excel-quickconvert-web",
  "private": true,
  "scripts": {
    "dev": "next dev -p 3000",
    "build": "next build",
    "start": "next start -p 3000",
    "lint": "next lint"
  },
  "dependencies": {
    "next": "14.2.4",
    "react": "18.3.1",
    "react-dom": "18.3.1",
    "zod": "3.23.8",
    "axios": "1.7.2"
  },
  "devDependencies": {
    "autoprefixer": "10.4.19",
    "postcss": "8.4.38",
    "tailwindcss": "3.4.7",
    "typescript": "5.5.3",
    "@types/react": "18.3.3",
    "@types/node": "20.12.12",
    "eslint": "8.57.0",
    "eslint-config-next": "14.2.4"
  }
}

## 6.3 루트 `.env.example`
# 공통
NODE_ENV=development

# Backend
API_PORT=8000
MAX_UPLOAD_MB=30
ALLOWED_ORIGINS=http://localhost:3000
TEMP_DIR=/tmp/excelqc

# Frontend
NEXT_PUBLIC_API_BASE=http://localhost:8000

---

# 7) 핵심 코드 스켈레톤

## 7.1 FastAPI 엔트리 – `apps/api/src/main.py
`
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import os, io, uuid
from dotenv import load_dotenv
from services.quick_convert import quick_convert
from services.io_utils import safe_filename, ensure_dirs, cleanup_later

load_dotenv()
app = FastAPI(title="Excel Quick Convert API")

origins = os.getenv("ALLOWED_ORIGINS", "http://localhost:3000").split(",")
app.add_middleware(
    CORSMiddleware, allow_origins=origins, allow_credentials=True,
    allow_methods=["*"], allow_headers=["*"]
)

TEMP_DIR = os.getenv("TEMP_DIR", "/tmp/excelqc")
MAX_MB = int(os.getenv("MAX_UPLOAD_MB", "30"))

ensure_dirs([TEMP_DIR])

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/api/convert")
async def convert(file: UploadFile = File(...), force_text: bool = False):
    if not file.filename:
        raise HTTPException(400, "파일명이 없습니다.")
    content = await file.read()
    if len(content) > MAX_MB * 1024 * 1024:
        raise HTTPException(413, f"파일이 너무 큽니다(≤ {MAX_MB}MB).")

    uid = str(uuid.uuid4())[:8]
    safe_name = safe_filename(file.filename)
    in_path = os.path.join(TEMP_DIR, f"{uid}_{safe_name}")
    out_path = os.path.splitext(in_path)[0] + "_복원.xlsx"

    with open(in_path, "wb") as f:
        f.write(content)

    try:
        quick_convert(in_path, out_path, force_text=force_text)
        cleanup_later([in_path, out_path], delay_sec=600)  # 10분 후 삭제
        data = open(out_path, "rb").read()
        fname = os.path.basename(out_path)
        headers = {"Content-Disposition": f'attachment; filename="{fname}"'}
        return StreamingResponse(io.BytesIO(data),
                                 media_type="application/octet-stream",
                                 headers=headers)
    except Exception as e:
        raise HTTPException(500, f"변환 실패: {e}")

## 7.2 변환 서비스 – `apps/api/src/services/quick_convert.py`

> 이전에 드린 `quick_convert.py`의 로직을 함수화·모듈화한 버전(요지만 표시)

import os, re
import chardet
from dateutil import parser as dparser
import pandas as pd

# ... (best_delimiter, parse_cell 등 유틸 포함)

def quick_convert(in_path: str, out_path: str, force_text: bool = False):
    if not force_text:
        try:
            sheets = try_pandas_parsers(in_path)
            write_dict_to_xlsx(sheets, out_path)
            return
        except Exception:
            pass
    # fallback: text reconstruction
    with open(in_path, "rb") as f:
        raw = f.read()
    text_reconstruct_to_xlsx(raw, out_path)
## 7.3 IO 유틸 – `apps/api/src/services/io_utils.py`

import os, re, threading, time

SAFE = re.compile(r"[^a-zA-Z0-9가-힣._-]")

def safe_filename(name: str) -> str:
    name = name.strip().replace(" ", "_")
    return SAFE.sub("_", name)

def ensure_dirs(paths):
    for p in paths:
        os.makedirs(p, exist_ok=True)

def cleanup_later(paths, delay_sec=600):
    def _run():
        time.sleep(delay_sec)
        for p in paths:
            try:
                if os.path.exists(p):
                    os.remove(p)
            except:
                pass
    threading.Thread(target=_run, daemon=True).start()
## 7.4 Next.js 업로드 UI – `apps/web/app/page.tsx`

"use client";
import { useState, useRef } from "react";
const API = process.env.NEXT_PUBLIC_API_BASE || "http://localhost:8000";

export default function Home() {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string| null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const onUpload = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    setError(null);
    const file = inputRef.current?.files?.[0];
    if (!file) return;
    setLoading(true);
    try {
      const fd = new FormData();
      fd.append("file", file);
      const res = await fetch(`${API}/api/convert`, { method: "POST", body: fd });
      if (!res.ok) {
        const msg = await res.text();
        throw new Error(msg || "변환 실패");
      }
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = file.name.replace(/\.[^.]+$/, "") + "_복원.xlsx";
      a.click();
      window.URL.revokeObjectURL(url);
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <main className="min-h-screen flex items-center justify-center p-6">
      <div className="w-full max-w-xl border rounded-2xl p-6 shadow-sm">
        <h1 className="text-2xl font-bold mb-4">Excel Quick Convert</h1>
        <form onSubmit={onUpload} className="space-y-4">
          <input ref={inputRef} type="file" accept=".xls,.xlsx,.csv,.tsv,.txt" className="w-full" />
          <button disabled={loading} className="px-4 py-2 rounded-xl border">
            {loading ? "변환 중..." : "변환하기"}
          </button>
        </form>
        {error && <p className="text-red-600 mt-4">{error}</p>}
        <p className="text-sm text-gray-500 mt-4">
          구엑셀/한컴/깨진 파일도 시도합니다. 민감정보 포함 파일 업로드 주의.
        </p>
      </div>
    </main>
  );
}

---

# 8) Docker & 실행

## 8.1 Backend Dockerfile – `docker/api.Dockerfile`

FROM python:3.11-slim
WORKDIR /app
COPY apps/api/pyproject.toml /app/pyproject.toml
RUN pip install --upgrade pip && \
    pip install .  --no-cache-dir
COPY apps/api/src /app/src
ENV PYTHONUNBUFFERED=1
EXPOSE 8000
CMD ["uvicorn", "src.main:app", "--host", "0.0.0.0", "--port", "8000"]
주의: pip install .을 사용하려면 pyproject.toml에 [build-system] 섹션을 추가하거나, 간단히 requirements.txt 방식을 써도 됩니다. 필요시 바꿔드릴게요.

## 8.2 Frontend Dockerfile – `docker/web.Dockerfile`

FROM node:20-alpine AS builder
WORKDIR /app
COPY apps/web/package.json apps/web/package-lock.json* ./ 
RUN npm ci
COPY apps/web/ .
RUN npm run build

FROM node:20-alpine
WORKDIR /app
ENV NODE_ENV=production
COPY --from=builder /app ./
EXPOSE 3000
CMD ["npm", "start"]

## 8.3 docker-compose – `compose.yaml`

services:
  api:
    build:
      context: .
      dockerfile: docker/api.Dockerfile
    env_file: .env
    environment:
      - ALLOWED_ORIGINS=http://localhost:3000
      - TEMP_DIR=/tmp/excelqc
      - MAX_UPLOAD_MB=30
    ports: ["8000:8000"]
  web:
    build:
      context: .
      dockerfile: docker/web.Dockerfile
    env_file: .env
    environment:
      - NEXT_PUBLIC_API_BASE=http://localhost:8000
    ports: ["3000:3000"]
    depends_on: [api]

로컬 실행:

cp .env.example .env
docker compose up --build
# http://localhost:3000 접속
---

# 9) README 초안 (루트 `README.md`)

# Excel Quick Convert (Web)

업로드한 구엑셀/한컴/깨진 엑셀/CSV/TSV 파일을 **최신 .xlsx**로 즉시 변환하는 웹앱.

## Features
- 표준 파서 우선: .xlsx/.xls/.csv/.tsv 자동 인식
- 실패 시 텍스트 기반 복구(구분자 추정, 숫자/날짜/퍼센트 캐스팅)
- EUC-KR/CP949/UTF-8 인코딩 감지
- 결과 파일: `<원본명>_복원.xlsx`

## Tech
- Frontend: Next.js 14, Tailwind
- Backend: FastAPI (pandas/openpyxl/xlrd/chardet)
- Docker & docker-compose

## Quick Start (Docker)
```bash
cp .env.example .env
docker compose up --build
# open http://localhost:3000

## Local Dev

- Frontend
cd apps/web
npm i
npm run dev
    
- Backend
cd apps/api
uvicorn src.main:app --reload --port 8000
    

## Env

- `NEXT_PUBLIC_API_BASE` (FE → BE 주소)
    
- `ALLOWED_ORIGINS` (CORS 허용 도메인, 쉼표 구분)
    
- `MAX_UPLOAD_MB` (기본 30)
    
- `TEMP_DIR` (임시 디렉터리)
    

## API

`POST /api/convert` (multipart/form-data)

- file: 업로드 파일 (필수)
    
- force_text: true시 텍스트 복구 우선
    

성공: 바이너리 .xlsx 다운로드  
실패: `{code, message}` JSON

## Limits & Security

- 크기 제한(기본 30MB)
    
- 임시파일 10분 내 삭제
    
- 도메인 기반 CORS
    
- (선택) IP당 Rate limit
    

## Roadmap

- 폴더 일괄 변환(Zip 업로드)
    
- 변환 로그 시트 포함 옵션
    
- 접근제어(로그인/토큰)
    
- S3/GCS 저장선택
# 10) 운영 체크리스트

- [ ] 파일 크기/형식 검증(화이트리스트 `.xls,.xlsx,.csv,.tsv,.txt`)  
- [ ] 한글/특수문자 파일명 안전화  
- [ ] 임시파일 삭제 타이머 동작 확인  
- [ ] 대용량 파일 시 타임아웃/메모리 사용량 모니터  
- [ ] 에러 메시지 한국어화 및 사용자 안내 가이드(“메모장 저장 후 재시도” 등)

---

# 11) 확장 포인트(선택)

- Zip 업로드(다중 파일 변환 → Zip로 재다운로드)  
- S3 업로드 후 서명 URL 다운로드  
- 변환 규칙 템플릿(열 서식 강제, 첫 행 헤더/무시 토글)  
- OCR(스캔 이미지) 경고 및 별도 경로

---

원하시면 이 구조 그대로 **초기 템플릿 저장소**를 만들어 드릴 수 있어요. 배포 대상(예: Render/EC2/도메인/HTTPS)만 알려주시면 배포 스크립트까지 포함해 드리겠습니다.