from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import openpyxl
import xlrd
import io
import logging
from typing import Optional
import chardet

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="Excel Converter Service", version="1.0.0")

# CORS 설정 (Vercel에서 호출 가능하도록)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 프로덕션에서는 특정 도메인으로 제한
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
async def root():
    return {"message": "Excel Converter Service is running!", "version": "1.0.0"}

@app.get("/health")
async def health_check():
    return {"status": "healthy", "service": "excel-converter"}

@app.post("/convert")
async def convert_excel(file: UploadFile = File(...)):
    """
    Excel 파일을 최신 .xlsx 형식으로 변환
    - 구 Excel 파일 (.xls) 지원
    - 한글 인코딩 문제 해결
    - 다양한 Excel 형식 호환
    """
    try:
        logger.info(f"📁 파일 업로드됨: {file.filename} ({file.content_type})")
        
        # 파일 내용 읽기
        file_content = await file.read()
        file_size = len(file_content)
        logger.info(f"📊 파일 크기: {file_size} bytes")
        
        # 파일 크기 제한 (50MB)
        if file_size > 50 * 1024 * 1024:
            raise HTTPException(status_code=413, detail="파일 크기가 50MB를 초과합니다.")
        
        # 파일 확장자 확인
        filename = file.filename.lower() if file.filename else ""
        logger.info(f"🔍 파일명 분석: {filename}")
        
        # DataFrame으로 변환 시도 (여러 방법)
        df = None
        conversion_method = ""
        
        # 방법 1: openpyxl 엔진 (.xlsx, .xlsm)
        try:
            logger.info("🔧 openpyxl 엔진으로 시도...")
            df = pd.read_excel(io.BytesIO(file_content), engine='openpyxl')
            conversion_method = "openpyxl"
            logger.info("✅ openpyxl 성공!")
        except Exception as e:
            logger.info(f"❌ openpyxl 실패: {str(e)}")
        
        # 방법 2: xlrd 엔진 (.xls, 구 Excel)
        if df is None:
            try:
                logger.info("🔧 xlrd 엔진으로 시도...")
                df = pd.read_excel(io.BytesIO(file_content), engine='xlrd')
                conversion_method = "xlrd"
                logger.info("✅ xlrd 성공!")
            except Exception as e:
                logger.info(f"❌ xlrd 실패: {str(e)}")
        
        # 방법 3: 인코딩 감지 후 CSV 시도
        if df is None:
            try:
                logger.info("🔧 인코딩 감지 후 CSV 시도...")
                # 인코딩 감지
                detected = chardet.detect(file_content)
                encoding = detected.get('encoding', 'utf-8')
                logger.info(f"🔍 감지된 인코딩: {encoding}")
                
                # CSV로 읽기 시도
                text_content = file_content.decode(encoding)
                df = pd.read_csv(io.StringIO(text_content))
                conversion_method = f"csv-{encoding}"
                logger.info("✅ CSV 변환 성공!")
            except Exception as e:
                logger.info(f"❌ CSV 변환 실패: {str(e)}")
        
        # 방법 4: 다양한 구분자로 CSV 시도
        if df is None:
            try:
                logger.info("🔧 다양한 구분자로 CSV 시도...")
                for sep in ['\t', ';', '|', ',']:
                    try:
                        for encoding in ['utf-8', 'cp949', 'euc-kr', 'latin1']:
                            text_content = file_content.decode(encoding)
                            df = pd.read_csv(io.StringIO(text_content), sep=sep)
                            if len(df.columns) > 1:  # 최소 2개 컬럼 이상
                                conversion_method = f"csv-{encoding}-{sep}"
                                logger.info(f"✅ CSV 변환 성공! (구분자: {sep}, 인코딩: {encoding})")
                                break
                    except:
                        continue
                    if df is not None:
                        break
            except Exception as e:
                logger.info(f"❌ 다양한 구분자 CSV 실패: {str(e)}")
        
        # 모든 방법 실패시 오류
        if df is None:
            logger.error("❌ 모든 변환 방법 실패")
            raise HTTPException(
                status_code=400, 
                detail="지원되지 않는 파일 형식이거나 손상된 파일입니다."
            )
        
        # DataFrame 정보 로깅
        rows, cols = df.shape
        logger.info(f"📊 변환 결과: {rows}행 × {cols}열 (방법: {conversion_method})")
        logger.info(f"📋 컬럼명: {list(df.columns)[:5]}...")  # 처음 5개만
        
        # 빈 값 처리
        df = df.fillna('')
        
        # 새로운 Excel 파일로 저장
        output_buffer = io.BytesIO()
        
        # Excel 작성 옵션
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            
            # 워크시트 스타일링 (선택사항)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # 헤더 스타일 적용
            from openpyxl.styles import Font, PatternFill
            header_font = Font(bold=True)
            header_fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
            
            for cell in worksheet[1]:  # 첫 번째 행 (헤더)
                cell.font = header_font
                cell.fill = header_fill
        
        output_buffer.seek(0)
        output_data = output_buffer.getvalue()
        
        # 변환된 파일명 생성 (한글 안전 처리)
        original_name = file.filename.rsplit('.', 1)[0] if file.filename else "converted"
        # 한글 파일명을 안전하게 처리
        safe_original_name = original_name.encode('utf-8', errors='ignore').decode('utf-8')
        converted_filename = f"{safe_original_name}_변환완료.xlsx"
        
        logger.info(f"✅ 변환 완료: {len(output_data)} bytes")
        logger.info(f"📁 변환된 파일명: {converted_filename}")
        
        # Excel 파일로 응답 (한글 파일명 URL 인코딩)
        from urllib.parse import quote
        encoded_filename = quote(converted_filename.encode('utf-8'))
        
        return Response(
            content=output_data,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}",
                "X-Conversion-Method": conversion_method,
                "X-Original-Rows": str(rows),
                "X-Original-Cols": str(cols)
            }
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"❌ 예상치 못한 오류: {str(e)}")
        raise HTTPException(status_code=500, detail=f"서버 오류: {str(e)}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
