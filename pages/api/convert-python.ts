import { NextApiRequest, NextApiResponse } from 'next';
import { IncomingForm, File } from 'formidable';
import { promises as fs } from 'fs';

// API 설정
export const config = {
  api: {
    bodyParser: false, // formidable을 사용하기 위해 비활성화
    responseLimit: false,
  },
};

// Python 서비스 URL (환경변수로 관리)
const PYTHON_SERVICE_URL = process.env.PYTHON_SERVICE_URL || 'http://localhost:8000';

// 에러 응답 타입
interface ErrorResponse {
  success: false;
  message: string;
  code?: string;
}

// 성공 응답 타입 (메타데이터용)
interface SuccessResponse {
  success: true;
  filename: string;
  originalSize: number;
  convertedSize: number;
  conversionMethod: string;
  warnings?: string[];
}

/**
 * 파일 업로드 파싱
 */
function parseFormData(req: NextApiRequest): Promise<{ fields: any; files: any }> {
  return new Promise((resolve, reject) => {
    const form = new IncomingForm({
      maxFileSize: 50 * 1024 * 1024, // 50MB
      keepExtensions: true,
      multiples: false,
    });

    form.parse(req, (err, fields, files) => {
      if (err) {
        reject(err);
      } else {
        resolve({ fields, files });
      }
    });
  });
}

/**
 * 파일을 Buffer로 읽기
 */
async function readFileToBuffer(file: File): Promise<Buffer> {
  const data = await fs.readFile(file.filepath);
  return Buffer.from(data);
}

/**
 * 에러 응답 전송
 */
function sendError(res: NextApiResponse, message: string, statusCode: number = 400, code?: string) {
  const errorResponse: ErrorResponse = {
    success: false,
    message,
    code,
  };
  
  res.status(statusCode).json(errorResponse);
}

/**
 * Python 서비스로 파일 변환 요청
 */
async function convertWithPythonService(fileBuffer: Buffer, filename: string): Promise<{
  buffer: Buffer;
  conversionMethod: string;
  originalRows: string;
  originalCols: string;
}> {
  console.log('🐍 Python 서비스 호출:', PYTHON_SERVICE_URL);

  // FormData 생성
  const formData = new FormData();
  const blob = new Blob([fileBuffer], { type: 'application/octet-stream' });
  formData.append('file', blob, filename);

  console.log('📤 Python 서비스로 파일 전송 중...');
  
  // Python FastAPI 서비스 호출
  const response = await fetch(`${PYTHON_SERVICE_URL}/convert`, {
    method: 'POST',
    body: formData,
    headers: {
      // FormData를 사용할 때는 Content-Type을 설정하지 않음 (브라우저가 자동 설정)
    },
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error('❌ Python 서비스 오류:', response.status, errorText);
    throw new Error(`Python 서비스 오류: ${response.status} - ${errorText}`);
  }

  // 변환된 파일 받기
  const convertedBuffer = Buffer.from(await response.arrayBuffer());
  
  // 응답 헤더에서 정보 추출
  const conversionMethod = response.headers.get('X-Conversion-Method') || 'unknown';
  const originalRows = response.headers.get('X-Original-Rows') || '0';
  const originalCols = response.headers.get('X-Original-Cols') || '0';
  
  console.log('✅ Python 서비스 변환 완료!');
  console.log('📊 변환 방법:', conversionMethod);
  console.log('📋 데이터 크기:', `${originalRows}행 × ${originalCols}열`);

  return {
    buffer: convertedBuffer,
    conversionMethod,
    originalRows,
    originalCols,
  };
}

/**
 * 기존 TypeScript 로직으로 fallback
 */
async function fallbackToTypeScript(fileBuffer: Buffer, filename: string): Promise<Buffer> {
  console.log('🔄 TypeScript 변환 로직으로 fallback');
  
  try {
    const { convertToXlsx } = await import('../../lib/converter');
    const convertedBuffer = await convertToXlsx(fileBuffer, filename);
    return convertedBuffer;
  } catch (error) {
    console.error('❌ TypeScript fallback도 실패:', error);
    throw new Error('모든 변환 방법이 실패했습니다.');
  }
}

/**
 * 메인 API 핸들러
 */
export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  console.log('🌐 Python API 호출됨:', req.method, req.url);
  
  // CORS 헤더 설정
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  // OPTIONS 요청 처리 (CORS preflight)
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  // POST 요청만 허용
  if (req.method !== 'POST') {
    return sendError(res, '지원하지 않는 HTTP 메서드입니다.', 405, 'METHOD_NOT_ALLOWED');
  }

  try {
    // 1. 폼 데이터 파싱
    const { fields, files } = await parseFormData(req);
    
    // 2. 파일 검증
    const uploadedFile = Array.isArray(files.file) ? files.file[0] : files.file;
    
    if (!uploadedFile) {
      return sendError(res, '파일이 업로드되지 않았습니다.', 400, 'NO_FILE');
    }

    if (!uploadedFile.originalFilename) {
      return sendError(res, '파일명이 없습니다.', 400, 'NO_FILENAME');
    }

    // 3. 옵션 파싱
    const metaOnly = fields.metaOnly === 'true'; // 메타데이터만 반환할지 여부

    console.log('🚀 변환 시작:', uploadedFile.originalFilename, '(' + uploadedFile.size + ' bytes)');

    // 4. 파일 읽기
    const fileBuffer = await readFileToBuffer(uploadedFile);

    let convertedBuffer: Buffer;
    let conversionMethod = 'python';
    let originalRows = '0';
    let originalCols = '0';

    try {
      // 5. Python 서비스로 변환 시도
      const result = await convertWithPythonService(fileBuffer, uploadedFile.originalFilename);
      convertedBuffer = result.buffer;
      conversionMethod = result.conversionMethod;
      originalRows = result.originalRows;
      originalCols = result.originalCols;
      
    } catch (pythonError) {
      console.error('❌ Python 서비스 실패:', pythonError);
      
      // 6. TypeScript 로직으로 fallback
      try {
        convertedBuffer = await fallbackToTypeScript(fileBuffer, uploadedFile.originalFilename);
        conversionMethod = 'typescript-fallback';
        console.log('✅ TypeScript fallback 성공');
      } catch (fallbackError) {
        console.error('❌ 모든 변환 방법 실패:', fallbackError);
        return sendError(res, '파일 변환에 실패했습니다.', 500, 'CONVERSION_FAILED');
      }
    }

    // 7. 변환된 파일명 생성
    const originalName = uploadedFile.originalFilename.replace(/\.[^.]+$/, '') || 'converted';
    const convertedFilename = `${originalName.replace(/[^\w가-힣\-_]/g, '_')}_변환완료.xlsx`;
    
    console.log('📁 파일 크기 비교: 원본', `${uploadedFile.size}bytes`, '→ 변환', `${convertedBuffer.length}bytes`, `(비율: ${(convertedBuffer.length / uploadedFile.size).toFixed(2)})`);
    console.log('변환 완료:', convertedFilename, `(${convertedBuffer.length} bytes)`);

    // 8. 응답 처리
    if (metaOnly) {
      // 메타데이터만 반환
      const successResponse: SuccessResponse = {
        success: true,
        filename: convertedFilename,
        originalSize: uploadedFile.size,
        convertedSize: convertedBuffer.length,
        conversionMethod,
      };
      
      res.status(200).json(successResponse);
    } else {
      // 파일 다운로드 응답
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(convertedFilename)}`);
      res.setHeader('Content-Length', convertedBuffer.length);
      
      // 추가 메타데이터 헤더
      res.setHeader('X-Original-Size', uploadedFile.size.toString());
      res.setHeader('X-Converted-Size', convertedBuffer.length.toString());
      res.setHeader('X-Conversion-Method', conversionMethod);
      res.setHeader('X-Original-Rows', originalRows);
      res.setHeader('X-Original-Cols', originalCols);

      res.status(200).send(convertedBuffer);
    }

  } catch (error) {
    console.error('API 에러:', error);
    
    // 파일 크기 초과 에러
    if (error instanceof Error && error.message && error.message.includes('maxFileSize')) {
      return sendError(res, '파일이 너무 큽니다. 최대 50MB까지 지원합니다.', 413, 'FILE_TOO_LARGE');
    }
    
    // 일반 에러
    return sendError(res, '서버 오류가 발생했습니다.', 500, 'INTERNAL_ERROR');
  }
}
