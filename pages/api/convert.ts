import { NextApiRequest, NextApiResponse } from 'next';
import { IncomingForm, File } from 'formidable';
import { promises as fs } from 'fs';
import { processFile } from '@/lib/converter';

// API 설정
export const config = {
  api: {
    bodyParser: false, // formidable을 사용하기 위해 비활성화
    responseLimit: false,
  },
};

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
  warnings?: string[];
}

/**
 * 파일 업로드 파싱
 */
function parseFormData(req: NextApiRequest): Promise<{ fields: any; files: any }> {
  return new Promise((resolve, reject) => {
    const form = new IncomingForm({
      maxFileSize: 50 * 1024 * 1024, // 50MB (긴 헤더 필드 고려)
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
 * 메인 API 핸들러
 */
export default async function handler(req: NextApiRequest, res: NextApiResponse) {
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
    const forceTextRecovery = fields.forceTextRecovery === 'true';
    const metaOnly = fields.metaOnly === 'true'; // 메타데이터만 반환할지 여부

    console.log(`변환 시작: ${uploadedFile.originalFilename} (${uploadedFile.size} bytes)`);

    // 4. 파일 읽기
    const buffer = await readFileToBuffer(uploadedFile);

    // 5. 변환 처리
    const result = await processFile(buffer, uploadedFile.originalFilename, forceTextRecovery);

    // 6. 변환 실패 처리
    if (!result.success) {
      console.error(`변환 실패: ${result.message}`);
      return sendError(res, result.message || '변환에 실패했습니다.', 500, 'CONVERSION_FAILED');
    }

    console.log(`변환 완료: ${result.filename} (${result.convertedSize} bytes)`);

    // 7. 응답 처리
    if (metaOnly) {
      // 메타데이터만 반환
      const successResponse: SuccessResponse = {
        success: true,
        filename: result.filename,
        originalSize: result.originalSize,
        convertedSize: result.convertedSize || 0,
        warnings: result.warnings,
      };
      
      res.status(200).json(successResponse);
    } else {
      // 파일 다운로드 응답
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(result.filename)}`);
      res.setHeader('Content-Length', result.buffer!.length);
      
      // 추가 메타데이터 헤더
      res.setHeader('X-Original-Size', result.originalSize.toString());
      res.setHeader('X-Converted-Size', result.convertedSize!.toString());
      
      if (result.warnings && result.warnings.length > 0) {
        res.setHeader('X-Warnings', encodeURIComponent(result.warnings.join('; ')));
      }

      res.status(200).send(result.buffer);
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
