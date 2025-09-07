import * as XLSX from 'xlsx';
import * as iconv from 'iconv-lite';
import { detect } from 'jschardet';
import { parseISO, isValid, format } from 'date-fns';

// 지원하는 파일 확장자
export const SUPPORTED_EXTENSIONS = ['.xls', '.xlsx', '.csv', '.tsv', '.txt'];

// 최대 파일 크기 (50MB로 증가 - 긴 헤더 필드 고려)
export const MAX_FILE_SIZE = 50 * 1024 * 1024;

/**
 * 파일 확장자 검증
 */
export function validateFileExtension(filename: string): boolean {
  const ext = filename.toLowerCase().substring(filename.lastIndexOf('.'));
  return SUPPORTED_EXTENSIONS.includes(ext);
}

/**
 * 파일 크기 검증
 */
export function validateFileSize(size: number): boolean {
  return size <= MAX_FILE_SIZE;
}

/**
 * 안전한 파일명 생성
 */
export function sanitizeFilename(filename: string): string {
  // 한글과 영문, 숫자, 일부 특수문자만 허용
  const sanitized = filename
    .replace(/[^\w가-힣.\-_]/g, '_')
    .replace(/_{2,}/g, '_')
    .trim();
  
  return sanitized || 'converted_file';
}

/**
 * 인코딩 감지 및 텍스트 디코딩
 */
function detectAndDecode(buffer: Buffer): string {
  try {
    // 1. UTF-8 시도
    const utf8Text = buffer.toString('utf8');
    if (!utf8Text.includes('\uFFFD')) {
      return utf8Text;
    }
  } catch (e) {
    // UTF-8 실패
  }

  try {
    // 2. 자동 감지
    const detected = detect(buffer);
    if (detected && detected.encoding && detected.confidence > 0.7) {
      const encoding = detected.encoding.toLowerCase();
      
      // 한국어 인코딩 우선 처리
      if (encoding.includes('euc-kr') || encoding.includes('cp949')) {
        return iconv.decode(buffer, 'euc-kr');
      }
      
      if (iconv.encodingExists(encoding)) {
        return iconv.decode(buffer, encoding);
      }
    }
  } catch (e) {
    // 자동 감지 실패
  }

  // 3. 한국어 인코딩들 순차 시도
  const encodings = ['euc-kr', 'cp949', 'utf8', 'latin1'];
  
  for (const encoding of encodings) {
    try {
      const decoded = iconv.decode(buffer, encoding);
      // 한글이 포함되어 있고 깨지지 않았다면 성공
      if (decoded && !decoded.includes('\uFFFD')) {
        return decoded;
      }
    } catch (e) {
      continue;
    }
  }

  // 4. 최후의 수단: latin1으로 디코딩
  return iconv.decode(buffer, 'latin1');
}

/**
 * CSV 구분자 추정 (개선된 버전)
 */
function detectDelimiter(text: string): string {
  const lines = text.split('\n').slice(0, 10).filter(line => line.trim()); // 더 많은 줄 검사
  const delimiters = ['\t', ',', ';', '|'];
  
  let bestDelimiter = ',';
  let maxScore = 0;
  
  for (const delimiter of delimiters) {
    let score = 0;
    let columnCounts: number[] = [];
    
    for (const line of lines) {
      if (line.trim()) {
        // 따옴표 안의 구분자는 무시하고 파싱
        const columns = parseCSVLine(line, delimiter);
        const columnCount = columns.length;
        
        if (columnCount > 1) {
          columnCounts.push(columnCount);
        }
      }
    }
    
    if (columnCounts.length > 0) {
      // 일관성 있는 컬럼 수를 가진 구분자에 높은 점수
      const avgColumns = columnCounts.reduce((a, b) => a + b, 0) / columnCounts.length;
      const consistency = 1 - (Math.max(...columnCounts) - Math.min(...columnCounts)) / avgColumns;
      score = avgColumns * consistency * columnCounts.length;
      
      if (score > maxScore) {
        maxScore = score;
        bestDelimiter = delimiter;
      }
    }
  }
  
  return bestDelimiter;
}

/**
 * CSV 라인 파싱 (따옴표 고려)
 */
function parseCSVLine(line: string, delimiter: string): string[] {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;
  let quoteChar = '';
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    
    if (!inQuotes) {
      if (char === '"' || char === "'") {
        inQuotes = true;
        quoteChar = char;
      } else if (char === delimiter) {
        result.push(current.trim());
        current = '';
      } else {
        current += char;
      }
    } else {
      if (char === quoteChar) {
        // 다음 문자가 같은 따옴표면 이스케이프된 따옴표
        if (i + 1 < line.length && line[i + 1] === quoteChar) {
          current += char;
          i++; // 다음 문자 건너뛰기
        } else {
          inQuotes = false;
          quoteChar = '';
        }
      } else {
        current += char;
      }
    }
  }
  
  result.push(current.trim());
  return result;
}

/**
 * 셀 값 정규화 (타입 변환)
 */
function normalizeCell(value: string): any {
  if (!value || value.trim() === '') {
    return null;
  }
  
  const trimmed = value.trim();
  
  // 1. 숫자 처리
  const numberMatch = trimmed.match(/^-?[\d,]+\.?\d*$/);
  if (numberMatch) {
    const cleaned = trimmed.replace(/,/g, '');
    const num = parseFloat(cleaned);
    if (!isNaN(num)) {
      return num;
    }
  }
  
  // 2. 퍼센트 처리
  const percentMatch = trimmed.match(/^(-?[\d,]+\.?\d*)\s*%$/);
  if (percentMatch) {
    const cleaned = percentMatch[1].replace(/,/g, '');
    const num = parseFloat(cleaned);
    if (!isNaN(num)) {
      return num / 100; // Excel 퍼센트 형식
    }
  }
  
  // 3. 날짜 처리
  const datePatterns = [
    /^\d{4}-\d{2}-\d{2}$/,
    /^\d{4}\/\d{1,2}\/\d{1,2}$/,
    /^\d{1,2}\/\d{1,2}\/\d{4}$/,
    /^\d{4}\.\d{1,2}\.\d{1,2}$/,
  ];
  
  for (const pattern of datePatterns) {
    if (pattern.test(trimmed)) {
      try {
        const date = parseISO(trimmed.replace(/\//g, '-').replace(/\./g, '-'));
        if (isValid(date)) {
          return date;
        }
      } catch (e) {
        // 날짜 파싱 실패, 계속 진행
      }
    }
  }
  
  // 4. 불린 값 처리
  const lowerValue = trimmed.toLowerCase();
  if (lowerValue === 'true' || lowerValue === '참' || lowerValue === 'yes') {
    return true;
  }
  if (lowerValue === 'false' || lowerValue === '거짓' || lowerValue === 'no') {
    return false;
  }
  
  // 5. 기본값: 문자열 그대로 반환
  return trimmed;
}

/**
 * 텍스트 기반 복구 (CSV/TSV 파싱) - 개선된 버전
 */
function textBasedRecovery(buffer: Buffer): XLSX.WorkBook {
  const text = detectAndDecode(buffer);
  const delimiter = detectDelimiter(text);
  
  console.log(`텍스트 복구 시작: 구분자="${delimiter}"`);
  
  const lines = text.split('\n').filter(line => line.trim());
  const data: any[][] = [];
  
  // 각 라인을 올바르게 파싱
  for (const line of lines) {
    try {
      const cells = parseCSVLine(line, delimiter).map(cell => normalizeCell(cell));
      
      // 빈 행이 아닌 경우에만 추가
      if (cells.some(cell => cell !== null && cell !== '')) {
        data.push(cells);
      }
    } catch (error) {
      console.warn('라인 파싱 실패:', line.substring(0, 100) + '...');
      // 파싱 실패한 라인은 단순 분할로 처리
      const cells = line.split(delimiter).map(cell => normalizeCell(cell.trim()));
      if (cells.some(cell => cell !== null && cell !== '')) {
        data.push(cells);
      }
    }
  }
  
  console.log(`파싱 완료: ${data.length}행, 평균 ${data.length > 0 ? Math.round(data.reduce((sum, row) => sum + row.length, 0) / data.length) : 0}열`);
  
  // 데이터가 없으면 에러
  if (data.length === 0) {
    throw new Error('파싱된 데이터가 없습니다. 파일 형식을 확인해주세요.');
  }
  
  // 빈 워크북 생성
  const workbook = XLSX.utils.book_new();
  
  // 데이터를 워크시트로 변환
  const worksheet = XLSX.utils.aoa_to_sheet(data);
  
  // 워크시트를 워크북에 추가
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  
  return workbook;
}

/**
 * 워크북 데이터 정규화
 */
function normalizeWorkbook(workbook: XLSX.WorkBook): XLSX.WorkBook {
  const normalizedWorkbook = XLSX.utils.book_new();
  
  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
      header: 1, 
      defval: null,
      raw: false 
    }) as any[][];
    
    // 각 셀 정규화
    const normalizedData = jsonData.map(row => 
      row.map(cell => typeof cell === 'string' ? normalizeCell(cell) : cell)
    );
    
    // 정규화된 데이터로 새 워크시트 생성
    const normalizedSheet = XLSX.utils.aoa_to_sheet(normalizedData);
    
    // 안전한 시트명 생성
    const safeSheetName = sanitizeFilename(sheetName).substring(0, 31);
    XLSX.utils.book_append_sheet(normalizedWorkbook, normalizedSheet, safeSheetName);
  });
  
  return normalizedWorkbook;
}

/**
 * 메인 변환 함수
 */
export async function convertToXlsx(
  buffer: Buffer, 
  filename: string,
  forceTextRecovery: boolean = false
): Promise<Buffer> {
  try {
    let workbook: XLSX.WorkBook;
    
    if (!forceTextRecovery) {
      try {
        // 1단계: 표준 파서 시도
        workbook = XLSX.read(buffer, {
          type: 'buffer',
          cellDates: true,
          cellNF: false,
          cellText: false,
          // 다양한 인코딩 시도
          codepage: 65001, // UTF-8
        });
        
        // 워크북이 비어있지 않은지 확인
        if (workbook.SheetNames.length === 0) {
          throw new Error('빈 워크북');
        }
        
      } catch (standardError) {
        console.log('표준 파서 실패, 텍스트 복구 시도:', standardError instanceof Error ? standardError.message : String(standardError));
        // 2단계: 텍스트 기반 복구
        workbook = textBasedRecovery(buffer);
      }
    } else {
      // 강제 텍스트 복구
      workbook = textBasedRecovery(buffer);
    }
    
    // 3단계: 데이터 정규화
    const normalizedWorkbook = normalizeWorkbook(workbook);
    
    // 4단계: .xlsx로 변환
    const xlsxBuffer = XLSX.write(normalizedWorkbook, {
      type: 'buffer',
      bookType: 'xlsx',
      compression: true,
      cellDates: true,
    });
    
    return Buffer.from(xlsxBuffer);
    
  } catch (error) {
    console.error('변환 실패:', error);
    throw new Error(`파일 변환에 실패했습니다: ${error instanceof Error ? error.message : String(error)}`);
  }
}

/**
 * 변환 결과 정보
 */
export interface ConversionResult {
  success: boolean;
  buffer?: Buffer;
  filename: string;
  originalSize: number;
  convertedSize?: number;
  message?: string;
  warnings?: string[];
}

/**
 * 파일 변환 (전체 프로세스)
 */
export async function processFile(
  buffer: Buffer,
  originalFilename: string,
  forceTextRecovery: boolean = false
): Promise<ConversionResult> {
  const warnings: string[] = [];
  
  try {
    // 파일 검증
    if (!validateFileExtension(originalFilename)) {
      throw new Error(`지원하지 않는 파일 형식입니다. 지원 형식: ${SUPPORTED_EXTENSIONS.join(', ')}`);
    }
    
    if (!validateFileSize(buffer.length)) {
      throw new Error(`파일이 너무 큽니다. 최대 크기: ${MAX_FILE_SIZE / 1024 / 1024}MB`);
    }
    
    // 변환 실행
    const convertedBuffer = await convertToXlsx(buffer, originalFilename, forceTextRecovery);
    
    // 결과 파일명 생성
    const baseName = originalFilename.replace(/\.[^.]+$/, '');
    const safeBaseName = sanitizeFilename(baseName);
    const resultFilename = `${safeBaseName}_변환완료.xlsx`;
    
    // 크기 비교 경고
    if (convertedBuffer.length > buffer.length * 2) {
      warnings.push('변환된 파일이 원본보다 상당히 큽니다. 데이터 확인을 권장합니다.');
    }
    
    return {
      success: true,
      buffer: convertedBuffer,
      filename: resultFilename,
      originalSize: buffer.length,
      convertedSize: convertedBuffer.length,
      warnings: warnings.length > 0 ? warnings : undefined,
    };
    
  } catch (error) {
    return {
      success: false,
      filename: originalFilename,
      originalSize: buffer.length,
      message: error instanceof Error ? error.message : String(error),
    };
  }
}
