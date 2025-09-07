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
 * CSV 구분자 추정 (개선된 버전 - 복잡한 헤더 고려)
 */
function detectDelimiter(text: string): string {
  const lines = text.split('\n').slice(0, 10).filter(line => line.trim()); // 더 많은 줄 검사
  const delimiters = ['\t', ',', ';', '|'];
  
  console.log('구분자 감지 시작, 첫 번째 라인:', lines[0]?.substring(0, 200));
  
  let bestDelimiter = ',';
  let maxScore = 0;
  
  for (const delimiter of delimiters) {
    let score = 0;
    let columnCounts: number[] = [];
    
    for (const line of lines) {
      if (line.trim()) {
        // 간단한 분할로 먼저 테스트 (성능상 이유)
        const simpleColumns = line.split(delimiter);
        const columnCount = simpleColumns.length;
        
        // 괄호가 포함된 복잡한 헤더의 경우 더 관대하게 처리
        if (columnCount > 1) {
          columnCounts.push(columnCount);
        }
      }
    }
    
    if (columnCounts.length > 0) {
      // 일관성 있는 컬럼 수를 가진 구분자에 높은 점수
      const avgColumns = columnCounts.reduce((a, b) => a + b, 0) / columnCounts.length;
      const maxCols = Math.max(...columnCounts);
      const minCols = Math.min(...columnCounts);
      
      // 복잡한 헤더의 경우 일관성 요구사항을 완화
      const consistency = maxCols > 10 ? 0.8 : (1 - (maxCols - minCols) / Math.max(avgColumns, 1));
      score = avgColumns * Math.max(consistency, 0.5) * columnCounts.length;
      
      console.log(`구분자 "${delimiter}": 평균 ${avgColumns.toFixed(1)}열, 일관성 ${consistency.toFixed(2)}, 점수 ${score.toFixed(1)}`);
      
      if (score > maxScore) {
        maxScore = score;
        bestDelimiter = delimiter;
      }
    }
  }
  
  console.log(`선택된 구분자: "${bestDelimiter}"`);
  return bestDelimiter;
}

/**
 * CSV 라인 파싱 (따옴표 및 특수문자 고려)
 */
function parseCSVLine(line: string, delimiter: string): string[] {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;
  let quoteChar = '';
  
  // 라인 전처리: 불필요한 공백 제거 및 정규화
  line = line.trim();
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    
    if (!inQuotes) {
      if (char === '"' || char === "'") {
        inQuotes = true;
        quoteChar = char;
      } else if (char === delimiter) {
        // 현재 셀 내용 정리 및 추가
        const cellContent = current.trim();
        result.push(cellContent);
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
  
  // 마지막 셀 추가
  const lastCell = current.trim();
  result.push(lastCell);
  
  // 빈 셀들을 null로 변환하지 않고 빈 문자열로 유지
  return result.map(cell => cell || '');
}

/**
 * 셀 값 정규화 (타입 변환) - 특수문자 처리 강화
 */
function normalizeCell(value: string): any {
  if (!value || value.trim() === '') {
    return '';  // null 대신 빈 문자열 반환
  }
  
  const trimmed = value.trim();
  
  // 특수문자가 많은 헤더 필드는 문자열로 유지
  if (trimmed.includes('(') && trimmed.includes(')')) {
    return trimmed;  // 괄호가 포함된 복잡한 텍스트는 그대로 유지
  }
  
  // 1. 숫자 처리 (더 엄격한 검증)
  const numberMatch = trimmed.match(/^-?[\d,]+\.?\d*$/);
  if (numberMatch && !trimmed.includes('(')) {  // 괄호가 없는 경우만
    const cleaned = trimmed.replace(/,/g, '');
    const num = parseFloat(cleaned);
    if (!isNaN(num) && isFinite(num)) {
      return num;
    }
  }
  
  // 2. 퍼센트 처리
  const percentMatch = trimmed.match(/^(-?[\d,]+\.?\d*)\s*%$/);
  if (percentMatch) {
    const cleaned = percentMatch[1].replace(/,/g, '');
    const num = parseFloat(cleaned);
    if (!isNaN(num) && isFinite(num)) {
      return num / 100; // Excel 퍼센트 형식
    }
  }
  
  // 3. 날짜 처리 (간단한 패턴만)
  const simpleDatePattern = /^\d{4}-\d{2}-\d{2}$/;
  if (simpleDatePattern.test(trimmed)) {
    try {
      const date = parseISO(trimmed);
      if (isValid(date)) {
        return date;
      }
    } catch (e) {
      // 날짜 파싱 실패, 문자열로 처리
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
  console.log(`전체 텍스트 길이: ${text.length}, 첫 200자:`, text.substring(0, 200));
  
  const lines = text.split('\n').filter(line => line.trim());
  console.log(`유효한 라인 수: ${lines.length}`);
  const data: any[][] = [];
  
  // 각 라인을 올바르게 파싱
  let maxColumns = 0;
  
  for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
    const line = lines[lineIndex];
    
    try {
      const cells = parseCSVLine(line, delimiter).map(cell => normalizeCell(cell));
      
      // 컬럼 수 추적
      maxColumns = Math.max(maxColumns, cells.length);
      
      // 빈 행이 아닌 경우에만 추가
      if (cells.some(cell => cell !== null && cell !== '')) {
        data.push(cells);
      }
    } catch (error) {
      console.warn(`라인 ${lineIndex + 1} 파싱 실패:`, line.substring(0, 100) + '...');
      
      // 파싱 실패한 라인은 여러 방법으로 시도
      let fallbackCells: any[] = [];
      
      // 방법 1: 단순 분할
      try {
        fallbackCells = line.split(delimiter).map(cell => normalizeCell(cell.trim()));
      } catch (e1) {
        console.warn('방법 1 실패, 방법 2 시도');
        
        // 방법 2: 쉼표로 분할 (delimiter가 다른 경우)
        try {
          fallbackCells = line.split(',').map(cell => normalizeCell(cell.trim()));
        } catch (e2) {
          console.warn('방법 2 실패, 방법 3 시도');
          
          // 방법 3: 탭으로 분할
          try {
            fallbackCells = line.split('\t').map(cell => normalizeCell(cell.trim()));
          } catch (e3) {
            console.warn('방법 3 실패, 방법 4 시도');
            
            // 방법 4: 공백으로 분할 (여러 공백은 하나로 처리)
            try {
              fallbackCells = line.split(/\s+/).filter(cell => cell.trim()).map(cell => normalizeCell(cell.trim()));
            } catch (e4) {
              console.warn('방법 4 실패, 전체를 하나의 셀로 처리');
              // 방법 5: 전체를 하나의 셀로 처리
              fallbackCells = [normalizeCell(line.trim())];
            }
          }
        }
      }
      
      if (fallbackCells.some(cell => cell !== null && cell !== '')) {
        data.push(fallbackCells);
        maxColumns = Math.max(maxColumns, fallbackCells.length);
      }
    }
  }
  
  // 모든 행의 컬럼 수를 맞춤 (빈 셀로 패딩)
  data.forEach(row => {
    while (row.length < maxColumns) {
      row.push('');
    }
  });
  
  // 헤더 정리 (빈 헤더만 처리, 중복은 그대로 유지)
  if (data.length > 0) {
    const headerRow = data[0];
    
    for (let i = 0; i < headerRow.length; i++) {
      let header = String(headerRow[i]).trim();
      
      // 빈 헤더만 처리 (중복은 그대로 유지)
      if (!header) {
        headerRow[i] = `컬럼${i + 1}`;
      } else {
        headerRow[i] = header;
      }
    }
    
    console.log('원본 헤더 유지:', headerRow.slice(0, 10)); // 처음 10개만 로그
  }
  
  console.log(`파싱 완료: ${data.length}행, ${maxColumns}열`);
  
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
