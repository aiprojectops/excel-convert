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
 * 헤더 필드명을 Excel 호환성을 위해 정규화
 */
function normalizeHeaderField(value: string): string {
  if (!value || value.trim() === '') {
    return '';
  }
  
  const trimmed = value.trim();
  
  // 1. 특수 문자 제거 (Excel 연결 시 문제가 될 수 있는 문자들)
  let normalized = trimmed
    .replace(/[\x00-\x1F\x7F-\x9F]/g, '') // 제어 문자 제거
    .replace(/[^\w\s가-힣()[\]{}.,\-_]/g, '') // 안전한 문자만 유지
    .trim();
  
  // 2. 연속된 공백을 하나로 통합
  normalized = normalized.replace(/\s+/g, ' ');
  
  // 3. 빈 문자열이 되면 기본값 설정
  if (!normalized) {
    return '컬럼';
  }
  
  return normalized;
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
  
  console.log('🔄 텍스트 복구 시작: 구분자="' + delimiter + '"');
  console.log('📄 전체 텍스트 길이:', text.length, '첫 200자:', text.substring(0, 200));
  console.log('📄 마지막 200자:', text.substring(Math.max(0, text.length - 200)));
  
  const lines = text.split('\n').filter(line => line.trim());
  console.log(`유효한 라인 수: ${lines.length}`);
  
  // 처음 3줄 로그로 확인
  lines.slice(0, 3).forEach((line, index) => {
    console.log(`라인 ${index + 1} (${line.length}자):`, line.substring(0, 100) + (line.length > 100 ? '...' : ''));
  });
  
  const data: any[][] = [];
  
  // 각 라인을 올바르게 파싱
  let maxColumns = 0;
  
  for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
    const line = lines[lineIndex];
    
    try {
      const cells = parseCSVLine(line, delimiter);
      
      // 헤더 행(첫 번째 행)은 정규화하지 않고 원본 유지
      const processedCells = lineIndex === 0 
        ? cells.map(cell => normalizeHeaderField(cell) || `컬럼${cells.indexOf(cell) + 1}`)  // 헤더는 정규화된 문자열로 처리
        : cells.map(cell => normalizeCell(cell));  // 데이터 행만 정규화
      
      // 컬럼 수 추적
      maxColumns = Math.max(maxColumns, processedCells.length);
      
      // 첫 번째 행(헤더)은 무조건 추가, 나머지는 빈 행이 아닌 경우에만 추가
      if (lineIndex === 0 || processedCells.some(cell => cell !== null && cell !== '')) {
        data.push(processedCells);
        console.log(`라인 ${lineIndex + 1} 파싱 성공: ${processedCells.length}개 셀`, 
          lineIndex === 0 ? `(헤더: ${processedCells.slice(0, 3).join(', ')}...)` : '');
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
      
      // Fallback 셀 처리 (헤더 고려)
      const processedFallbackCells = lineIndex === 0 
        ? fallbackCells.map(cell => String(cell).trim() || `컬럼${fallbackCells.indexOf(cell) + 1}`)  // 헤더는 문자열로 유지
        : fallbackCells.map(cell => normalizeCell(String(cell)));  // 데이터 행만 정규화
      
      // 첫 번째 행(헤더)은 무조건 추가, 나머지는 빈 행이 아닌 경우에만 추가
      if (lineIndex === 0 || processedFallbackCells.some(cell => cell !== null && cell !== '')) {
        data.push(processedFallbackCells);
        maxColumns = Math.max(maxColumns, processedFallbackCells.length);
        console.log(`라인 ${lineIndex + 1} fallback 파싱 성공: ${processedFallbackCells.length}개 셀`);
      }
    }
  }
  
  // 모든 행의 컬럼 수를 맞춤 (빈 셀로 패딩)
  data.forEach(row => {
    while (row.length < maxColumns) {
      row.push('');
    }
  });
  
  // 헤더 확인 로그
  if (data.length > 0) {
    console.log('최종 헤더 확인:', data[0].slice(0, 5)); // 처음 5개만 로그
    console.log('헤더 개수:', data[0].length);
    console.log('전체 데이터 행 수:', data.length);
    
    // 데이터 샘플 확인
    if (data.length > 1) {
      console.log('두 번째 행 샘플:', data[1].slice(0, 5));
    }
  } else {
    console.error('❌ 파싱된 데이터가 없습니다!');
  }
  
  console.log(`파싱 완료: ${data.length}행, ${maxColumns}열`);
  
  // 데이터가 없으면 에러
  if (data.length === 0) {
    throw new Error('파싱된 데이터가 없습니다. 파일 형식을 확인해주세요.');
  }
  
  // 빈 워크북 생성
  const workbook = XLSX.utils.book_new();
  
  // 데이터를 워크시트로 변환
  console.log('워크시트 생성 중... 데이터 크기:', data.length, 'x', data[0]?.length || 0);
  const worksheet = XLSX.utils.aoa_to_sheet(data);
  
  // 워크시트 범위 확인
  console.log('워크시트 범위:', worksheet['!ref']);
  
  // 워크시트를 워크북에 추가
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  console.log('워크북에 시트 추가 완료');
  
  return workbook;
}

/**
 * 워크북 데이터 정규화
 */
function normalizeWorkbook(workbook: XLSX.WorkBook): XLSX.WorkBook {
  console.log('🔧 normalizeWorkbook 시작, 시트 수:', workbook.SheetNames.length);
  const normalizedWorkbook = XLSX.utils.book_new();
  
  workbook.SheetNames.forEach((sheetName, index) => {
    console.log(`🔧 시트 ${index + 1} 처리 중: "${sheetName}"`);
    const worksheet = workbook.Sheets[sheetName];
    
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
      header: 1, 
      defval: null,
      raw: false 
    }) as any[][];
    
    console.log(`🔧 시트 "${sheetName}" 데이터: ${jsonData.length}행`);
    if (jsonData.length > 0) {
      console.log(`🔧 첫 번째 행 (헤더): ${jsonData[0].length}개 컬럼`);
      console.log(`🔧 헤더 내용 (처음 5개):`, jsonData[0].slice(0, 5));
    }
    
    // 각 셀 정규화
    const normalizedData = jsonData.map(row => 
      row.map(cell => typeof cell === 'string' ? normalizeCell(cell) : cell)
    );
    
    // 정규화된 데이터로 새 워크시트 생성
    const normalizedSheet = XLSX.utils.aoa_to_sheet(normalizedData);
    
    // 안전한 시트명 생성
    const safeSheetName = sanitizeFilename(sheetName).substring(0, 31);
    XLSX.utils.book_append_sheet(normalizedWorkbook, normalizedSheet, safeSheetName);
    console.log(`🔧 시트 "${safeSheetName}" 정규화 완료`);
  });
  
  return normalizedWorkbook;
}

/**
 * 바이너리 파일을 JSON 형태로 변환한 후 .xlsx로 저장
 */
async function convertBinaryToJsonToXlsx(buffer: Buffer, filename: string): Promise<Buffer> {
  console.log('🔧 바이너리 → JSON → .xlsx 변환 시작');
  
  try {
    // 1. 바이너리 데이터를 여러 인코딩으로 시도하여 텍스트 추출
    const encodings = ['utf8', 'latin1', 'ascii', 'base64'];
    const extractedData: any[] = [];
    
    for (const encoding of encodings) {
      try {
        console.log(`🔧 ${encoding} 인코딩으로 텍스트 추출 시도`);
        const textContent = buffer.toString(encoding as BufferEncoding);
        
        // Excel 데이터 패턴 찾기 (실제 셀 값들)
        const meaningfulLines = textContent
          .split(/[\r\n]+/)
          .map(line => line.trim())
          .filter(line => {
            // XML 태그나 바이너리 데이터가 아닌 실제 데이터 찾기
            return line.length > 0 && 
                   line.length < 500 && // 너무 긴 라인 제외
                   !/^[<>]/.test(line) && // XML 태그 제외
                   !/^PK|^[A-Za-z0-9+/]{20,}$/.test(line) && // ZIP 헤더나 Base64 제외
                   /[가-힣a-zA-Z0-9\s\-_()[\]{}.,]/.test(line) && // 의미있는 문자 포함
                   !/^[\x00-\x1F\x7F-\xFF]{5,}/.test(line); // 제어문자나 바이너리 데이터 제외
          })
          .slice(0, 100); // 더 많은 라인 확인
        
        if (meaningfulLines.length > 0) {
          console.log(`✅ ${encoding}에서 ${meaningfulLines.length}개 의미있는 라인 발견`);
          extractedData.push({
            encoding,
            lines: meaningfulLines,
            sample: meaningfulLines.slice(0, 3)
          });
        }
      } catch (e) {
        console.log(`❌ ${encoding} 인코딩 실패`);
      }
    }
    
    // 2. 추출된 데이터를 JSON 구조로 변환
    const jsonData = {
      filename,
      originalSize: buffer.length,
      extractedAt: new Date().toISOString(),
      encodings: extractedData
    };
    
    console.log('📋 JSON 구조 생성 완료:', {
      encodings: extractedData.length,
      totalLines: extractedData.reduce((sum, item) => sum + item.lines.length, 0)
    });
    
    // 3. 추출된 데이터를 실제 Excel 표 형태로 구조화
    console.log('🔧 Excel 표 구조 생성 시작');
    
    // 가장 많은 데이터가 있는 인코딩 선택
    const bestEncoding = extractedData.reduce((best, current) => 
      current.lines.length > best.lines.length ? current : best
    );
    
    console.log(`📊 최적 인코딩: ${bestEncoding.encoding} (${bestEncoding.lines.length}개 라인)`);
    
    // 실제 데이터를 표 형태로 변환
    const excelData: string[][] = [];
    
    // 헤더 생성 (첫 번째 라인이 헤더일 가능성이 높음)
    if (bestEncoding.lines.length > 0) {
      const firstLine = bestEncoding.lines[0];
      
      // 구분자로 분할 시도 (탭, 쉼표, 공백 등)
      const delimiters = ['\t', ',', ';', '|', ' '];
      let bestSplit = [firstLine];
      let maxColumns = 1;
      
      for (const delimiter of delimiters) {
        const split = firstLine.split(delimiter).filter((cell: string) => cell.trim());
        if (split.length > maxColumns && split.length <= 20) { // 너무 많은 컬럼은 제외
          bestSplit = split;
          maxColumns = split.length;
        }
      }
      
      // 헤더 추가 (정규화 적용)
      const headers = bestSplit.map(header => normalizeHeaderField(header.trim()));
      if (headers.length > 1) {
        excelData.push(headers);
        console.log(`📋 헤더 생성: ${headers.length}개 컬럼`, headers.slice(0, 5));
      } else {
        // 헤더를 찾을 수 없으면 기본 헤더 생성
        excelData.push(['데이터', '내용', '인코딩', '라인번호']);
      }
    }
    
    // 데이터 행 추가
    extractedData.forEach((item, encodingIndex) => {
      item.lines.forEach((line: string, lineIndex: number) => {
        // 헤더는 이미 처리했으므로 스킵
        if (encodingIndex === 0 && lineIndex === 0 && excelData.length > 0) return;
        
        // 라인을 구분자로 분할하여 여러 컬럼으로 만들기
        const delimiters = ['\t', ',', ';', '|'];
        let cells = [line];
        
        for (const delimiter of delimiters) {
          const split = line.split(delimiter);
          if (split.length > cells.length && split.length <= 20) {
            cells = split.map(cell => cell.trim());
          }
        }
        
        // 컬럼 수를 헤더와 맞춤
        const headerLength = excelData[0]?.length || 4;
        while (cells.length < headerLength) {
          cells.push('');
        }
        cells = cells.slice(0, headerLength);
        
        // 빈 라인이 아닌 경우에만 추가
        if (cells.some(cell => cell && cell.trim())) {
          excelData.push(cells);
        }
      });
    });
    
    console.log(`📊 Excel 표 생성 완료: ${excelData.length}행 x ${excelData[0]?.length || 0}열`);
    
    // 4. Excel 파일 생성
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(excelData);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'ExtractedData');
    
    const outputBuffer = XLSX.write(workbook, { 
      type: 'buffer', 
      bookType: 'xlsx',
      compression: true 
    });
    
    console.log(`✅ 바이너리 → JSON → .xlsx 변환 완료: ${outputBuffer.length} bytes`);
    return outputBuffer;
    
  } catch (error) {
    console.error('❌ JSON 변환 실패:', error instanceof Error ? error.message : String(error));
    
    // 최종 Fallback: 원본 파일 정보만 저장
    const fallbackData = [
      ['파일명', '크기', '상태', '비고'],
      [filename, `${buffer.length} bytes`, '변환 실패', '바이너리 파일 - 수동 확인 필요']
    ];
    
    const fallbackWorkbook = XLSX.utils.book_new();
    const fallbackWorksheet = XLSX.utils.aoa_to_sheet(fallbackData);
    XLSX.utils.book_append_sheet(fallbackWorkbook, fallbackWorksheet, 'FileInfo');
    
    return XLSX.write(fallbackWorkbook, { 
      type: 'buffer', 
      bookType: 'xlsx',
      compression: true 
    });
  }
}

/**
 * 압축된 데이터 해제 시도 (여러 방법)
 */
async function tryDecompressData(compressedData: string): Promise<string | null> {
  console.log('🔧 압축 해제 시도 중...');
  
  try {
    // 방법 1: 이미 압축 해제된 데이터인지 확인
    if (compressedData.includes('<') && compressedData.includes('>')) {
      console.log('✅ 이미 압축 해제된 XML 데이터 발견');
      return compressedData;
    }
    
    // 방법 2: Node.js zlib으로 압축 해제 시도
    const zlib = require('zlib');
    const compressedBuffer = Buffer.from(compressedData, 'latin1');
    
    // DEFLATE 압축 해제 시도
    try {
      const inflated = zlib.inflateRawSync(compressedBuffer);
      const result = inflated.toString('utf8');
      if (result.includes('<') && result.includes('>')) {
        console.log('✅ DEFLATE 압축 해제 성공');
        return result;
      }
    } catch (e) {
      console.log('❌ DEFLATE 압축 해제 실패');
    }
    
    // GZIP 압축 해제 시도
    try {
      const gunzipped = zlib.gunzipSync(compressedBuffer);
      const result = gunzipped.toString('utf8');
      if (result.includes('<') && result.includes('>')) {
        console.log('✅ GZIP 압축 해제 성공');
        return result;
      }
    } catch (e) {
      console.log('❌ GZIP 압축 해제 실패');
    }
    
    // 방법 3: 다른 인코딩으로 시도
    const encodings = ['utf8', 'ascii', 'base64'];
    for (const encoding of encodings) {
      try {
        const decoded = Buffer.from(compressedData, 'latin1').toString(encoding as BufferEncoding);
        if (decoded.includes('<') && decoded.includes('>') && decoded.includes('xml')) {
          console.log(`✅ ${encoding} 인코딩으로 XML 데이터 발견`);
          return decoded;
        }
      } catch (e) {
        // 무시
      }
    }
    
    console.log('❌ 모든 압축 해제 방법 실패');
    return null;
    
  } catch (error) {
    console.log('❌ 압축 해제 중 오류:', error instanceof Error ? error.message : String(error));
    return null;
  }
}

/**
 * 의미있는 텍스트들로 Excel 파일 생성
 */
async function createExcelFromMeaningfulTexts(texts: string[], filename: string): Promise<Buffer> {
  console.log('🔧 의미있는 텍스트로 Excel 생성 시작');
  console.log('📋 사용할 텍스트들:', texts);
  
  try {
    const workbook = XLSX.utils.book_new();
    const excelData: string[][] = [];
    
    // 헤더 생성 (파일명 기반)
    const cleanFilename = filename.replace(/\.[^.]+$/, '').replace(/[_\-\[\]().\s]+/g, ' ').trim();
    excelData.push([`발견된 텍스트 (${cleanFilename})`]);
    
    // 텍스트 분류 및 정리
    const koreanTexts = texts.filter(text => /[가-힣]/.test(text));
    const englishTexts = texts.filter(text => /^[a-zA-Z]+$/.test(text));
    const numberTexts = texts.filter(text => /^\d+$/.test(text));
    const mixedTexts = texts.filter(text => 
      !koreanTexts.includes(text) && 
      !englishTexts.includes(text) && 
      !numberTexts.includes(text)
    );
    
    // 분류별로 데이터 추가
    if (koreanTexts.length > 0) {
      excelData.push(['한글 텍스트', '분류', '길이', '설명']);
      koreanTexts.forEach(text => {
        excelData.push([text, '한글', text.length.toString(), '한글 텍스트']);
      });
      excelData.push(['']); // 빈 줄
    }
    
    if (englishTexts.length > 0) {
      excelData.push(['영문 텍스트', '분류', '길이', '설명']);
      englishTexts.forEach(text => {
        excelData.push([text, '영문', text.length.toString(), '영문 텍스트']);
      });
      excelData.push(['']); // 빈 줄
    }
    
    if (numberTexts.length > 0) {
      excelData.push(['숫자 텍스트', '분류', '길이', '설명']);
      numberTexts.forEach(text => {
        excelData.push([text, '숫자', text.length.toString(), '숫자 텍스트']);
      });
      excelData.push(['']); // 빈 줄
    }
    
    if (mixedTexts.length > 0) {
      excelData.push(['혼합 텍스트', '분류', '길이', '설명']);
      mixedTexts.forEach(text => {
        excelData.push([text, '혼합', text.length.toString(), '혼합 텍스트']);
      });
    }
    
    // 데이터가 없으면 기본 데이터 추가
    if (excelData.length <= 1) {
      excelData.push(['데이터 없음', '오류', '0', '추출된 텍스트가 없습니다']);
    }
    
    const worksheet = XLSX.utils.aoa_to_sheet(excelData);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    
    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    console.log('✅ 의미있는 텍스트 Excel 생성 완료:', buffer.length, 'bytes');
    
    return buffer;
  } catch (error) {
    console.log('❌ Excel 생성 실패:', error instanceof Error ? error.message : String(error));
    
    // 최소한의 오류 파일 생성
    const workbook = XLSX.utils.book_new();
    const errorData = [
      ['오류 발생'],
      ['파일명', filename],
      ['오류 내용', error instanceof Error ? error.message : String(error)],
      ['텍스트 수', texts.length.toString()],
      ['텍스트 샘플', texts.slice(0, 5).join(', ')]
    ];
    const worksheet = XLSX.utils.aoa_to_sheet(errorData);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Error');
    
    return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  }
}

/**
 * SharedStrings에서 실제 데이터 추출
 */
async function extractDataFromSharedStrings(buffer: Buffer, filename: string, sharedStringsContent: string): Promise<Buffer> {
  console.log('🔧 SharedStrings에서 데이터 추출 시작');
  console.log('📄 SharedStrings 내용 (처음 500자):', sharedStringsContent.substring(0, 500));
  
  try {
    const extractedTexts: string[] = [];
    
    // 방법 1: 표준 <t> 태그 추출
    const textMatches = sharedStringsContent.match(/<t[^>]*>([^<]*)<\/t>/g);
    if (textMatches) {
      textMatches.forEach(match => {
        const textContent = match.replace(/<[^>]*>/g, '').trim();
        if (textContent && textContent.length > 0) {
          extractedTexts.push(textContent);
        }
      });
    }
    
    // 방법 2: 더 유연한 텍스트 추출 (단계적 필터링)
    if (extractedTexts.length === 0) {
      console.log('🔧 표준 <t> 태그 없음, 유연한 텍스트 추출 시도');
      
      // 1단계: 기본 정리 (XML 태그 제거)
      const withoutTags = sharedStringsContent.replace(/<[^>]*>/g, ' ');
      
      // 2단계: 단어 분할 및 기본 필터링
      const allWords = withoutTags
        .split(/[\s\n\r\t]+/)
        .map(word => word.trim())
        .filter(word => word.length > 0);
      
      console.log(`🔧 전체 단어 수: ${allWords.length}, 샘플:`, allWords.slice(0, 10));
      
      // 3단계: 점진적 필터링 (여러 레벨로 시도)
      const filterLevels = [
        // 레벨 1: 가장 엄격한 필터링
        (word: string) => {
          return word.length >= 2 && 
                 word.length <= 50 &&
                 /^[가-힣a-zA-Z0-9\-_()[\]{}.,]+$/.test(word) &&
                 !/^[0-9.]+$/.test(word) &&
                 !/^xl\/|\.xml$|^PK|^Content|^Types|sharedStrings/i.test(word);
        },
        // 레벨 2: 중간 필터링 (특수문자 일부 허용)
        (word: string) => {
          return word.length >= 2 && 
                 word.length <= 50 &&
                 /[가-힣a-zA-Z0-9]/.test(word) && // 최소한 의미있는 문자 포함
                 !/^xl\/|\.xml$|^PK|^Content|^Types/i.test(word) &&
                 !/[\x00-\x08\x0E-\x1F\x7F-\x9F]/.test(word); // 심각한 제어문자만 제외
        },
        // 레벨 3: 관대한 필터링 (길이와 기본 문자만 체크)
        (word: string) => {
          return word.length >= 1 && 
                 word.length <= 100 &&
                 /[가-힣a-zA-Z0-9]/.test(word) && // 최소한 의미있는 문자 포함
                 !/^PK$|^xl$|xml$/i.test(word); // 명확한 시스템 키워드만 제외
        }
      ];
      
      // 각 레벨별로 시도 (충분한 데이터가 나올 때까지)
      for (let level = 0; level < filterLevels.length; level++) {
        const filtered = allWords.filter(filterLevels[level]);
        console.log(`🔧 레벨 ${level + 1} 필터링 결과: ${filtered.length}개`);
        
        if (filtered.length > 0) {
          const unique = Array.from(new Set(filtered)).slice(0, 50);
          console.log(`📋 레벨 ${level + 1} 샘플:`, unique.slice(0, 10));
          
          // 충분한 데이터가 있거나 마지막 레벨이면 사용
          if (unique.length >= 5 || level === filterLevels.length - 1) {
            extractedTexts.push(...unique);
            console.log(`✅ 레벨 ${level + 1}에서 ${unique.length}개 텍스트 최종 선택`);
            break;
          } else {
            console.log(`🔧 레벨 ${level + 1} 결과가 부족함 (${unique.length}개), 다음 레벨 시도`);
          }
        }
      }
    }
    
    // 방법 3: 특정 패턴 찾기 (관대한 접근)
    if (extractedTexts.length < 5) { // 5개 미만이면 패턴 추출도 시도
      console.log(`🔧 패턴 기반 텍스트 추출 시도 (현재 ${extractedTexts.length}개)`);
      
      const patterns = [
        /[가-힣]{1,}/g, // 한글 1글자 이상
        /[a-zA-Z]{2,}/g, // 영문 2글자 이상
        /[0-9]{4,}/g, // 숫자 4글자 이상 (전화번호, 주문번호 등)
        /[가-힣a-zA-Z0-9]{1,}/g // 모든 의미있는 문자
      ];
      
      for (const pattern of patterns) {
        const matches = sharedStringsContent.match(pattern);
        if (matches) {
          console.log(`🔧 패턴 매칭 결과: ${matches.length}개 발견`);
          
          matches.forEach(match => {
            const cleaned = match.trim();
            // 관대한 필터링
            if (cleaned.length >= 1 && 
                cleaned.length <= 100 && 
                !/^xl$|^PK$|xml$/i.test(cleaned)) { // 명확한 시스템 키워드만 제외
              extractedTexts.push(cleaned);
            }
          });
          
          if (extractedTexts.length > 0) {
            // 기존 데이터와 합치기
            const allTexts = Array.from(new Set([...extractedTexts])); // 기존 데이터 유지
            const unique = Array.from(new Set(allTexts)).slice(0, 50);
            console.log(`✅ 패턴에서 총 ${unique.length}개 고유 텍스트 (기존 포함):`, unique.slice(0, 10));
            extractedTexts.length = 0; // 기존 배열 초기화
            extractedTexts.push(...unique);
            break;
          }
        }
      }
    }
    
    console.log(`📋 SharedStrings에서 ${extractedTexts.length}개 텍스트 추출`);
    console.log('📄 추출된 텍스트 샘플:', extractedTexts.slice(0, 10));
    
    // 의미있는 한글 텍스트가 있는지 확인
    const meaningfulTexts = extractedTexts.filter(text => {
      return /[가-힣]{2,}/.test(text) || // 한글 2글자 이상
             /\b(주문|배송|연락|전화|주소|번호|회사|고객|상품)\b/.test(text) || // 업무 키워드
             /\b\d{4,}\b/.test(text); // 의미있는 숫자
    });
    
    console.log(`📋 의미있는 텍스트: ${meaningfulTexts.length}개`, meaningfulTexts.slice(0, 5));
    
    if (meaningfulTexts.length > 0) {
      console.log('✅ SharedStrings에서 의미있는 텍스트 발견, 사용');
    } else {
      console.log('❌ SharedStrings에 의미있는 텍스트 없음, 전체 파일 스캔으로 전환');
      throw new Error('의미있는 텍스트 없음');
    }
    
    if (extractedTexts.length > 0) {
      // 추출된 텍스트를 Excel 형태로 구성
      const excelData: string[][] = [];
      
      // 첫 번째 텍스트를 헤더로 사용하거나 적절히 분할
      const headerCount = Math.min(extractedTexts.length, 10); // 최대 10개 컬럼
      const headers = extractedTexts.slice(0, headerCount).map(text => normalizeHeaderField(text));
      excelData.push(headers);
      
      // 나머지 데이터를 행으로 구성
      for (let i = headerCount; i < extractedTexts.length; i += headerCount) {
        const row = extractedTexts.slice(i, i + headerCount);
        while (row.length < headerCount) row.push('');
        excelData.push(row);
      }
      
      // Excel 파일 생성
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.aoa_to_sheet(excelData);
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
      
      const outputBuffer = XLSX.write(workbook, { 
        type: 'buffer', 
        bookType: 'xlsx',
        compression: true 
      });
      
      console.log(`✅ SharedStrings 데이터 변환 완료: ${outputBuffer.length} bytes`);
      return outputBuffer;
    }
  } catch (error) {
    console.error('❌ SharedStrings 추출 실패:', error instanceof Error ? error.message : String(error));
  }
  
  // 실패 시 원본 파일에서 실제 한글/의미있는 텍스트 찾기
  console.log('🔧 SharedStrings에서 의미있는 텍스트 없음, 원본 파일 전체 스캔 시도');
  return await findMeaningfulTextInBuffer(buffer, filename);
}

/**
 * 원본 파일에서 실제 의미있는 텍스트 찾기 (한글, 실제 단어 중심)
 */
async function findMeaningfulTextInBuffer(buffer: Buffer, filename: string): Promise<Buffer> {
  console.log('🔧 원본 파일에서 실제 의미있는 텍스트 스캔 시작');
  
  try {
    const meaningfulTexts: string[] = [];
    
    // 1. 파일명에서 의미있는 텍스트 추출
    console.log('🔧 파일명에서 텍스트 추출:', filename);
    const filenameTexts = filename
      .replace(/\.[^.]+$/, '') // 확장자 제거
      .split(/[_\-\[\]().\s]+/) // 구분자로 분할
      .filter(text => text.length > 0)
      .filter(text => /[가-힣]/.test(text) || /[a-zA-Z]{2,}/.test(text) || /\d{4,}/.test(text));
    
    console.log('📋 파일명에서 추출된 텍스트:', filenameTexts);
    meaningfulTexts.push(...filenameTexts);
    
    // 파일명 텍스트는 백업용으로만 사용, 실제 Excel 데이터 추출 우선
    
    // 2. 다양한 인코딩으로 스캔 (EUC-KR 우선)
    const encodings = ['euc-kr', 'cp949', 'utf8', 'utf16le', 'latin1', 'ascii'] as const;
    
    for (const encoding of encodings) {
      try {
        console.log(`🔧 ${encoding} 인코딩으로 전체 파일 스캔`);
        const textContent = buffer.toString(encoding as BufferEncoding);
        
        // 1. 한글 단어 찾기 (2글자 이상의 완전한 한글 단어)
        const koreanWords = textContent.match(/[가-힣]{2,}/g);
        if (koreanWords) {
          const validKorean = koreanWords
            .filter(word => {
              return word.length >= 2 && 
                     word.length <= 20 &&
                     /^[가-힣]+$/.test(word) && // 순수 한글만
                     !word.includes('ㅏ') && !word.includes('ㅓ') && // 불완전한 글자 제외
                     word !== '가가' && word !== '나나'; // 의미없는 반복 제외
            });
          
          if (validKorean.length > 0) {
            meaningfulTexts.push(...validKorean);
            console.log(`📋 ${encoding}에서 한글 단어 ${validKorean.length}개 발견:`, validKorean.slice(0, 10));
          }
        }
        
        // 2. 실제 업무 관련 키워드 찾기
        const businessKeywords = [
          /주문번호/g, /송장번호/g, /운송장/g, /배송/g, /택배/g,
          /연락처/g, /전화번호/g, /휴대폰/g, /핸드폰/g,
          /주소/g, /도로명/g, /지번/g, /우편번호/g,
          /수량/g, /금액/g, /가격/g, /합계/g, /총액/g,
          /고객/g, /업체/g, /회사/g, /상호/g, /법인/g,
          /날짜/g, /시간/g, /년/g, /월/g, /일/g,
          /상품/g, /제품/g, /품목/g, /아이템/g,
          /부내사업/g, /지정송하인/g, /오쎄/g
        ];
        
        businessKeywords.forEach(pattern => {
          const matches = textContent.match(pattern);
          if (matches) {
            meaningfulTexts.push(...matches);
            console.log(`📋 ${encoding}에서 업무 키워드 발견:`, matches);
          }
        });
        
        // 3. 영문 단어 (업무 관련)
        const businessEnglishWords = textContent.match(/\b(order|delivery|phone|address|company|customer|date|time|product|item|type|number)\b/gi);
        if (businessEnglishWords) {
          meaningfulTexts.push(...businessEnglishWords);
          console.log(`📋 ${encoding}에서 영문 업무 단어 ${businessEnglishWords.length}개 발견:`, businessEnglishWords.slice(0, 5));
        }
        
        // 4. 숫자 패턴 (전화번호, 우편번호, 주문번호 등)
        const numberPatterns = [
          /\b\d{2,4}-\d{2,4}-\d{4}\b/g, // 전화번호
          /\b\d{5}\b/g, // 우편번호
          /\b20\d{8,10}\b/g, // 주문번호 (2025072369 같은)
          /\b\d{4,}/g // 기타 긴 숫자
        ];
        
        numberPatterns.forEach(pattern => {
          const matches = textContent.match(pattern);
          if (matches) {
            meaningfulTexts.push(...matches);
            console.log(`📋 ${encoding}에서 숫자 패턴 발견:`, matches.slice(0, 5));
          }
        });
        
      } catch (error) {
        console.log(`❌ ${encoding} 인코딩 스캔 실패`);
      }
    }
    
    // 중복 제거 및 정리
    const uniqueTexts = Array.from(new Set(meaningfulTexts))
      .filter(text => text && text.trim().length > 0)
      .slice(0, 100);
    
    console.log(`📋 총 ${uniqueTexts.length}개 의미있는 텍스트 발견`);
    console.log('📄 발견된 텍스트 샘플:', uniqueTexts.slice(0, 20));
    
    if (uniqueTexts.length > 0) {
      // Excel 형태로 구성
      const excelData: string[][] = [];
      
      // 헤더
      excelData.push(['발견된 텍스트', '타입', '길이', '설명']);
      
      // 데이터 분류 및 추가
      uniqueTexts.forEach(text => {
        let type = '기타';
        let description = '';
        
        if (/^[가-힣]+$/.test(text)) {
          type = '한글';
          if (text.includes('주문') || text.includes('번호')) description = '주문 관련';
          else if (text.includes('배송') || text.includes('택배')) description = '배송 관련';
          else if (text.includes('연락') || text.includes('전화')) description = '연락처 관련';
          else description = '한글 텍스트';
        } else if (/^[a-zA-Z]+$/.test(text)) {
          type = '영문';
          description = '업무 키워드';
        } else if (/^\d+$/.test(text)) {
          type = '숫자';
          if (text.length >= 8) description = '주문번호/ID';
          else if (text.length === 5) description = '우편번호';
          else description = '기타 숫자';
        } else {
          type = '복합';
          description = '전화번호/주소 등';
        }
        
        excelData.push([
          normalizeHeaderField(text),
          type,
          String(text.length),
          description
        ]);
      });
      
      // Excel 파일 생성
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.aoa_to_sheet(excelData);
      XLSX.utils.book_append_sheet(workbook, worksheet, 'MeaningfulData');
      
      const outputBuffer = XLSX.write(workbook, { 
        type: 'buffer', 
        bookType: 'xlsx',
        compression: true 
      });
      
      console.log(`✅ 의미있는 텍스트 추출 완료: ${outputBuffer.length} bytes`);
      return outputBuffer;
    }
    
  } catch (error) {
    console.error('❌ 의미있는 텍스트 스캔 실패:', error instanceof Error ? error.message : String(error));
  }
  
  // 최종 실패 시 - 파일이 손상되었음을 알림
  const fallbackData = [
    ['상태', '설명'],
    ['파일 손상', '의미있는 텍스트를 찾을 수 없습니다'],
    ['파일명', filename],
    ['파일 크기', `${buffer.length} bytes`],
    ['권장사항', '원본 파일을 다시 확인해주세요']
  ];
  
  const fallbackWorkbook = XLSX.utils.book_new();
  const fallbackWorksheet = XLSX.utils.aoa_to_sheet(fallbackData);
  XLSX.utils.book_append_sheet(fallbackWorkbook, fallbackWorksheet, 'FileStatus');
  
  return XLSX.write(fallbackWorkbook, { 
    type: 'buffer', 
    bookType: 'xlsx',
    compression: true 
  });
}

/**
 * 원본 파일에서 직접 텍스트 추출 (최후의 수단)
 */
async function extractTextDirectlyFromBuffer(buffer: Buffer, filename: string): Promise<Buffer> {
  console.log('🔧 원본 파일에서 직접 텍스트 추출 시작');
  
  try {
    const extractedTexts: string[] = [];
    
    // 여러 인코딩으로 텍스트 추출
    const encodings = ['utf8', 'latin1', 'ascii'] as const;
    
    for (const encoding of encodings) {
      try {
        const textContent = buffer.toString(encoding);
        console.log(`🔧 ${encoding} 인코딩으로 텍스트 추출 시도`);
        
        // 한글 텍스트 패턴 찾기 (더 엄격한 필터링)
        const koreanMatches = textContent.match(/[가-힣]{2,}/g);
        if (koreanMatches) {
          const uniqueKorean = Array.from(new Set(koreanMatches))
            .filter(text => {
              return text.length >= 2 && 
                     text.length <= 50 &&
                     !/[\x00-\x1F\x7F-\xFF]/.test(text) && // 바이너리 문자 제외
                     /^[가-힣]+$/.test(text); // 순수 한글만
            })
            .slice(0, 20);
          
          if (uniqueKorean.length > 0) {
            extractedTexts.push(...uniqueKorean);
            console.log(`📋 ${encoding}에서 한글 텍스트 ${uniqueKorean.length}개 발견:`, uniqueKorean.slice(0, 5));
          }
        }
        
        // 영문 텍스트 패턴 찾기 (더 엄격한 필터링)
        const englishMatches = textContent.match(/[a-zA-Z]{3,}/g);
        if (englishMatches) {
          const uniqueEnglish = Array.from(new Set(englishMatches))
            .filter(text => {
              return text.length >= 3 && 
                     text.length <= 50 &&
                     !/[\x00-\x1F\x7F-\xFF]/.test(text) && // 바이너리 문자 제외
                     /^[a-zA-Z]+$/.test(text) && // 순수 영문만
                     !/^(PK|xml|Content|Types|rels|docProps|app|core|workbook|worksheet|sharedStrings|styles|DEFLATE|GZIP|ascii|utf|latin)$/i.test(text); // XML/시스템 키워드 제외
            })
            .slice(0, 20);
          
          if (uniqueEnglish.length > 0) {
            extractedTexts.push(...uniqueEnglish);
            console.log(`📋 ${encoding}에서 영문 텍스트 ${uniqueEnglish.length}개 발견:`, uniqueEnglish.slice(0, 5));
          }
        }
        
        // 숫자 패턴 찾기 (전화번호, 주문번호 등)
        const numberMatches = textContent.match(/[0-9]{4,}/g);
        if (numberMatches) {
          const uniqueNumbers = Array.from(new Set(numberMatches))
            .filter(text => text.length >= 4 && text.length <= 20)
            .slice(0, 10);
          extractedTexts.push(...uniqueNumbers);
          console.log(`📋 ${encoding}에서 숫자 패턴 ${uniqueNumbers.length}개 발견:`, uniqueNumbers.slice(0, 3));
        }
        
        if (extractedTexts.length >= 10) break; // 충분한 데이터가 있으면 중단
        
      } catch (error) {
        console.log(`❌ ${encoding} 인코딩 실패`);
      }
    }
    
    console.log(`📋 총 ${extractedTexts.length}개 텍스트 추출 완료`);
    
    if (extractedTexts.length > 0) {
      // 중복 제거 및 정리
      const uniqueTexts = Array.from(new Set(extractedTexts));
      
      // Excel 형태로 구성
      const excelData: string[][] = [];
      
      // 헤더 생성
      const headers = ['추출된 데이터', '타입', '길이', '인코딩'];
      excelData.push(headers);
      
      // 데이터 추가
      uniqueTexts.forEach((text, index) => {
        const type = /[가-힣]/.test(text) ? '한글' : 
                    /[a-zA-Z]/.test(text) ? '영문' : 
                    /[0-9]/.test(text) ? '숫자' : '기타';
        excelData.push([
          normalizeHeaderField(text), 
          type, 
          String(text.length),
          '직접추출'
        ]);
      });
      
      // Excel 파일 생성
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.aoa_to_sheet(excelData);
      XLSX.utils.book_append_sheet(workbook, worksheet, 'ExtractedData');
      
      const outputBuffer = XLSX.write(workbook, { 
        type: 'buffer', 
        bookType: 'xlsx',
        compression: true 
      });
      
      console.log(`✅ 직접 텍스트 추출 완료: ${outputBuffer.length} bytes`);
      return outputBuffer;
    }
    
  } catch (error) {
    console.error('❌ 직접 텍스트 추출 실패:', error instanceof Error ? error.message : String(error));
  }
  
  // 최종 실패 시 기본 오류 파일 생성
  const fallbackData = [
    ['파일명', '상태', '크기'],
    [filename, '텍스트 추출 실패', `${buffer.length} bytes`]
  ];
  
  const fallbackWorkbook = XLSX.utils.book_new();
  const fallbackWorksheet = XLSX.utils.aoa_to_sheet(fallbackData);
  XLSX.utils.book_append_sheet(fallbackWorkbook, fallbackWorksheet, 'Error');
  
  return XLSX.write(fallbackWorkbook, { 
    type: 'buffer', 
    bookType: 'xlsx',
    compression: true 
  });
}

/**
 * Worksheet에서 실제 데이터 추출
 */
async function extractDataFromWorksheet(buffer: Buffer, filename: string, worksheetContent: string): Promise<Buffer> {
  console.log('🔧 Worksheet에서 데이터 추출 시작');
  
  try {
    // <c> 태그 (셀) 안의 <v> 태그 (값) 추출
    const cellMatches = worksheetContent.match(/<c[^>]*>.*?<\/c>/g);
    const extractedValues: string[] = [];
    
    if (cellMatches) {
      cellMatches.forEach(cellMatch => {
        const valueMatch = cellMatch.match(/<v[^>]*>([^<]*)<\/v>/);
        if (valueMatch && valueMatch[1]) {
          extractedValues.push(valueMatch[1].trim());
        }
      });
    }
    
    console.log(`📋 Worksheet에서 ${extractedValues.length}개 값 추출`);
    console.log('📄 추출된 값 샘플:', extractedValues.slice(0, 10));
    
    if (extractedValues.length > 0) {
      // 추출된 값을 Excel 형태로 구성
      const excelData: string[][] = [];
      
      // 적절한 컬럼 수 추정 (보통 5-15개)
      const estimatedColumns = Math.min(Math.max(Math.floor(Math.sqrt(extractedValues.length)), 3), 15);
      
      for (let i = 0; i < extractedValues.length; i += estimatedColumns) {
        const row = extractedValues.slice(i, i + estimatedColumns);
        while (row.length < estimatedColumns) row.push('');
        
        // 첫 번째 행은 헤더로 정규화
        if (i === 0) {
          const headers = row.map(cell => normalizeHeaderField(cell) || `컬럼${row.indexOf(cell) + 1}`);
          excelData.push(headers);
        } else {
          excelData.push(row);
        }
      }
      
      // Excel 파일 생성
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.aoa_to_sheet(excelData);
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
      
      const outputBuffer = XLSX.write(workbook, { 
        type: 'buffer', 
        bookType: 'xlsx',
        compression: true 
      });
      
      console.log(`✅ Worksheet 데이터 변환 완료: ${outputBuffer.length} bytes`);
      return outputBuffer;
    }
  } catch (error) {
    console.error('❌ Worksheet 추출 실패:', error instanceof Error ? error.message : String(error));
  }
  
  // 실패 시 기본 처리로 넘김
  throw new Error('Worksheet 추출 실패');
}

/**
 * Excel 파일 전용 처리 함수 (구 Excel, 한셀, 최신 Excel 모두 지원)
 */
async function handleExcelFile(buffer: Buffer, filename: string): Promise<Buffer> {
  console.log('🔧 Excel 파일 전용 처리 시작');
  
  // 1. 먼저 강력한 XLSX.read 옵션들로 시도
  console.log('🔧 강화된 XLSX.read 옵션들로 시도');
  const readOptions = [
    { cellText: false, cellDates: true, raw: false },
    { cellText: true, cellDates: true, raw: false },
    { cellText: false, cellDates: true, raw: true },
    { cellText: true, cellDates: true, raw: true },
    { cellText: false, cellDates: false, raw: false, codepage: 949 },
    { cellText: true, cellDates: false, raw: false, codepage: 949 },
    { cellText: false, cellDates: false, raw: true, codepage: 949 },
    { cellText: false, cellDates: true, raw: false, codepage: 1200 },
    { cellText: false, cellDates: true, raw: false, codepage: 65001 },
    { cellText: false, cellDates: true, raw: false, bookVBA: true },
    { cellText: false, cellDates: true, raw: false, bookSheets: true, bookProps: true }
  ];
  
  for (let i = 0; i < readOptions.length; i++) {
    try {
      console.log(`🔧 XLSX.read 옵션 ${i + 1}/${readOptions.length} 시도:`, readOptions[i]);
      const workbook = XLSX.read(buffer, readOptions[i]);
      
      if (workbook && workbook.SheetNames && workbook.SheetNames.length > 0) {
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: null }) as any[][];
        
        console.log(`📊 옵션 ${i + 1} 결과: ${jsonData.length}행, 첫 행: ${jsonData[0]?.length || 0}개 컬럼`);
        
        // 실제 데이터가 있는지 확인 (1행이라도 유효한 헤더가 있으면 성공)
        if (jsonData.length > 0 && jsonData[0] && jsonData[0].length > 0) {
          // 첫 번째 행이 모두 null/undefined가 아닌지 확인
          const hasValidHeader = jsonData[0].some(cell => cell !== null && cell !== undefined && cell !== '');
          
          if (hasValidHeader) {
            console.log('✅ 유효한 데이터 발견! 헤더:', jsonData[0].slice(0, 5));
            console.log('📋 전체 헤더:', jsonData[0]);
            console.log(`📊 총 ${jsonData.length}행, ${jsonData[0].length}개 컬럼`);
            
            // 1행만 있어도 유효한 Excel 파일로 처리
            const normalizedWorkbook = normalizeWorkbook(workbook);
            const outputBuffer = XLSX.write(normalizedWorkbook, { type: 'buffer', bookType: 'xlsx' });
            console.log('✅ XLSX.read 성공으로 변환 완료');
            return outputBuffer;
          } else {
            console.log('❌ 헤더가 모두 비어있음:', jsonData[0]);
          }
        } else {
          console.log('❌ 데이터가 없거나 구조가 잘못됨');
        }
      }
    } catch (error) {
      console.log(`❌ 옵션 ${i + 1} 실패:`, error instanceof Error ? error.message : String(error));
    }
  }
  
  console.log('❌ 모든 XLSX.read 옵션 실패, ZIP 구조 분석으로 전환');
  
  // 2. ZIP 구조 분석 (기존 로직)
  console.log('🔍 파일 구조 분석 시작');
  const header = buffer.slice(0, 100);
  console.log('📄 파일 헤더 (처음 100바이트):', header.toString('hex').substring(0, 200));
  console.log('📄 파일 헤더 (ASCII):', header.toString('ascii').replace(/[^\x20-\x7E]/g, '.'));
  
  // ZIP 파일 내부 구조 확인 및 압축 해제
  if (buffer[0] === 0x50 && buffer[1] === 0x4B) {
    console.log('🔍 ZIP 파일 구조 분석 및 압축 해제 시도');
    try {
      // JSZip 라이브러리 대신 간단한 ZIP 엔트리 찾기
      const textContent = buffer.toString('latin1'); // 바이너리 데이터 보존
      
      // SharedStrings.xml 파일 찾기
      const sharedStringsIndex = textContent.indexOf('xl/sharedStrings.xml');
      if (sharedStringsIndex !== -1) {
        console.log('📋 SharedStrings.xml 파일 위치 발견:', sharedStringsIndex);
        
        // ZIP 엔트리 헤더 분석하여 압축된 데이터 위치 찾기
        const zipEntryStart = textContent.lastIndexOf('PK\x03\x04', sharedStringsIndex);
        if (zipEntryStart !== -1) {
          console.log('📋 SharedStrings ZIP 엔트리 시작 위치:', zipEntryStart);
          
          // 압축된 데이터 추출 시도 (간단한 방법)
          const afterEntry = textContent.substring(zipEntryStart + 30); // ZIP 헤더 스킵
          const nextPKIndex = afterEntry.indexOf('PK');
          const compressedData = nextPKIndex !== -1 ? afterEntry.substring(0, nextPKIndex) : afterEntry.substring(0, 1000);
          
          console.log('📋 압축된 데이터 길이:', compressedData.length);
          
          // 압축 해제 시도 (여러 방법)
          const decompressedData = await tryDecompressData(compressedData);
          if (decompressedData) {
            console.log('✅ 압축 해제 성공, XML 파싱 시도');
            return await extractDataFromSharedStrings(buffer, filename, decompressedData);
          }
        }
      }
      
      // Worksheet 파일 찾기
      const worksheetIndex = textContent.indexOf('xl/worksheets/sheet1.xml');
      if (worksheetIndex !== -1) {
        console.log('📋 Worksheet.xml 파일 위치 발견:', worksheetIndex);
        
        const zipEntryStart = textContent.lastIndexOf('PK\x03\x04', worksheetIndex);
        if (zipEntryStart !== -1) {
          const afterEntry = textContent.substring(zipEntryStart + 30);
          const nextPKIndex = afterEntry.indexOf('PK');
          const compressedData = nextPKIndex !== -1 ? afterEntry.substring(0, nextPKIndex) : afterEntry.substring(0, 2000);
          
          const decompressedData = await tryDecompressData(compressedData);
          if (decompressedData) {
            console.log('✅ Worksheet 압축 해제 성공, XML 파싱 시도');
            return await extractDataFromWorksheet(buffer, filename, decompressedData);
          }
        }
      }
      
    } catch (error) {
      console.log('❌ ZIP 구조 분석 실패:', error instanceof Error ? error.message : String(error));
    }
  }
  
  // 다양한 Excel 파서 옵션으로 시도 (백업)
  const fallbackReadOptions = [
    // 구 Excel 파일용 옵션들
    { type: 'buffer' as const, codepage: 949, cellText: false, cellDates: true }, // CP949/EUC-KR
    { type: 'buffer' as const, codepage: 1200, cellText: false, cellDates: true }, // UTF-16
    { type: 'buffer' as const, codepage: 65001, cellText: false, cellDates: true }, // UTF-8
    { type: 'buffer' as const, cellText: false, cellDates: true, raw: true }, // 원본 데이터
    { type: 'buffer' as const, cellText: true, cellDates: false, raw: false }, // 텍스트 변환
    { type: 'buffer' as const, cellFormula: false, cellHTML: false }, // 기본 옵션
  ];
  
  for (let i = 0; i < fallbackReadOptions.length; i++) {
    try {
      console.log(`🔧 Excel 파서 옵션 ${i + 1} 시도:`, fallbackReadOptions[i]);
      const workbook = XLSX.read(buffer, fallbackReadOptions[i]);
      
      if (workbook.SheetNames.length === 0) {
        console.log(`❌ 옵션 ${i + 1}: 시트가 없음`);
        continue;
      }
      
      // 첫 번째 시트 데이터 확인
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const testData = XLSX.utils.sheet_to_json(firstSheet, { 
        header: 1, 
        defval: null,
        raw: false,
        blankrows: false
      }) as any[][];
      
      console.log(`📊 옵션 ${i + 1} 결과: ${testData.length}행`);
      if (testData.length > 0 && testData[0]) {
        console.log(`📋 첫 행 샘플:`, testData[0].slice(0, 5));
      }
      
      // 유효한 데이터가 있는지 확인
      if (testData.length > 0 && testData[0] && 
          testData[0].some(cell => cell && String(cell).trim() !== '')) {
        console.log(`✅ 옵션 ${i + 1}로 Excel 파일 읽기 성공!`);
        
        // 정규화 처리
        console.log('🔧 Excel 데이터 정규화 시작');
        const normalizedWorkbook = normalizeWorkbook(workbook);
        
        // .xlsx 파일로 출력
        const outputBuffer = XLSX.write(normalizedWorkbook, { 
          type: 'buffer', 
          bookType: 'xlsx',
          compression: true 
        });
        
        console.log(`✅ Excel 파일 정규화 완료: ${outputBuffer.length} bytes`);
        return outputBuffer;
      }
    } catch (error) {
      console.log(`❌ 옵션 ${i + 1} 실패:`, error instanceof Error ? error.message : String(error));
    }
  }
  
  // 모든 표준 파서 실패 시 바이너리 처리로 fallback
  console.log('🔧 모든 Excel 파서 실패, 바이너리 처리로 전환');
  return await handleBinaryExcelFile(buffer, filename);
}

/**
 * 바이너리 Excel 파일을 텍스트로 변환하여 새로운 Excel 파일로 저장
 */
async function handleBinaryExcelFile(buffer: Buffer, filename: string): Promise<Buffer> {
  console.log('🔧 바이너리 Excel 파일 처리 시작');
  
  try {
    // 표준 Excel 파서로 직접 읽기 (다양한 옵션 시도)
    console.log('🔧 표준 Excel 파서로 직접 읽기 시도');
    
    const readOptions = [
      { type: 'buffer' as const, cellText: false, cellDates: true, raw: true },
      { type: 'buffer' as const, cellText: true, cellDates: false, raw: false },
      { type: 'buffer' as const, cellText: false, cellDates: false, raw: false },
      { type: 'buffer' as const, cellFormula: false, cellHTML: false }
    ];
    
    let workbook: XLSX.WorkBook | null = null;
    let successfulOption = -1;
    
    for (let i = 0; i < readOptions.length; i++) {
      try {
        console.log(`🔧 옵션 ${i + 1} 시도:`, readOptions[i]);
        const testWorkbook = XLSX.read(buffer, readOptions[i]);
        
        // 데이터가 있는지 확인
        const firstSheet = testWorkbook.Sheets[testWorkbook.SheetNames[0]];
        const testData = XLSX.utils.sheet_to_json(firstSheet, { 
          header: 1, 
          defval: null,
          raw: false,
          blankrows: false
        }) as any[][];
        
        console.log(`📊 옵션 ${i + 1} 결과: ${testData.length}행, 첫 행:`, testData[0]?.slice(0, 3));
        
        if (testData.length > 0 && testData[0] && testData[0].some(cell => cell && String(cell).trim() !== '')) {
          console.log(`✅ 옵션 ${i + 1}로 성공! 유효한 데이터 발견`);
          workbook = testWorkbook;
          successfulOption = i;
          break;
        }
      } catch (error) {
        console.log(`❌ 옵션 ${i + 1} 실패:`, error instanceof Error ? error.message : String(error));
      }
    }
    
    if (!workbook) {
      console.log('🔧 모든 표준 파서 옵션 실패, JSON 변환 방식 시도');
      return await convertBinaryToJsonToXlsx(buffer, filename);
    }
    
    console.log(`✅ 표준 파서 성공 (옵션 ${successfulOption + 1})`);
    
    // 정상적인 Excel 데이터 처리 및 정규화
    console.log('🔧 Excel 데이터 정규화 시작');
    
    // 정규화된 워크북 생성
    const normalizedWorkbook = normalizeWorkbook(workbook);
    
    console.log('✅ Excel 데이터 정규화 완료');
    
    // 정규화된 워크북을 .xlsx 파일로 출력
    const outputBuffer = XLSX.write(normalizedWorkbook, { 
      type: 'buffer', 
      bookType: 'xlsx',
      compression: true 
    });
    
    console.log(`✅ 정규화된 Excel 파일 생성 완료: ${outputBuffer.length} bytes`);
    return outputBuffer;
    
  } catch (error) {
    console.error('❌ 바이너리 파일 처리 실패:', error instanceof Error ? error.message : String(error));
    
    // 실패 시 원본 파일의 내용을 텍스트로 표시
    console.log('🔄 원본 파일 내용을 텍스트로 변환 시도');
    
    const textContent = buffer.toString('utf8', 0, Math.min(buffer.length, 10000)); // 처음 10KB만
    const lines = textContent.split('\n').slice(0, 100); // 처음 100줄만
    
    const fallbackData = [
      ['파일명', '내용', '비고'],
      [filename, '바이너리 파일 - 텍스트 변환', `원본 크기: ${buffer.length} bytes`],
      ...lines.map((line, index) => [`라인 ${index + 1}`, line.substring(0, 500), ''])
    ];
    
    const fallbackWorkbook = XLSX.utils.book_new();
    const fallbackWorksheet = XLSX.utils.aoa_to_sheet(fallbackData);
    XLSX.utils.book_append_sheet(fallbackWorkbook, fallbackWorksheet, 'Sheet1');
    
    return XLSX.write(fallbackWorkbook, { 
      type: 'buffer', 
      bookType: 'xlsx',
      compression: true 
    });
  }
}

/**
 * 메인 변환 함수
 */
export async function convertToXlsx(
  buffer: Buffer, 
  filename: string,
  forceTextRecovery: boolean = false
): Promise<Buffer> {
  console.log('🔧 convertToXlsx 함수 시작, 파일명:', filename, '크기:', buffer.length, 'bytes');
  console.log('🔧 forceTextRecovery:', forceTextRecovery);
  console.log('🔧 모든 파일을 호환성 향상을 위해 정규화 처리합니다');

  // 파일 형식 감지 및 적절한 처리 방식 선택
  const fileExtension = filename.toLowerCase().split('.').pop();
  const isZipFile = buffer.length >= 4 && 
    buffer[0] === 0x50 && buffer[1] === 0x4B && 
    (buffer[2] === 0x03 || buffer[2] === 0x05 || buffer[2] === 0x07);
  
  console.log(`🔧 파일 확장자: ${fileExtension}, ZIP 시그니처: ${isZipFile}`);
  
  // 1. 먼저 표준 Excel 파서로 직접 시도 (모든 Excel 파일)
  if (fileExtension === 'xlsx' || fileExtension === 'xls' || isZipFile) {
    console.log('🔧 Excel 파일 감지됨, 표준 파서 우선 시도');
    return await handleExcelFile(buffer, filename);
  }
  
  try {
    let workbook: XLSX.WorkBook;
    
    if (!forceTextRecovery) {
      try {
        console.log('🔧 표준 파서 시도 중...');
        // 1단계: 표준 파서 시도
        workbook = XLSX.read(buffer, {
          type: 'buffer',
          cellDates: true,
          cellNF: false,
          cellText: false,
          // 다양한 인코딩 시도
          codepage: 65001, // UTF-8
        });
        console.log('🔧 XLSX.read 완료');
        
        // 워크북이 비어있지 않은지 확인
        if (workbook.SheetNames.length === 0) {
          throw new Error('빈 워크북');
        }
        
        // 첫 번째 시트의 실제 데이터 확인
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { 
          header: 1, 
          defval: null,
          raw: false 
        }) as any[][];
        
        // 헤더가 모두 비어있거나 null이면 텍스트 복구로 전환
        const hasValidHeader = jsonData.length > 0 && 
          jsonData[0].some(cell => cell && String(cell).trim() !== '');
        
        if (!hasValidHeader) {
          console.log('⚠️ 표준 파서로 읽은 헤더가 모두 비어있음, 텍스트 복구로 전환');
          throw new Error('헤더가 비어있음 - 텍스트 복구 필요');
        }
        
        console.log('✅ 표준 파서 성공! 시트 수:', workbook.SheetNames.length);
        console.log('📋 시트 이름들:', workbook.SheetNames);
        const range = firstSheet['!ref'];
        console.log('📊 첫 번째 시트 범위:', range);
        
        // 첫 번째 행(헤더) 확인
        if (range) {
          const firstRowCells = [];
          const endCol = range.split(':')[1]?.charAt(0) || 'A';
          const endColCode = endCol.charCodeAt(0);
          
          for (let i = 65; i <= Math.min(endColCode, 75); i++) { // A~K까지만 확인
            const cellAddr = String.fromCharCode(i) + '1';
            const cell = firstSheet[cellAddr];
            if (cell) {
              firstRowCells.push(cell.v || cell.w || '');
            }
          }
          console.log('🏷️ 표준 파서로 읽은 헤더 (처음 10개):', firstRowCells.slice(0, 10));
        }
        
      } catch (standardError) {
        console.log('⚠️ 표준 파서 실패, 텍스트 복구 시도:', standardError instanceof Error ? standardError.message : String(standardError));
        // 2단계: 텍스트 기반 복구
        workbook = textBasedRecovery(buffer);
      }
    } else {
      // 강제 텍스트 복구
      workbook = textBasedRecovery(buffer);
    }
    
    // 3단계: 데이터 정규화
    console.log('🔧 데이터 정규화 시작...');
    const normalizedWorkbook = normalizeWorkbook(workbook);
    console.log('🔧 데이터 정규화 완료');
    
    // 4단계: .xlsx로 변환
    console.log('🔧 .xlsx 버퍼 생성 시작...');
    const xlsxBuffer = XLSX.write(normalizedWorkbook, {
      type: 'buffer',
      bookType: 'xlsx',
      compression: true,
      cellDates: true,
    });
    console.log('🔧 .xlsx 버퍼 생성 완료, 크기:', xlsxBuffer.length, 'bytes');
    
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
    
    // 크기 비교 경고 (더 정확한 기준)
    const sizeRatio = convertedBuffer.length / buffer.length;
    if (sizeRatio > 3) {
      warnings.push('변환된 파일이 원본보다 상당히 큽니다. 데이터 확인을 권장합니다.');
    } else if (sizeRatio < 0.1 && buffer.length > 1000) {
      warnings.push('변환된 파일이 원본보다 상당히 작습니다. 데이터 손실이 있을 수 있습니다.');
    }
    
    console.log(`파일 크기 비교: 원본 ${buffer.length}bytes → 변환 ${convertedBuffer.length}bytes (비율: ${sizeRatio.toFixed(2)})`);
    
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
