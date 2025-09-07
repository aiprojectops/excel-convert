'use client';

import React, { useState, useCallback, useRef } from 'react';
import { Upload, FileText, AlertCircle, CheckCircle, Download, RotateCcw } from 'lucide-react';
import { clsx } from 'clsx';

// 지원하는 파일 확장자
const SUPPORTED_EXTENSIONS = ['.xls', '.xlsx', '.csv', '.tsv', '.txt'];
const MAX_FILE_SIZE = 50 * 1024 * 1024; // 50MB

// 파일 상태 타입
type FileStatus = 'idle' | 'uploading' | 'converting' | 'success' | 'error';

// 변환 결과 타입
interface ConversionResult {
  success: boolean;
  filename: string;
  originalSize: number;
  convertedSize: number;
  warnings?: string[];
  downloadUrl?: string;
}

// 에러 타입
interface ConversionError {
  message: string;
  code?: string;
}

export default function FileUploader() {
  // 상태 관리
  const [status, setStatus] = useState<FileStatus>('idle');
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [result, setResult] = useState<ConversionResult | null>(null);
  const [error, setError] = useState<ConversionError | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const [progress, setProgress] = useState(0);
  const [forceTextRecovery, setForceTextRecovery] = useState(false);
  const [usePythonService, setUsePythonService] = useState(true); // Python 서비스 사용 여부

  // 참조
  const fileInputRef = useRef<HTMLInputElement>(null);
  const downloadLinkRef = useRef<HTMLAnchorElement>(null);

  /**
   * 파일 검증
   */
  const validateFile = useCallback((file: File): string | null => {
    // 확장자 검증
    const ext = file.name.toLowerCase().substring(file.name.lastIndexOf('.'));
    if (!SUPPORTED_EXTENSIONS.includes(ext)) {
      return `지원하지 않는 파일 형식입니다. 지원 형식: ${SUPPORTED_EXTENSIONS.join(', ')}`;
    }

    // 크기 검증
    if (file.size > MAX_FILE_SIZE) {
      return `파일이 너무 큽니다. 최대 크기: ${Math.round(MAX_FILE_SIZE / 1024 / 1024)}MB`;
    }

    return null;
  }, []);

  /**
   * 파일 선택 처리
   */
  const handleFileSelect = useCallback((file: File) => {
    const validationError = validateFile(file);
    if (validationError) {
      setError({ message: validationError });
      setStatus('error');
      return;
    }

    setSelectedFile(file);
    setStatus('idle');
    setError(null);
    setResult(null);
    setProgress(0);
  }, [validateFile]);

  /**
   * 파일 입력 변경 처리
   */
  const handleFileInputChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      handleFileSelect(file);
    }
  }, [handleFileSelect]);

  /**
   * 드래그 앤 드롭 처리
   */
  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);

    const file = e.dataTransfer.files[0];
    if (file) {
      handleFileSelect(file);
    }
  }, [handleFileSelect]);

  /**
   * 파일 변환 처리
   */
  const handleConvert = useCallback(async () => {
    if (!selectedFile) return;

    console.log('🚀 클라이언트: 변환 시작', selectedFile.name, selectedFile.size + ' bytes');
    setStatus('uploading');
    setError(null);
    setProgress(10);

    try {
      // FormData 생성
      const formData = new FormData();
      formData.append('file', selectedFile);
      formData.append('forceTextRecovery', forceTextRecovery.toString());

      console.log('📤 클라이언트: 서버로 파일 전송 중...');
      setStatus('converting');
      setProgress(50);

      // API 호출 (Python 서비스 또는 기존 TypeScript)
      const apiEndpoint = usePythonService ? '/api/convert-python' : '/api/convert';
      console.log('🔧 사용할 API:', apiEndpoint);
      
      const response = await fetch(apiEndpoint, {
        method: 'POST',
        body: formData,
      });

      console.log('📥 클라이언트: 서버 응답 받음', response.status, response.statusText);
      setProgress(80);

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.message || '변환에 실패했습니다.');
      }

      // 응답 헤더에서 메타데이터 추출
      const originalSize = parseInt(response.headers.get('X-Original-Size') || '0');
      const convertedSize = parseInt(response.headers.get('X-Converted-Size') || '0');
      const conversionMethod = response.headers.get('X-Conversion-Method') || 'unknown';
      const warningsHeader = response.headers.get('X-Warnings');
      const warnings = warningsHeader ? decodeURIComponent(warningsHeader).split('; ') : undefined;
      
      console.log('🔧 변환 방법:', conversionMethod);

      // 파일 다운로드 준비
      const blob = await response.blob();
      console.log('💾 클라이언트: 변환된 파일 크기', blob.size + ' bytes');
      
      const downloadUrl = window.URL.createObjectURL(blob);

      // 결과 설정
      const baseName = selectedFile.name.replace(/\.[^.]+$/, '');
      const resultFilename = `${baseName}_변환완료.xlsx`;

      console.log('✅ 클라이언트: 변환 완료!', resultFilename);
      console.log('📊 클라이언트: 크기 비교 - 원본:', originalSize, '→ 변환:', convertedSize);

      setResult({
        success: true,
        filename: resultFilename,
        originalSize,
        convertedSize,
        warnings,
        downloadUrl,
      });

      setStatus('success');
      setProgress(100);

      // 자동 다운로드 트리거
      setTimeout(() => {
        if (downloadLinkRef.current) {
          downloadLinkRef.current.href = downloadUrl;
          downloadLinkRef.current.download = resultFilename;
          downloadLinkRef.current.click();
        }
      }, 500);

    } catch (err) {
      console.error('❌ 클라이언트: 변환 에러', err);
      setError({
        message: err instanceof Error ? err.message : '알 수 없는 오류가 발생했습니다.',
      });
      setStatus('error');
      setProgress(0);
    }
  }, [selectedFile, forceTextRecovery]);

  /**
   * 초기화
   */
  const handleReset = useCallback(() => {
    setSelectedFile(null);
    setStatus('idle');
    setError(null);
    setResult(null);
    setProgress(0);
    setForceTextRecovery(false);
    
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }

    // 다운로드 URL 정리
    if (result?.downloadUrl) {
      window.URL.revokeObjectURL(result.downloadUrl);
    }
  }, [result?.downloadUrl]);

  /**
   * 파일 크기 포맷팅
   */
  const formatFileSize = useCallback((bytes: number): string => {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }, []);

  return (
    <div className="w-full max-w-2xl mx-auto p-6">
      {/* 헤더 */}
      <div className="text-center mb-8">
        <h1 className="text-3xl font-bold text-gray-900 mb-2">
          Excel Quick Convert
        </h1>
        <p className="text-gray-600">
          구엑셀, 한컴오피스, 깨진 파일을 최신 Excel(.xlsx)로 빠르게 변환
        </p>
      </div>

      {/* 파일 업로드 영역 */}
      <div
        className={clsx(
          'border-2 border-dashed rounded-xl p-8 text-center transition-all duration-200',
          {
            'border-primary-300 bg-primary-50': dragOver,
            'border-gray-300 bg-gray-50': !dragOver && status === 'idle',
            'border-primary-500 bg-primary-50': selectedFile && status === 'idle',
            'border-yellow-300 bg-yellow-50': status === 'uploading' || status === 'converting',
            'border-green-300 bg-green-50': status === 'success',
            'border-red-300 bg-red-50': status === 'error',
          }
        )}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        {/* 아이콘 */}
        <div className="mb-4">
          {status === 'idle' && (
            <Upload className="w-12 h-12 mx-auto text-gray-400" />
          )}
          {(status === 'uploading' || status === 'converting') && (
            <div className="w-12 h-12 mx-auto">
              <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-primary-600"></div>
            </div>
          )}
          {status === 'success' && (
            <CheckCircle className="w-12 h-12 mx-auto text-green-600" />
          )}
          {status === 'error' && (
            <AlertCircle className="w-12 h-12 mx-auto text-red-600" />
          )}
        </div>

        {/* 메시지 */}
        <div className="mb-4">
          {status === 'idle' && !selectedFile && (
            <>
              <p className="text-lg font-medium text-gray-700 mb-2">
                파일을 드래그하거나 클릭하여 선택하세요
              </p>
              <p className="text-sm text-gray-500">
                지원 형식: .xls, .xlsx, .csv, .tsv, .txt (최대 50MB)
              </p>
            </>
          )}
          
          {status === 'idle' && selectedFile && (
            <>
              <FileText className="w-8 h-8 mx-auto text-primary-600 mb-2" />
              <p className="text-lg font-medium text-gray-700">
                {selectedFile.name}
              </p>
              <p className="text-sm text-gray-500">
                {formatFileSize(selectedFile.size)}
              </p>
            </>
          )}

          {status === 'uploading' && (
            <p className="text-lg font-medium text-yellow-700">
              파일 업로드 중...
            </p>
          )}

          {status === 'converting' && (
            <p className="text-lg font-medium text-yellow-700">
              변환 중... 잠시만 기다려주세요
            </p>
          )}

          {status === 'success' && result && (
            <>
              <p className="text-lg font-medium text-green-700 mb-2">
                변환 완료!
              </p>
              <p className="text-sm text-gray-600">
                {result.filename} ({formatFileSize(result.convertedSize)})
              </p>
              {result.warnings && result.warnings.length > 0 && (
                <div className="mt-2 p-2 bg-yellow-100 rounded text-sm text-yellow-800">
                  <p className="font-medium">주의사항:</p>
                  {result.warnings.map((warning, index) => (
                    <p key={index}>• {warning}</p>
                  ))}
                </div>
              )}
            </>
          )}

          {status === 'error' && error && (
            <>
              <p className="text-lg font-medium text-red-700 mb-2">
                변환 실패
              </p>
              <p className="text-sm text-red-600">
                {error.message}
              </p>
            </>
          )}
        </div>

        {/* 진행률 바 */}
        {(status === 'uploading' || status === 'converting') && (
          <div className="w-full bg-gray-200 rounded-full h-2 mb-4">
            <div
              className="bg-primary-600 h-2 rounded-full transition-all duration-300"
              style={{ width: `${progress}%` }}
            ></div>
          </div>
        )}

        {/* 파일 선택 버튼 */}
        {status === 'idle' && (
          <div>
            <input
              ref={fileInputRef}
              type="file"
              accept={SUPPORTED_EXTENSIONS.join(',')}
              onChange={handleFileInputChange}
              className="hidden"
            />
            <button
              onClick={() => fileInputRef.current?.click()}
              className="px-6 py-2 bg-primary-600 text-white rounded-lg hover:bg-primary-700 transition-colors"
            >
              파일 선택
            </button>
          </div>
        )}
      </div>

      {/* 옵션 */}
      {selectedFile && status === 'idle' && (
        <div className="mt-4 p-4 bg-gray-50 rounded-lg space-y-3">
          <label className="flex items-center space-x-2 text-sm">
            <input
              type="checkbox"
              checked={usePythonService}
              onChange={(e) => setUsePythonService(e.target.checked)}
              className="rounded border-gray-300"
            />
            <span className="font-medium">🐍 Python 서비스 사용 (권장)</span>
          </label>
          <p className="text-xs text-gray-600 ml-6">
            더 강력한 Excel 처리 엔진으로 구 Excel 파일과 한글 인코딩을 완벽 지원합니다.
          </p>
          
          {!usePythonService && (
            <label className="flex items-center space-x-2 text-sm">
              <input
                type="checkbox"
                checked={forceTextRecovery}
                onChange={(e) => setForceTextRecovery(e.target.checked)}
                className="rounded border-gray-300"
              />
              <span>텍스트 기반 복구 강제 실행 (표준 파서가 실패할 때 유용)</span>
            </label>
          )}
        </div>
      )}

      {/* 액션 버튼 */}
      {selectedFile && (
        <div className="mt-6 flex gap-3 justify-center">
          {status === 'idle' && (
            <button
              onClick={handleConvert}
              className="px-8 py-3 bg-primary-600 text-white rounded-lg hover:bg-primary-700 transition-colors font-medium"
            >
              변환하기
            </button>
          )}

          {status === 'success' && result?.downloadUrl && (
            <button
              onClick={() => downloadLinkRef.current?.click()}
              className="px-8 py-3 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors font-medium flex items-center gap-2"
            >
              <Download className="w-4 h-4" />
              다시 다운로드
            </button>
          )}

          <button
            onClick={handleReset}
            className="px-6 py-3 bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition-colors font-medium flex items-center gap-2"
          >
            <RotateCcw className="w-4 h-4" />
            다시 시작
          </button>
        </div>
      )}

      {/* 숨겨진 다운로드 링크 */}
      <a ref={downloadLinkRef} className="hidden" />

      {/* 도움말 */}
      <div className="mt-8 p-4 bg-blue-50 rounded-lg">
        <h3 className="font-medium text-blue-900 mb-2">💡 사용 팁</h3>
        <ul className="text-sm text-blue-800 space-y-1">
          <li>• 파일이 열리지 않으면 먼저 표준 변환을 시도해보세요</li>
          <li>• 여전히 문제가 있다면 "텍스트 기반 복구" 옵션을 체크하세요</li>
          <li>• 한글 파일명과 EUC-KR/CP949 인코딩을 자동으로 처리합니다</li>
          <li>• 민감한 정보가 포함된 파일 업로드 시 주의하세요</li>
        </ul>
      </div>
    </div>
  );
}
