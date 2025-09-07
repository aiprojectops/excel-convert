import Head from 'next/head';
import FileUploader from '@/components/FileUploader';

export default function Home() {
  return (
    <>
      <Head>
        <title>Excel Quick Convert - 구엑셀 빠른 변환</title>
        <meta 
          name="description" 
          content="구엑셀(.xls), 한컴오피스, 깨진 Excel 파일을 최신 .xlsx 형식으로 빠르게 변환하는 무료 웹 서비스" 
        />
        <meta name="keywords" content="엑셀변환, xls, xlsx, 한컴오피스, 파일변환, Excel converter" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <link rel="icon" href="/favicon.ico" />
        
        {/* Open Graph */}
        <meta property="og:title" content="Excel Quick Convert - 구엑셀 빠른 변환" />
        <meta property="og:description" content="구엑셀, 한컴오피스 파일을 최신 Excel로 무료 변환" />
        <meta property="og:type" content="website" />
        
        {/* Twitter Card */}
        <meta name="twitter:card" content="summary_large_image" />
        <meta name="twitter:title" content="Excel Quick Convert" />
        <meta name="twitter:description" content="구엑셀 파일을 최신 Excel로 빠르게 변환" />
      </Head>

      <main className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100">
        {/* 헤더 */}
        <header className="bg-white shadow-sm">
          <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4">
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-3">
                <div className="w-8 h-8 bg-primary-600 rounded-lg flex items-center justify-center">
                  <span className="text-white font-bold text-sm">EC</span>
                </div>
                <h1 className="text-xl font-bold text-gray-900">Excel Quick Convert</h1>
              </div>
              <nav className="hidden sm:flex space-x-4">
                <a 
                  href="#features" 
                  className="text-gray-600 hover:text-gray-900 transition-colors"
                >
                  기능
                </a>
                <a 
                  href="#guide" 
                  className="text-gray-600 hover:text-gray-900 transition-colors"
                >
                  사용법
                </a>
                <a 
                  href="https://github.com" 
                  target="_blank" 
                  rel="noopener noreferrer"
                  className="text-gray-600 hover:text-gray-900 transition-colors"
                >
                  GitHub
                </a>
              </nav>
            </div>
          </div>
        </header>

        {/* 메인 컨텐츠 */}
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-12">
          <FileUploader />
        </div>

        {/* 기능 소개 */}
        <section id="features" className="bg-white py-16">
          <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div className="text-center mb-12">
              <h2 className="text-3xl font-bold text-gray-900 mb-4">
                왜 Excel Quick Convert인가요?
              </h2>
              <p className="text-lg text-gray-600 max-w-2xl mx-auto">
                구버전 Excel, 한컴오피스, 깨진 파일까지 모두 최신 Excel에서 열리도록 변환합니다
              </p>
            </div>

            <div className="grid md:grid-cols-3 gap-8">
              <div className="text-center p-6">
                <div className="w-16 h-16 bg-primary-100 rounded-full flex items-center justify-center mx-auto mb-4">
                  <svg className="w-8 h-8 text-primary-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 10V3L4 14h7v7l9-11h-7z" />
                  </svg>
                </div>
                <h3 className="text-xl font-semibold text-gray-900 mb-2">빠른 변환</h3>
                <p className="text-gray-600">
                  업로드부터 다운로드까지 단 몇 초. 복잡한 설치나 설정 없이 바로 사용 가능합니다.
                </p>
              </div>

              <div className="text-center p-6">
                <div className="w-16 h-16 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
                  <svg className="w-8 h-8 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                  </svg>
                </div>
                <h3 className="text-xl font-semibold text-gray-900 mb-2">높은 호환성</h3>
                <p className="text-gray-600">
                  구 Excel(.xls), 한컴오피스, CSV, TSV 등 다양한 형식을 지원합니다.
                </p>
              </div>

              <div className="text-center p-6">
                <div className="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center mx-auto mb-4">
                  <svg className="w-8 h-8 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z" />
                  </svg>
                </div>
                <h3 className="text-xl font-semibold text-gray-900 mb-2">안전한 처리</h3>
                <p className="text-gray-600">
                  파일은 변환 후 자동으로 삭제되며, 브라우저에서 직접 처리되어 안전합니다.
                </p>
              </div>
            </div>
          </div>
        </section>

        {/* 사용 가이드 */}
        <section id="guide" className="bg-gray-50 py-16">
          <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div className="text-center mb-12">
              <h2 className="text-3xl font-bold text-gray-900 mb-4">
                사용법
              </h2>
              <p className="text-lg text-gray-600">
                3단계로 간단하게 파일을 변환하세요
              </p>
            </div>

            <div className="grid md:grid-cols-3 gap-8">
              <div className="bg-white p-6 rounded-xl shadow-sm">
                <div className="w-12 h-12 bg-primary-600 text-white rounded-lg flex items-center justify-center font-bold text-xl mb-4">
                  1
                </div>
                <h3 className="text-xl font-semibold text-gray-900 mb-2">파일 업로드</h3>
                <p className="text-gray-600">
                  변환하고 싶은 Excel 파일을 드래그하거나 클릭하여 선택합니다.
                </p>
              </div>

              <div className="bg-white p-6 rounded-xl shadow-sm">
                <div className="w-12 h-12 bg-primary-600 text-white rounded-lg flex items-center justify-center font-bold text-xl mb-4">
                  2
                </div>
                <h3 className="text-xl font-semibold text-gray-900 mb-2">변환 실행</h3>
                <p className="text-gray-600">
                  "변환하기" 버튼을 클릭하면 자동으로 최신 Excel 형식으로 변환됩니다.
                </p>
              </div>

              <div className="bg-white p-6 rounded-xl shadow-sm">
                <div className="w-12 h-12 bg-primary-600 text-white rounded-lg flex items-center justify-center font-bold text-xl mb-4">
                  3
                </div>
                <h3 className="text-xl font-semibold text-gray-900 mb-2">다운로드</h3>
                <p className="text-gray-600">
                  변환이 완료되면 자동으로 다운로드되며, 최신 Excel에서 바로 열 수 있습니다.
                </p>
              </div>
            </div>
          </div>
        </section>

        {/* 푸터 */}
        <footer className="bg-white border-t py-8">
          <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div className="flex flex-col md:flex-row justify-between items-center">
              <div className="flex items-center space-x-3 mb-4 md:mb-0">
                <div className="w-6 h-6 bg-primary-600 rounded flex items-center justify-center">
                  <span className="text-white font-bold text-xs">EC</span>
                </div>
                <span className="text-gray-600">Excel Quick Convert</span>
              </div>
              
              <div className="flex space-x-6 text-sm text-gray-500">
                <span>© 2024 Excel Quick Convert</span>
                <a href="#" className="hover:text-gray-700">개인정보처리방침</a>
                <a href="#" className="hover:text-gray-700">이용약관</a>
              </div>
            </div>
          </div>
        </footer>
      </main>
    </>
  );
}
