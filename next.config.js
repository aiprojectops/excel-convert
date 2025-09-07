/** @type {import('next').NextConfig} */
const nextConfig = {
  experimental: {
    serverComponentsExternalPackages: ['xlsx']
  },
  webpack: (config) => {
    // xlsx 라이브러리 최적화
    config.resolve.fallback = {
      ...config.resolve.fallback,
      fs: false,
      path: false,
      crypto: false,
    };
    return config;
  },
}

module.exports = nextConfig
