/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  
  // Biar pptxgenjs & @google/generative-ai bisa jalan di server
  serverExternalPackages: ['pptxgenjs', '@google/generative-ai'],
  
  // Naikin limit body request (buat prompt panjang)
  experimental: {
    serverActions: {
      bodySizeLimit: '2mb',
    },
  },
};

module.exports = nextConfig;
