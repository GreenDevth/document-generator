import { defineConfig } from 'vite';
import { viteSingleFile } from 'vite-plugin-singlefile';

// ตั้งค่าโครงการสำหรับ Vite
export default defineConfig({
  plugins: [
    viteSingleFile() // รวมไฟล์ CSS และ JS ทั้งหมดลงในไฟล์ HTML ผลลัพธ์เพียงไฟล์เดียว
  ],
  base: './', // ใช้ relative path สำหรับการดึงข้อมูลบน GitHub Pages
  build: {
    target: 'es2015', // เพื่อรองรับเบราว์เซอร์และสภาพแวดล้อมที่หลากหลาย
    assetsInlineLimit: 100000000, // บังคับให้ asset ทุกตัวถูกฝังลงในไฟล์ HTML (inline)
    cssCodeSplit: false, // ป้องกันการแยกไฟล์ CSS ออกจาก HTML
    rollupOptions: {
      output: {
        inlineDynamicImports: true // บังคับให้ JS dynamic imports ถูกรวมอยู่ในไฟล์เดียวกัน
      }
    }
  }
});
