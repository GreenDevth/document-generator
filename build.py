import os
import re

def bundle():
    # สร้างโฟลเดอร์ dist หากยังไม่มี
    if not os.path.exists('dist'):
        os.makedirs('dist')

    # อ่านไฟล์ต้นฉบับ
    try:
        with open('index.html', 'r', encoding='utf-8') as f:
            html = f.read()
        
        with open('style.css', 'r', encoding='utf-8') as f:
            css = f.read()

        with open('script.js', 'r', encoding='utf-8') as f:
            js = f.read()

        # 1. แทนที่ <link rel="stylesheet" href="style.css"> ด้วย <style>...</style>
        css_bundle = f'<style>\n{css}\n</style>'
        html = re.sub(r'<link rel="stylesheet" href="style.css">', css_bundle, html)

        # 2. แทนที่ <script src="script.js" defer></script> ด้วย <script>...</script>
        js_bundle = f'<script>\n{js}\n</script>'
        html = re.sub(r'<script src="script.js" defer></script>', js_bundle, html)

        # บันทึกไฟล์ผลลัพธ์ไปที่ dist/index.html
        with open('dist/index.html', 'w', encoding='utf-8') as f:
            f.write(html)

        print("✅ Build Success! ไฟล์รวมร่างสำหรับ GAS อยู่ที่: dist/index.html")
        print("🌐 ส่วนไฟล์ด้านนอก (index.html, style.css, script.js) พร้อมใช้งานบน GitHub Pages แล้วครับ!")

    except FileNotFoundError as e:
        print(f"❌ Error: ไม่พบไฟล์ {e.filename} กรุณาตรวจสอบว่าอยู่ในโฟลเดอร์ที่ถูกต้อง")

if __name__ == "__main__":
    bundle()
