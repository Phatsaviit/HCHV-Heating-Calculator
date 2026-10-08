"""
สคริปต์คอมไพล์โปรแกรม Heating Cycle Analyzer เป็นไฟล์ .exe และบีบอัดเป็น ZIP อัตโนมัติ
"""
import os
import shutil
import subprocess
import zipfile

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))

print("1. กำลังคอมไพล์โปรแกรมด้วย PyInstaller...")
cmd = [
    "pyinstaller",
    "--noconfirm",
    "--onefile",
    "--windowed",
    "--name", "HeatingCycleAnalyzer",
    os.path.join(PROJECT_DIR, "cycle_analyzer.py")
]
subprocess.run(cmd, check=True, cwd=PROJECT_DIR)

print("2. จัดเตรียมโฟลเดอร์ Release...")
release_dir = os.path.join(PROJECT_DIR, "release", "HeatingCycleAnalyzer")
if os.path.exists(release_dir):
    shutil.rmtree(release_dir)
os.makedirs(release_dir, exist_ok=True)

# Copy executable
exe_src = os.path.join(PROJECT_DIR, "dist", "HeatingCycleAnalyzer.exe")
shutil.copy2(exe_src, release_dir)

# Copy sample excel if present
sample_excel = os.path.join(PROJECT_DIR, "HCHV 115 kV.xlsx")
if os.path.exists(sample_excel):
    shutil.copy2(sample_excel, release_dir)

# Readme
readme_content = """========================================================================
โปรแกรมสรุปผลการทดสอบ Heating Cycle Voltage Test (IEC 60840 / 60502-2)
========================================================================

วิธีใช้งาน:
1. ดับเบิลคลิกที่ไฟล์ 'HeatingCycleAnalyzer.exe' เพื่อเปิดโปรแกรมทันที
2. กดปุ่ม 'เลือกไฟล์...' เพื่อเลือกไฟล์ Excel รายงานผลการทดสอบ (เช่น HCHV 115 kV.xlsx)
3. เลือกมาตรฐานการทดสอบ:
   - ตรวจจับอัตโนมัติ (Auto-detect จากค่าแรงดันทดสอบในไฟล์)
   - สายไฟฟ้าแรงสูง HV (รอบละ 24 ชม. - ร้อน 8 ชม. / ระบาย 16 ชม.)
   - สายไฟฟ้าแรงดันปานกลาง MV (ร้อน 8 ชม. / ระบายตามธรรมชาติ)
4. กดปุ่ม '[ เริ่มประมวลผล และแทรก Sheet สรุปใน Excel ]'
5. ระบบจะเพิ่มชีต 'Cycle_Summary' เข้าไปเป็นหน้าแรกของไฟล์ Excel โดยอัตโนมัติ
   พร้อมสรุปผลเวลา, สถานะผ่าน/ไม่ผ่าน, และค่ากระแสเฉลี่ย (I hold avg, I ramp avg, I heat avg)
6. สามารถกดปุ่ม 'เปิดดูไฟล์ Excel' เพื่อเปิดตรวจผลลัพธ์ได้ทันที

หมายเหตุ:
- โปรแกรมนี้สามารถนำไปใช้งานบนเครื่องคอมพิวเตอร์ Windows เครื่องใดก็ได้
  โดยไม่จำเป็นต้องติดตั้ง Python หรือโปรแกรมเสริมอื่นๆ เพิ่มเติม
========================================================================
"""
with open(os.path.join(release_dir, "คำแนะนำการใช้งาน.txt"), "w", encoding="utf-8") as f:
    f.write(readme_content)

print("3. กำลังสร้างไฟล์ ZIP สำหรับนำไปแตกไฟล์ใช้งาน...")
zip_path = os.path.join(PROJECT_DIR, "release", "HeatingCycleAnalyzer_Portable.zip")
with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
    for root, dirs, files in os.walk(release_dir):
        for file in files:
            full_path = os.path.join(root, file)
            arcname = os.path.relpath(full_path, release_dir)
            zipf.write(full_path, arcname)

print("สำเร็จเรียบร้อย!")
print(f"• ไฟล์ .exe: {os.path.join(release_dir, 'HeatingCycleAnalyzer.exe')}")
print(f"• ไฟล์ .zip: {zip_path}")
