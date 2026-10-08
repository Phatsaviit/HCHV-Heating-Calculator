"""
สคริปต์คอมไพล์โปรแกรม Heating Cycle Analyzer เป็นไฟล์ .exe และบีบอัดเป็น ZIP อัตโนมัติ
"""
import os
import shutil
import subprocess
import zipfile

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))

# Kill any existing instance if running
subprocess.run(["taskkill", "/f", "/im", "HeatingCycleAnalyzer.exe"], capture_output=True)

print("1. กำลังคอมไพล์โปรแกรมด้วย PyInstaller...")
cmd = [
    "pyinstaller",
    "--clean",
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

# Copy readme / user guide
readme_src = os.path.join(PROJECT_DIR, "คำแนะนำการใช้งาน.txt")
if os.path.exists(readme_src):
    shutil.copy2(readme_src, release_dir)

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
