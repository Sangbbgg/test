# build_script.py - 자동 빌드 스크립트
import os
import shutil
import subprocess
import sys

def build_exe():
    print("EXE 빌드 시작...")
    
    # 이전 빌드 파일 정리
    if os.path.exists('build'):
        shutil.rmtree('build')
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    
    # PyInstaller 실행
    cmd = [
        'pyinstaller', 
        '--onefile', 
        '--windowed',
        '--name=엑셀생성기',
        'excel_creator.py'
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode == 0:
        print("빌드 성공!")
        print("실행 파일: dist/엑셀생성기.exe")
    else:
        print("빌드 실패:", result.stderr)
        sys.exit(1)

if __name__ == "__main__":
    build_exe()