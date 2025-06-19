@echo off
echo 엑셀 생성기 EXE 빌드 (수정된 버전)

echo 1. 패키지 설치 중...
pip install openpyxl pyinstaller

echo 2. 기존 빌드 파일 정리...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del *.spec

echo 3. EXE 파일 생성 중...
pyinstaller --onefile --windowed --hidden-import=openpyxl --hidden-import=openpyxl.styles --hidden-import=openpyxl.workbook excel_creator.py

echo 4. 빌드 완료!
if exist "dist\excel_creator.exe" (
    echo 성공: dist\excel_creator.exe 파일이 생성되었습니다.
) else (
    echo 오류: EXE 파일 생성에 실패했습니다.
)

pause