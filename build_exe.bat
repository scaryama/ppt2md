@echo off
REM PPT to Markdown GUI를 exe 파일로 빌드하는 스크립트

echo ========================================
echo PPT to Markdown GUI 빌드 스크립트
echo ========================================
echo.

REM 가상 환경 활성화
call venv\Scripts\activate.bat

REM PyInstaller 설치 확인 및 설치
echo PyInstaller 설치 확인 중...
pip install pyinstaller --quiet

echo.
echo exe 파일 빌드 시작...
echo.

REM PyInstaller로 빌드
pyinstaller --name="PPT2Markdown" ^
    --windowed ^
    --onefile ^
    --icon=NONE ^
    --add-data "ppt2md.py;." ^
    --hidden-import=PyQt6.QtCore ^
    --hidden-import=PyQt6.QtGui ^
    --hidden-import=PyQt6.QtWidgets ^
    --hidden-import=pptx ^
    --hidden-import=pptx.shapes ^
    --hidden-import=pptx.shapes.base ^
    --hidden-import=pptx.shapes.autoshape ^
    --hidden-import=pptx.shapes.group ^
    --hidden-import=pptx.shapes.picture ^
    --hidden-import=pptx.shapes.table ^
    --hidden-import=pptx.shapes.connector ^
    --hidden-import=pptx.shapes.freeform ^
    --hidden-import=pptx.shapes.ole ^
    --hidden-import=pptx.shapes.placeholder ^
    --hidden-import=pptx.enum.shapes ^
    --hidden-import=pptx.enum.text ^
    --hidden-import=pptx.parts ^
    --hidden-import=pptx.parts.image ^
    --hidden-import=pptx.parts.slide ^
    --hidden-import=pptx.parts.presentation ^
    --hidden-import=pptx.oxml ^
    --hidden-import=pptx.oxml.ns ^
    --hidden-import=pptx.oxml.presentation ^
    --hidden-import=pptx.oxml.slide ^
    --hidden-import=pptx.oxml.shapes ^
    --hidden-import=pptx.oxml.table ^
    --hidden-import=pptx.oxml.text ^
    --hidden-import=pptx.util ^
    --hidden-import=pptx.compat ^
    --hidden-import=pptx.exc ^
    --hidden-import=lxml ^
    --hidden-import=lxml.etree ^
    --hidden-import=lxml._elementpath ^
    --hidden-import=ppt2md ^
    ppt2md_gui.py

echo.
echo ========================================
echo 빌드 완료!
echo exe 파일 위치: dist\PPT2Markdown.exe
echo ========================================
echo.

pause

