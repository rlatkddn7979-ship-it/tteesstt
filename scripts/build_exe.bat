@echo off
REM 두원전자통신 물품 조회 도구 모음을 실행파일(.exe) 하나로 빌드합니다.
REM 이 배치파일을 scripts 폴더 안에서 더블클릭하거나, 이 폴더에서 실행하세요.

pip install -r requirements.txt
if errorlevel 1 (
    echo pip install 실패. 위 오류 메시지를 확인해주세요.
    pause
    exit /b 1
)

pyinstaller --onefile --noconsole --name 두원전자통신물품조회 launcher.py
if errorlevel 1 (
    echo PyInstaller 빌드 실패. 위 오류 메시지를 확인해주세요.
    pause
    exit /b 1
)

echo.
echo 빌드 완료: dist\두원전자통신물품조회.exe
echo 이 exe 파일 하나만 다른 폴더로 복사하거나 다른 PC로 옮겨서 실행하면 됩니다.
pause
