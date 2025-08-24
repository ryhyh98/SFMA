
# 스마트공장 수준 진단 프로그램

## 1. 프로젝트 개요

본 프로젝트는 웹 기반 설문지를 통해 기업의 스마트공장 수준을 진단하고, 결과를 분석하여 리포트(차트, 표 포함)를 제공하는 애플리케이션입니다. 진단 결과는 PDF 또는 Excel 파일로 저장할 수 있습니다.

## 2. 기술 스택

*   **백엔드:** Python, Flask
*   **프론트엔드:** HTML, CSS, JavaScript, Bootstrap, Chart.js
*   **데이터 처리:** Pandas, openpyxl
*   **파일 생성:** fpdf2, matplotlib
*   **빌드:** PyInstaller

## 3. 개발 환경 설정

1.  **Python 설치:** Python 3.11 이상을 설치합니다.
2.  **의존성 설치:** 프로젝트 루트 디렉토리에서 다음 명령어를 실행하여 필요한 라이브러리를 모두 설치합니다.

    ```bash
    pip install -r backend/requirements.txt
    ```

## 4. 애플리케이션 실행 (개발 모드)

1.  다음 명령어를 실행하여 Flask 개발 서버를 시작합니다.

    ```bash
    python backend/app.py
    ```

2.  웹 브라우저에서 `http://127.0.0.1:5000` 주소로 접속합니다.

## 5. 독립 실행 파일 빌드 (Build)

`PyInstaller`를 사용하여 웹 애플리케이션을 하나의 독립적인 실행 파일(.exe)로 빌드할 수 있습니다.

1.  **PyInstaller 설치:**

    ```bash
    pip install pyinstaller
    ```

2.  **빌드 실행:**
    프로젝트 루트 디렉토리에서 다음 명령어를 실행합니다. 이 명령어는 `backend/app.py`를 메인 스크립트로 하여, 필요한 모든 데이터 파일(HTML, CSV, 폰트 등)을 포함시키고, 콘솔 창 없이(`--noconsole`) `SFactoryDiagnosis.exe`라는 이름의 단일 실행 파일을 생성합니다.

    ```bash
    pyinstaller --onefile --noconsole --name SFactoryDiagnosis --add-data "index.html;." --add-data "스마트팩토리수준진단_input.csv;." --add-data "backend\NanumGothic.ttf;backend"
    ```
    
    **참고:** `--add-data` 옵션의 구분자는 운영체제에 따라 다를 수 있습니다. Windows에서는 `;`를, macOS/Linux에서는 `:`를 사용합니다. 위 명령어는 Windows 기준입니다.

3.  **빌드 결과 확인:**
    빌드가 성공적으로 완료되면, 프로젝트 루트 디렉토리에 `dist` 폴더가 생성되고, 그 안에 `SFactoryDiagnosis.exe` 파일이 만들어집니다.

## 6. 빌드된 애플리케이션 실행

`dist` 폴더 안에 있는 `SFactoryDiagnosis.exe` 파일을 더블클릭하여 실행합니다. 잠시 후 자동으로 웹 브라우저가 열리면서 애플리케이션이 실행됩니다. (만약 자동으로 열리지 않으면, 웹 브라우저에서 `http://127.0.0.1:5000` 주소로 접속하세요.)
