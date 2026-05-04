# MarkItDown GUI

Microsoft MarkItDown 변환 옵션을 테스트하기 위한 간단한 로컬 데스크톱 앱이다.

## 실행 준비

먼저 가상 환경을 만들고 활성화한다.

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
```

## 설치

필요한 기능에 맞춰 패키지를 설치한다.

### 기본 문서 변환

PDF, DOCX, XLSX, PPTX, HTML, CSV, JSON, XML, ZIP, EPUB 같은 일반 문서 변환만 테스트할 때 사용한다.

```powershell
pip install 'markitdown[all]'
```

### LLM 이미지 설명

이미지를 LLM으로 설명하는 기능을 테스트할 때 사용한다.

```powershell
pip install 'markitdown[all]' openai
```

앱 실행 전에 `OPENAI_API_KEY`를 설정한다.

```powershell
$env:OPENAI_API_KEY = "your-api-key"
```

### OCR 플러그인

`markitdown-ocr` 플러그인을 테스트할 때 사용한다.

```powershell
pip install 'markitdown[all]' openai markitdown-ocr
```

`requirements.txt`는 모든 테스트 의존성을 한 번에 설치하고 싶을 때만 사용한다.

```powershell
pip install -r requirements.txt
```

## 실행

```powershell
python app.py
```

## 테스트 항목

- PDF를 Markdown으로 변환
- DOCX 제목과 목록 변환
- XLSX 표 변환
- PPTX 슬라이드 텍스트 추출
- 이미지 메타데이터 또는 텍스트 추출
- 플러그인 활성화 및 개별 플러그인 선택
- `keep_data_uris` 옵션
- 확장자, MIME 타입, 문자 인코딩 힌트
- `OPENAI_API_KEY`를 사용한 LLM 이미지 설명
- exiftool 경로 지정
- DOCX style map 설정

앱은 변환된 Markdown을 선택한 출력 경로에 저장하고, 창 안에서 미리보기를 보여준다.
