# Institution Total Report (기관 순매수 상위 종목 집계)

한국투자증권 API를 사용하여 기관 순매수 상위 종목을 집계하고 텔레그램으로 결과를 전송하는 프로젝트입니다.

## 기능

- 한국투자증권 API를 통한 기관 순매수 상위 종목 조회
- 이미지 형태로 결과 저장
- 텔레그램으로 결과 전송
- 웹 서버 API로 결과 전송

## 설치 방법

1. 필수 라이브러리 설치:
```
pip install requests python-dotenv pandas imgkit pillow
```
또는 requirements.txt를 사용하여 설치:
```
pip install -r requirements.txt
```

2. wkhtmltoimage 설치:
   - Windows: https://wkhtmltopdf.org/downloads.html
   - Linux: `sudo apt-get install wkhtmltopdf`
   - Mac: `brew install wkhtmltopdf`

3. 환경 변수 설정:
   - `.env.sample` 파일을 복사하여 `.env` 파일 생성
   - 필요한 API 키와 설정 값 입력
```
cp .env.sample .env
# 이후 .env 파일을 열어 필요한 값을 입력
```

4. 토큰 파일 설정 (선택 사항):
   - 초기 실행 시 자동으로 생성되므로 필수는 아님
   - 필요한 경우 `token.json.sample` 파일을 복사하여 `token.json` 생성
```
cp token.json.sample token.json
# 필요시 token.json 파일 수정
```

## 보안 주의사항

이 프로젝트는 다음과 같은 민감한 정보를 사용합니다:
- 텔레그램 봇 토큰
- 한국투자증권 API 키 및 시크릿
- API 액세스 토큰

**중요**: 
- `.env` 및 `token.json` 파일은 GitHub에 올리지 마세요. (`.gitignore`에 이미 포함됨)
- 샘플 파일(`.env.sample`, `token.json.sample`)만 버전 관리에 포함하세요.
- 실제 API 키는 로컬에만 저장하고 안전하게 관리하세요.

## 사용 방법

```
python main.py
```

## 디렉토리 구조

- `/img`: 생성된 이미지 저장 디렉토리
- `/utils`: 유틸리티 함수들 (API, 텔레그램, 로깅)
- `/log`: 로그 파일 저장 디렉토리
- `.env.sample`: 환경 변수 샘플 파일
- `token.json.sample`: 토큰 정보 샘플 파일