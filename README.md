# 정기고사 관리 시스템

학교가 소유한 Google Sheets·Apps Script 백엔드와 연결되는 GitHub Pages 프런트엔드입니다. 원본 서비스의 소스나 데이터를 복사하지 않고 교사용 결시 입력, 고사본부, 자리배치, 설정 흐름을 독립 구현했습니다.

## 데모

개발 서버 또는 배포 주소에 `?demo=1`을 붙이면 실제 데이터 없이 전체 화면을 확인할 수 있습니다.

- 학교코드: 아무 값
- 데모 관리자 암호: `demo-admin`

데모 변경 내용은 브라우저를 새로고침하면 초기화되며 Google Sheets에는 쓰지 않습니다.

## 학교 백엔드 연결

1. `exam-management-template` 저장소의 배포용 원본 `.xlsx`를 학교 Google Drive에서 Google 스프레드시트로 변환합니다.
2. 같은 저장소의 `apps-script` 파일을 시트에 연결된 Apps Script 프로젝트에 추가합니다.
3. 시트 메뉴 `정기고사 관리 → 초기 설정`에서 학교코드와 관리자 암호를 등록합니다.
4. Apps Script를 웹 앱으로 배포합니다.
5. [`public/assets/config.js`](public/assets/config.js)의 `API_URL`에 웹 앱 주소를 입력하고 다시 배포합니다.

학교코드·관리자 암호·학생 데이터는 공개 저장소에 커밋하지 않습니다.

## 로컬 실행

```bash
pnpm install
pnpm dev
```

검증:

```bash
pnpm test
pnpm check
pnpm build
```

GitHub Pages는 검증된 `dist` 결과를 `gh-pages` 브랜치에 게시하는 방식으로 운영합니다.

## 오픈소스

- React, Vite
- Tabler Icons
- Pretendard, Noto Serif KR
- SheetJS Community Edition (`public/vendor/xlsx-LICENSE.txt`)
