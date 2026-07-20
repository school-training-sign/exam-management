# 정기고사 관리 시스템

학교가 소유한 Google Sheets·Apps Script 백엔드와 연결되는 GitHub Pages 프런트엔드입니다. 원본 서비스의 소스나 데이터를 복사하지 않고 교사용 결시 입력, 고사본부, 자리배치, 설정 흐름을 독립 구현했습니다.

참고 서비스에서 확인한 입력 순서와 의도적 차이는 [`source-parity.md`](source-parity.md)에 정리했습니다. 참고 화면과 조사 중 만든 가상 자료는 공개 저장소에 포함하지 않습니다.

## 데모

개발 서버 또는 배포 주소에 `?demo=1`을 붙이면 실제 데이터 없이 전체 화면을 확인할 수 있습니다.

- 학교코드: 아무 값
- 데모 관리자 암호: `demo-admin`

데모 변경 내용은 브라우저를 새로고침하면 초기화되며 Google Sheets에는 쓰지 않습니다.

## 운영 기능

- 결시 명단 불러오기, 사유 편집, revision 충돌을 검사하는 일괄 제출·재수정
- 상단 연결 상태와 현재 화면을 제외한 최근 90초 이내 활성 접속자 수
- 고사일·학년·학급·교시별 고사본부 현황, 제출 상태, 기간·사유·학급·일자·학생별 통계
- 교시 현황과 기간 합산 통계의 분리된 Excel 내보내기
- 학생 개별 수정·비활성화, 명단 교체 시 같은 학급·번호의 학생 ID 유지
- 한 교시의 복수 공통·선택과목, 적용 학급, 학생별 선택과목·호실
- 학년 전체 또는 특정 학급의 개인 시간표 A4 일괄 출력
- 6행×5열 별실·각자교실 자리배치, 가로 순번·세로 순번 선택, 제외 좌석, 결시 표시, 저장·갱신·불러오기·인쇄
- 5분 미사용 로그아웃, API 제한시간·선별 재시도, 관리자 세션 만료

## 학교 백엔드 연결

1. `exam-management-template` 저장소의 배포용 원본 `.xlsx`를 학교 Google Drive에서 Google 스프레드시트로 변환합니다.
2. 같은 저장소의 `apps-script` 파일을 시트에 연결된 Apps Script 프로젝트에 추가합니다.
3. 시트 메뉴 `정기고사 관리 → 초기 설정`에서 학교코드와 관리자 암호를 등록합니다.
4. Apps Script를 웹 앱으로 배포합니다.
5. [`public/assets/config.js`](public/assets/config.js)의 `API_URL`에 웹 앱 주소를 입력하고 다시 배포합니다.

학교코드·관리자 암호·학생 데이터는 공개 저장소에 커밋하지 않습니다.

관리자 암호를 바꾸면 이전 관리자 세션이 모두 만료됩니다. 운영 화면은 세션 토큰만 `sessionStorage`에 보관하며, 데모 모드는 운영 API를 호출하지 않습니다.

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
