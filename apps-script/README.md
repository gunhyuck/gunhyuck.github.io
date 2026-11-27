# 합동 비서실 웹페이지 (카림애비뉴 · 스타타워)

- 카림애비뉴: https://gunhyuck.github.io/karim/
- 스타타워: https://gunhyuck.github.io/startower/
- 데이터: 구글 시트 **피클 합동 비서실 DB**
  https://docs.google.com/spreadsheets/d/1bWeBT0SEThBbFqP_87QkqgclWGYzk62wtM5iRYibu6k/edit

## 구성

| 경로 | 설명 |
| --- | --- |
| `karim/index.html` | 카림애비뉴 합동 비서실 |
| `startower/index.html` | 스타타워 비서실 |
| `assets/office.css` | 두 페이지 공용 스타일 |
| `assets/office.js` | 안건·메모·필터·연동 로직 |
| `assets/data.js` | 초기 스냅샷(시트와 동일). 연동 전에도 화면이 바로 채워집니다 |
| `apps-script/Code.gs` | 구글 시트 양방향 연동용 Apps Script (시트 ID 포함) |

## 구글 시트 연동 (5분)

1. 위 스프레드시트 → **확장 프로그램 → Apps Script**
2. `apps-script/Code.gs` 내용 전체를 붙여넣고 저장
3. **배포 → 새 배포 → 유형: 웹 앱**, 실행 계정 `나`, 액세스 `모든 사용자` → 배포
4. 나오는 `https://script.google.com/macros/s/.../exec` 주소를 페이지 우측 상단 **⚙** 에 붙여넣고 저장
   (또는 `?api=주소` 로 열기)

주소를 넣기 전에는 브라우저(localStorage)에만 저장되고, 넣은 뒤에는
안건 추가 · 상태 변경 · 메모가 시트에 바로 기록되며 60초마다 자동 새로고침됩니다.

## 시트 탭

- `karim_tasks` / `startower_tasks`: id, title, category, owner, status, due, summary, updated_by, updated_at
- `karim_memos` / `startower_memos`: id, author, role, content, created_at
- 상태 값: `대기` `진행중` `검토중` `완료`

## 메모

- 계좌번호·비밀번호·개인 연락처·입주자 분쟁 금액 등 민감정보는 넣지 않습니다. 이 페이지는 공개 주소입니다.
