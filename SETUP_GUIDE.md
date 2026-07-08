# 🔧 로펌 비용 대시보드 — 세팅 가이드

> 비개발자 기준으로 작성했습니다. 처음부터 끝까지 따라하시면 됩니다.
> 전체 소요시간: 약 30~40분

---

## 전체 구조 (먼저 이해하기)

```
[구글 드라이브]          [GitHub]              [Streamlit Cloud]
 인보이스 PDF 저장  →   코드 저장소    →     웹 대시보드 실행
 (지금처럼 사용)       (1회 업로드)         (자동 배포 + URL 생성)
```

유진님이 평소처럼 구글 드라이브에 인보이스를 올리면, 대시보드 URL에 접속할 때마다 자동으로 새 파일을 읽어서 집계합니다.


---

## STEP 1: GitHub 계정 만들기 (5분)

1. https://github.com 접속
2. 우측 상단 [Sign up] 클릭
3. 이메일, 비밀번호, 사용자명 입력 (사용자명은 영문, 예: `dw-legal-yj`)
4. 이메일 인증 완료


---

## STEP 2: GitHub에 코드 올리기 (5분)

1. GitHub 로그인 후 우측 상단 [+] → [New repository] 클릭
2. 설정:
   - Repository name: `legal-invoice-dashboard`
   - ⚠️ **반드시 Private 선택** (공개하면 안 됩니다)
   - [Create repository] 클릭

3. 생성된 페이지에서 [uploading an existing file] 링크 클릭

4. 제가 만들어드린 파일 3개를 드래그 앤 드롭:
   - `app.py`
   - `requirements.txt`
   - `.streamlit/config.toml` (이 파일은 없어도 동작합니다)

5. 하단 [Commit changes] 클릭


---

## STEP 3: 구글 서비스 계정 만들기 (10분)

> 이 단계가 가장 복잡합니다. 천천히 따라하세요.

### 3-1. Google Cloud Console 접속
1. https://console.cloud.google.com 접속 (개인 Gmail: dw.syj0826@gmail.com 으로 로그인)
2. 상단 바에서 프로젝트 선택 드롭다운 클릭 → [새 프로젝트]
3. 프로젝트 이름: `legal-dashboard` → [만들기]

### 3-2. Google Drive API 활성화
1. 왼쪽 메뉴 → [API 및 서비스] → [라이브러리]
2. 검색창에 "Google Drive API" 입력
3. [Google Drive API] 클릭 → [사용] 클릭

### 3-3. 서비스 계정 생성
1. [API 및 서비스] → [사용자 인증 정보]
2. 상단 [+ 사용자 인증 정보 만들기] → [서비스 계정]
3. 서비스 계정 이름: `dashboard-reader` → [완료]
4. 생성된 서비스 계정 이메일이 표시됩니다 (예: `dashboard-reader@legal-dashboard.iam.gserviceaccount.com`)
   → **이 이메일을 복사해 두세요**

### 3-4. JSON 키 다운로드
1. 서비스 계정 목록에서 방금 만든 계정 클릭
2. [키] 탭 → [키 추가] → [새 키 만들기]
3. [JSON] 선택 → [만들기]
4. JSON 파일이 자동 다운로드됩니다 → **이 파일을 잘 보관하세요**


---

## STEP 4: 구글 드라이브 폴더 공유 (2분)

> 서비스 계정이 인보이스 폴더를 읽을 수 있도록 권한을 줘야 합니다.

1. 구글 드라이브에서 인보이스 상위 폴더 (광장, 김앤장 등이 들어있는 폴더) 우클릭
2. [공유] → [공유]
3. STEP 3-3에서 복사한 서비스 계정 이메일 붙여넣기
4. 권한: **뷰어** (보기 전용)
5. [보내기] 클릭

⚠️ 주의: "뷰어"만 선택하세요. 편집자 권한을 주면 안 됩니다.


---

## STEP 5: Streamlit Cloud 설정 (10분)

### 5-1. Streamlit Cloud 가입
1. https://share.streamlit.io 접속
2. [Continue with GitHub] 로 로그인 (STEP 1에서 만든 GitHub 계정)
3. 권한 허용

### 5-2. 앱 배포
1. [New app] 클릭
2. 설정:
   - Repository: `(본인 사용자명)/legal-invoice-dashboard`
   - Branch: `main`
   - Main file path: `app.py`
3. [Advanced settings] 클릭 → **Secrets** 탭

### 5-3. Secrets 입력 (가장 중요!)

Secrets 입력창에 아래 내용을 붙여넣으세요:

```toml
ROOT_FOLDER_ID = "1HleAj4z6DH9KxMjyuf56lRD4b-c7pOyh"

GOOGLE_SERVICE_ACCOUNT = '''
여기에 STEP 3-4에서 다운로드한 JSON 파일 내용 전체를 붙여넣으세요
'''
```

JSON 파일 내용 붙여넣는 방법:
1. 다운로드된 JSON 파일을 메모장(또는 텍스트 편집기)으로 열기
2. 전체 내용 복사 (Ctrl+A → Ctrl+C)
3. 위 `'''` 사이에 붙여넣기

### 5-4. 배포
1. [Deploy!] 클릭
2. 1~2분 후 대시보드 URL이 생성됩니다
3. 이 URL을 북마크하세요!


---

## 완료 후 사용 방법

### 평소 (매월 반복)
1. 인보이스 PDF를 구글 드라이브 해당 로펌 폴더에 업로드 (지금처럼)
2. 대시보드 URL 접속 (북마크)
3. 자동으로 새 파일 인식 → 집계 업데이트

### 새로고침이 필요할 때
- 사이드바의 [🔄 새로고침] 버튼 클릭
- 데이터는 5분 단위로 캐시되어 있어서, 파일 업로드 직후에는 새로고침을 눌러주세요

### 절전 모드
- 며칠간 아무도 접속하지 않으면 앱이 절전 모드에 들어갑니다
- 접속하면 "This app is waking up" 메시지 → 20~30초 후 정상 로드


---

## 자주 묻는 질문

**Q: 로펌이 추가되면?**
A: 구글 드라이브에 새 폴더를 만들고 인보이스를 넣으면 자동 인식됩니다. 코드 수정 불필요.

**Q: 금액이 잘못 추출되면?**
A: 사이드바의 "파싱 이슈"를 확인하고, Claude에게 해당 로펌 PDF를 보여주며 수정 요청하세요.

**Q: 보안은?**
A: GitHub Private 저장소 + 서비스 계정(읽기 전용) + Streamlit Cloud(GitHub 인증). 정성욱님 대시보드의 보안 문제(Public 저장소, 이메일 노출, 접근 통제 없음)가 모두 해결된 구조입니다.

**Q: 대시보드 URL을 다른 팀원에게 공유해도 되나요?**
A: 네. 다만 URL만 알면 누구나 볼 수 있으므로, 팀 내부에서만 공유하세요. 접근 제한을 추가하려면 Streamlit의 인증 기능 설정이 필요합니다 (추후 안내 가능).


---

## 문제 발생 시 체크리스트

| 증상 | 확인 사항 |
|---|---|
| "인보이스를 찾을 수 없습니다" | 서비스 계정에 폴더 공유했는지 확인 |
| 특정 로펌만 안 나옴 | 해당 폴더가 상위 폴더 안에 있는지 확인 |
| 금액이 0 또는 이상 | "파싱 이슈" 확인 → PDF 형식 변경 가능성 |
| 앱이 안 열림 | GitHub 저장소에 app.py 있는지 확인 |
| "waking up" 계속 뜸 | 1~2분 대기, 안 되면 Streamlit Cloud에서 Reboot |
