"""
⚖️ 법무팀 외부 로펌 비용 대시보드
대웅제약 법무1팀 · 구글 드라이브 인보이스 자동 집계

구조: Google Drive(인보이스 PDF/xlsx) → 자동 파싱 → Streamlit 대시보드
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import pdfplumber
import openpyxl
import io
import re
import json

# ============================================================
# 페이지 설정
# ============================================================
st.set_page_config(
    page_title="정기자문 비용 현황",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# 상수
# ============================================================

# 루트 폴더 ID (Streamlit secrets에서 가져오거나 기본값 사용)
ROOT_FOLDER_ID = st.secrets.get("ROOT_FOLDER_ID", "1HleAj4z6DH9KxMjyuf56lRD4b-c7pOyh")

# 로펌별 차트 색상 (파스텔 톤, 통일감 있게)
FIRM_COLORS = {
    "광장": "#7BA7D9",       # 파스텔 블루
    "김앤장": "#E8998D",     # 파스텔 코랄
    "지평": "#A3C9A8",       # 파스텔 그린
    "D&A": "#B8A0D9",        # 파스텔 퍼플
    "율촌": "#F0C987",       # 파스텔 오렌지
    "율촌(영문)": "#E8D687",  # 파스텔 옐로우
    "SLP": "#9EC5D9",        # 파스텔 스카이
    "DLS": "#D9A87B",        # 파스텔 캐러멜
    "세종": "#A0D4C8",       # 파스텔 민트
    "세종_인도네시아": "#8FC7B9",  # 파스텔 틸
}

# 인보이스 파일 판별 키워드
INVOICE_KEYWORDS = ["자문료", "인보이스", "invoice", "청구", "보수금"]
EXCLUDE_KEYWORDS = ["할인 전", "할인전", "이체확인증", "납부확인서", "실비"]

# 폴더명 → 표시명 매핑
FIRM_DISPLAY_NAMES = {
    "율촌_일반 자문업무 (정성무, 유일한, 안유선)": "율촌",
    "율촌_일반 자문업무 (영문계약, 인니)": "율촌(영문)",
}


def extract_file_date(filename):
    """파일명에서 6자리 날짜(YYMMDD) 추출 → 정수 반환. 없으면 0."""
    m = re.search(r"(?<!\d)(2[56]\d{4})(?!\d)", filename)
    if m:
        return int(m.group(1))
    return 0


def extract_month_from_pdf_text(text):
    """
    PDF 본문에서 자문 수행 월 추출.
    광장: "보수금(2026 년 4 월)"
    율촌: "2026-05-01부터 2026-05-31까지"
    D&A: "1월 법률자문료" 또는 "(2026. 1.)"
    지평/김앤장/SLP: 유사 패턴
    """
    if not text:
        return None
    t = re.sub(r"\s+", " ", text)

    # 광장: "보수금(2026년 4월)" 또는 "보수금(2026 년 4 월)"
    m = re.search(r"보수금\s*\(\s*(202[4-9])\s*년\s*(\d{1,2})\s*월\s*\)", t)
    if m:
        return (int(m.group(1)), int(m.group(2)))

    # 율촌: "2026-05-01부터 2026-05-31까지"
    m = re.search(r"(202[4-9])-(\d{2})-\d{2}\s*부터\s*(202[4-9])-(\d{2})-\d{2}", t)
    if m:
        return (int(m.group(1)), int(m.group(2)))

    # "N월 법률자문료" / "N월 자문료" (D&A 스타일)
    m = re.search(r"(\d{1,2})\s*월\s*(?:법률\s*)?자문료", t)
    if m:
        return (2026, int(m.group(1)))

    # "(2026. N.)" / "(2026.N.)" (D&A Description Of Services 스타일)
    m = re.search(r"\(\s*(202[4-9])\.\s*(\d{1,2})\.\s*\)", t)
    if m:
        return (int(m.group(1)), int(m.group(2)))

    # 김앤장: "2026 년 5 월 1 일부터 2026 년 5 월 31 일까지"
    m = re.search(r"(202[4-9])\s*년\s*(\d{1,2})\s*월\s*\d+\s*일\s*부터", t)
    if m:
        return (int(m.group(1)), int(m.group(2)))

    # 지평/일반: "2026년 N월"
    m = re.search(r"(202[4-9])\s*년\s*(\d{1,2})\s*월", t)
    if m:
        return (int(m.group(1)), int(m.group(2)))

    return None


def prev_month(period_str):
    """(사용 안 함) '2026-01' → '2025-12'"""
    y, m = period_str.split("-")
    y, m = int(y), int(m)
    if m == 1:
        return f"{y-1}-12"
    return f"{y}-{m-1:02d}"


# ============================================================
# Google Drive 인증
# ============================================================

@st.cache_resource
def get_drive_service():
    """Google Drive API 서비스 생성 (서비스 계정 인증)"""
    creds_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    creds = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    return build("drive", "v3", credentials=creds)


# ============================================================
# Google Drive 파일 조작
# ============================================================

def list_folder(service, folder_id):
    """폴더 내 모든 파일/하위폴더 목록"""
    items = []
    page_token = None
    while True:
        resp = (
            service.files()
            .list(
                q=f"'{folder_id}' in parents and trashed = false",
                fields="nextPageToken, files(id, name, mimeType, webViewLink)",
                pageToken=page_token,
                orderBy="name",
            )
            .execute()
        )
        items.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return items


def download_file(service, file_id):
    """파일 바이트 다운로드"""
    req = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    buf.seek(0)
    return buf


# ============================================================
# 파일 선택 로직
# ============================================================

def _has_keyword(filename):
    """인보이스 키워드 포함 여부"""
    low = filename.lower()
    return any(kw.lower() in low for kw in INVOICE_KEYWORDS)


def _is_excluded(filename):
    """제외 대상 여부"""
    return any(kw in filename for kw in EXCLUDE_KEYWORDS)


def select_invoices(files):
    """
    파일 목록에서 인보이스 후보 반환.
    '할인 전' 등 제외 키워드만 필터링하고, 모든 유효한 PDF/xlsx를 반환.
    같은 월+로펌 중복은 이후 파일명 날짜 기준으로 제거됨.
    """
    pdfs = [f for f in files if f["name"].lower().endswith(".pdf") and not _is_excluded(f["name"])]
    xlsxs = [f for f in files if f["name"].lower().endswith(".xlsx") and not _is_excluded(f["name"])]

    # PDF가 있으면 PDF, 없으면 xlsx
    if pdfs:
        return pdfs
    return xlsxs


# ============================================================
# 파일명에서 청구 월 추출
# ============================================================

def extract_month(filename):
    """파일명/폴더명 → (year, month) 또는 None"""

    # "N월" / "N월분"
    m = re.search(r"(\d{1,2})\s*월", filename)
    if m:
        month = int(m.group(1))
        y = re.search(r"(202[4-9])", filename)
        return (int(y.group(1)) if y else 2026, month)

    # "2026.06" / "2026. 06"
    m = re.search(r"(202[4-9])[.\s]+(\d{2})", filename)
    if m:
        return (int(m.group(1)), int(m.group(2)))

    # "(2604)" — DLS 스타일
    m = re.search(r"\(26(\d{2})\)", filename)
    if m:
        return (2026, int(m.group(1)))

    # "26.04" — 율촌 스타일
    m = re.search(r"(?<!\d)26\.(\d{2})", filename)
    if m:
        return (2026, int(m.group(1)))

    # "26MMDDDD" — 율촌 ID 스타일 (B-26051222)
    m = re.search(r"26(\d{2})\d{4}", filename)
    if m:
        return (2026, int(m.group(1)))

    # "20260520"
    m = re.search(r"(202[4-9])(\d{2})\d{2}", filename)
    if m:
        return (int(m.group(1)), int(m.group(2)))

    return None


# ============================================================
# PDF / XLSX 금액 파싱
# ============================================================

def _clean(s):
    """금액 문자열 → 정수"""
    return int(re.sub(r"[^\d]", "", s))


def parse_pdf_amount(text, firm):
    """
    로펌별 PDF 텍스트에서 VAT 제외 최종 금액 추출.
    각 로펌 인보이스 양식에 맞춘 패턴 사용.
    """
    # 공백 정규화
    t = re.sub(r"\s+", " ", text)

    # --- 로펌별 패턴 ---

    if firm == "광장":
        # "보수금합계 : 19,976,550 원"
        m = re.search(r"보수금\s*합\s*계\s*[:：]\s*(?:금\s*)?([\d,]+)\s*원", t)
        if m:
            return _clean(m.group(1))

    elif firm == "김앤장":
        # "자 문 료 ￦ 3,000,000"
        m = re.search(r"자\s*문\s*료\s*[￦₩W]?\s*([\d,]+)", t)
        if m:
            return _clean(m.group(1))

    elif firm == "지평":
        # "소 계 ￦ 16,029,280"
        m = re.search(r"소\s*계\s*[￦₩W]?\s*([\d,]+)", t)
        if m:
            return _clean(m.group(1))

    elif firm == "D&A":
        # D&A 표준 양식: "공급가 W 18,880,000" (VAT 제외 정상가)
        # 여러 통화 기호 대응: W, ₩, ￦, \, ¥ 등
        m = re.search(r"공급가\s*[Ww₩￦\\￥]?\s*([\d,]{7,})", t)
        if m:
            return _clean(m.group(1))
        # 할인 있는 경우: "자문료 합계 (20% 할인 후) 34,269,600"
        m = re.search(r"자문료\s*합계\s*\(\s*\d+\s*%\s*할인\s*후\s*\)\s*([\d,]+)", t)
        if m:
            return _clean(m.group(1))
        # "자문료 합계" 단독
        m = re.search(r"자문료\s*합계\s*[:：]?\s*([\d,]{7,})", t)
        if m:
            return _clean(m.group(1))
        # "총 법률자문료(3+4) 20,768,000" 은 VAT 포함이라 마지막 순위
        m = re.search(r"총\s*법률자문료\s*\([^)]+\)\s*([\d,]+)", t)
        if m:
            return int(round(_clean(m.group(1)) / 1.1))

    elif firm.startswith("율촌"):
        # "1. 보수금 : 27,243,750원" (첫 번째 보수금, 보수금합계가 아닌 것)
        m = re.search(r"(?<!합)보수금\s*[:：]\s*([\d,]+)\s*원", t)
        if m:
            return _clean(m.group(1))

    elif firm == "SLP":
        # "법률 자문료 15,512,500"
        m = re.search(r"법률\s*자문료\s*([\d,]+)", t)
        if m:
            return _clean(m.group(1))
        # 새로운 SLP 양식 (2601_법무검토인보이스): TOTAL 컬럼 합산 필요
        # PDF에서는 표 파싱이 어려우니 마지막 큰 숫자 찾기
        amounts = re.findall(r"₩\s*([\d,]+)", t)
        if amounts:
            # 마지막 것이 총합계일 가능성이 높음
            try:
                total = _clean(amounts[-1])
                if total > 100_000:
                    return total
            except Exception:
                pass

    elif firm == "세종_인도네시아" or firm == "세종":
        # 세종은 띄어쓰기가 있고 원화기호가 \(백슬래시)로 표시됨
        # 라인 아이템: "법률고문료 \ 5,000,000" (VAT 제외)
        # 상단 문구: "...법률고문료 ₩ 5,500,000 (10% 부가가치세 포함)..." (VAT 포함, 제외 필요)
        t_ns = t.replace(" ", "")
        # 백슬래시 \ 뒤에 오는 법률고문료 금액 우선 (라인 아이템)
        m = re.search(r"법률고문료\\([\d,]+)", t_ns)
        if m:
            return _clean(m.group(1))
        # 백슬래시가 없으면 부가가치세 포함액에서 VAT 제거
        m = re.search(r"법률고문료\s*[₩￦W]?\s*([\d,]+)\s*\(10%\s*부가가치세\s*포함\)", t_ns)
        if m:
            return int(round(_clean(m.group(1)) / 1.1))
        # 기존 원 단위 패턴 (예비)
        m = re.search(r"청구\s*금액\s*[:：]?\s*([\d,]+)\s*원", t)
        if m:
            return _clean(m.group(1))
        m = re.search(r"자문료\s*합계\s*[:：]?\s*([\d,]+)\s*원", t)
        if m:
            return _clean(m.group(1))

    # --- 범용 폴백 ---
    # "소계", "공급가액", "보수금" 뒤 금액
    m = re.search(r"(?:소\s*계|공급가액?|보수금|자문료)\s*[￦₩W:：]?\s*([\d,]+)", t)
    if m:
        val = _clean(m.group(1))
        if val > 100_000:  # 최소 10만원 이상이어야 의미 있는 금액
            return val

    return None


def parse_xlsx_amount(buf):
    """
    DLS 등 xlsx 파일에서 금액 추출.
    TOTAL 컬럼 합산 또는 '총합계' 셀 값 사용.
    """
    try:
        wb = openpyxl.load_workbook(buf, data_only=True)
        best_total = 0

        for ws in wb.worksheets:
            # 템플릿 시트 스킵
            if "양식" in (ws.title or ""):
                continue

            # 방법 1: "총합계" 옆의 금액 찾기
            for row in ws.iter_rows(values_only=False):
                for i, cell in enumerate(row):
                    if cell.value and "총합계" in str(cell.value):
                        # 같은 행의 다른 셀에서 숫자 찾기
                        for other in row:
                            v = other.value
                            if isinstance(v, (int, float)) and v > 10_000:
                                best_total = max(best_total, int(v))
                            elif isinstance(v, str):
                                cleaned = re.sub(r"[^\d]", "", v)
                                if cleaned and int(cleaned) > 10_000:
                                    best_total = max(best_total, int(cleaned))

            # 방법 2: TOTAL 컬럼 합산
            if best_total == 0:
                header_row = None
                total_col = None
                for row in ws.iter_rows(values_only=False):
                    for cell in row:
                        if cell.value == "TOTAL":
                            header_row = cell.row
                            total_col = cell.column
                            break
                    if header_row:
                        break

                if header_row and total_col:
                    col_total = 0
                    for row in ws.iter_rows(
                        min_row=header_row + 1,
                        min_col=total_col,
                        max_col=total_col,
                    ):
                        v = row[0].value
                        if isinstance(v, (int, float)) and v > 0:
                            col_total += int(v)
                        elif isinstance(v, str):
                            cleaned = re.sub(r"[^\d]", "", v)
                            if cleaned:
                                col_total += int(cleaned)
                    if col_total > 0:
                        best_total = max(best_total, col_total)

        return best_total if best_total > 0 else None
    except Exception:
        return None


# ============================================================
# 데이터 수집 파이프라인
# ============================================================

def get_display_name(folder_name):
    """폴더명 → 대시보드 표시명"""
    return FIRM_DISPLAY_NAMES.get(folder_name, folder_name)


@st.cache_data(ttl=300, show_spinner=False)
def collect_invoices():
    """
    전체 로펌 폴더를 스캔하여 인보이스 데이터 수집.
    반환: (records: list[dict], errors: list[str])
    """
    service = get_drive_service()
    records = []
    errors = []

    # 루트 폴더의 하위 폴더(= 로펌별 폴더) 조회
    firm_folders = [
        f for f in list_folder(service, ROOT_FOLDER_ID)
        if f["mimeType"] == "application/vnd.google-apps.folder"
    ]

    for folder in firm_folders:
        firm_raw = folder["name"]
        firm = get_display_name(firm_raw)
        contents = list_folder(service, folder["id"])

        # 파일과 하위폴더 분리
        sub_folders = [c for c in contents if c["mimeType"] == "application/vnd.google-apps.folder"]
        direct_files = [c for c in contents if c["mimeType"] != "application/vnd.google-apps.folder"]

        # 탐색 대상 그룹: (소스명, 파일목록)
        groups = []
        if direct_files:
            groups.append(("direct", direct_files))
        for sf in sub_folders:
            sf_files = [
                f for f in list_folder(service, sf["id"])
                if f["mimeType"] != "application/vnd.google-apps.folder"
            ]
            if sf_files:
                groups.append((sf["name"], sf_files))

        for source, files in groups:
            invoices = select_invoices(files)
            for inv in invoices:
                fname = inv["name"]

                # 파일명 우선 월 추출 (백업용)
                month_info_from_name = extract_month(fname)
                if not month_info_from_name and source != "direct":
                    month_info_from_name = extract_month(source)

                # 금액 + PDF 텍스트 기반 월 추출
                amount = None
                pdf_month_info = None
                try:
                    if fname.lower().endswith(".pdf"):
                        buf = download_file(service, inv["id"])
                        with pdfplumber.open(buf) as pdf:
                            text = "\n".join(
                                (p.extract_text() or "") for p in pdf.pages
                            )
                        # PDF 본문에서 자문 월 추출 (최우선)
                        pdf_month_info = extract_month_from_pdf_text(text)
                        amount = parse_pdf_amount(text, firm)
                    elif fname.lower().endswith(".xlsx"):
                        buf = download_file(service, inv["id"])
                        amount = parse_xlsx_amount(buf)
                except Exception as e:
                    errors.append(f"[{firm}] 파싱 오류 ({fname}): {e}")

                # 월 결정: PDF 텍스트 → 파일명 순
                month_info = pdf_month_info or month_info_from_name
                if not month_info:
                    errors.append(f"[{firm}] 월 추출 실패: {fname}")
                    continue

                year, month = month_info

                if amount and amount > 0:
                    records.append(
                        {
                            "로펌": firm,
                            "연도": year,
                            "월": month,
                            "기간": f"{year}-{month:02d}",
                            "금액": amount,
                            "파일명": fname,
                            "링크": inv.get("webViewLink", ""),
                            "파일날짜": extract_file_date(fname),
                            "우선순위": (1 if ("최종" in fname or "할인 후" in fname) else 0),
                        }
                    )
                elif amount is None:
                    errors.append(f"[{firm}] 금액 추출 실패: {fname}")

    # 중복 제거: 같은 로펌+같은 월 → (1) '최종' 우선 (2) 파일명 날짜 최신 우선
    if records:
        df_tmp = pd.DataFrame(records)
        df_tmp = df_tmp.sort_values(["우선순위", "파일날짜"], ascending=[False, False])
        df_tmp = df_tmp.drop_duplicates(subset=["로펌", "기간"], keep="first")
        df_tmp = df_tmp.drop(columns=["우선순위"])
        records = df_tmp.to_dict("records")

    # 표시기간 = 실제 자문 수행 월 (PDF에서 추출된 그대로)
    for r in records:
        r["표시기간"] = r["기간"]
        r["표시연도"] = r["연도"]
        r["표시월"] = r["월"]

    return records, errors


# ============================================================
# 대시보드 UI
# ============================================================

def apply_custom_css():
    """대시보드 커스텀 스타일"""
    st.markdown(
        """
        <style>
        /* 전체 폰트 크기 */
        html, body, [class*="css"] {
            font-size: 16px;
        }
        /* KPI 카드 */
        div[data-testid="stMetric"] {
            background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
            border-radius: 12px;
            padding: 16px 20px;
            border-left: 4px solid #1B4F72;
        }
        div[data-testid="stMetric"] label {
            font-size: 0.95rem !important;
            color: #495057;
        }
        div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
            font-size: 1.6rem !important;
            font-weight: 700;
            color: #1B4F72;
        }
        /* 사이드바 */
        section[data-testid="stSidebar"] {
            background: #f8f9fa;
        }
        /* multiselect 태그 색상 (빨강 → 진한 회색) */
        span[data-baseweb="tag"] {
            background-color: #5a5a5a !important;
        }
        span[data-baseweb="tag"] span {
            color: white !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def check_password():
    """비밀번호 인증. 통과 시 True 반환."""
    if st.session_state.get("password_correct", False):
        return True

    expected = st.secrets.get("DASHBOARD_PASSWORD", "dw2026")

    st.title("🔒 정기자문 비용 현황")
    st.caption("대웅제약 법무1팀 · 접근 인증")
    st.markdown("")

    pw = st.text_input("비밀번호", type="password", key="pw_input")
    if st.button("확인", type="primary"):
        if pw == expected:
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("비밀번호가 일치하지 않습니다.")
    return False


def main():
    if not check_password():
        st.stop()

    apply_custom_css()

    st.title("⚖️ 정기자문 비용 현황")
    st.caption("대웅제약 법무1팀 · 구글 드라이브 인보이스 자동 집계")

    # ---- 데이터 로드 ----
    with st.spinner("📂 구글 드라이브에서 인보이스를 읽는 중..."):
        records, errors = collect_invoices()

    if not records:
        st.error("인보이스 데이터를 찾을 수 없습니다. 폴더 공유 설정을 확인해 주세요.")
        if errors:
            with st.expander("⚠️ 오류 상세"):
                for e in errors:
                    st.text(e)
        st.stop()

    df = pd.DataFrame(records)

    # ---- 사이드바 필터 ----
    with st.sidebar:
        st.header("📊 조회 설정")

        period_unit = st.radio("집계 단위", ["월별", "분기별", "반기별", "연도별"])

        years = sorted(df["표시연도"].unique())
        sel_year = st.selectbox("연도", years, index=len(years) - 1)

        all_firms = sorted(df["로펌"].unique())

        # 세션 상태로 로펌 선택 관리
        if "sel_firms_state" not in st.session_state:
            st.session_state["sel_firms_state"] = all_firms

        bc1, bc2 = st.columns(2)
        if bc1.button("전체 선택", use_container_width=True):
            st.session_state["sel_firms_state"] = all_firms
            st.rerun()
        if bc2.button("전체 해제", use_container_width=True):
            st.session_state["sel_firms_state"] = []
            st.rerun()

        sel_firms = st.multiselect(
            "로펌",
            all_firms,
            default=st.session_state["sel_firms_state"],
            key="sel_firms_widget",
        )
        st.session_state["sel_firms_state"] = sel_firms

        st.divider()
        if st.button("🔄 새로고침", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

        if errors:
            with st.expander(f"⚠️ 파싱 이슈 ({len(errors)}건)"):
                for e in errors:
                    st.caption(e)

    # ---- 필터 적용 ----
    fdf = df[(df["표시연도"] == sel_year) & (df["로펌"].isin(sel_firms))].copy()
    if fdf.empty:
        st.warning("선택한 조건에 해당하는 데이터가 없습니다.")
        st.stop()

    # ---- 기간 집계 ----
    agg = fdf.groupby(["로펌", "표시연도", "표시월", "표시기간"], as_index=False)["금액"].sum()

    if period_unit == "월별":
        agg["집계"] = agg["표시기간"]
    elif period_unit == "분기별":
        agg["집계"] = agg.apply(lambda r: f"{r['표시연도']}-Q{(r['표시월']-1)//3+1}", axis=1)
    elif period_unit == "반기별":
        agg["집계"] = agg.apply(
            lambda r: f"{r['표시연도']}-{'상반기' if r['표시월'] <= 6 else '하반기'}", axis=1
        )
    else:
        agg["집계"] = agg["표시연도"].astype(str)

    chart_df = agg.groupby(["집계", "로펌"], as_index=False)["금액"].sum()

    # ---- KPI ----
    c1, c2, c3, c4 = st.columns(4)
    total = fdf["금액"].sum()
    monthly_totals = fdf.groupby("표시기간")["금액"].sum()
    avg = monthly_totals.mean()
    peak = monthly_totals.idxmax() if not monthly_totals.empty else "-"
    n_firms = fdf["로펌"].nunique()

    c1.metric("총 비용 (VAT 제외)", f"₩{total:,.0f}")
    c2.metric("월평균", f"₩{avg:,.0f}")
    c3.metric("최고 지출월", str(peak))
    c4.metric("활성 로펌", f"{n_firms}곳")

    st.divider()

    # ---- 최근 3개월 요약 (클릭 가능 HTML 테이블) ----
    recent_months = sorted(fdf["표시기간"].unique(), reverse=True)[:3]
    firms_sorted = sorted(fdf["로펌"].unique())

    html = '<table style="width:100%; border-collapse:collapse; font-size:15px;">'
    html += '<tr style="background:#f1f3f5; border-bottom:2px solid #dee2e6;">'
    html += '<th style="padding:10px; text-align:center;">비용발생월</th>'
    for firm in firms_sorted:
        html += f'<th style="padding:10px; text-align:center;">{firm}</th>'
    html += '</tr>'

    for period in recent_months:
        html += '<tr style="border-bottom:1px solid #dee2e6;">'
        html += f'<td style="padding:10px; font-weight:600; text-align:center;">{period}</td>'
        for firm in firms_sorted:
            row_data = fdf[(fdf["로펌"] == firm) & (fdf["표시기간"] == period)]
            if not row_data.empty:
                amt = row_data["금액"].sum()
                link = row_data.iloc[0]["링크"]
                html += f'<td style="padding:10px; text-align:center;">'
                html += f'<a href="{link}" target="_blank" style="color:#1B4F72; text-decoration:none; font-weight:500;">₩{amt:,.0f}</a>'
                html += '</td>'
            else:
                html += '<td style="padding:10px; text-align:center; color:#ccc;">-</td>'
        html += '</tr>'
    html += '</table>'

    st.markdown(html, unsafe_allow_html=True)
    st.caption("💡 금액을 클릭하면 인보이스 PDF가 새 탭에서 열립니다.")

    st.markdown("<br>", unsafe_allow_html=True)
    st.divider()

    # ---- 막대 차트 ----
    periods_sorted = sorted(chart_df["집계"].unique())

    fig = go.Figure()
    for firm in sorted(chart_df["로펌"].unique()):
        fd = chart_df[chart_df["로펌"] == firm]
        fig.add_trace(
            go.Bar(
                x=fd["집계"],
                y=fd["금액"],
                name=firm,
                marker_color=FIRM_COLORS.get(firm, "#95A5A6"),
                text=fd["금액"].apply(lambda v: f"{v/10_000:,.0f}만"),
                textposition="inside",
                textangle=0,
                constraintext="none",
                textfont=dict(size=11, color="white"),
                hovertemplate="%{x}<br>%{fullData.name}: ₩%{y:,.0f}<extra></extra>",
            )
        )

    fig.update_layout(
        barmode="group",
        xaxis=dict(title="비용발생월", categoryorder="array", categoryarray=periods_sorted, tickfont=dict(size=13)),
        yaxis=dict(title="금액 (원)", tickformat=",", tickfont=dict(size=13)),
        legend=dict(orientation="h", y=-0.18, x=0.5, xanchor="center", font=dict(size=14)),
        height=480,
        margin=dict(t=20, b=80),
        template="plotly_white",
        hoverlabel=dict(font_size=14),
    )

    # 월 사이 세로 점선 구분선 (카테고리 축)
    shapes = []
    for i in range(len(periods_sorted) - 1):
        shapes.append(
            dict(
                type="line",
                xref="x",
                yref="paper",
                x0=i + 0.5,
                x1=i + 0.5,
                y0=0,
                y1=1,
                line=dict(color="#ccc", width=1, dash="dot"),
            )
        )
    fig.update_layout(shapes=shapes)

    st.plotly_chart(fig, use_container_width=True)

    # ---- 상세 내역 (필터 포함) ----
    st.subheader("📋 상세 내역")

    fc1, fc2 = st.columns(2)
    with fc1:
        detail_firms = st.multiselect(
            "로펌 필터", firms_sorted, default=firms_sorted, key="detail_firm"
        )
    with fc2:
        all_months = sorted(fdf["표시기간"].unique(), reverse=True)
        detail_months = st.multiselect(
            "월 필터", all_months, default=all_months, key="detail_month"
        )

    tbl = fdf[
        (fdf["로펌"].isin(detail_firms)) & (fdf["표시기간"].isin(detail_months))
    ][["로펌", "표시기간", "금액", "파일명", "링크"]].copy()
    tbl = tbl.sort_values(["표시기간", "로펌"], ascending=[False, True]).reset_index(drop=True)
    tbl["금액(VAT제외)"] = tbl["금액"].apply(lambda x: f"₩{x:,.0f}")
    tbl = tbl.drop(columns=["금액"])
    tbl = tbl.rename(columns={"표시기간": "비용발생월"})

    st.dataframe(
        tbl,
        column_config={
            "링크": st.column_config.LinkColumn("PDF", display_text="📄 열기"),
            "파일명": st.column_config.TextColumn("파일명", width="large"),
        },
        hide_index=True,
        use_container_width=True,
    )

    # ---- 로펌×기간 피벗 테이블 ----
    st.subheader("📊 로펌별 기간 합계")

    pivot = fdf.pivot_table(
        values="금액",
        index="로펌",
        columns="표시기간",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="합계",
    )
    cols = sorted([c for c in pivot.columns if c != "합계"]) + ["합계"]
    pivot = pivot.reindex(columns=cols)

    st.dataframe(
        pivot.style.format("₩{:,.0f}").map(
            lambda v: "color: #ccc" if v == 0 else "", subset=pivot.columns
        ),
        use_container_width=True,
    )


# ============================================================
# 실행
# ============================================================
if __name__ == "__main__":
    main()
