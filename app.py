"""
⚖️ 법무팀 외부 로펌 비용 대시보드
대웅제약 법무1팀 · 구글 시트 + 드라이브 PDF 연동

데이터: 구글 시트 "정기자문 & 일반자문" 섹션 (원 단위)
PDF: 구글 드라이브 로펌별 폴더에서 자동 매칭
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from google.oauth2 import service_account
from googleapiclient.discovery import build
import re
import json

# ============================================================
# 페이지 설정
# ============================================================
st.set_page_config(
    page_title="법무팀 로펌 비용 대시보드",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# 상수
# ============================================================

SPREADSHEET_ID = st.secrets.get(
    "SPREADSHEET_ID", "1KiSI0ZdXvp_SHeiiz2DJX0Cl-Wv0CnJiRhzp72f9LgY"
)
SHEET_NAME = st.secrets.get("SHEET_NAME", "2026")
ROOT_FOLDER_ID = st.secrets.get(
    "ROOT_FOLDER_ID", "1HleAj4z6DH9KxMjyuf56lRD4b-c7pOyh"
)

# 시트 업체명 → 드라이브 폴더명 매핑 (이름이 다른 경우만)
FIRM_NAME_MAP = {
    "디.엘.에스": "DLS",
    "대륙아주": "D&A",
    "에스엘파트너스": "SLP",
    "세종_인니": "세종_인도네시아",
}

# 로펌별 차트 색상
FIRM_COLORS = {
    "광장": "#1B4F72",
    "김앤장": "#922B21",
    "지평": "#196F3D",
    "D&A": "#7D3C98",
    "대륙아주": "#7D3C98",
    "율촌": "#B9770E",
    "SLP": "#2E86C1",
    "에스엘파트너스": "#2E86C1",
    "DLS": "#D35400",
    "디.엘.에스": "#D35400",
    "세종": "#16A085",
    "세종_인니": "#148F77",
    "세종_인도네시아": "#148F77",
    "비앤에이치": "#5D6D7E",
    "CCL": "#AAB7B8",
    "셰퍼드멀린": "#717D7E",
}

INVOICE_KEYWORDS = ["자문료", "인보이스", "invoice", "청구", "보수금", "INVOICE"]


# ============================================================
# Google API 인증
# ============================================================

@st.cache_resource
def get_services():
    """Google Drive + Sheets API 서비스 생성"""
    creds_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    creds = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=[
            "https://www.googleapis.com/auth/drive.readonly",
            "https://www.googleapis.com/auth/spreadsheets.readonly",
        ],
    )
    drive = build("drive", "v3", credentials=creds)
    sheets = build("sheets", "v4", credentials=creds)
    return drive, sheets


# ============================================================
# 구글 시트 데이터 읽기
# ============================================================

@st.cache_data(ttl=300, show_spinner=False)
def read_sheet_data():
    """
    구글 시트에서 '정기자문 & 일반자문' 섹션 읽기.
    반환: (records: list[dict], subtotal_row: dict, errors: list[str])
    """
    _, sheets = get_services()
    errors = []

    try:
        result = (
            sheets.spreadsheets()
            .values()
            .get(
                spreadsheetId=SPREADSHEET_ID,
                range=f"{SHEET_NAME}!A1:S80",
                valueRenderOption="UNFORMATTED_VALUE",
            )
            .execute()
        )
    except Exception as e:
        return [], {}, [f"시트 읽기 실패: {e}"]

    rows = result.get("values", [])
    if not rows:
        return [], {}, ["시트에 데이터가 없습니다."]

    # 헤더에서 월 컬럼 위치 찾기 (1~12)
    month_cols = {}  # {월번호: 열인덱스}
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            if val == 1 or val == "1":
                # 이 행이 월 번호 헤더인지 확인 (연속된 1~12)
                is_header = True
                for k in range(1, 12):
                    if j + k < len(row) and (row[j + k] == k + 1 or row[j + k] == str(k + 1)):
                        continue
                    else:
                        is_header = False
                        break
                if is_header:
                    for m in range(12):
                        month_cols[m + 1] = j + m
                    break
        if month_cols:
            break

    if not month_cols:
        return [], {}, ["시트에서 월 헤더를 찾을 수 없습니다."]

    # '정기자문' 섹션 찾기
    start_idx = None
    end_idx = None
    subtotal_idx = None
    firm_col = None  # 업체명 컬럼

    for i, row in enumerate(rows):
        row_text = " ".join(str(v) for v in row)
        if "정기자문" in row_text and "일반자문" in row_text:
            start_idx = i
            # 업체명은 보통 시작 행과 같은 행 또는 다음 행부터
            continue
        if start_idx is not None and "소계" in row_text:
            subtotal_idx = i
            end_idx = i
            break

    if start_idx is None:
        return [], {}, ["시트에서 '정기자문 & 일반자문' 섹션을 찾을 수 없습니다."]

    # 업체명 컬럼 추정 (정기자문 행의 다음 열들에서 업체명이 있는 컬럼)
    # 보통 C열(인덱스 2)
    firm_col = 2
    content_col = 3  # 내용 컬럼 (D열)

    # 개별 로펌 행 읽기
    records = []
    for i in range(start_idx, end_idx):
        row = rows[i] if i < len(rows) else []
        if len(row) <= firm_col:
            continue
        firm_name = str(row[firm_col]).strip() if row[firm_col] else ""
        if not firm_name or firm_name == "소계" or "정기자문" in firm_name:
            continue

        content = str(row[content_col]).strip() if len(row) > content_col and row[content_col] else ""

        for month_num, col_idx in month_cols.items():
            val = row[col_idx] if col_idx < len(row) else 0
            if isinstance(val, (int, float)):
                amount = int(val) if val != 0 else 0
            else:
                amount = 0
            records.append({
                "업체명": firm_name,
                "내용": content,
                "월": month_num,
                "기간": f"2026-{month_num:02d}",
                "금액": amount,
            })

    # 소계 행 읽기
    subtotal = {}
    if subtotal_idx is not None and subtotal_idx < len(rows):
        srow = rows[subtotal_idx]
        for month_num, col_idx in month_cols.items():
            val = srow[col_idx] if col_idx < len(srow) else 0
            if isinstance(val, (int, float)):
                subtotal[month_num] = int(val)
            else:
                subtotal[month_num] = 0

    return records, subtotal, errors


# ============================================================
# 구글 드라이브 PDF 스캔
# ============================================================

def extract_month_from_filename(filename):
    """파일명에서 청구 월 추출 → month 정수 또는 None"""
    m = re.search(r"(\d{1,2})\s*월", filename)
    if m:
        return int(m.group(1))
    m = re.search(r"2026[.\s]+(\d{2})", filename)
    if m:
        return int(m.group(1))
    m = re.search(r"\(26(\d{2})\)", filename)
    if m:
        return int(m.group(1))
    m = re.search(r"(?<!\d)26\.(\d{2})", filename)
    if m:
        return int(m.group(1))
    m = re.search(r"26(\d{2})\d{4}", filename)
    if m:
        return int(m.group(1))
    return None


def extract_file_date(filename):
    """파일명에서 6자리 날짜 추출 → 정수. 없으면 0."""
    m = re.search(r"(?<!\d)(2[56]\d{4})(?!\d)", filename)
    return int(m.group(1)) if m else 0


@st.cache_data(ttl=300, show_spinner=False)
def scan_drive_pdfs():
    """
    드라이브에서 로펌별 PDF 목록 스캔.
    반환: {폴더명: {월: {name, link, file_date}}}
    """
    drive, _ = get_services()
    pdf_map = {}

    try:
        # 루트 폴더의 하위 폴더 목록
        resp = drive.files().list(
            q=f"'{ROOT_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
            fields="files(id, name)",
        ).execute()
        folders = resp.get("files", [])
    except Exception:
        return {}

    for folder in folders:
        fname = folder["name"]
        pdf_map[fname] = {}

        # 폴더 내 파일 목록
        try:
            items = []
            page_token = None
            while True:
                resp = drive.files().list(
                    q=f"'{folder['id']}' in parents and trashed=false",
                    fields="nextPageToken, files(id, name, mimeType, webViewLink)",
                    pageToken=page_token,
                ).execute()
                items.extend(resp.get("files", []))
                page_token = resp.get("nextPageToken")
                if not page_token:
                    break
        except Exception:
            continue

        # 하위폴더도 탐색 (SLP 같은 경우)
        sub_files = []
        for item in items:
            if item["mimeType"] == "application/vnd.google-apps.folder":
                try:
                    sr = drive.files().list(
                        q=f"'{item['id']}' in parents and trashed=false",
                        fields="files(id, name, mimeType, webViewLink)",
                    ).execute()
                    sub_files.extend(sr.get("files", []))
                except Exception:
                    pass
            else:
                sub_files.append(item)

        # PDF 파일에서 월 추출
        for f in sub_files:
            if not f["name"].lower().endswith((".pdf", ".xlsx")):
                continue
            month = extract_month_from_filename(f["name"])
            if month is None:
                continue
            fdate = extract_file_date(f["name"])

            # 같은 월에 여러 파일이면 파일명 날짜가 가장 큰 것만 유지
            if month not in pdf_map[fname] or fdate > pdf_map[fname][month].get("file_date", 0):
                pdf_map[fname][month] = {
                    "name": f["name"],
                    "link": f.get("webViewLink", ""),
                    "file_date": fdate,
                }

    return pdf_map


# ============================================================
# 데이터 병합
# ============================================================

def get_drive_folder(sheet_firm):
    """시트 업체명 → 드라이브 폴더명"""
    return FIRM_NAME_MAP.get(sheet_firm, sheet_firm)


@st.cache_data(ttl=300, show_spinner=False)
def build_dashboard_data():
    """시트 데이터 + 드라이브 PDF → 대시보드용 데이터 생성"""
    sheet_records, subtotal, sheet_errors = read_sheet_data()
    pdf_map = scan_drive_pdfs()
    errors = list(sheet_errors)

    rows = []
    firms_seen = set()

    for rec in sheet_records:
        firm = rec["업체명"]
        month = rec["월"]
        amount = rec["금액"]
        drive_folder = get_drive_folder(firm)
        firms_seen.add(firm)

        # PDF 존재 여부 확인
        pdf_info = pdf_map.get(drive_folder, {}).get(month)
        has_pdf = pdf_info is not None
        pdf_link = pdf_info["link"] if pdf_info else ""
        pdf_name = pdf_info["name"] if pdf_info else ""

        # 판정
        if amount > 0 and has_pdf:
            status = "확정"
        elif amount == 0 and has_pdf:
            status = "⚠️ 법무비용 업데이트"
            amount = 0
        elif amount > 0 and not has_pdf:
            status = "예측"  # 무시
            continue
        else:
            continue  # 금액 0 + PDF 없음 → 스킵

        rows.append({
            "로펌": firm,
            "월": month,
            "기간": rec["기간"],
            "금액": amount,
            "상태": status,
            "PDF링크": pdf_link,
            "파일명": pdf_name,
        })

    # 드라이브에 PDF 있는데 시트에 해당 로펌이 아예 없는 경우 체크
    for folder_name, months in pdf_map.items():
        # 역매핑: 드라이브 폴더명 → 시트 업체명
        reverse_map = {v: k for k, v in FIRM_NAME_MAP.items()}
        sheet_name = reverse_map.get(folder_name, folder_name)
        if sheet_name not in firms_seen:
            for month, info in months.items():
                rows.append({
                    "로펌": sheet_name,
                    "월": month,
                    "기간": f"2026-{month:02d}",
                    "금액": 0,
                    "상태": "⚠️ 법무비용 업데이트",
                    "PDF링크": info["link"],
                    "파일명": info["name"],
                })

    # 소계 검증
    if subtotal:
        df_tmp = pd.DataFrame(rows)
        if not df_tmp.empty:
            confirmed = df_tmp[df_tmp["상태"] == "확정"]
            for month_num, expected in subtotal.items():
                if expected == 0:
                    continue
                actual = confirmed[confirmed["월"] == month_num]["금액"].sum()
                if actual > 0 and abs(actual - expected) > 100:
                    errors.append(
                        f"⚠️ 소계 불일치 {month_num}월: 시트 소계 ₩{expected:,.0f} vs 확정 합산 ₩{actual:,.0f} (차이: ₩{abs(actual-expected):,.0f})"
                    )

    return rows, errors


# ============================================================
# 대시보드 UI
# ============================================================

def apply_custom_css():
    st.markdown(
        """
        <style>
        html, body, [class*="css"] {
            font-size: 16px;
        }
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
        section[data-testid="stSidebar"] {
            background: #f8f9fa;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def main():
    apply_custom_css()

    st.title("⚖️ 외부 로펌 비용 대시보드")
    st.caption("대웅제약 법무1팀 · 구글 시트 + 드라이브 PDF 연동")

    # ---- 데이터 로드 ----
    with st.spinner("📂 데이터를 읽는 중..."):
        rows, errors = build_dashboard_data()

    if not rows:
        st.error("데이터를 찾을 수 없습니다. 시트/폴더 공유 설정을 확인해 주세요.")
        if errors:
            with st.expander("⚠️ 오류 상세"):
                for e in errors:
                    st.text(e)
        st.stop()

    df = pd.DataFrame(rows)

    # ---- 사이드바 필터 ----
    with st.sidebar:
        st.header("📊 조회 설정")

        period_unit = st.radio("집계 단위", ["월별", "분기별", "반기별", "연도별"])

        all_firms = sorted(df["로펌"].unique())
        sel_firms = st.multiselect("로펌", all_firms, default=all_firms)

        st.divider()
        if st.button("🔄 새로고침", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

        if errors:
            with st.expander(f"⚠️ 이슈 ({len(errors)}건)"):
                for e in errors:
                    st.caption(e)

    # ---- 필터 적용 ----
    fdf = df[df["로펌"].isin(sel_firms)].copy()
    confirmed = fdf[fdf["상태"] == "확정"].copy()

    if confirmed.empty:
        st.warning("확정된 데이터가 없습니다.")
        # 법무비용 업데이트 필요 항목 표시
        updates = fdf[fdf["상태"] == "⚠️ 법무비용 업데이트"]
        if not updates.empty:
            st.subheader("⚠️ 법무비용 업데이트 필요")
            st.dataframe(updates[["로펌", "기간", "파일명"]], hide_index=True)
        st.stop()

    # ---- 기간 집계 ----
    agg = confirmed.groupby(["로펌", "월", "기간"], as_index=False)["금액"].sum()

    if period_unit == "월별":
        agg["집계"] = agg["기간"]
    elif period_unit == "분기별":
        agg["집계"] = agg["월"].apply(lambda m: f"2026-Q{(m-1)//3+1}")
    elif period_unit == "반기별":
        agg["집계"] = agg["월"].apply(lambda m: f"2026-{'상반기' if m <= 6 else '하반기'}")
    else:
        agg["집계"] = "2026"

    chart_df = agg.groupby(["집계", "로펌"], as_index=False)["금액"].sum()

    # ---- KPI ----
    c1, c2, c3, c4 = st.columns(4)
    total = confirmed["금액"].sum()
    monthly_totals = confirmed.groupby("기간")["금액"].sum()
    avg = monthly_totals.mean()
    peak = monthly_totals.idxmax() if not monthly_totals.empty else "-"
    n_firms = confirmed["로펌"].nunique()

    c1.metric("총 비용 (VAT 제외)", f"₩{total:,.0f}")
    c2.metric("월평균", f"₩{avg:,.0f}")
    c3.metric("최고 지출월", str(peak))
    c4.metric("활성 로펌", f"{n_firms}곳")

    st.divider()

    # ---- 최근 3개월 요약 (클릭 가능 HTML 테이블) ----
    recent_months = sorted(confirmed["기간"].unique(), reverse=True)[:3]
    firms_sorted = sorted(confirmed["로펌"].unique())

    html = '<table style="width:100%; border-collapse:collapse; font-size:15px;">'
    html += '<tr style="background:#f1f3f5; border-bottom:2px solid #dee2e6;">'
    html += '<th style="padding:10px; text-align:center;">기간</th>'
    for firm in firms_sorted:
        html += f'<th style="padding:10px; text-align:center;">{firm}</th>'
    html += '</tr>'

    for period in recent_months:
        html += '<tr style="border-bottom:1px solid #dee2e6;">'
        html += f'<td style="padding:10px; font-weight:600; text-align:center;">{period}</td>'
        for firm in firms_sorted:
            row_data = confirmed[(confirmed["로펌"] == firm) & (confirmed["기간"] == period)]
            if not row_data.empty:
                amt = row_data["금액"].sum()
                link = row_data.iloc[0]["PDF링크"]
                if link:
                    html += f'<td style="padding:10px; text-align:center;">'
                    html += f'<a href="{link}" target="_blank" style="color:#1B4F72; text-decoration:none; font-weight:500;">₩{amt:,.0f}</a>'
                    html += '</td>'
                else:
                    html += f'<td style="padding:10px; text-align:center;">₩{amt:,.0f}</td>'
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
        xaxis=dict(title="기간", categoryorder="array", categoryarray=periods_sorted, tickfont=dict(size=13)),
        yaxis=dict(title="금액 (원)", tickformat=",", tickfont=dict(size=13)),
        legend=dict(orientation="h", y=-0.18, x=0.5, xanchor="center", font=dict(size=14)),
        height=480,
        margin=dict(t=20, b=80),
        template="plotly_white",
        hoverlabel=dict(font_size=14),
    )
    st.plotly_chart(fig, use_container_width=True)

    # ---- ⚠️ 법무비용 업데이트 필요 ----
    updates = fdf[fdf["상태"] == "⚠️ 법무비용 업데이트"]
    if not updates.empty:
        st.subheader("⚠️ 법무비용 업데이트 필요")
        st.caption("PDF가 드라이브에 있으나 시트에 금액이 미입력된 항목입니다.")
        utbl = updates[["로펌", "기간", "파일명", "PDF링크"]].copy()
        utbl = utbl.sort_values(["기간", "로펌"]).reset_index(drop=True)
        st.dataframe(
            utbl,
            column_config={
                "PDF링크": st.column_config.LinkColumn("PDF", display_text="📄 열기"),
            },
            hide_index=True,
            use_container_width=True,
        )

    # ---- 상세 내역 (필터 포함) ----
    st.subheader("📋 상세 내역")

    fc1, fc2 = st.columns(2)
    with fc1:
        detail_firms = st.multiselect(
            "로펌 필터", firms_sorted, default=firms_sorted, key="detail_firm"
        )
    with fc2:
        all_months = sorted(confirmed["기간"].unique(), reverse=True)
        detail_months = st.multiselect(
            "월 필터", all_months, default=all_months, key="detail_month"
        )

    tbl = confirmed[
        (confirmed["로펌"].isin(detail_firms)) & (confirmed["기간"].isin(detail_months))
    ][["로펌", "기간", "금액", "파일명", "PDF링크"]].copy()
    tbl = tbl.sort_values(["기간", "로펌"], ascending=[False, True]).reset_index(drop=True)
    tbl["금액(VAT제외)"] = tbl["금액"].apply(lambda x: f"₩{x:,.0f}")
    tbl = tbl.drop(columns=["금액"])

    st.dataframe(
        tbl,
        column_config={
            "PDF링크": st.column_config.LinkColumn("PDF", display_text="📄 열기"),
            "파일명": st.column_config.TextColumn("파일명", width="large"),
        },
        hide_index=True,
        use_container_width=True,
    )

    # ---- 로펌×기간 피벗 ----
    st.subheader("📊 로펌별 기간 합계")

    pivot = confirmed.pivot_table(
        values="금액",
        index="로펌",
        columns="기간",
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


if __name__ == "__main__":
    main()
