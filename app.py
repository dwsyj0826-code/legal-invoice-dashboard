"""
⚖️ 법무팀 외부 로펌 비용 대시보드
대웅제약 법무1팀 · 구글 시트 + 드라이브 PDF 연동
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from google.oauth2 import service_account
from googleapiclient.discovery import build
import re
import json

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
SHEET_NAME = st.secrets.get("SHEET_NAME", "2026년")
ROOT_FOLDER_ID = st.secrets.get(
    "ROOT_FOLDER_ID", "1HleAj4z6DH9KxMjyuf56lRD4b-c7pOyh"
)

# 시트 업체명 → 드라이브 폴더명 매핑 (다른 경우만)
FIRM_NAME_MAP = {
    "디.엘.에스": "DLS",
    "대륙아주": "D&A",
    "에스엘파트너스": "SLP",
    "세종_인니": "세종_인도네시아",
}

FIRM_COLORS = {
    "광장": "#1B4F72",
    "김앤장": "#922B21",
    "지평": "#196F3D",
    "디.엘.에스": "#D35400",
    "대륙아주": "#7D3C98",
    "에스엘파트너스": "#2E86C1",
    "율촌": "#B9770E",
    "세종": "#16A085",
    "세종_인니": "#148F77",
    "비앤에이치": "#5D6D7E",
    "CCL": "#AAB7B8",
    "셰퍼드멀린": "#717D7E",
}


# ============================================================
# Google API 인증
# ============================================================
@st.cache_resource
def get_services():
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
# 구글 시트 읽기
# ============================================================
@st.cache_data(ttl=300, show_spinner=False)
def read_sheet_data():
    """
    시트 구조:
    - B열: 구분 (정기자문 & 일반자문)
    - C열: 업체명
    - D열: 내용
    - E~P열: 1~12월 (억 단위, 예: 0.07 = 7,000,000원)
    - 소계 행은 원 단위 (예: 24,860,500)
    """
    _, sheets = get_services()
    errors = []

    try:
        result = (
            sheets.spreadsheets().values().get(
                spreadsheetId=SPREADSHEET_ID,
                range=f"'{SHEET_NAME}'!A1:S80",
                valueRenderOption="UNFORMATTED_VALUE",
            ).execute()
        )
    except Exception as e:
        return [], {}, [f"시트 읽기 실패: {e}"]

    rows = result.get("values", [])
    if not rows:
        return [], {}, ["시트에 데이터가 없습니다."]

    # 월 헤더 행 찾기 (1,2,3,...,12 가 연속으로 나오는 행)
    month_row_idx = None
    month_start_col = None
    for i, row in enumerate(rows):
        for j in range(len(row) - 11):
            try:
                vals = [int(row[j + k]) for k in range(12)]
                if vals == list(range(1, 13)):
                    month_row_idx = i
                    month_start_col = j
                    break
            except (ValueError, TypeError):
                continue
        if month_row_idx is not None:
            break

    if month_row_idx is None:
        return [], {}, ["월 헤더(1~12)를 찾을 수 없습니다."]

    # 정기자문 & 일반자문 섹션 찾기
    section_start = None
    section_end = None
    for i in range(month_row_idx + 1, len(rows)):
        row = rows[i]
        b_val = str(row[1]) if len(row) > 1 else ""
        if "정기자문" in b_val and "일반자문" in b_val:
            section_start = i
            break

    if section_start is None:
        # 이전 행에 "정기자문 & 일반자문"이 있을 수도 있음 (병합셀)
        # C열에 업체명이 있고 이전 소계 이후 시작하는 지점을 찾자
        subtotal_indices = []
        for i in range(month_row_idx + 1, len(rows)):
            row = rows[i]
            b_val = str(row[1]) if len(row) > 1 else ""
            if b_val == "소계":
                subtotal_indices.append(i)

        # 첫 번째 소계 이후가 정기자문 시작
        if subtotal_indices:
            section_start = subtotal_indices[0] + 1

    if section_start is None:
        return [], {}, ["정기자문 & 일반자문 섹션을 찾을 수 없습니다."]

    # 섹션 끝(다음 소계) 찾기
    for i in range(section_start, len(rows)):
        row = rows[i]
        b_val = str(row[1]) if len(row) > 1 else ""
        if b_val == "소계":
            section_end = i
            break

    if section_end is None:
        section_end = len(rows)

    # 개별 로펌 행 파싱
    records = []
    firm_col = 2  # C열
    content_col = 3  # D열

    for i in range(section_start, section_end):
        row = rows[i]
        if len(row) <= firm_col:
            continue
        firm_name = str(row[firm_col]).strip() if row[firm_col] else ""
        if not firm_name or "정기자문" in firm_name or firm_name == "소계":
            continue
        content = str(row[content_col]).strip() if len(row) > content_col and row[content_col] else ""

        for m in range(12):
            col_idx = month_start_col + m
            if col_idx >= len(row):
                continue
            val = row[col_idx]
            if isinstance(val, (int, float)) and val > 0:
                # 억 단위 → 원 단위 (0.07 → 7,000,000)
                amount = int(round(val * 100_000_000))
            else:
                amount = 0
            records.append({
                "업체명": firm_name,
                "내용": content,
                "월": m + 1,
                "기간": f"2026-{m+1:02d}",
                "금액": amount,
            })

    # 소계 행 (원 단위)
    subtotal = {}
    if section_end < len(rows):
        srow = rows[section_end]
        for m in range(12):
            col_idx = month_start_col + m
            if col_idx < len(srow):
                v = srow[col_idx]
                if isinstance(v, (int, float)):
                    subtotal[m + 1] = int(v)

    return records, subtotal, errors


# ============================================================
# 드라이브 PDF 스캔
# ============================================================
def extract_month_from_filename(fn):
    m = re.search(r"(\d{1,2})\s*월", fn)
    if m: return int(m.group(1))
    m = re.search(r"2026[.\s]+(\d{2})", fn)
    if m: return int(m.group(1))
    m = re.search(r"\(26(\d{2})\)", fn)
    if m: return int(m.group(1))
    m = re.search(r"(?<!\d)26\.(\d{2})", fn)
    if m: return int(m.group(1))
    m = re.search(r"26(\d{2})\d{4}", fn)
    if m: return int(m.group(1))
    return None


def extract_file_date(fn):
    m = re.search(r"(?<!\d)(2[56]\d{4})(?!\d)", fn)
    return int(m.group(1)) if m else 0


@st.cache_data(ttl=300, show_spinner=False)
def scan_drive_pdfs():
    drive, _ = get_services()
    pdf_map = {}
    try:
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
        items = []
        try:
            resp = drive.files().list(
                q=f"'{folder['id']}' in parents and trashed=false",
                fields="files(id, name, mimeType, webViewLink)",
            ).execute()
            items.extend(resp.get("files", []))
        except Exception:
            continue

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

        for f in sub_files:
            if not f["name"].lower().endswith((".pdf", ".xlsx")):
                continue
            month = extract_month_from_filename(f["name"])
            if month is None:
                continue
            fdate = extract_file_date(f["name"])
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
    return FIRM_NAME_MAP.get(sheet_firm, sheet_firm)


@st.cache_data(ttl=300, show_spinner=False)
def build_dashboard_data():
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
        pdf_info = pdf_map.get(drive_folder, {}).get(month)
        has_pdf = pdf_info is not None
        pdf_link = pdf_info["link"] if pdf_info else ""
        pdf_name = pdf_info["name"] if pdf_info else ""

        if amount > 0 and has_pdf:
            status = "확정"
        elif amount == 0 and has_pdf:
            status = "⚠️ 법무비용 업데이트"
            amount = 0
        elif amount > 0 and not has_pdf:
            continue  # 예측 무시
        else:
            continue

        rows.append({
            "로펌": firm, "월": month, "기간": rec["기간"],
            "금액": amount, "상태": status,
            "PDF링크": pdf_link, "파일명": pdf_name,
        })

    # 드라이브에 PDF는 있는데 시트에 아예 없는 로펌
    reverse_map = {v: k for k, v in FIRM_NAME_MAP.items()}
    for folder_name, months in pdf_map.items():
        sheet_name = reverse_map.get(folder_name, folder_name)
        if sheet_name not in firms_seen:
            for month, info in months.items():
                rows.append({
                    "로펌": sheet_name, "월": month, "기간": f"2026-{month:02d}",
                    "금액": 0, "상태": "⚠️ 법무비용 업데이트",
                    "PDF링크": info["link"], "파일명": info["name"],
                })

    # 소계 검증
    if subtotal and rows:
        df_tmp = pd.DataFrame(rows)
        confirmed = df_tmp[df_tmp["상태"] == "확정"]
        for m, expected in subtotal.items():
            if expected == 0:
                continue
            actual = confirmed[confirmed["월"] == m]["금액"].sum()
            if actual > 0 and abs(actual - expected) > 100:
                errors.append(
                    f"⚠️ 소계 불일치 {m}월: 시트 ₩{expected:,.0f} vs 확정 합산 ₩{actual:,.0f} (차이: ₩{abs(actual-expected):,.0f})"
                )

    return rows, errors


# ============================================================
# UI
# ============================================================
def apply_custom_css():
    st.markdown(
        """
        <style>
        html, body, [class*="css"] { font-size: 16px; }
        div[data-testid="stMetric"] {
            background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
            border-radius: 12px; padding: 16px 20px;
            border-left: 4px solid #1B4F72;
        }
        div[data-testid="stMetric"] label { font-size: 0.95rem !important; color: #495057; }
        div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
            font-size: 1.6rem !important; font-weight: 700; color: #1B4F72;
        }
        section[data-testid="stSidebar"] { background: #f8f9fa; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def main():
    apply_custom_css()
    st.title("⚖️ 외부 로펌 비용 대시보드")
    st.caption("대웅제약 법무1팀 · 구글 시트 + 드라이브 PDF 연동")

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

    fdf = df[df["로펌"].isin(sel_firms)].copy()
    confirmed = fdf[fdf["상태"] == "확정"].copy()

    if confirmed.empty:
        st.warning("확정된 데이터가 없습니다.")
        updates = fdf[fdf["상태"] == "⚠️ 법무비용 업데이트"]
        if not updates.empty:
            st.subheader("⚠️ 법무비용 업데이트 필요")
            st.dataframe(
                updates[["로펌", "기간", "파일명", "PDF링크"]],
                column_config={"PDF링크": st.column_config.LinkColumn("PDF", display_text="📄 열기")},
                hide_index=True, use_container_width=True,
            )
        st.stop()

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

    # 최근 3개월 요약
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
                    html += f'<td style="padding:10px; text-align:center;"><a href="{link}" target="_blank" style="color:#1B4F72; text-decoration:none; font-weight:500;">₩{amt:,.0f}</a></td>'
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

    # 막대 차트
    periods_sorted = sorted(chart_df["집계"].unique())
    fig = go.Figure()
    for firm in sorted(chart_df["로펌"].unique()):
        fd = chart_df[chart_df["로펌"] == firm]
        fig.add_trace(
            go.Bar(
                x=fd["집계"], y=fd["금액"], name=firm,
                marker_color=FIRM_COLORS.get(firm, "#95A5A6"),
                text=fd["금액"].apply(lambda v: f"{v/10_000:,.0f}만"),
                textposition="inside", textangle=0, constraintext="none",
                textfont=dict(size=11, color="white"),
                hovertemplate="%{x}<br>%{fullData.name}: ₩%{y:,.0f}<extra></extra>",
            )
        )
    fig.update_layout(
        barmode="group",
        xaxis=dict(title="기간", categoryorder="array", categoryarray=periods_sorted, tickfont=dict(size=13)),
        yaxis=dict(title="금액 (원)", tickformat=",", tickfont=dict(size=13)),
        legend=dict(orientation="h", y=-0.18, x=0.5, xanchor="center", font=dict(size=14)),
        height=480, margin=dict(t=20, b=80), template="plotly_white",
        hoverlabel=dict(font_size=14),
    )
    st.plotly_chart(fig, use_container_width=True)

    # 업데이트 필요
    updates = fdf[fdf["상태"] == "⚠️ 법무비용 업데이트"]
    if not updates.empty:
        st.subheader("⚠️ 법무비용 업데이트 필요")
        st.caption("PDF가 드라이브에 있으나 시트에 금액이 미입력된 항목입니다.")
        utbl = updates[["로펌", "기간", "파일명", "PDF링크"]].copy()
        utbl = utbl.sort_values(["기간", "로펌"]).reset_index(drop=True)
        st.dataframe(
            utbl,
            column_config={"PDF링크": st.column_config.LinkColumn("PDF", display_text="📄 열기")},
            hide_index=True, use_container_width=True,
        )

    # 상세 내역
    st.subheader("📋 상세 내역")
    fc1, fc2 = st.columns(2)
    with fc1:
        detail_firms = st.multiselect("로펌 필터", firms_sorted, default=firms_sorted, key="detail_firm")
    with fc2:
        all_months = sorted(confirmed["기간"].unique(), reverse=True)
        detail_months = st.multiselect("월 필터", all_months, default=all_months, key="detail_month")

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
        hide_index=True, use_container_width=True,
    )

    # 피벗
    st.subheader("📊 로펌별 기간 합계")
    pivot = confirmed.pivot_table(
        values="금액", index="로펌", columns="기간",
        aggfunc="sum", fill_value=0, margins=True, margins_name="합계",
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
