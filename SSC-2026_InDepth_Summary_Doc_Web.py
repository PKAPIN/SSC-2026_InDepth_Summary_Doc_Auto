import io
import os
import glob
import zipfile
import datetime
import re
import ssl
import html
import xml.etree.ElementTree as ET

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import streamlit as st

# =========================================================================
# [웹 호스팅용] 심층면담 회의록 및 점검 로그 자동 생성 Streamlit 웹 앱
# =========================================================================

ssl._create_default_https_context = ssl._create_unverified_context

st.set_page_config(page_title="심층면담 회의록 자동 생성 시스템", page_icon="📄", layout="wide")

st.title("📄 심층면담 회의록 문서 및 로그 자동 생성기")
st.markdown("🔗 **연동 시트**: [구글 스프레드시트 (심층면담_회의록)](https://docs.google.com/spreadsheets/d/1ws9JTAdRXwbp--NhrjWwelNorSTv1_LIJW7DijUtJLU/edit?gid=770556375#gid=770556375)")
st.caption("구글 시트의 최신 데이터를 실시간으로 읽어와 지정된 HWPX 회의록 서식으로 개별 문서 및 점검 로그를 일괄 생성합니다.")

st.divider()

SHEET_ID = "1ws9JTAdRXwbp--NhrjWwelNorSTv1_LIJW7DijUtJLU"
GID = "770556375"
GOOGLE_SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"

FIELD_PATTERN = re.compile(r"\{\{([^{}]+)\}\}")

# -------------------------------------------------------------------------
# HWPX XML 안전 치환 및 줄바꿈 헬퍼 함수
# -------------------------------------------------------------------------
def xml_escape(text):
    """XML 텍스트 노드용 이스케이프"""
    return html.escape(str(text), quote=False)

def render_text_with_linebreaks(text):
    """\n 줄바꿈을 hp:lineBreak로 변환 (hp:p / hp:run 구조 보존)"""
    lines = str(text).replace("\r\n", "\n").replace("\r", "\n").split("\n")
    result = []
    for i, line in enumerate(lines):
        result.append(f"<hp:t>{xml_escape(line)}</hp:t>")
        if i < len(lines) - 1:
            result.append("<hp:lineBreak/>")
    return "".join(result)

def replace_fields_inside_t_nodes(xml_content, field_values):
    """<hp:t> 노드 내부의 {{필드명}}만 선택적으로 치환하여 XML 구조 파괴 방지"""
    replacement_count = {field: 0 for field in field_values.keys()}
    t_pattern = re.compile(r"(<hp:t(?:\s[^>]*)?>)(.*?)(</hp:t>)", re.DOTALL)

    def replace_one_t(match):
        raw_text = match.group(2)
        plain_text = html.unescape(raw_text)
        matches = list(FIELD_PATTERN.finditer(plain_text))
        
        if not matches:
            return match.group(0)

        result = []
        cursor = 0
        for field_match in matches:
            before = plain_text[cursor:field_match.start()]
            field_name = field_match.group(1).strip()

            if field_name not in field_values:
                result.append(xml_escape(before))
                result.append(xml_escape(field_match.group(0)))
                cursor = field_match.end()
                continue

            result.append(xml_escape(before))
            value = field_values[field_name]
            replacement_count[field_name] += 1
            result.append(render_text_with_linebreaks(value))
            cursor = field_match.end()

        result.append(xml_escape(plain_text[cursor:]))
        return "".join(result)

    return t_pattern.sub(replace_one_t, xml_content), replacement_count

# -------------------------------------------------------------------------
# XML 및 ZIP 안전 검사 함수
# -------------------------------------------------------------------------
def validate_xml(xml_content):
    try:
        ET.fromstring(xml_content.encode("utf-8"))
        return True, ""
    except Exception as e:
        return False, str(e)

def validate_hwpx(hwpx_bytes):
    try:
        with zipfile.ZipFile(io.BytesIO(hwpx_bytes), "r") as test_zip:
            if test_zip.testzip() is not None:
                return False, f"ZIP 손상 파일 발견"
            if "Contents/section0.xml" not in test_zip.namelist():
                return False, "Contents/section0.xml 누락"
        return True, ""
    except Exception as e:
        return False, str(e)

# =========================================================================
# 메인 실행 로직
# =========================================================================
if st.button("🚀 실시간 데이터 읽기 및 회의록 자동 생성 시작", type="primary", use_container_width=True):
    
    # 1. HWPX 템플릿 검색
    hwpx_files = [f for f in glob.glob("*.hwpx") if not os.path.basename(f).startswith("~$")]
    if not hwpx_files:
        st.error("❌ 저장소 내에서 .hwpx 템플릿 파일을 찾을 수 없습니다.")
        st.stop()

    template_path = max(hwpx_files, key=os.path.getmtime)
    st.info(f"📄 사용 템플릿: `{os.path.basename(template_path)}`")
    
    with open(template_path, "rb") as f:
        hwpx_bytes = f.read()

    # 2. 구글 시트 데이터 읽기
    with st.spinner("🔄 구글 시트에서 최신 데이터를 불러오는 중..."):
        try:
            df_raw = pd.read_csv(GOOGLE_SHEET_URL)
            data_df = df_raw.iloc[2:].dropna(subset=["문서ID"]).copy()
        except Exception as e:
            st.error(f"❌ 구글 시트 데이터를 읽는 중 오류가 발생했습니다: {e}")
            st.stop()

    if len(data_df) == 0:
        st.error("❌ 생성할 데이터가 없습니다.")
        st.stop()

    # 3. 원본 HWPX ZIP 열기
    try:
        with zipfile.ZipFile(io.BytesIO(hwpx_bytes), "r") as hwpx_zip:
            template_infolist = hwpx_zip.infolist()
            template_files = {info.filename: hwpx_zip.read(info.filename) for info in template_infolist}
    except Exception as e:
        st.error(f"❌ HWPX 템플릿을 열 수 없습니다: {e}")
        st.stop()

    section_path = "Contents/section0.xml"
    if section_path not in template_files:
        st.error("❌ HWPX 내부에서 Contents/section0.xml을 찾을 수 없습니다.")
        st.stop()

    original_xml = template_files[section_path].decode("utf-8")

    # 4. 필드 탐색
    detected_fields = [f.strip() for f in FIELD_PATTERN.findall(original_xml) if f.strip()]
    merge_fields = []
    for field in detected_fields:
        if field not in merge_fields:
            merge_fields.append(field)

    commands = re.findall(r'<hp:stringParam name="Command">(.*?)</hp:stringParam>', original_xml)
    for cmd in commands:
        cmd = re.sub(r"^\{\{|\}\}$", "", html.unescape(cmd).strip()).strip()
        if cmd and cmd not in merge_fields:
            merge_fields.append(cmd)

    if not merge_fields:
        st.error("❌ section0.xml에서 `{{필드명}}` 형태의 치환 필드를 찾지 못했습니다.")
        st.stop()

    st.success(f"🔎 HWPX 확인된 필드: {', '.join(merge_fields)}")

    # 5. 정렬 및 생성 준비
    data_df["일시_dt"] = pd.to_datetime(data_df["일시"], errors="coerce") if "일시" in data_df.columns else pd.NaT
    sort_cols = [c for c in ["일시_dt", "시작시간", "학교명"] if c in data_df.columns]
    sorted_df = data_df.sort_values(by=sort_cols, na_position="last").reset_index(drop=True) if sort_cols else data_df.reset_index(drop=True)

    version_str = f"v.{datetime.datetime.now().strftime('%y%m%d_%H%M')}"
    folder_name = f"결과물_심층면담_회의록_{version_str}"
    doc_bytes_dict, log_records = {}, []

    progress_bar = st.progress(0)
    progress_text = st.empty()
    total_rows = len(sorted_df)

    # 6. 개별 문서 생성 루프
    for idx, (_, row) in enumerate(sorted_df.iterrows(), start=1):
        doc_id = str(row.get("문서ID", "")).strip()
        school = str(row.get("학교명", "")).strip()
        round_num = str(row.get("회차", "")).strip()

        field_values, missing_fields = {}, []
        for col in data_df.columns:
            if col == "일시_dt": continue
            val = row[col]
            if pd.isna(val):
                value_text = ""
            elif isinstance(val, (pd.Timestamp, datetime.datetime)):
                value_text = val.strftime("%Y-%m-%d")
            elif isinstance(val, datetime.time):
                value_text = val.strftime("%H:%M")
            else:
                value_text = str(val)

            if value_text.strip() == "":
                value_text = ""
                if col in merge_fields: missing_fields.append(col)

            field_values[col] = value_text

        # 정밀 XML 치환 및 문법 검사
        xml_content, _ = replace_fields_inside_t_nodes(original_xml, field_values)
        xml_ok, xml_error = validate_xml(xml_content)
        if not xml_ok:
            st.error(f"❌ {doc_id} XML 검사 실패: {xml_error}")
            st.stop()

        # HWPX ZIP 재압축 및 검사
        doc_buffer = io.BytesIO()
        with zipfile.ZipFile(doc_buffer, "w") as z_out:
            for info in template_infolist:
                content_bytes = xml_content.encode("utf-8") if info.filename == section_path else template_files[info.filename]
                z_out.writestr(info, content_bytes)

        generated_bytes = doc_buffer.getvalue()
        hwpx_ok, hwpx_error = validate_hwpx(generated_bytes)
        if not hwpx_ok:
            st.error(f"❌ {doc_id}.hwpx 검사 실패: {hwpx_error}")
            st.stop()

        doc_bytes_dict[doc_id] = generated_bytes

        # 로그 기록
        missing_cnt = len(missing_fields)
        row_dict = {
            "선택": False, "생성번호": idx, "문서ID": doc_id, "학교명": school,
            "회의일시": field_values.get("일시", ""), "회차": round_num,
            "생성상태": "정상" if missing_cnt == 0 else "일부항목누락",
            "누락현황(누락/전체)": f"{missing_cnt}건 / {len(merge_fields)}건"
        }
        for field in merge_fields:
            row_dict[f"[점검]{field}"] = f"{field} N/A" if field in missing_fields else "-"
        log_records.append(row_dict)

        progress_bar.progress(idx / total_rows)
        progress_text.text(f"⚡ HWPX 회의록 생성 중... [{idx}/{total_rows}] {doc_id}.hwpx")

    # 7. 엑셀/CSV 로그 생성
    log_df = pd.DataFrame(log_records)
    excel_log_df = log_df.drop(columns=["선택"])
    excel_buffer = io.BytesIO()
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "생성로그"
    headers = list(excel_log_df.columns)
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    red_font = Font(color="C00000", bold=True)
    warning_font = Font(color="9C6500", bold=True)
    thin_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal="center", vertical="center")

    for r_idx, row_data in excel_log_df.iterrows():
        excel_r = r_idx + 2
        for c_idx, col_name in enumerate(headers, start=1):
            val = row_data[col_name]
            cell = ws.cell(row=excel_r, column=c_idx, value=val)
            cell.border = thin_border
            if col_name in ["생성번호", "회의일시", "회차", "생성상태", "누락현황(누락/전체)"]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if col_name.startswith("[점검]") and str(val).endswith("N/A"):
                cell.font, cell.alignment = red_font, Alignment(horizontal="center", vertical="center")
            elif col_name == "생성상태" and val == "일부항목누락":
                cell.font, cell.alignment = warning_font, Alignment(horizontal="center", vertical="center")

    wb.save(excel_buffer)
    excel_bytes = excel_buffer.getvalue()
    csv_bytes = excel_log_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

    # 8. 전체 ZIP 압축 및 세션 저장
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as main_zip:
        for doc_id, b_data in doc_bytes_dict.items():
            main_zip.writestr(f"{folder_name}/{doc_id}.hwpx", b_data)
        main_zip.writestr(f"{folder_name}/심층면담_서류_Log_{version_str}.xlsx", excel_bytes)
        main_zip.writestr(f"{folder_name}/심층면담_서류_Log_{version_str}.csv", csv_bytes)

    st.session_state["generated_data"] = {
        "folder_name": folder_name, "version_str": version_str, "doc_bytes_dict": doc_bytes_dict,
        "all_zip_bytes": zip_buffer.getvalue(), "excel_bytes": excel_bytes, "csv_bytes": csv_bytes,
        "log_df": log_df, "doc_ids": list(doc_bytes_dict.keys())
    }

    st.success(f"🎉 총 **{len(doc_bytes_dict)}건**의 HWPX 회의록 생성 완료 (XML 및 ZIP 검사 통과)")

# =========================================================================
# 생성 결과 화면 및 선택 다운로드 UI
# =========================================================================
if "generated_data" in st.session_state:
    data = st.session_state["generated_data"]
    st.subheader("📋 생성로그 미리보기 & 선택 문서 다운로드")
    st.markdown("아래 표의 맨 앞 **`선택` 체크박스**를 클릭하여 다운로드할 회의록 문서를 지정할 수 있습니다.")

    col_btn1, col_btn2, col_btn3, col_btn4 = st.columns([1.5, 1, 1, 1.5])

    with col_btn1:
        st.download_button(
            label=f"📦 전체 일괄 다운로드 ({len(data['doc_ids'])}건 ZIP)",
            data=data["all_zip_bytes"], file_name=f"{data['folder_name']}.zip",
            mime="application/zip", use_container_width=True, key="btn_download_all"
        )

    with col_btn2:
        if st.button("☑️ 전체 선택", use_container_width=True):
            data["log_df"]["선택"] = True
            st.rerun()

    with col_btn3:
        if st.button("⬜ 전체 해제", use_container_width=True):
            data["log_df"]["선택"] = False
            st.rerun()

    edited_df = st.data_editor(
        data["log_df"],
        column_config={"선택": st.column_config.CheckboxColumn("선택", help="다운로드할 문서를 선택하세요", default=False)},
        disabled=[col for col in data["log_df"].columns if col != "선택"],
        hide_index=True, use_container_width=True, key="log_data_editor"
    )

    selected_rows = edited_df[edited_df["선택"] == True]
    selected_doc_ids = selected_rows["문서ID"].tolist()

    with col_btn4:
        if len(selected_doc_ids) == 0:
            st.button("📄 선택 문서 다운로드 (0건)", disabled=True, use_container_width=True)
        elif len(selected_doc_ids) == 1:
            target_id = selected_doc_ids[0]
            st.download_button(
                label="📄 선택 문서 다운로드 (1건 .hwpx)",
                data=data["doc_bytes_dict"][target_id], file_name=f"{target_id}.hwpx",
                mime="application/hwp+zip", use_container_width=True, key="btn_download_single_sel"
            )
        else:
            sub_zip_buffer = io.BytesIO()
            with zipfile.ZipFile(sub_zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as sub_zip:
                for target_id in selected_doc_ids:
                    if target_id in data["doc_bytes_dict"]:
                        sub_zip.writestr(f"{data['folder_name']}_선택문서/{target_id}.hwpx", data["doc_bytes_dict"][target_id])
            
            st.download_button(
                label=f"📦 선택 문서 다운로드 ({len(selected_doc_ids)}건 ZIP)",
                data=sub_zip_buffer.getvalue(), file_name=f"{data['folder_name']}_선택문서_{len(selected_doc_ids)}건.zip",
                mime="application/zip", use_container_width=True, key="btn_download_multi_sel"
            )

    st.caption(f"💡 현재 **{len(selected_doc_ids)}개**의 문서가 선택되었습니다.")
