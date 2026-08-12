# =========================================================================
# [웹 호스팅용] 심층면담 회의록 및 점검 로그 자동 생성 Streamlit 웹 앱
# - 모든 줄바꿈(\n) 시 첫 행 스타일(10pt/맑은고딕/120%) 100% 완벽 통일판
# =========================================================================
import io
import os
import glob
import zipfile
import datetime
import re
import ssl
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import streamlit as st

ssl._create_default_https_context = ssl._create_unverified_context

st.set_page_config(page_title="심층면담 회의록 자동 생성 시스템", page_icon="📄", layout="wide")

st.title("📄 심층면담 회의록 문서 및 로그 자동 생성기")
st.markdown("🔗 **연동 시트**: [구글 스프레드시트 (심층면담_회의록)](https://docs.google.com/spreadsheets/d/1ws9JTAdRXwbp--NhrjWwelNorSTv1_LIJW7DijUtJLU/edit?gid=770556375#gid=770556375)")
st.caption("구글 시트의 최신 데이터를 실시간으로 읽어와 지정된 HWPX 회의록 서식으로 개별 문서 및 점검 로그(Excel/CSV)를 일괄 생성합니다.")

st.divider()

SHEET_ID = "1ws9JTAdRXwbp--NhrjWwelNorSTv1_LIJW7DijUtJLU"
GID = "770556375"
GOOGLE_SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"

if st.button("🚀 실시간 데이터 읽기 및 회의록 자동 생성 시작", type="primary", use_container_width=True):
    
    hwpx_files = [f for f in glob.glob("*.hwpx") if not os.path.basename(f).startswith("~$")]
    if not hwpx_files:
        st.error("❌ 저장소 내에서 지정된 .hwpx 템플릿 파일을 찾을 수 없습니다. GitHub 저장소에 .hwpx 파일을 올려주세요.")
        st.stop()
        
    template_path = max(hwpx_files, key=os.path.getmtime)
    with open(template_path, "rb") as f:
        hwpx_bytes = f.read()

    with st.spinner("🔄 구글 시트에서 최신 데이터를 불러오는 중..."):
        try:
            df_raw = pd.read_csv(GOOGLE_SHEET_URL)
            data_df = df_raw.iloc[2:].dropna(subset=['문서ID']).copy()
        except Exception as e:
            st.error(f"❌ 구글 시트 데이터를 읽어오는 중 오류가 발생했습니다: {e}")
            st.stop()

    hwpx_zip = zipfile.ZipFile(io.BytesIO(hwpx_bytes), 'r')
    template_infolist = hwpx_zip.infolist()
    template_files = {info.filename: hwpx_zip.read(info.filename) for info in template_infolist}
    
    sec0_text = template_files['Contents/section0.xml'].decode('utf-8')
    commands = re.findall(r'<hp:stringParam name="Command">(.*?)</hp:stringParam>', sec0_text)
    
    merge_fields = []
    for cmd in commands:
        if cmd not in merge_fields:
            merge_fields.append(cmd)
    total_merge_cnt = len(merge_fields)

    data_df['일시_dt'] = pd.to_datetime(data_df['일시'], errors='coerce')
    sorted_df = data_df.sort_values(by=['일시_dt', '시작시간', '학교명'], na_position='last').reset_index(drop=True)

    version_str = f"v.{datetime.datetime.now().strftime('%y%m%d_%H%M')}"
    folder_name = f"결과물_심층면담_회의록_{version_str}"

    doc_bytes_dict = {}
    log_records = []
    
    progress_bar = st.progress(0)
    progress_text = st.empty()
    total_rows = len(sorted_df)

    for idx, (_, row) in enumerate(sorted_df.iterrows(), start=1):
        doc_id = str(row['문서ID']).strip()
        school = str(row['학교명']).strip() if pd.notna(row['학교명']) else ""
        round_num = str(row['회차']).strip() if pd.notna(row['회차']) else ""
        
        date_val = row['일시']
        date_str = date_val.strftime('%Y-%m-%d') if isinstance(date_val, (pd.Timestamp, datetime.datetime)) else (str(date_val).strip() if pd.notna(date_val) else "")
            
        xml_content = template_files['Contents/section0.xml'].decode('utf-8')
        missing_fields = []
        
        for col in data_df.columns:
            if col == '일시_dt': continue
            val = row[col]
            if pd.isna(val) or str(val).strip() == "" or str(str(val)).strip().lower() == "nan":
                val_str = ""
                if col in merge_fields: missing_fields.append(col)
            elif isinstance(val, (pd.Timestamp, datetime.datetime)): val_str = val.strftime('%Y-%m-%d')
            elif isinstance(val, datetime.time): val_str = val.strftime('%H:%M')
            else: val_str = str(val).strip()
                
            # 특수문자 이스케이프 및 줄바꿈 정돈
            val_str = val_str.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            val_str = val_str.replace("\r\n", "\n").replace("\r", "\n")
            val_str = val_str.replace("\n\n", "\n")
            
            target_field = f"{{{{{col}}}}}"
            if target_field in xml_content:
                field_pos = xml_content.find(target_field)
                
                # 🎯 [핵심] 치환 필드가 위치한 직전의 <hp:p ...> 및 <hp:run ...> 속성 태그를 정확히 추출
                p_matches = list(re.finditer(r'<hp:p\b[^>]*>', xml_content[:field_pos]))
                open_p_tag = p_matches[-1].group(0) if p_matches else '<hp:p>'
                
                run_matches = list(re.finditer(r'<hp:run\b[^>]*>', xml_content[:field_pos]))
                open_run_tag = run_matches[-1].group(0) if run_matches else '<hp:run>'
                
                # 엔터(\n) 발생 시 추출한 속성 태그(10pt / 맑은고딕 / 120%)를 그대로 주입하여 똑같은 스타일로 문단 생성
                paragraph_replace = f'</hp:t></hp:run></hp:p>{open_p_tag}{open_run_tag}<hp:t>'
                val_str = val_str.replace("\n", paragraph_replace)
                
                xml_content = xml_content.replace(target_field, val_str)
            
        # 🧹 메일머지 표시 태그 및 레이아웃 캐시 정돈
        xml_content = re.sub(r'<hp:ctrl><hp:fieldBegin.*?</hp:ctrl>', '', xml_content, flags=re.DOTALL)
        xml_content = re.sub(r'<hp:ctrl><hp:fieldEnd.*?</hp:ctrl>', '', xml_content, flags=re.DOTALL)
        xml_content = re.sub(r'<hp:linesegarray>.*?</hp:linesegarray>', '<hp:linesegarray/>', xml_content, flags=re.DOTALL)
        xml_content = re.sub(r'(<hp:t>\s*</hp:t>\s*<hp:t>\s*,\s*</hp:t>)+', '', xml_content)

        doc_buffer = io.BytesIO()
        with zipfile.ZipFile(doc_buffer, 'w') as z_out:
            for info in template_infolist:
                fname = info.filename
                content_bytes = xml_content.encode('utf-8') if fname == 'Contents/section0.xml' else template_files[fname]
                new_info = zipfile.ZipInfo(fname)
                new_info.compress_type = info.compress_type
                z_out.writestr(new_info, content_bytes)
                
        doc_bytes_dict[doc_id] = doc_buffer.getvalue()
        
        missing_cnt = len(missing_fields)
        ratio_str = f"{missing_cnt}건 / {total_merge_cnt}건"
        status = "정상" if missing_cnt == 0 else "일부항목누락"
        
        row_dict = {
            "선택": False, "생성번호": idx, "문서ID": doc_id, "학교명": school,
            "회의일시": date_str, "회차": round_num, "생성상태": status, "누락현황(누락/전체)": ratio_str
        }
        for field in merge_fields:
            row_dict[f"[점검]{field}"] = f"{field} N/A" if field in missing_fields else "-"
        log_records.append(row_dict)
        
        progress_bar.progress(idx / total_rows)
        progress_text.text(f"⚡ HWPX 회의록 자동 생성 중... [{idx}/{total_rows}] {doc_id}.hwpx")

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
    red_text_font = Font(color="C00000", bold=True)
    warning_text_font = Font(color="9C6500", bold=True)
    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill; cell.font = header_font; cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for r_idx, row_data in excel_log_df.iterrows():
        excel_r = r_idx + 2
        for c_idx, col_name in enumerate(headers, start=1):
            val = row_data[col_name]
            cell = ws.cell(row=excel_r, column=c_idx, value=val)
            cell.border = thin_border
            if col_name in ["생성번호", "회의일시", "회차", "생성상태", "누락현황(누락/전체)"]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if col_name.startswith("[점검]") and str(val).endswith("N/A"):
                cell.font = red_text_font; cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name == "생성상태" and val == "일부항목누락":
                cell.font = warning_text_font; cell.alignment = Alignment(horizontal="center", vertical="center")
                
    wb.save(excel_buffer)
    excel_bytes = excel_buffer.getvalue()
    csv_bytes = excel_log_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', compression=zipfile.ZIP_DEFLATED) as main_zip:
        for doc_id, b_data in doc_bytes_dict.items():
            main_zip.writestr(f"{folder_name}/{doc_id}.hwpx", b_data)
            
        main_zip.writestr(f"{folder_name}/심층면담_서류_Log_{version_str}.xlsx", excel_bytes)
        main_zip.writestr(f"{folder_name}/심층면담_서류_Log_{version_str}.csv", csv_bytes)

    st.session_state['generated_data'] = {
        'folder_name': folder_name, 'version_str': version_str,
        'doc_bytes_dict': doc_bytes_dict, 'all_zip_bytes': zip_buffer.getvalue(),
        'excel_bytes': excel_bytes, 'csv_bytes': csv_bytes,
        'log_df': log_df, 'doc_ids': list(doc_bytes_dict.keys())
    }

if 'generated_data' in st.session_state:
    data = st.session_state['generated_data']
    st.success(f"🎉 구글 시트 연결 성공! 총 **{len(data['doc_ids'])}건**의 회의록 및 점검 로그가 시간순으로 생성되었습니다.")

    st.subheader("📋 생성로그 미리보기 {시간순} & 선택 문서 다운로드")
    col_btn1, col_btn2, col_btn3, col_btn4 = st.columns([1.5, 1, 1, 1.5])
    
    with col_btn1:
        st.download_button(
            label=f"📦 전체 일괄 다운로드 ({len(data['doc_ids'])}건 ZIP)",
            data=data['all_zip_bytes'], file_name=f"{data['folder_name']}.zip",
            mime="application/zip", use_container_width=True, key="btn_download_all"
        )
        
    with col_btn2:
        if st.button("☑️ 전체 선택", use_container_width=True): data['log_df']['선택'] = True
            
    with col_btn3:
        if st.button("⬜ 전체 해제", use_container_width=True): data['log_df']['선택'] = False

    edited_df = st.data_editor(
        data['log_df'],
        column_config={"선택": st.column_config.CheckboxColumn("선택", help="다운로드할 문서 항목을 체크하세요", default=False)},
        disabled=[col for col in data['log_df'].columns if col != "선택"],
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
                label=f"📄 선택 문서 다운로드 (1건 .hwpx)",
                data=data['doc_bytes_dict'][target_id], file_name=f"{target_id}.hwpx",
                mime="application/hwp+zip", use_container_width=True, key="btn_download_single_sel"
            )
        else:
            sub_zip_buffer = io.BytesIO()
            with zipfile.ZipFile(sub_zip_buffer, 'w', compression=zipfile.ZIP_DEFLATED) as sub_zip:
                for target_id in selected_doc_ids:
                    if target_id in data['doc_bytes_dict']:
                        sub_zip.writestr(f"{data['folder_name']}_선택문서/{target_id}.hwpx", data['doc_bytes_dict'][target_id])
            
            st.download_button(
                label=f"📦 선택 문서 다운로드 ({len(selected_doc_ids)}건 ZIP)",
                data=sub_zip_buffer.getvalue(), file_name=f"{data['folder_name']}_선택문서_{len(selected_doc_ids)}건.zip",
                mime="application/zip", use_container_width=True, key="btn_download_multi_sel"
            )

    st.caption(f"💡 현재 **{len(selected_doc_ids)}개**의 문서가 선택되었습니다.")
