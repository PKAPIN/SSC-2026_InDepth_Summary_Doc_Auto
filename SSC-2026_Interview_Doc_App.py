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

# SSL 인증서 검증 우회 (Mac 및 클라우드 네트워크 연결 오류 방지)
ssl._create_default_https_context = ssl._create_unverified_context

# Streamlit 웹 화면 기본 설정
st.set_page_config(
    page_title="심층면담 회의록 자동 생성 시스템",
    page_icon="📄",
    layout="wide"
)

st.title("📄 심층면담 회의록 문서 및 로그 자동 생성기")
st.markdown("🔗 **연동 시트**: [구글 스프레드시트 (심층면담_회의록)](https://docs.google.com/spreadsheets/d/1ws9JTAdRXwbp--NhrjWwelNorSTv1_LIJW7DijUtJLU/edit?gid=770556375#gid=770556375)")
st.caption("구글 시트의 최신 데이터를 실시간으로 읽어와 지정된 HWPX 회의록 서식으로 개별 문서 및 점검 로그(Excel/CSV)를 일괄 생성합니다.")

st.divider()

# 구글 시트 URL 설정
SHEET_ID = "1ws9JTAdRXwbp--NhrjWwelNorSTv1_LIJW7DijUtJLU"
GID = "770556375"  # '심층면담_회의록' 시트 GID
GOOGLE_SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"

# 실행 버튼
if st.button("🚀 실시간 데이터 읽기 및 회의록 자동 생성 시작", type="primary", use_container_width=True):
    
    # 지정된 저장소 내 HWPX 템플릿 탐색 (새로운 템플릿 업로드 기능 삭제)
    hwpx_files = glob.glob("*.hwpx")
    if not hwpx_files:
        st.error("❌ 저장소 내에서 지정된 .hwpx 템플릿 파일을 찾을 수 없습니다. GitHub 저장소에 .hwpx 파일을 올려주세요.")
        st.stop()
        
    template_path = hwpx_files[0]
    with open(template_path, "rb") as f:
        hwpx_bytes = f.read()

    with st.spinner("🔄 구글 시트에서 최신 데이터를 불러오는 중..."):
        try:
            df_raw = pd.read_csv(GOOGLE_SHEET_URL)
            data_df = df_raw.iloc[2:].dropna(subset=['문서ID']).copy()
        except Exception as e:
            st.error(f"❌ 구글 시트 데이터를 읽어오는 중 오류가 발생했습니다: {e}")
            st.stop()

    # HWPX 템플릿 구조 및 ZipInfo 보존 해석 (윈도우 한글 보안/손상 오류 방지)
    hwpx_zip = zipfile.ZipFile(io.BytesIO(hwpx_bytes), 'r')
    template_infolist = hwpx_zip.infolist()
    template_files = {info.filename: hwpx_zip.read(info.filename) for info in template_infolist}
    
    sec0_text = template_files['Contents/section0.xml'].decode('utf-8')
    pattern = r'<hp:stringParam name="Command">(.*?)</hp:stringParam>'
    commands = re.findall(pattern, sec0_text)
    
    merge_fields = []
    for cmd in commands:
        if cmd not in merge_fields:
            merge_fields.append(cmd)
            
    total_merge_cnt = len(merge_fields)  # 14개

    # 회의진행일시 순 정렬
    data_df['일시_dt'] = pd.to_datetime(data_df['일시'], errors='coerce')
    sorted_df = data_df.sort_values(by=['일시_dt', '시작시간', '학교명'], na_position='last').reset_index(drop=True)

    version_str = f"v.{datetime.datetime.now().strftime('%y%m%d_%H%M')}"
    folder_name = f"결과물_심층면담_회의록_{version_str}"

    # 개별 HWPX 바이너리 보관용 딕셔너리
    doc_bytes_dict = {}
    log_records = []
    
    progress_bar = st.progress(0)
    progress_text = st.empty()
    total_rows = len(sorted_df)

    # 1. 개별 HWPX 문서 생성 (ZipInfo 원본 압축방식 보존)
    for idx, (_, row) in enumerate(sorted_df.iterrows(), start=1):
        doc_id = str(row['문서ID']).strip()
        school = str(row['학교명']).strip() if pd.notna(row['학교명']) else ""
        round_num = str(row['회차']).strip() if pd.notna(row['회차']) else ""
        
        date_val = row['일시']
        if pd.isna(date_val):
            date_str = ""
        elif isinstance(date_val, (pd.Timestamp, datetime.datetime)):
            date_str = date_val.strftime('%Y-%m-%d')
        else:
            date_str = str(date_val).strip()
            
        xml_content = template_files['Contents/section0.xml'].decode('utf-8')
        missing_fields = []
        
        for col in data_df.columns:
            if col == '일시_dt':
                continue
            val = row[col]
            
            if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == "nan":
                val_str = ""
                if col in merge_fields:
                    missing_fields.append(col)
            elif isinstance(val, (pd.Timestamp, datetime.datetime)):
                val_str = val.strftime('%Y-%m-%d')
            elif isinstance(val, datetime.time):
                val_str = val.strftime('%H:%M')
            else:
                val_str = str(val).strip()
                
            val_str = val_str.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            xml_content = xml_content.replace(f"{{{{{col}}}}}", val_str)
            
        # HWPX 원본 압축구조(mimetype STORED 등) 그대로 재압축하여 윈도우 한글 보안경고 완전 방지
        doc_buffer = io.BytesIO()
        with zipfile.ZipFile(doc_buffer, 'w') as z_out:
            for info in template_infolist:
                fname = info.filename
                if fname == 'Contents/section0.xml':
                    content_bytes = xml_content.encode('utf-8')
                else:
                    content_bytes = template_files[fname]
                    
                new_info = zipfile.ZipInfo(fname)
                new_info.compress_type = info.compress_type
                z_out.writestr(new_info, content_bytes)
                
        doc_bytes_dict[doc_id] = doc_buffer.getvalue()
        
        missing_cnt = len(missing_fields)
        ratio_str = f"{missing_cnt}건 / {total_merge_cnt}건"
        status = "정상" if missing_cnt == 0 else "일부항목누락"
        
        row_dict = {
            "생성번호": idx,
            "문서ID": doc_id,
            "학교명": school,
            "회의일시": date_str,
            "회차": round_num,
            "생성상태": status,
            "누락현황(누락/전체)": ratio_str
        }
        
        for field in merge_fields:
            if field in missing_fields:
                row_dict[f"[점검]{field}"] = f"{field} N/A"
            else:
                row_dict[f"[점검]{field}"] = "-"
                
        log_records.append(row_dict)
        
        progress_bar.progress(idx / total_rows)
        progress_text.text(f"⚡ HWPX 회의록 생성 중... [{idx}/{total_rows}] {doc_id}.hwpx")

    log_df = pd.DataFrame(log_records)
    
    # 2. 엑셀 로그 생성 (빨간 글씨 서식 적용)
    excel_buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "생성로그"
    
    headers = list(log_df.columns)
    ws.append(headers)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    red_text_font = Font(color="C00000", bold=True)
    warning_text_font = Font(color="9C6500", bold=True)
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for r_idx, row_data in enumerate(log_records, start=2):
        for c_idx, col_name in enumerate(headers, start=1):
            val = row_data[col_name]
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            
            if col_name in ["생성번호", "회의일시", "회차", "생성상태", "누락현황(누락/전체)"]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
            if col_name.startswith("[점검]") and str(val).endswith("N/A"):
                cell.font = red_text_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name == "생성상태" and val == "일부항목누락":
                cell.font = warning_text_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
    wb.save(excel_buffer)
    excel_bytes = excel_buffer.getvalue()
    csv_bytes = log_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')

    # 3. 전체 ZIP 패키징 생성
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', compression=zipfile.ZIP_DEFLATED) as main_zip:
        for doc_id, b_data in doc_bytes_dict.items():
            main_zip.writestr(f"{folder_name}/{doc_id}.hwpx", b_data)
            
        excel_filename = f"심층면담_서류_Log_{version_str}.xlsx"
        csv_filename = f"심층면담_서류_Log_{version_str}.csv"
        
        main_zip.writestr(f"{folder_name}/{excel_filename}", excel_bytes)
        main_zip.writestr(f"{folder_name}/{csv_filename}", csv_bytes)

    # st.session_state에 저장하여 다운로드 시 화면 자동 리셋(로그 사라짐) 방지!
    st.session_state['generated_data'] = {
        'folder_name': folder_name,
        'version_str': version_str,
        'doc_bytes_dict': doc_bytes_dict,
        'all_zip_bytes': zip_buffer.getvalue(),
        'excel_bytes': excel_bytes,
        'csv_bytes': csv_bytes,
        'log_df': log_df,
        'doc_ids': list(doc_bytes_dict.keys())
    }

# 생성 결과 화면 출력 (st.session_state 이용으로 다운로드 시에도 사라지지 않고 영구 유지)
if 'generated_data' in st.session_state:
    data = st.session_state['generated_data']
    st.success(f"🎉 구글 시트 연결 성공! 총 **{len(data['doc_ids'])}건**의 회의록 및 점검 로그 생성이 완료되었습니다.")

    st.subheader("📥 다운로드 옵션 선택")
    
    download_tab1, download_tab2 = st.tabs(["📦 전체 일괄 다운로드 (ZIP)", "📄 개별/선택 문서 다운로드"])
    
    with download_tab1:
        st.markdown(f"**전체 {len(data['doc_ids'])}개 HWPX 회의록 문서 + 엑셀/CSV 로그**가 하나의 압축파일에 포함되어 있습니다.")
        st.download_button(
            label=f"📦 전체 일괄 다운로드 ({data['folder_name']}.zip)",
            data=data['all_zip_bytes'],
            file_name=f"{data['folder_name']}.zip",
            mime="application/zip",
            use_container_width=True,
            key="btn_download_all"
        )
        
    with download_tab2:
        st.markdown("다운로드받고자 하는 문서를 직접 선택하세요 (단일 파일 또는 선택 패키지 ZIP 다운로드).")
        
        selected_docs = st.multiselect(
            "다운로드할 회의록 문서를 선택하세요:",
            options=data['doc_ids'],
            default=data['doc_ids'][:1] if data['doc_ids'] else []
        )
        
        if len(selected_docs) == 1:
            # 단일 선택 시 바로 HWPX 다운로드
            target_id = selected_docs[0]
            st.download_button(
                label=f"📄 [{target_id}.hwpx] 선택 문서 다운로드",
                data=data['doc_bytes_dict'][target_id],
                file_name=f"{target_id}.hwpx",
                mime="application/hwp+zip",
                use_container_width=True,
                key="btn_download_single"
            )
        elif len(selected_docs) > 1:
            # 여러 개 선택 시 선택 항목 모음 ZIP 다운로드
            sub_zip_buffer = io.BytesIO()
            with zipfile.ZipFile(sub_zip_buffer, 'w', compression=zipfile.ZIP_DEFLATED) as sub_zip:
                for target_id in selected_docs:
                    if target_id in data['doc_bytes_dict']:
                        sub_zip.writestr(f"{data['folder_name']}_선택문서/{target_id}.hwpx", data['doc_bytes_dict'][target_id])
            
            st.download_button(
                label=f"📦 선택한 문서 {len(selected_docs)}개 모음 다운로드 (.zip)",
                data=sub_zip_buffer.getvalue(),
                file_name=f"{data['folder_name']}_선택문서_{len(selected_docs)}건.zip",
                mime="application/zip",
                use_container_width=True,
                key="btn_download_multi"
            )
        else:
            st.warning("⚠️ 다운로드할 문서를 1개 이상 선택해 주세요.")

    st.divider()
    
    # 생성 로그 미리보기 표 (화면 리셋 없이 계속 유지됨)
    st.subheader("📋 생성 로그 및 점검 현황 미리보기")
    st.dataframe(data['log_df'], use_container_width=True)