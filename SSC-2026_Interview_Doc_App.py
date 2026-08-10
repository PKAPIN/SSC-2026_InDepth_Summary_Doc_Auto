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

# Streamlit 웹 화면 설정
st.set_page_config(
    page_title="심층면담 회의록 자동 생성기",
    page_icon="📄",
    layout="wide"
)

st.title("📄 심층면담 회의록 문서 자동화 {로그 생성}")
st.markdown("🔗 연동 시트: **[구글 스프레드시트 (심층면담_회의록)](https://docs.google.com/spreadsheets/d/1ws9JTAdRXwbp--NhrjWwelNorSTv1_LIJW7DijUtJLU/edit?gid=770556375#gid=770556375)**")
st.caption("구글 시트의 심층면담 일정의 최신 데이터를 실시간으로 읽어와 HWPX 회의록과 점검 로그를 일괄 생성합니다.")

st.divider()

# 구글 시트 정보
SHEET_ID = "1ws9JTAdRXwbp--NhrjWwelNorSTv1_LIJW7DijUtJLU"
GID = "770556375"  # '심층면담_회의록' 시트 GID
GOOGLE_SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"

# 파일 업로드 옵션 (선택 사항)
uploaded_file = st.file_uploader("새로운 .hwpx 템플릿 파일이 있다면 업로드하세요 (없으면 기본 템플릿 자동 사용)", type=["hwpx"])

# 실행 버튼
if st.button("🚀 실시간 데이터 읽기 및 회의록 자동 생성 시작", type="primary", use_container_width=True):
    
    # 템플릿 파일 탐색 (1. 업로드 파일 우선 -> 2. 저장소 내 .hwpx 파일 탐색)
    if uploaded_file is not None:
        hwpx_bytes = uploaded_file.getvalue()
        st.info("💡 업로드된 템플릿 파일을 사용합니다.")
    else:
        hwpx_files = glob.glob("*.hwpx")
        if hwpx_files:
            with open(hwpx_files[0], "rb") as f:
                hwpx_bytes = f.read()
            st.info(f"💡 저장소 기본 템플릿 사용: {os.path.basename(hwpx_files[0])}")
        else:
            st.error("❌ .hwpx 템플릿 파일을 찾을 수 없습니다. GitHub 저장소에 .hwpx 파일을 업로드해 주세요.")
            st.stop()

    with st.spinner("🔄 구글 시트에서 최신 데이터를 불러오는 중..."):
        try:
            df_raw = pd.read_csv(GOOGLE_SHEET_URL)
            data_df = df_raw.iloc[2:].dropna(subset=['문서ID']).copy()
        except Exception as e:
            st.error(f"❌ 구글 시트 데이터를 읽어오는 중 오류가 발생했습니다: {e}")
            st.stop()
            
    st.success(f"✅ 구글 시트 연결 성공! 총 **{len(data_df)}건**의 회의록 데이터를 불러왔습니다.")

    # HWPX 템플릿 구조 해석
    hwpx_zip = zipfile.ZipFile(io.BytesIO(hwpx_bytes), 'r')
    template_files = {name: hwpx_zip.read(name) for name in hwpx_zip.namelist()}
    
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

    # 메모리 압축 생성
    zip_buffer = io.BytesIO()
    log_records = []
    
    progress_bar = st.progress(0)
    progress_text = st.empty()

    with zipfile.ZipFile(zip_buffer, 'w', compression=zipfile.ZIP_DEFLATED) as main_zip:
        total_rows = len(sorted_df)
        
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
                
            # HWPX 메모리 저장
            doc_buffer = io.BytesIO()
            with zipfile.ZipFile(doc_buffer, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:
                for name, data in template_files.items():
                    if name == 'Contents/section0.xml':
                        z_out.writestr(name, xml_content.encode('utf-8'))
                    else:
                        z_out.writestr(name, data)
                        
            main_zip.writestr(f"{folder_name}/{doc_id}.hwpx", doc_buffer.getvalue())
            
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
        
        # 엑셀 로그 생성 (빨간 글씨 서식)
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
        
        # 로그 파일 2종 ZIP 포함
        excel_filename = f"심층면담_서류_Log_{version_str}.xlsx"
        csv_filename = f"심층면담_서류_Log_{version_str}.csv"
        
        main_zip.writestr(f"{folder_name}/{excel_filename}", excel_buffer.getvalue())
        main_zip.writestr(f"{folder_name}/{csv_filename}", log_df.to_csv(index=False, encoding='utf-8-sig'))

    st.success("🎉 회의록 58개 및 점검 로그 완성이 완료되었습니다!")
    
    st.download_button(
        label=f"📦 {folder_name}.zip 압축 파일 다운로드",
        data=zip_buffer.getvalue(),
        file_name=f"{folder_name}.zip",
        mime="application/zip",
        use_container_width=True
    )
    
    with st.expander("📋 생성 로그 데이터 미리보기"):
        st.dataframe(log_df)