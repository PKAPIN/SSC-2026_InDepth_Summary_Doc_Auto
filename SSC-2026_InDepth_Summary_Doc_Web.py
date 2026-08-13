# =========================================================================
# [웹 호스팅용] 심층면담 회의록 및 점검 로그 자동 생성 Streamlit 웹 앱
# - 회의내용/회의결과 상호 공간 재배분(유동 높이 흡수) 최적화판
# - 1페이지 표 틀 및 하단 로고 위치 완전 고정
# - 백그라운드 비동기 멀티스레딩 적용 (15% 대기 멈춤 및 Read timed out 완벽 해결)
# - timeout=600(10분) 설정으로 대량 문서 업데이트 타임아웃 방지
# - 구글 앱스크립트(GS) 실행 과정 실시간 4단계 동기화 메시지 롤링 디스플레이
# =========================================================================
import io
import os
import glob
import zipfile
import datetime
import re
import ssl
import time
import threading
import requests
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import streamlit as st

ssl._create_default_https_context = ssl._create_unverified_context

st.set_page_config(page_title="심층면담 회의록 자동 생성 시스템", page_icon="📄", layout="wide")

# -------------------------------------------------------------------------
# ★ [지정해주신 GS 웹 앱 URL 적용 완료]
# -------------------------------------------------------------------------
APPS_SCRIPT_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbzQmhwyMmSAw-zreGgkVXTbqnHuaxPkdwIdcP5E_iEOTb3m3VVATB4w3hh55a1k85Qc/exec"

# 초록색 버튼 지정 커스텀 CSS
st.markdown("""
    <style>
    div.stButton > button[key="btn_update_data"] {
        background-color: #28a745 !important;
        color: white !important;
        border-color: #28a745 !important;
        font-weight: bold !important;
    }
    div.stButton > button[key="btn_update_data"]:hover {
        background-color: #218838 !important;
        border-color: #1e7e34 !important;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------------------
# [1] 상단 레이아웃 (좌측: 타이틀 및 연동 시트 / 우측 상단: 초록색 버튼 및 안내 박스)
# -------------------------------------------------------------------------
col_title, col_top_btn = st.columns([3, 1.3])

with col_title:
    st.title("📄 심층면담 회의록 문서 및 로그 자동 생성기")
    st.markdown("🔗 **연동 시트**: [구글 스프레드시트 (심층면담_회의록)](https://docs.google.com/spreadsheets/d/1ws9JTAdRXwbp--NhrjWwelNorSTv1_LIJW7DijUtJLU/edit?gid=770556375#gid=770556375)")
    st.caption("구글 시트의 최신 데이터를 실시간으로 읽어와 지정된 HWPX 회의록 서식으로 개별 문서 및 점검 로그(Excel/CSV)를 일괄 생성합니다.")

with col_top_btn:
    st.write(" ")
    st.info("💡 **안내**: 심층면담 기록지가 추가로 들어왔을 경우 아래 버튼을 통해 업데이트를 할 수 있습니다. 다운로드 전 업데이트 부탁드립니다.")
    update_clicked = st.button("🔄 자료 업데이트", key="btn_update_data", use_container_width=True)
    st.caption("구글 드라이브에 업로드된 문서내용을 스프레드 시트로 불러옵니다.")

# 자료 업데이트 버튼 클릭 시 비동기 스레드(timeout 600s) + 실시간 4단계 동기화 실행 로그 디스플레이
if update_clicked:
    st.write("🚀 **구글 드라이브 심층면담 기록지를 읽어 구글 시트(DB)에 AI 요약을 기입 중입니다...**")
    st.caption("💡 업데이트 분량이 많으면 1~2분 정도의 시간이 소요될 수 있습니다.")
    
    ai_progress_bar = st.progress(0)
    ai_status_text = st.empty()
    
    api_result = {"response": None, "error": None, "done": False}
    
    # 백그라운드 앱스크립트 비동기 호출 함수 (timeout=600으로 타임아웃 방지)
    def fetch_apps_script():
        try:
            res = requests.get(APPS_SCRIPT_WEBAPP_URL, timeout=600)
            api_result["response"] = res
        except Exception as e:
            api_result["error"] = e
        finally:
            api_result["done"] = True

    thread = threading.Thread(target=fetch_apps_script)
    thread.start()
    
    start_time = time.time()
    current_pct = 0
    
    # 백엔드 완료 시점까지 UI 메인 스레드에서 실시간 진행 상태와 4단계 로그 메시지 롤링
    while not api_result["done"]:
        time.sleep(0.8)
        elapsed = int(time.time() - start_time)
        
        if current_pct < 95:
            current_pct += 1
            
        ai_progress_bar.progress(current_pct)
        
        # 앱스크립트 진행 단계별 메시지 정의
        if current_pct < 20:
            stage_msg = "🔍 [1/4단계] 구글 드라이브 내 심층면담 기록지 파일 탐색 및 HWP/HWPX 구조 진단 중..."
        elif current_pct < 45:
            stage_msg = "📄 [2/4단계] HWPX 문서 내 XML 텍스트 파싱 및 연동 데이터 추출 중..."
        elif current_pct < 80:
            stage_msg = "🤖 [3/4단계] Gemini AI 모델 호출 및 심층면담 안건/내용/결과 요약 분석 중..."
        else:
            stage_msg = "📊 [4/4단계] 구글 스프레드시트(SSC-2026_AUTO_DB) 기입 및 데이터 동기화 중..."
            
        ai_status_text.text(f"⏳ {stage_msg} [{current_pct}% / 100%] (경과 시간: {elapsed}초)")

    response = api_result["response"]
    error = api_result["error"]
    
    if response and response.status_code == 200:
        ai_progress_bar.progress(100)
        ai_status_text.success("✅ 구글 드라이브의 문서 내용이 구글 시트(DB)로 성공적으로 자동 요약되어 동기화되었습니다! (100% 완료)")
        st.session_state.clear()
        
        try:
            res_json = response.json()
            gs_result_msg = res_json.get("result", "상세 로그를 불러올 수 없습니다.")
            st.session_state['gs_update_log'] = gs_result_msg
        except Exception:
            st.session_state['gs_update_log'] = response.text
            
    else:
        status_err = response.status_code if response else f"오류 발생: {error}"
        st.error(f"❌ 앱스크립트 동기화 호출 실패 ({status_err})")

# 구글 앱스크립트 업데이트 상세 결과 로그 박스 (로그 내 우측 하단 버전 표기 포함)
if 'gs_update_log' in st.session_state:
    with st.expander("📋 업데이트 상세 로그 보기", expanded=True):
        st.code(st.session_state['gs_update_log'], language="text")

st.divider()

# -------------------------------------------------------------------------
# [2] 회의록 문서 및 생성로그 일괄 생성 영역
# -------------------------------------------------------------------------
def format_section_headers(text: str) -> str:
    if not text or not isinstance(text, str):
        return ""
    cleaned = text.strip()
    def replace_header(match):
        return f"\n\n**{match.group(0)}**"
    return re.sub(r'<[^>]+>', replace_header, cleaned).strip()

SHEET_ID = "1ws9JTAdRXwbp--NhrjWwelNorSTv1_LIJW7DijUtJLU"
GID = "770556375"

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
            nocache_url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}&_nocache={int(time.time())}"
            df_raw = pd.read_csv(nocache_url)
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

        content_text = str(row.get('회의내용', '')).strip() if pd.notna(row.get('회의내용')) else ""
        result_text = str(row.get('회의결과', '')).strip() if pd.notna(row.get('회의결과')) else ""

        c_lines = max(1, len(content_text.splitlines()))
        r_lines = max(1, len(result_text.splitlines()))

        TOTAL_ALLOWANCE = 25000
        LINE_HEIGHT_C = 1100
        LINE_HEIGHT_R = 1000

        needed_c_height = max(3500, c_lines * LINE_HEIGHT_C + 1500)
        assigned_r_height = max(3500, TOTAL_ALLOWANCE - needed_c_height)
        assigned_c_height = TOTAL_ALLOWANCE - assigned_r_height

        for col in data_df.columns:
            if col == '일시_dt': continue
            val = row[col]
            if pd.isna(val) or str(val).strip() == "" or str(str(val)).strip().lower() == "nan":
                val_str = ""
                if col in merge_fields: missing_fields.append(col)
            elif isinstance(val, (pd.Timestamp, datetime.datetime)): val_str = val.strftime('%Y-%m-%d')
            elif isinstance(val, datetime.time): val_str = val.strftime('%H:%M')
            else: val_str = str(val).strip()
            
            if col in ['회의내용', '회의결과']:
                val_str = format_section_headers(val_str)
                val_str = val_str.replace("**", "")
                
            val_str = val_str.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            val_str = val_str.replace("\r\n", "\n").replace("\r", "\n")
            val_str = val_str.replace("\n\n", "\n")
            
            target_field = f"{{{{{col}}}}}"
            if target_field in xml_content:
                field_pos = xml_content.find(target_field)
                p_matches = list(re.finditer(r'<hp:p\b[^>]*>', xml_content[:field_pos]))
                open_p_tag = p_matches[-1].group(0) if p_matches else '<hp:p>'
                
                run_matches = list(re.finditer(r'<hp:run\b[^>]*>', xml_content[:field_pos]))
                open_run_tag = run_matches[-1].group(0) if run_matches else '<hp:run>'
                
                paragraph_replace = f'</hp:t></hp:run></hp:p>{open_p_tag}{open_run_tag}<hp:t>'
                val_str = val_str.replace("\n", paragraph_replace)
                xml_content = xml_content.replace(target_field, val_str)

        xml_content = re.sub(
            r'(<hp:tc\b[^>]*?)(height="\d+")([^>]*?>[\s\S]*?\{\{회의내용\}\}|<hp:tc\b[^>]*?)(height="\d+")([^>]*?>[\s\S]*?회의내용)',
            rf'\1height="{assigned_c_height}"\3',
            xml_content
        )
        xml_content = re.sub(
            r'(<hp:tc\b[^>]*?)(height="\d+")([^>]*?>[\s\S]*?\{\{회의결과\}\}|<hp:tc\b[^>]*?)(height="\d+")([^>]*?>[\s\S]*?회의결과)',
            rf'\1height="{assigned_r_height}"\3',
            xml_content
        )

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
