import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
import io, json, os, re
from PIL import Image as PILImage
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# --- 1. 구글 드라이브 서비스 연결 함수 ---
def get_drive_service():
    # Streamlit Cloud의 Secrets에 저장된 정보를 사용
    creds_info = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_info, scopes=["https://www.googleapis.com/auth/drive"])
    return build('drive', 'v3', credentials=creds)

# Secrets에서 폴더 ID 가져오기
FOLDER_ID = st.secrets["google_drive"]["folder_id"]
st.write(f"디버깅용 - 현재 폴더 ID: [{FOLDER_ID}]") # 이 줄을 추가해서 따옴표나 공백이 있는지 확인

# --- 2. 구글 드라이브 프리셋/양식 처리 로직 ---
def load_presets_from_drive():
    try:
        service = get_drive_service()
        query = f"name = 'presets.json' and '{FOLDER_ID}' in parents and trashed = false"
        results = service.files().list(q=query).execute().get('files', [])
        
        if not results: return {}
        
        file_id = results[0]['id']
        request = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        return json.loads(fh.getvalue().decode('utf-8'))
    except Exception as e:
        st.error(f"프리셋 로드 실패: {e}")
        return {}

def save_presets_to_drive(presets):
    try:
        service = get_drive_service()
        content = json.dumps(presets, indent=4, ensure_ascii=False)
        fh = io.BytesIO(content.encode('utf-8'))
        media = MediaIoBaseUpload(fh, mimetype='application/json')
        
        # save_presets_to_drive 함수 내 query 부분 수정 제안
        query = f"name = 'presets.json' and '{FOLDER_ID}' in parents and trashed = false"
        results = service.files().list(
            q=query,
            fields="files(id, name)" # 불필요한 옵션을 제거하고 핵심 필드만 요청
        ).execute().get('files', [])
        
        if results:
            # 업데이트 시
            service.files().update(
                fileId=results[0]['id'], 
                media_body=media,
                supportsAllDrives=True
            ).execute()
        else:
            # 생성 시 (body 구조 확인!)
            file_metadata = {
                'name': 'presets.json', 
                'parents': [FOLDER_ID] 
            }
            service.files().create(
                body=file_metadata, 
                media_body=media,
                supportsAllDrives=True,
                fields='id'
            ).execute()
        return True
    except Exception as e:
        st.error(f"프리셋 저장 실패: {e}")
        return False

def upload_template_to_drive(file_name, file_data):
    try:
        service = get_drive_service()
        fh = io.BytesIO(file_data)
        media = MediaIoBaseUpload(fh, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        # 1. 일단 파일이 있는지 찾습니다.
        query = f"name = '{file_name}' and '{FOLDER_ID}' in parents and trashed = false"
        results = service.files().list(q=query).execute().get('files', [])
        
        if results:
            # 파일이 있으면 수정(Update) -> 이건 무조건 성공!
            service.files().update(fileId=results[0]['id'], media_body=media).execute()
            return True
        else:
            # 파일이 없으면 생성(Create) 시도
            # 여기서 에러가 난다면, 구글 드라이브 폴더 설정에서 
            # '편집자가 권한을 변경하고 항목을 공유할 수 있음'이 켜져 있는지 확인해야 합니다.
            file_metadata = {'name': file_name, 'parents': [FOLDER_ID]}
            service.files().create(body=file_metadata, media_body=media).execute()
            return True
    except Exception as e:
        # 만약 여기서 storageQuotaExceeded(용량부족) 에러가 뜨면 
        # 서비스 계정의 근본적 한계라 처음 한 번은 파일을 던져줘야 합니다 ㅠ
        st.error(f"양식 저장 실패: {e}")
        return False

def download_template_from_drive(file_name):
    try:
        service = get_drive_service()
        query = f"name = '{file_name}' and '{FOLDER_ID}' in parents and trashed = false"
        results = service.files().list(q=query).execute().get('files', [])
        
        if not results: return None
        
        request = service.files().get_media(fileId=results[0]['id'])
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        return fh.getvalue()
    except:
        return None

# --- 3. 이미지 처리 로직 (고화질 유지) ---
def fit_image_to_merged_cell(ws, img_data, cell_addr):
    try:
        input_img = PILImage.open(io.BytesIO(img_data))
        # 해상도 최적화 (가로 1600 기준)
        if input_img.width > 1600:
            ratio = 1600 / float(input_img.width)
            hsize = int((float(input_img.height) * float(ratio)))
            input_img = input_img.resize((1600, hsize), PILImage.Resampling.LANCZOS)
        
        img_byte_arr = io.BytesIO()
        input_img.convert("RGB").save(img_byte_arr, format='JPEG', quality=85)
        img_byte_arr.seek(0)
        
        img = Image(img_byte_arr)
        target_range = None
        for merged_range in ws.merged_cells.ranges:
            if cell_addr.upper() in merged_range:
                target_range = merged_range
                break
        
        if target_range:
            sc, sr = target_range.min_col - 1, target_range.min_row - 1
            ec, er = target_range.max_col, target_range.max_row
            img.anchor = TwoCellAnchor('twoCell', AnchorMarker(sc, 0, sr, 0), AnchorMarker(ec, 0, er, 0))
        else:
            img.anchor = cell_addr.upper()
        ws.add_image(img)
    except Exception as e:
        st.error(f"이미지 삽입 실패 ({cell_addr}): {e}")

# --- 4. 메인 UI 및 앱 로직 ---
st.set_page_config(page_title="이미지 업로드 보고서 시스템", layout="wide")
st.title("이미지 업로드 보고서 시스템")

# 초기 프리셋 로드
if 'presets' not in st.session_state:
    st.session_state.presets = load_presets_from_drive()
if 'temp_cells' not in st.session_state:
    st.session_state.temp_cells = ""

# 프리셋 관리
st.sidebar.header("💾 구글 프리셋 매니저")
preset_list = list(st.session_state.presets.keys())
selected_preset = st.sidebar.selectbox("설정 불러오기", ["직접 입력"] + preset_list)

# --- 삭제 기능 ---
if selected_preset != "직접 입력":
    if st.sidebar.button("현재 프리셋 삭제", use_container_width=True):
        if selected_preset in st.session_state.presets:
            # 세션에서 삭제
            del st.session_state.presets[selected_preset]
            # 구글 드라이브에 업데이트 (전체 목록 다시 저장)
            if save_presets_to_drive(st.session_state.presets):
                st.sidebar.success(f"'{selected_preset}' 삭제 완료!")
                st.rerun() # 화면 새로고침하여 목록 갱신
# ----------------------

d_name, d_cells, active_temp_data, active_temp_name = "", "", None, None

if selected_preset != "직접 입력" and selected_preset in st.session_state.presets:
    p = st.session_state.presets[selected_preset]
    d_name = p.get("filename", "")
    d_cells = p.get("cells", "")
    st.session_state.temp_cells = d_cells
    t_name = p.get("template_name")
    if t_name:
        active_temp_data = download_template_from_drive(t_name)
        if active_temp_data:
            active_temp_name = t_name
            st.sidebar.success(f"✅ 양식 연결됨: {t_name}")
else:
    
    d_name, d_cells, active_temp_data, active_temp_name = "", "", None, None
# 메인 화면
main_col1, main_col2 = st.columns([1, 1])

with main_col1:
    st.subheader("1. 위치 및 설정")
    custom_filename = st.text_input("결과 파일명", value=d_name)
    
    st.write("**위치 생성**")
    w_col1, w_col2, w_col3 = st.columns([1, 1, 1.5])
    with w_col1:
        s_num = st.number_input("시트", min_value=1, step=1)
    with w_col2:
        all_cols = [chr(i) for i in range(ord('A'), ord('Z') + 1)]
        c_let = st.selectbox("열", all_cols, index=all_cols.index("B"))
    with w_col3:
        r_num = st.number_input("행 번호", min_value=1, step=1)
    
    cur_addr = f"{s_num}:{c_let}{r_num}"
    
    b_col1, b_col2, b_col3 = st.columns(3)
    with b_col1:
        if st.button("위치 추가", use_container_width=True):
            if st.session_state.temp_cells: st.session_state.temp_cells += f", {cur_addr}"
            else: st.session_state.temp_cells = cur_addr
    with b_col2:
        if st.button("+16행 추가"):
            st.session_state.temp_cells += "+16"
    with b_col3:
        if st.button("전체 초기화", use_container_width=True):
            st.session_state.temp_cells = ""
            st.rerun()

    cell_input = st.text_area("최종 위치 리스트", value=st.session_state.temp_cells, height=100)
    st.session_state.temp_cells = cell_input

    uploaded_excel = st.file_uploader("새 엑셀 양식 업로드", type=['xlsx'])

    new_p_name = st.sidebar.text_input("신규 프리셋 이름")
    if st.sidebar.button("구글 드라이브에 저장", use_container_width=True):
        if new_p_name and custom_filename and cell_input:
            target_t_name = active_temp_name  # 기본적으로 기존 연결된 양식 이름 사용
            
            # 1. 새로운 엑셀 파일이 업로드된 경우 
            if uploaded_excel:
                target_t_name = uploaded_excel.name
                # 파일을 바이트 데이터로 변환
                excel_data = uploaded_excel.getvalue()
                
                # 드라이브 업로드 함수 호출
                if upload_template_to_drive(target_t_name, excel_data):
                    st.sidebar.success(f"📦 양식 파일({target_t_name}) 업로드 성공!")
                else:
                    # 여기서 실패하면 아래 프리셋 저장으로 넘어가지 않도록 방어
                    st.stop() 
            
            # 2. 프리셋(설정값) 정보 업데이트
            st.session_state.presets[new_p_name] = {
                "filename": custom_filename,
                "cells": cell_input,
                "template_name": target_t_name
            }
            
            # 3. 프리셋 파일(json)을 드라이브에 최종 저장
            if save_presets_to_drive(st.session_state.presets):
                st.sidebar.success(f"💾 '{new_p_name}' 설정 저장 완료!")
                st.rerun()

with main_col2:
    st.subheader("2. 사진 업로드 및 생성")
    uploaded_imgs = st.file_uploader("작업 사진 선택", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)
    
    if st.button("엑셀 보고서 즉시 생성", use_container_width=True, type="primary"):
        final_excel_data = uploaded_excel.getvalue() if uploaded_excel else active_temp_data
        
        # 위치 해석
        final_cells = []
        parts = [p.strip() for p in cell_input.split(",") if p.strip()]
        for p in parts:
            if "+" in p:
                try:
                    base, inc = p.split("+")
                    s_idx, s_cell = base.split(":")
                    row_inc = int(inc)
                    match = re.match(r"([a-zA-Z]+)([0-9]+)", s_cell)
                    c_l, r_s = match.group(1), int(match.group(2))
                    for k in range(50):
                        final_cells.append(f"{s_idx}:{c_l}{r_s + (k * row_inc)}")
                except: continue
            else:
                if ":" in p: final_cells.append(p)

        if final_excel_data and final_cells and uploaded_imgs:
            with st.spinner("구글에서 양식을 불러와 이미지를 배치 중..."):
                try:
                    wb = openpyxl.load_workbook(io.BytesIO(final_excel_data))
                    for i, img_file in enumerate(uploaded_imgs):
                        if i >= len(final_cells): break
                        p_idx_str, cell_addr = final_cells[i].split(":")
                        ws = wb[wb.sheetnames[int(p_idx_str)-1]]
                        fit_image_to_merged_cell(ws, img_file.read(), cell_addr)
                    
                    out = io.BytesIO()
                    wb.save(out)
                    out.seek(0)
                    st.success(f"✅ 배치 완료! (총 {min(len(uploaded_imgs), len(final_cells))}장)")
                    st.download_button("결과 엑셀 다운로드", out, f"{custom_filename}.xlsx", use_container_width=True)
                except Exception as e:
                    st.error(f"오류 발생: {e}")
        else:
            st.warning("양식, 위치 리스트, 사진이 모두 준비되어야 합니다.")

st.markdown("---")
st.info(""""💡 **팁:** 구글 드라이브 연동으로 서버가 꺼져도 내 프리셋과 양식 파일이 안전하게 보존됩니다."
- **자동증가:** `1:B2+16`으로 입력하면 B2, B18, B34 순으로 사진이 들어갑니다.
- **프리셋:** 한 번 저장해두면 다음부터는 사진만 올리면 끝납니다.
""")
