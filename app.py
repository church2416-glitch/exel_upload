import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
import io, json, re
from PIL import Image as PILImage
from supabase import create_client, Client

# --- 1. Supabase 연결 설정 ---
@st.cache_resource
def get_supabase() -> Client:
    url = st.secrets["supabase"]["url"]
    key = st.secrets["supabase"]["key"]
    return create_client(url, key)

supabase = get_supabase()
BUCKET_NAME = "templates" # Supabase Storage에 만든 버킷 이름

# --- 2. Supabase 프리셋/양식 처리 로직 ---

def load_presets_from_supabase():
    """DB에서 프리셋 정보를 가져옵니다 (JSON 파일 대신 DB 테이블 권장하나 기존 로직 유지 위해 파일로 처리)"""
    try:
        # Storage에서 presets.json 다운로드
        res = supabase.storage.from_(BUCKET_NAME).download("presets.json")
        return json.loads(res.decode('utf-8'))
    except Exception:
        return {}

def save_presets_to_supabase(presets):
    """DB 또는 Storage에 프리셋 정보를 저장합니다"""
    try:
        content = json.dumps(presets, indent=4, ensure_ascii=False)
        content_bytes = content.encode('utf-8')
        
        # Supabase Storage에 덮어쓰기 (upsert=True)
        supabase.storage.from_(BUCKET_NAME).upload(
            path="presets.json",
            file=content_bytes,
            file_options={"cache-control": "3600", "upsert": "true"}
        )
        return True
    except Exception as e:
        st.error(f"프리셋 저장 실패: {e}")
        return False

def upload_template_to_supabase(file_name, file_data):
    """엑셀 양식 파일을 Supabase Storage에 업로드합니다"""
    try:
        supabase.storage.from_(BUCKET_NAME).upload(
            path=file_name,
            file=file_data,
            file_options={"cache-control": "3600", "upsert": "true"}
        )
        return True
    except Exception as e:
        st.error(f"양식 업로드 실패: {e}")
        return False

def download_template_from_supabase(file_name):
    """Supabase Storage에서 양식 파일을 다운로드합니다"""
    try:
        return supabase.storage.from_(BUCKET_NAME).download(file_name)
    except Exception:
        return None

# --- 3. 이미지 처리 로직 (기존과 동일) ---
def fit_image_to_merged_cell(ws, img_data, cell_addr):
    try:
        input_img = PILImage.open(io.BytesIO(img_data))
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
st.set_page_config(page_title="이미지 업로드 보고서 시스템 (Supabase)", layout="wide")
st.title("이미지 업로드 보고서 시스템")

if 'presets' not in st.session_state:
    st.session_state.presets = load_presets_from_supabase()
if 'temp_cells' not in st.session_state:
    st.session_state.temp_cells = ""

# 프리셋 관리 사이드바
st.sidebar.header("💾 Supabase 프리셋 매니저")
preset_list = list(st.session_state.presets.keys())
selected_preset = st.sidebar.selectbox("설정 불러오기", ["직접 입력"] + preset_list)

if selected_preset != "직접 입력":
    if st.sidebar.button("현재 프리셋 삭제", use_container_width=True):
        del st.session_state.presets[selected_preset]
        if save_presets_to_supabase(st.session_state.presets):
            st.sidebar.success(f"'{selected_preset}' 삭제 완료!")
            st.rerun()

# 프리셋 데이터 매핑
d_name, d_cells, active_temp_data, active_temp_name = "", "", None, None

if selected_preset != "직접 입력" and selected_preset in st.session_state.presets:
    p = st.session_state.presets[selected_preset]
    d_name = p.get("filename", "")
    d_cells = p.get("cells", "")
    st.session_state.temp_cells = d_cells
    t_name = p.get("template_name")
    if t_name:
        active_temp_data = download_template_from_supabase(t_name)
        if active_temp_data:
            active_temp_name = t_name
            st.sidebar.success(f"✅ 양식 연결됨: {t_name}")

# 메인 UI 구성
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
    if st.sidebar.button("Supabase에 저장", use_container_width=True):
        if new_p_name and custom_filename and cell_input:
            target_t_name = active_temp_name
            if uploaded_excel:
                target_t_name = uploaded_excel.name
                if upload_template_to_supabase(target_t_name, uploaded_excel.getvalue()):
                    st.sidebar.success(f"📦 양식 업로드 성공!")
                else: st.stop()
            
            st.session_state.presets[new_p_name] = {
                "filename": custom_filename,
                "cells": cell_input,
                "template_name": target_t_name
            }
            if save_presets_to_supabase(st.session_state.presets):
                st.sidebar.success(f"💾 '{new_p_name}' 저장 완료!")
                st.rerun()

with main_col2:
    st.subheader("2. 사진 업로드 및 생성")
    uploaded_imgs = st.file_uploader("작업 사진 선택", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)
    
    if st.button("엑셀 보고서 즉시 생성", use_container_width=True, type="primary"):
        final_excel_data = uploaded_excel.getvalue() if uploaded_excel else active_temp_data
        
        # 위치 해석 로직
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
            with st.spinner("Supabase에서 데이터를 가져와 생성 중..."):
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
                    st.success(f"✅ 생성 완료!")
                    st.download_button("결과 엑셀 다운로드", out, f"{custom_filename}.xlsx", use_container_width=True)
                except Exception as e:
                    st.error(f"오류 발생: {e}")

st.markdown("---")
st.info("""💡 **Supabase 연동으로 서버가 꺼져도 내 프리셋과 양식 파일이 안전하게 보존됩니다."
- **자동증가:** `1:B2+16`으로 입력하면 B2, B18, B34 순으로 사진이 들어갑니다.
- **프리셋:** 한 번 저장해두면 다음부터는 사진만 올리면 끝납니다.
""")
