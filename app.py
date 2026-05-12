import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
import io, json, re
from PIL import Image as PILImage
from supabase import create_client, Client
import re
import uuid
import os

# --- Supabase 설정 ---
@st.cache_resource
def get_supabase() -> Client:
    url = st.secrets["supabase"]["url"]
    key = st.secrets["supabase"]["key"]
    return create_client(url, key)

supabase = get_supabase()
BUCKET_NAME = "templates" # Supabase Storage 버킷 이름

# --- 2. Supabase 프리셋/양식 ---

def load_presets_from_db():
    try:
        # excel_presets 테이블의 모든 데이터 가져오기
        response = supabase.table("excel_presets").select("*").execute()
        
        # UI { '프리셋명': {데이터} } 형태로 변환
        presets = {}
        for row in response.data:
            presets[row['preset_name']] = {
                "filename": row['filename'],
                "cells": row['cell_positions'],
                "template_name": row['template_path']
            }
        return presets
    except Exception as e:
        st.error(f"DB 불러오기 실패: {e}")
        return {}

def save_preset_to_db(name, filename, cells, t_name):
    try:
        data = {
            "preset_name": name,
            "filename": filename,
            "cell_positions": cells,
            "template_path": t_name
        }
        supabase.table("excel_presets").upsert(data, on_conflict="preset_name").execute()
        return True
    except Exception as e:
        st.error(f"DB 저장 실패: {e}")
        return False

def upload_template_to_supabase(file_name, file_data):
    """파일명을 안전한 고유 ID로 바꿔서 업로드합니다."""
    try:
        extension = os.path.splitext(file_name)[1]
        # 한글/공백 에러를 방지하기 위해 파일명을 고유값으로 변경
        safe_db_name = f"template_{uuid.uuid4().hex[:8]}{extension}" 
        
        supabase.storage.from_(BUCKET_NAME).upload(
            path=safe_db_name, 
            file=file_data,
            file_options={"cache-control": "3600", "upsert": "true"}
        )
        return safe_db_name # 변경된 안전한 이름을 반환
    except Exception as e:
        st.error(f"업로드 실패: {e}")
        return None

def download_template_from_supabase(file_name):
    """Supabase Storage에서 양식 파일을 다운로드합니다"""
    try:
        # storage.from_("버킷이름").download("파일명")
        return supabase.storage.from_(BUCKET_NAME).download(file_name)
    except Exception as e:
        return None
    
    # 사이드바에 비밀번호 입력창을 만들어 관리자 인증 구현 (간단한 예시)
admin_mode = st.sidebar.checkbox("관리자 모드 접속")

if admin_mode:
    password = st.sidebar.text_input("Admin Password", type="password")
    if password == "@tlavmf123": # 실제 비밀번호로 변경
        st.subheader("관리자 대시보드")
        
        # 현재 DB 상태를 표로 보여줌
        res = supabase.table("excel_presets").select("*").execute()
        if res.data:
            st.table(res.data) # 데이터 시각화
            
            # 특정 데이터 삭제 기능
            target = st.selectbox("삭제할 프리셋 선택", [r['preset_name'] for r in res.data])
            if st.button("DB에서 영구 삭제"):
                supabase.table("excel_presets").delete().eq("preset_name", target).execute()
                st.success(f"{target} 삭제 완료!")
                st.rerun()
    else:
        st.sidebar.warning("비밀번호가 틀렸습니다.")

# --- 3. 이미지 처리 로직  ---
def fit_image_to_merged_cell(ws, img_data, cell_addr):
    try:
        try:
            from pillow_heif import register_heif_opener
            register_heif_opener()
        except ImportError:
            pass

        input_img = PILImage.open(io.BytesIO(img_data))
        
        if input_img.width > 1600:
            ratio = 1600 / float(input_img.width)
            hsize = int(float(input_img.height) * ratio)
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
           # st.write(f"✅ 병합셀 찾음: {target_range}") 
        else:
            img.anchor = cell_addr.upper()
            #st.write(f"⚠️ 병합셀 못찾음, 단순앵커: {cell_addr}") 
        
        ws.add_image(img)
       # st.write(f"✅ 이미지 추가완료: {cell_addr}")
        
    except Exception as e:
        st.error(f"❌ 이미지 삽입 실패 ({cell_addr}): {type(e).__name__}: {e}")

# --- 4. 메인 UI 및 앱 ---
st.set_page_config(page_title="EXEL UPLOAD (by.Simroot)", layout="wide")
st.title("EXEL UPLOAD")

if 'presets' not in st.session_state:
    st.session_state.presets = load_presets_from_supabase()
if 'temp_cells' not in st.session_state:
    st.session_state.temp_cells = ""

# 프리셋 관리 사이드바
st.sidebar.header("프리셋")
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
    st.subheader("설정")
    custom_filename = st.text_input("결과 파일명", value=d_name)
    
    st.write("**위치 설정**")
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

    cell_input = st.text_area("최종 위치", value=st.session_state.temp_cells, height=100)
    st.session_state.temp_cells = cell_input
    uploaded_excel = st.file_uploader("엑셀 양식 업로드", type=['xlsx'])

    new_p_name = st.sidebar.text_input("신규 프리셋 이름")
    if st.sidebar.button("저장", use_container_width=True):
        if new_p_name and custom_filename and cell_input:
            target_t_name = active_temp_name
            if uploaded_excel:
                excel_data = uploaded_excel.getvalue()
                safe_name = upload_template_to_supabase(uploaded_excel.name, excel_data)
                if safe_name:
                   target_t_name = safe_name # 정제된 파일명으로 프리셋 저장
                   st.sidebar.success(f"📦 양식 업로드 성공!")
                else:
                    st.stop()   
            
            st.session_state.presets[new_p_name] = {
                "filename": custom_filename,
                "cells": cell_input,
                "template_name": target_t_name
            }
            if save_presets_to_supabase(st.session_state.presets):
                st.sidebar.success(f"'{new_p_name}' 저장 완료!")
                st.rerun()

with main_col2:
    st.subheader("사진 업로드")
    uploaded_imgs = st.file_uploader("작업 사진 선택", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)
    
    if st.button("생성", use_container_width=True, type="primary"):
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
            with st.spinner("엑셀 보고서를 생성 중입니다..."):
                try:
                    wb = openpyxl.load_workbook(io.BytesIO(final_excel_data))

                    img_bytes_list = [img_file.read() for img_file in uploaded_imgs]

                    for i, img_bytes in enumerate(img_bytes_list):
                        if i >= len(final_cells):
                            break
    
                        #st.write(f"디버그 {i+1}번: {len(img_bytes)} bytes") 
    
                        if not img_bytes:
                            st.warning(f"{i+1}번째 이미지 비어있음")
                            continue
                            
                        p_idx_str, cell_addr = final_cells[i].split(":")
                        sheet_idx = int(p_idx_str) - 1
                        
                        if sheet_idx < 0 or sheet_idx >= len(wb.sheetnames):
                            st.error(f"❌{p_idx_str}번 시트를 찾을 수 없습니다.")
                            continue
                            
                        ws = wb[wb.sheetnames[sheet_idx]]
                
                        fit_image_to_merged_cell(ws, img_bytes, cell_addr)
                    
                    out = io.BytesIO()
                    wb.save(out)
                    out.seek(0)
                    
                    st.success(f"✅ 생성 완료!")
                    st.download_button(
                        label="결과 엑셀 다운로드",
                        data=out,
                        file_name=f"{custom_filename}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"오류 발생: {e}")
# --- 관리자 전용 ---
st.sidebar.markdown("---")
admin_key = st.sidebar.checkbox("관리자 모드 활성화")

if admin_key:
    st.header("관리자 데이터베이스 제어판")
    
    # 1. DB에서 현재 저장된 모든 프리셋 데이터 직접 불러오기
    try:
        # Supabase 'presets' 테이블의 모든 행 가져오기
        res = supabase.table("presets").select("*").execute()
        db_data = res.data
        
        if db_data:
            st.subheader("현재 저장된 프리셋 목록 (DB)")
            # 표 형태로 출력
            st.dataframe(db_data, use_container_width=True)
            
            # 2. 특정 데이터 삭제 기능
            st.subheader("데이터 삭제")
            target_id = st.selectbox("삭제할 데이터의 ID를 선택하세요", [r['id'] for r in db_data])
            
            if st.button("선택한 데이터 영구 삭제", type="secondary"):
                supabase.table("presets").delete().eq("id", target_id).execute()
                st.success(f"ID {target_id} 데이터가 삭제되었습니다.")
                st.rerun()
        else:
            st.info("DB에 저장된 데이터가 없습니다.")
            
    except Exception as e:
        st.error(f"관리자 모드 로드 실패: {e}")
        st.info("💡 먼저 Supabase SQL Editor에서 테이블을 생성해야 합니다.")

st.markdown("---")
st.info("""💡 **Supabase 연동으로 서버가 꺼져도 내 프리셋과 양식 파일이 안전하게 보존됩니다."
- **자동증가:** `1:B2+16`으로 입력하면 B2, B18, B34 순으로 사진이 들어갑니다.
- **프리셋:** 한 번 저장해두면 다음부터는 사진만 올리면 끝납니다.
""")
