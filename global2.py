import streamlit as st
import pandas as pd
import numpy as np
#import plotly.express as px
from datetime import datetime, timedelta
import io
from PIL import Image
import json
import os

# 1. 페이지 설정
st.set_page_config(page_title="LNG선공사팀 글로벌인력관리", layout="wide")

# 엑셀 다운로드 함수
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    return output.getvalue()

# 데이터 저장 함수
def save_data():
    """모든 데이터를 JSON 파일로 저장"""
    data_dir = "data"
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)
    
    with open(f"{data_dir}/workers.json", "w", encoding="utf-8") as f:
        json.dump(st.session_state.workers, f, ensure_ascii=False, indent=2)
    
    with open(f"{data_dir}/history.json", "w", encoding="utf-8") as f:
        json.dump(st.session_state.history, f, ensure_ascii=False, indent=2)
    
    with open(f"{data_dir}/korean_certs.json", "w", encoding="utf-8") as f:
        json.dump(st.session_state.korean_certs, f, ensure_ascii=False, indent=2)
    
    with open(f"{data_dir}/korean_classes.json", "w", encoding="utf-8") as f:
        json.dump(st.session_state.korean_classes, f, ensure_ascii=False, indent=2)

# 데이터 로드 함수
def load_data():
    """JSON 파일에서 데이터 로드"""
    data_dir = "data"
    
    try:
        if os.path.exists(f"{data_dir}/workers.json"):
            with open(f"{data_dir}/workers.json", "r", encoding="utf-8") as f:
                st.session_state.workers = json.load(f)
    except (json.JSONDecodeError, FileNotFoundError):
        st.session_state.workers = []
    
    try:
        if os.path.exists(f"{data_dir}/history.json"):
            with open(f"{data_dir}/history.json", "r", encoding="utf-8") as f:
                st.session_state.history = json.load(f)
    except (json.JSONDecodeError, FileNotFoundError):
        st.session_state.history = []
    
    try:
        if os.path.exists(f"{data_dir}/korean_certs.json"):
            with open(f"{data_dir}/korean_certs.json", "r", encoding="utf-8") as f:
                st.session_state.korean_certs = json.load(f)
    except (json.JSONDecodeError, FileNotFoundError):
        st.session_state.korean_certs = []
    
    try:
        if os.path.exists(f"{data_dir}/korean_classes.json"):
            with open(f"{data_dir}/korean_classes.json", "r", encoding="utf-8") as f:
                st.session_state.korean_classes = json.load(f)
    except (json.JSONDecodeError, FileNotFoundError):
        st.session_state.korean_classes = []

# 엑셀 데이터 정제 함수
def clean_excel_data(df):
    """엑셀 데이터에서 시간정보 제거 및 날짜 포맷팅, 사번을 문자열로 변환"""
    df_clean = df.copy()
    
    # 사번을 문자열로 변환 (형식 통일)
    if '사번' in df_clean.columns:
        df_clean['사번'] = df_clean['사번'].astype(str).str.strip()
    
    # 날짜 관련 컬럼 처리
    date_columns = ["입국일", "입국일자", "취득일"]
    for col in date_columns:
        if col in df_clean.columns:
            # Timestamp 또는 datetime 타입을 문자열로 변환 (시간 제외)
            df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce').dt.strftime('%Y-%m-%d')
    
    return df_clean

# ---------------------------------------------------------
# [데이터 초기화] Session State 설정
# ---------------------------------------------------------
if 'workers' not in st.session_state:
    st.session_state.workers = []
if 'history' not in st.session_state:
    st.session_state.history = []
if 'korean_certs' not in st.session_state:
    st.session_state.korean_certs = []
if 'korean_classes' not in st.session_state:
    st.session_state.korean_classes = []
if 'show_update_confirm' not in st.session_state:
    st.session_state.show_update_confirm = False
if 'pending_worker_data' not in st.session_state:
    st.session_state.pending_worker_data = None
if 'worker_photos' not in st.session_state:
    st.session_state.worker_photos = {}
if 'confirm_reset_mode' not in st.session_state:
    st.session_state.confirm_reset_mode = False
if 'selected_employee_data' not in st.session_state:
    st.session_state.selected_employee_data = None
if 'confirm_delete_mode' not in st.session_state:
    st.session_state.confirm_delete_mode = False
if 'pending_delete_emp_id' not in st.session_state:
    st.session_state.pending_delete_emp_id = None
if 'pending_delete_name' not in st.session_state:
    st.session_state.pending_delete_name = None
if 'edit_selected_emp_id' not in st.session_state:
    st.session_state.edit_selected_emp_id = None

# 저장된 데이터 로드
load_data()

# 조직 구조 정의
dept_structure = {
    "공사1부5과": {
        "반": ["1직1반", "1직2반", "1직3반", "2직1반", "2직2반", "2직3반"],
        "직종": ["수동본딩", "ABM", "TBP"]
    },
    "공사2부3과": {
        "반": ["설치직1반", "설치직2반", "설치직3반", "설치직4반", "용접직1반", "용접직2반", "용접직3반"],
        "직종": ["MB설치", "MB수동용접", "MB자동용접", "MB리웰딩"]
    },
    "3부의장과": {
        "반": ["2직1반", "2직2반", "2직3반"],
        "직종": ["의장", "LNGTIG"]
    }
}

# ---------------------------------------------------------
# [사이드바] 메뉴 구성
# ---------------------------------------------------------
st.sidebar.title("🌍 LNG선공사팀 글로벌인력관리")
menu = st.sidebar.radio("메뉴 선택", ["📊 통합 대시보드", "👤 인력 등록/업데이트", "📚 한국어 교육/자격 관리"])

# 사이드바 - 데이터 관리
st.sidebar.write("---")
st.sidebar.subheader("⚙️ 데이터 관리")

if st.sidebar.button("💾 데이터 저장", use_container_width=True):
    save_data()
    st.sidebar.success("✅ 데이터가 저장되었습니다!")

if st.sidebar.button("🔄 데이터 초기화", use_container_width=True):
    st.session_state.confirm_reset_mode = True

# 초기화 확인 UI
if st.session_state.confirm_reset_mode:
    st.sidebar.warning("⚠️ **정말로 모든 데이터를 삭제하시겠습니까?**")
    col_yn1, col_yn2 = st.sidebar.columns(2)
    
    with col_yn1:
        if st.button("✅ 예, 삭제합니다", use_container_width=True):
            # 데이터 초기화
            st.session_state.workers = []
            st.session_state.history = []
            st.session_state.korean_certs = []
            st.session_state.korean_classes = []
            st.session_state.worker_photos = {}
            
            # 저장된 파일도 삭제
            data_dir = "data"
            if os.path.exists(f"{data_dir}/workers.json"):
                os.remove(f"{data_dir}/workers.json")
            if os.path.exists(f"{data_dir}/history.json"):
                os.remove(f"{data_dir}/history.json")
            if os.path.exists(f"{data_dir}/korean_certs.json"):
                os.remove(f"{data_dir}/korean_certs.json")
            if os.path.exists(f"{data_dir}/korean_classes.json"):
                os.remove(f"{data_dir}/korean_classes.json")
            
            st.session_state.confirm_reset_mode = False
            st.sidebar.success("✅ 모든 데이터가 초기화되었습니다!")
            st.rerun()
    
    with col_yn2:
        if st.button("❌ 아니오, 취소", use_container_width=True):
            st.session_state.confirm_reset_mode = False
            st.rerun()

# ---------------------------------------------------------
# 메뉴 1: 📊 통합 대시보드 (분석 및 시각화)
# ---------------------------------------------------------
if menu == "📊 통합 대시보드":
    st.header("📊 글로벌 인력 현황 및 데이터 업로드")
    
    # 데이터 병합 (인력 + 자격증)
    with st.expander("📥 기존 인력 엑셀 파일 일괄 등록", expanded=False):
        uploaded_file = st.file_uploader("엑셀 파일(.xlsx)을 업로드하세요", type=["xlsx"])
        if uploaded_file:
            try:
                df_upload = pd.read_excel(uploaded_file)
                df_upload = clean_excel_data(df_upload)  # 날짜 데이터 정제
                if st.button("시스템에 데이터 반영하기"):
                    uploaded_data = df_upload.to_dict('records')
                    
                    # 중복 체크 및 추가
                    added_count = 0
                    duplicate_count = 0
                    
                    for new_worker in uploaded_data:
                        # 사번이 이미 존재하는지 확인
                        existing = next((w for w in st.session_state.workers if w.get('사번') == new_worker.get('사번')), None)
                        if existing:
                            duplicate_count += 1
                        else:
                            st.session_state.workers.append(new_worker)
                            added_count += 1
                    
                    save_data()  # 데이터 자동 저장
                    
                    if duplicate_count > 0:
                        st.warning(f"⚠️ {duplicate_count}명은 이미 존재합니다. {added_count}명의 새로운 데이터가 추가되었습니다.")
                    else:
                        st.success(f"✅ {added_count}명의 데이터가 추가되었습니다.")
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")

    st.write("---")
    
    if st.session_state.workers:
        df = pd.DataFrame(st.session_state.workers)
        # 사번 열을 문자열로 명시적 지정
        if '사번' in df.columns:
            df['사번'] = df['사번'].astype(str)
        
        c1, c2, c3 = st.columns(3)
        c1.metric("총 인원", f"{len(df)}명")
        c2.metric("방글라데시", f"{len(df[df['국적']=='방글라데시'])}명")
        c3.metric("파키스탄", f"{len(df[df['국적']=='파키스탄'])}명")
        
        st.subheader("📋 전체 인력 명부")
        st.dataframe(df, width='stretch')
        
        # 테이블에서 사번을 클릭하여 수정하기
        st.write("---")
        st.subheader("🔍 인력 정보 조회 및 수정")
        col_search1, col_search2 = st.columns([3, 1])
        
        with col_search1:
            # 사번 선택
            emp_ids_list = df['사번'].unique().tolist()
            selected_emp_for_edit = st.selectbox("수정할 사번 선택", emp_ids_list, key="emp_select_for_edit")
        
        with col_search2:
            st.write("")  # 공간 맞추기
            if st.button("👉 수정하기", use_container_width=True):
                st.session_state.edit_selected_emp_id = selected_emp_for_edit
                st.rerun()
        
        # 선택한 인력 정보 표시 및 수정 폼
        if st.session_state.edit_selected_emp_id:
            selected_worker = next((w for w in st.session_state.workers if w.get('사번') == st.session_state.edit_selected_emp_id), None)
            if selected_worker:
                st.session_state.selected_employee_data = selected_worker
                st.info(f"📋 선택된 인력: **{selected_worker.get('이름')}** (사번: {st.session_state.edit_selected_emp_id})")
                
                st.write("---")
                st.subheader("✏️ 정보 수정")
                
                # 부서 선택
                default_dept = selected_worker.get('부서') if selected_worker else None
                dept_list = list(dept_structure.keys())
                dept_index = dept_list.index(default_dept) if default_dept and default_dept in dept_list else 0
                edit_selected_dept = st.selectbox("소속 부서", dept_list, index=dept_index, key="edit_dept_select")
                    
                edit_col1, edit_col2 = st.columns(2)
                
                # 기본값 설정
                edit_default_emp_id = selected_worker.get('사번', '')
                edit_default_name = selected_worker.get('이름', '')
                edit_default_eng_name = selected_worker.get('영어이름', '')
                edit_default_nation = selected_worker.get('국적', '방글라데시')
                edit_default_entry = selected_worker.get('입국일', None)
                if edit_default_entry and isinstance(edit_default_entry, str):
                    edit_default_entry = datetime.strptime(edit_default_entry, '%Y-%m-%d').date()
                else:
                    edit_default_entry = datetime.now().date()
                
                with edit_col1:
                    edit_emp_id_input = st.text_input("사번", value=edit_default_emp_id, max_chars=6, key="edit_emp_id_input")
                    edit_name = st.text_input("이름", value=edit_default_name, key="edit_name")
                    edit_eng_name = st.text_input("영어 이름", value=edit_default_eng_name, key="edit_eng_name")
                    edit_nation = st.selectbox("국적", ["방글라데시", "파키스탄"], index=(0 if edit_default_nation == "방글라데시" else 1), key="edit_nation")
                    edit_photo = st.file_uploader("인물 사진 (선택 사항)", type=["jpg", "jpeg", "png"], key="edit_worker_photo")

                with edit_col2:
                    Unit_options = dept_structure[edit_selected_dept]["반"]
                    job_options = dept_structure[edit_selected_dept]["직종"]
                    
                    edit_default_unit = selected_worker.get('반') if selected_worker else None
                    edit_unit_index = Unit_options.index(edit_default_unit) if edit_default_unit and edit_default_unit in Unit_options else 0
                    edit_unit = st.selectbox("상세 조직(반)", Unit_options, index=edit_unit_index, key="edit_unit_select")
                    
                    edit_default_job = selected_worker.get('직종') if selected_worker else None
                    edit_job_index = job_options.index(edit_default_job) if edit_default_job and edit_default_job in job_options else 0
                    edit_job = st.selectbox("직종", job_options, index=edit_job_index, key="edit_job_select")
                    
                    edit_entry_date = st.date_input("입국일자", value=edit_default_entry, key="edit_entry_date")

                    # 입국일자부터 오늘까지의 개월수 계산 (소수점)
                    today = datetime.now().date()
                    edit_days_passed = (today - edit_entry_date).days
                    edit_calculated_months = round(edit_days_passed / 30.44, 1)
                    st.write(f"**근속개월**: {edit_calculated_months}개월")
                    edit_service_month = edit_calculated_months

                st.write("---")
                
                # 상세 정보 기본값 설정
                edit_default_h_type = selected_worker.get('숙소구분', '기숙사')
                edit_default_h_addr = selected_worker.get('주소', '')
                edit_default_c_type = selected_worker.get('계약', '부동산')
                edit_default_has_fam = selected_worker.get('가족동반', 'X')
                edit_default_fam_note = selected_worker.get('비고', '')
                
                with st.expander("📋 상세 정보"):
                    st.subheader("숙소 및 가족")
                    edit_col5, edit_col6 = st.columns(2)
                       
                    with edit_col5:
                        edit_h_type_index = 0 if edit_default_h_type == "기숙사" else 1
                        edit_h_type = st.radio("숙소 구분", ["기숙사", "사외"], index=edit_h_type_index, horizontal=True, key="edit_h_type")
                        edit_h_addr = st.text_input("숙소 상세 주소", value=edit_default_h_addr, key="edit_h_addr")

                    with edit_col6:
                        edit_c_type_index = 0 if edit_default_c_type == "부동산" else 1
                        edit_c_type = st.selectbox("계약 구분", ["부동산", "개인"], index=edit_c_type_index, key="edit_c_type")
                        edit_has_fam_index = 0 if edit_default_has_fam == "X" else 1
                        edit_has_fam = st.radio("가족 동반 여부", ["X", "O"], index=edit_has_fam_index, horizontal=True, key="edit_has_fam")
                           
                    edit_fam_note = st.text_area("비고 (가족 상세 및 기타 특이사항)", value=edit_default_fam_note, key="edit_fam_note")

                st.write("---")
                
                edit_col_btn1, edit_col_btn2 = st.columns(2)
                
                with edit_col_btn1:
                    if st.button("💾 수정 저장", use_container_width=True):
                        edit_updated_worker = {
                            "사번": edit_emp_id_input,
                            "이름": edit_name,
                            "영어이름": edit_eng_name,
                            "국적": edit_nation,
                            "부서": edit_selected_dept,
                            "반": edit_unit,
                            "직종": edit_job,
                            "근속개월": edit_service_month,
                            "입국일": edit_entry_date.strftime("%Y-%m-%d"),
                            "숙소구분": edit_h_type,
                            "주소": edit_h_addr,
                            "계약": edit_c_type,
                            "가족동반": edit_has_fam,
                            "비고": edit_fam_note
                        }
                        
                        # 사진 처리
                        if edit_photo:
                            try:
                                img = Image.open(edit_photo)
                                img_resized = img.resize((300, 500), Image.Resampling.LANCZOS)
                                st.session_state.worker_photos[edit_emp_id_input] = img_resized
                            except Exception as e:
                                st.warning(f"사진 처리 중 오류: {e}")
                        
                        # 기존 정보 업데이트
                        for i, worker in enumerate(st.session_state.workers):
                            if worker.get('사번') == st.session_state.edit_selected_emp_id:
                                st.session_state.workers[i] = edit_updated_worker
                                st.session_state.history.append({
                                    "사번": edit_emp_id_input,
                                    "변경일": datetime.now().strftime("%Y-%m-%d"),
                                    "내용": f"{edit_selected_dept} {edit_unit} 정보 수정"
                                })
                                break
                        
                        save_data()
                        st.session_state.selected_employee_data = None
                        st.session_state.edit_selected_emp_id = None
                        st.success(f"✅ 사번 {edit_emp_id_input} 정보가 수정되었습니다.")
                        st.rerun()
                
                with edit_col_btn2:
                    if st.button("❌ 취소", use_container_width=True):
                        st.session_state.selected_employee_data = None
                        st.session_state.edit_selected_emp_id = None
                        st.rerun()
        
        st.write("---")
        st.subheader("🗑️ 인력 삭제")
        col_del1, col_del2 = st.columns([3, 1])
        
        with col_del1:
            delete_emp_id = st.text_input("삭제할 사번 입력", max_chars=6, key="delete_emp_id")
        
        with col_del2:
            st.write("")  # 공간 맞추기
            if st.button("❌ 삭제하기", use_container_width=True):
                if delete_emp_id:
                    existing_worker = next((w for w in st.session_state.workers if w.get("사번") == delete_emp_id), None)
                    if existing_worker:
                        # 삭제 확인
                        st.session_state.confirm_delete_mode = True
                        st.session_state.pending_delete_emp_id = delete_emp_id
                        st.session_state.pending_delete_name = existing_worker.get('이름', '?')
                    else:
                        st.error(f"❌ 사번 {delete_emp_id}를 찾을 수 없습니다.")
                else:
                    st.error("사번을 입력해주세요.")
        
        # 삭제 확인 팝업
        if st.session_state.confirm_delete_mode:
            st.warning(f"⚠️ **사번 {st.session_state.pending_delete_emp_id} - {st.session_state.pending_delete_name} 정보를 삭제하시겠습니까?**")
            col_del_confirm1, col_del_confirm2 = st.columns(2)
            
            with col_del_confirm1:
                if st.button("✅ 네, 삭제하겠습니다", use_container_width=True, key="btn_confirm_delete"):
                    # 해당 사번의 인력 삭제
                    st.session_state.workers = [w for w in st.session_state.workers if w.get("사번") != st.session_state.pending_delete_emp_id]
                    
                    # 사진도 삭제
                    if st.session_state.pending_delete_emp_id in st.session_state.worker_photos:
                        del st.session_state.worker_photos[st.session_state.pending_delete_emp_id]
                    
                    save_data()
                    st.session_state.confirm_delete_mode = False
                    st.success(f"✅ 사번 {st.session_state.pending_delete_emp_id}가 삭제되었습니다.")
                    st.rerun()
            
            with col_del_confirm2:
                if st.button("❌ 아니오, 취소", use_container_width=True, key="btn_cancel_delete"):
                    st.session_state.confirm_delete_mode = False
                    st.rerun()
       
  
        st.write("---")
        csv = df.to_csv(index=False).encode('utf-8-sig')
        excel_data = to_excel(df)

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                "📥 현재 명부 다운로드(CSV)",
                data=csv,
                file_name=f"global_list_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        with col_dl2:
            st.download_button(
                "📥 현재 명부 다운로드(Excel)",
                data=excel_data,
                file_name=f"global_list_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
#        st.write("---")
#        st.subheader("🔄 조직 변경 이력")
#        if st.session_state.history:
#            history_df = pd.DataFrame(st.session_state.history)
#            st.table(history_df)
#
#            history_csv = history_df.to_csv(index=False).encode("utf-8-sig")
#            history_excel = to_excel(history_df)
#
#            h1, h2 = st.columns(2)
#            with h1:
#                st.download_button(
#                    "📥 변경 이력 CSV 다운로드",
#                    data=history_csv,
#                    file_name=f"history_{datetime.now().strftime('%Y%m%d')}.csv",
#                    mime="text/csv"
#                )
#            with h2:
#                st.download_button(
#                    "📥 변경 이력 Excel 다운로드",
#                    data=history_excel,
#                    file_name=f"history_{datetime.now().strftime('%Y%m%d')}.xlsx",
#                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#                )
#    else:
        st.info("등록된 인력이 없습니다. 엑셀 업로드나 개별 등록을 이용해 주세요.")

# ---------------------------------------------------------
# 메뉴 2: 👤 인력 등록/업데이트
# ---------------------------------------------------------
elif menu == "👤 인력 등록/업데이트":
    st.header("👤 인력 정보 관리")
    
    st.subheader("기본 정보")
    
    # edit_selected_emp_id가 있으면 그 정보로 폼 채우기
    if st.session_state.edit_selected_emp_id:
        edit_worker = next((w for w in st.session_state.workers if w.get('사번') == st.session_state.edit_selected_emp_id), None)
        if edit_worker:
            st.session_state.selected_employee_data = edit_worker
            st.info(f"📝 **{edit_worker.get('이름')}** (사번: {st.session_state.edit_selected_emp_id})의 정보를 수정 중입니다.")
    
    # 부서 선택
    default_dept = st.session_state.selected_employee_data.get('부서') if st.session_state.selected_employee_data else None
    dept_list = list(dept_structure.keys())
    dept_index = dept_list.index(default_dept) if default_dept and default_dept in dept_list else 0
    selected_dept = st.selectbox("소속 부서", dept_list, index=dept_index, key="dept_select_main")
        
    col1, col2 = st.columns(2)
    
    # 기본값 설정
    if st.session_state.selected_employee_data:
        default_emp_id = st.session_state.selected_employee_data.get('사번', '')
        default_name = st.session_state.selected_employee_data.get('이름', '')
        default_eng_name = st.session_state.selected_employee_data.get('영어이름', '')
        default_nation = st.session_state.selected_employee_data.get('국적', '방글라데시')
        default_entry = st.session_state.selected_employee_data.get('입국일', None)
        if default_entry and isinstance(default_entry, str):
            default_entry = datetime.strptime(default_entry, '%Y-%m-%d').date()
        else:
            default_entry = datetime.now().date()
    else:
        default_emp_id = ''
        default_name = ''
        default_eng_name = ''
        default_nation = '방글라데시'
        default_entry = datetime.now().date()
    
    with col1:
        emp_id_input = st.text_input("사번", value=default_emp_id, max_chars=6, key="emp_id_input")
        name = st.text_input("이름", value=default_name, key="name")
        eng_name = st.text_input("영어 이름", value=default_eng_name, key="eng_name")
        nation = st.selectbox("국적", ["방글라데시", "파키스탄"], index=(0 if default_nation == "방글라데시" else 1), key="nation")
        
        # 인물 사진 업로드
        photo = st.file_uploader("인물 사진 (선택 사항)", type=["jpg", "jpeg", "png"], key="worker_photo")

    with col2:
        # 선택된 부서를 사용하여 반과 직종 동적 업데이트
        Unit_options = dept_structure[selected_dept]["반"]
        job_options = dept_structure[selected_dept]["직종"]
        
        # 기본값 설정
        default_unit = st.session_state.selected_employee_data.get('반') if st.session_state.selected_employee_data else None
        unit_index = Unit_options.index(default_unit) if default_unit and default_unit in Unit_options else 0
        unit = st.selectbox("상세 조직(반)", Unit_options, index=unit_index, key="unit_select")
        
        default_job = st.session_state.selected_employee_data.get('직종') if st.session_state.selected_employee_data else None
        job_index = job_options.index(default_job) if default_job and default_job in job_options else 0
        job = st.selectbox("직종", job_options, index=job_index, key="job_select")
        
        entry_date = st.date_input("입국일자", value=default_entry, key="entry_date")

        # 입국일자부터 오늘까지의 개월수 계산 (소수점)
        today = datetime.now().date()
        days_passed = (today - entry_date).days
        calculated_months = round(days_passed / 30.44, 1)
        st.write(f"**근속개월**: {calculated_months}개월")
        service_month = calculated_months

    st.write("---")
    
    # 상세 정보 기본값 설정
    if st.session_state.selected_employee_data:
        default_h_type = st.session_state.selected_employee_data.get('숙소구분', '기숙사')
        default_h_addr = st.session_state.selected_employee_data.get('주소', '')
        default_c_type = st.session_state.selected_employee_data.get('계약', '부동산')
        default_has_fam = st.session_state.selected_employee_data.get('가족동반', 'X')
        default_fam_note = st.session_state.selected_employee_data.get('비고', '')
    else:
        default_h_type = '기숙사'
        default_h_addr = ''
        default_c_type = '부동산'
        default_has_fam = 'X'
        default_fam_note = ''
    
    with st.expander("📋 상세 정보"):
        st.subheader("숙소 및 가족")
        col5, col6 = st.columns(2)
           
        with col5:
            h_type_index = 0 if default_h_type == "기숙사" else 1
            h_type = st.radio("숙소 구분", ["기숙사", "사외"], index=h_type_index, horizontal=True, key="h_type")
            h_addr = st.text_input("숙소 상세 주소", value=default_h_addr, key="h_addr")

        with col6:
            c_type_index = 0 if default_c_type == "부동산" else 1
            c_type = st.selectbox("계약 구분", ["부동산", "개인"], index=c_type_index, key="c_type")
            has_fam_index = 0 if default_has_fam == "X" else 1
            has_fam = st.radio("가족 동반 여부", ["X", "O"], index=has_fam_index, horizontal=True, key="has_fam")
               
        fam_note = st.text_area("비고 (가족 상세 및 기타 특이사항)", value=default_fam_note, key="fam_note")

    st.write("---")
    
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        button_label = "인력 수정" if st.session_state.edit_selected_emp_id else "인력 등록"
        if st.button(button_label, use_container_width=True):
            if len(emp_id_input) == 6:
                # 동일한 사번이 있는지 확인
                existing_worker = next((w for w in st.session_state.workers if w.get("사번") == emp_id_input), None)
                
                new_worker = {
                    "사번": emp_id_input,
                    "이름": name,
                    "영어이름": eng_name,
                    "국적": nation,
                    "부서": selected_dept,
                    "반": unit,
                    "직종": job,
                    "근속개월": service_month,
                    "입국일": entry_date.strftime("%Y-%m-%d"),
                    "숙소구분": h_type,
                    "주소": h_addr,
                    "계약": c_type,
                    "가족동반": has_fam,
                    "비고": fam_note
                }
                
                # 사진 처리
                if photo:
                    try:
                        img = Image.open(photo)
                        # 3:5 비율로 리사이징 (width:300, height:500)
                        img_resized = img.resize((300, 500), Image.Resampling.LANCZOS)
                        st.session_state.worker_photos[emp_id_input] = img_resized
                    except Exception as e:
                        st.warning(f"사진 처리 중 오류: {e}")
                
                if existing_worker:
                    # 기존 정보가 있으면 확인 메시지 표시
                    st.session_state.show_update_confirm = True
                    st.session_state.pending_worker_data = new_worker
                else:
                    # 기존 정보가 없으면 신규 등록
                    st.session_state.workers.append(new_worker)
                    st.session_state.history.append({
                        "사번": emp_id_input, "변경일": datetime.now().strftime("%Y-%m-%d"),
                        "내용": f"{selected_dept} {unit} 신규 배정"
                    })
                    save_data()  # 데이터 자동 저장
                    st.session_state.selected_employee_data = None
                    st.session_state.edit_selected_emp_id = None
                    st.success(f"사번 {emp_id_input} 등록 완료!")
                    st.rerun()
            else:
                st.error("사번은 반드시 6자리여야 합니다.")
    
    with col_btn2:
        if st.button("🔄 초기화", use_container_width=True):
            st.session_state.selected_employee_data = None
            st.session_state.edit_selected_emp_id = None
            st.rerun()
    
    # 업데이트 확인 팝업
    if st.session_state.show_update_confirm:
        st.warning("⚠️ **기존 정보가 존재합니다. 기존 정보를 업데이트 하시겠습니까?**")
        col_confirm1, col_confirm2 = st.columns(2)
        
        with col_confirm1:
            if st.button("✅ 예, 업데이트합니다", key="btn_yes"):
                # 기존 정보 찾기 및 업데이트
                for i, worker in enumerate(st.session_state.workers):
                    if worker.get("사번") == st.session_state.pending_worker_data["사번"]:
                        st.session_state.workers[i] = st.session_state.pending_worker_data
                        st.session_state.history.append({
                            "사번": st.session_state.pending_worker_data["사번"],
                            "변경일": datetime.now().strftime("%Y-%m-%d"),
                            "내용": f"{st.session_state.pending_worker_data['부서']} {st.session_state.pending_worker_data['반']} 정보 업데이트"
                        })
                        save_data()  # 데이터 자동 저장
                        st.success(f"사번 {st.session_state.pending_worker_data['사번']} 정보가 업데이트되었습니다!")
                        break
                
                st.session_state.show_update_confirm = False
                st.session_state.pending_worker_data = None
                st.session_state.selected_employee_data = None
                st.rerun()
        
        with col_confirm2:
            if st.button("❌ 아니오, 취소합니다", key="btn_no"):
                st.session_state.show_update_confirm = False
                st.session_state.pending_worker_data = None
                st.info("업데이트가 취소되었습니다.")
                st.rerun()
    
    st.write("---")
    
    # 등록된 인력 현황 표시
    if st.session_state.workers:
        st.subheader("📋 등록된 인력 현황")
        
        # 데이터 테이블 표시
        df_workers = pd.DataFrame(st.session_state.workers)
        st.dataframe(df_workers, use_container_width=True)
        
        # 인물 카드 형식으로 사진과 정보 함께 표시
        st.subheader("👥 인물 카드")
        for idx, worker in enumerate(st.session_state.workers):
            col_photo, col_info = st.columns([1, 3])
            
            with col_photo:
                emp_id_display = worker.get("사번")
                if emp_id_display in st.session_state.worker_photos:
                    st.image(st.session_state.worker_photos[emp_id_display], use_container_width=True)
                else:
                    st.info("📷 사진 없음")
            
            with col_info:
                st.write(f"**사번**: {worker.get('사번')}")
                st.write(f"**이름**: {worker.get('이름')} ({worker.get('영어이름')})")
                st.write(f"**국적**: {worker.get('국적')}")
                st.write(f"**부서**: {worker.get('부서')} / **반**: {worker.get('반')} / **직종**: {worker.get('직종')}")
                st.write(f"**근속개월**: {worker.get('근속개월')}개월 | **입국일**: {worker.get('입국일')}")
                st.write(f"**주소**: {worker.get('주소')}")
                st.write(f"**비고**: {worker.get('비고')}")
            
            st.divider()
        
        # 다운로드 기능
        csv_data = df_workers.to_csv(index=False).encode('utf-8-sig')
        excel_data = to_excel(df_workers)
        
        col_down1, col_down2 = st.columns(2)
        with col_down1:
            st.download_button(
                "📥 인력 목록 다운로드 (CSV)",
                data=csv_data,
                file_name=f"workers_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        with col_down2:
            st.download_button(
                "📥 인력 목록 다운로드 (Excel)",
                data=excel_data,
                file_name=f"workers_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    st.write("---")
    
    # 엑셀 업로드 기능
    with st.expander("📤 기존 인력 엑셀 파일 일괄 등록", expanded=False):
        uploaded_file = st.file_uploader("엑셀 파일(.xlsx)을 업로드하세요", type=["xlsx"], key="worker_upload")
        if uploaded_file:
            try:
                df_upload = pd.read_excel(uploaded_file)
                df_upload = clean_excel_data(df_upload)  # 날짜 데이터 정제
                st.dataframe(df_upload, width='stretch')
                
                if st.button("시스템에 데이터 반영하기", key="confirm_upload"):
                    uploaded_data = df_upload.to_dict('records')
                    
                    # 중복 체크 및 추가
                    added_count = 0
                    duplicate_count = 0
                    
                    for new_worker in uploaded_data:
                        # 사번이 이미 존재하는지 확인
                        existing = next((w for w in st.session_state.workers if w.get('사번') == new_worker.get('사번')), None)
                        if existing:
                            duplicate_count += 1
                        else:
                            st.session_state.workers.append(new_worker)
                            added_count += 1
                    
                    save_data()  # 데이터 자동 저장
                    
                    if duplicate_count > 0:
                        st.warning(f"⚠️ {duplicate_count}명은 이미 존재합니다. {added_count}명의 새로운 데이터가 추가되었습니다.")
                    else:
                        st.success(f"✅ {added_count}명의 데이터가 추가되었습니다.")
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")
# ---------------------------------------------------------
# 메뉴 3: 📚 한국어 교육/자격 관리
# ---------------------------------------------------------
elif menu == "📚 한국어 교육/자격 관리":
    st.header("📚 한국어 역량")
    
    t1, t2 = st.tabs(["🎓 자격증 취득 정보 등록", "🏫 교육 수업 참여"])
    
    with t1:
        st.subheader("🎓 자격증 취득 정보 등록")

        c_id = st.text_input("사번 입력", key="cert_id")
        exam_type = st.selectbox("시험 종류", ["사회통합프로그램", "사전평가", "TOPIK"], key="exam_type")

        if exam_type == "사회통합프로그램":
            level_options = [f"{i}단계" for i in range(1, 7)]
            label_text = "단계 선택"
        elif exam_type == "사전평가":
            level_options = [f"{i}급" for i in range(1, 6)]
            label_text = "급수 선택"
        else:
            level_options = [f"{i}급" for i in range(1, 7)]
            label_text = "급수 선택"

        c_level = st.selectbox(label_text, level_options, key="cert_level")
        c_date = st.date_input("취득일", key="cert_date")

        if st.button("자격증 저장", key="save_cert"):
            if c_id:
                st.session_state.korean_certs.append({
                    "사번": c_id,
                    "시험종류": exam_type,
                    "자격": c_level,
                    "취득일": c_date.strftime("%Y-%m-%d")
                })
                save_data()  # 데이터 자동 저장
                st.success(f"✅ 저장 완료: {exam_type} {c_level}")
            else:
                st.error("사번을 입력해 주세요.")

    with t2:
        st.subheader("🏫 교육 수업 참여")

        with st.form("class_form"):
            cl_id = st.text_input("사번 입력", key="class_id")
            cl_name = st.selectbox("수업명", ["사내 사통교육", "사외 사통교육", "통역사 주관 수업"], key="class_name")
            cl_date = st.selectbox("시기", ["2026년 상반기", "2026년 하반기"], key="class_date")

            if st.form_submit_button("수업 이력 저장"):
                if cl_id:
                    st.session_state.korean_classes.append({
                        "사번": cl_id,
                        "수업명": cl_name,
                        "시기": cl_date
                    })
                    save_data()  # 데이터 자동 저장
                    st.success("저장되었습니다.")
                else:
                    st.error("사번을 입력해 주세요.")