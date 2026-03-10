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
menu = st.sidebar.radio("메뉴 선택", ["📊 통합 대시보드", "👤 인력 정보 관리", "📚 한국어 교육/자격 관리"])

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
# 메뉴 2: 👤 인력 정보 관리
# ---------------------------------------------------------
elif menu == "👤 인력 정보 관리":
    st.header("👤 인력 정보 관리")
    
    # 탭 구성
    tab1, tab2 = st.tabs(["📥 데이터 및 사진 업로드", "👥 인력 정보 조회"])
    
    with tab1:
        st.subheader("📥 엑셀 데이터 일괄 업로드")
        st.write("엑셀 파일(사번, 이름, 영어이름, ... 등)을 업로드합니다.")
        
        uploaded_excel = st.file_uploader("엑셀 파일(.xlsx)을 업로드하세요", type=["xlsx"], key="upload_excel_menu2")
        if uploaded_excel:
            try:
                df_excel = pd.read_excel(uploaded_excel)
                df_excel = clean_excel_data(df_excel)  # 날짜 데이터 정제
                st.dataframe(df_excel, width='stretch')
                
                if st.button("✅ 데이터 저장하기", key="save_excel_menu2"):
                    excel_data = df_excel.to_dict('records')
                    
                    # 중복 체크 및 추가
                    added_count = 0
                    updated_count = 0
                    duplicate_count = 0
                    
                    for new_worker in excel_data:
                        emp_id = new_worker.get('사번')
                        existing = next((w for w in st.session_state.workers if w.get('사번') == emp_id), None)
                        
                        if existing:
                            # 기존 데이터 업데이트
                            for i, w in enumerate(st.session_state.workers):
                                if w.get('사번') == emp_id:
                                    st.session_state.workers[i] = new_worker
                                    st.session_state.history.append({
                                        "사번": emp_id,
                                        "변경일": datetime.now().strftime("%Y-%m-%d"),
                                        "내용": "엑셀 일괄 업로드로 정보 업데이트"
                                    })
                                    updated_count += 1
                                    break
                        else:
                            # 신규 등록
                            st.session_state.workers.append(new_worker)
                            st.session_state.history.append({
                                "사번": emp_id,
                                "변경일": datetime.now().strftime("%Y-%m-%d"),
                                "내용": "엑셀 일괄 업로드로 신규 등록"
                            })
                            added_count += 1
                    
                    save_data()
                    
                    st.success(f"✅ 저장 완료: 신규 {added_count}명, 업데이트 {updated_count}명")
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")
        
        st.write("---")
        st.subheader("📷 사진 일괄 업로드")
        st.write("사번.jpg 형식으로 이름이 지정된 사진들을 업로드합니다. (예: 111111.jpg, 222222.jpg, ...)")
        
        uploaded_photos = st.file_uploader(
            "사진 파일을 선택하세요 (JPG, PNG)", 
            type=["jpg", "jpeg", "png"], 
            accept_multiple_files=True,
            key="upload_photos_menu2"
        )
        
        if uploaded_photos:
            photo_preview = []
            for photo_file in uploaded_photos:
                # 파일명에서 사번 추출 (확장자 제거)
                file_name = photo_file.name
                emp_id_from_file = file_name.split('.')[0]  # '111111.jpg' → '111111'
                photo_preview.append({
                    "파일명": file_name,
                    "사번": emp_id_from_file
                })
            
            preview_df = pd.DataFrame(photo_preview)
            st.dataframe(preview_df, width='stretch', use_container_width=True)
            
            if st.button("✅ 사진 저장하기", key="save_photos_menu2"):
                success_count = 0
                error_count = 0
                
                for photo_file in uploaded_photos:
                    try:
                        file_name = photo_file.name
                        emp_id = file_name.split('.')[0]
                        
                        # 해당 사번 존재 확인
                        existing = next((w for w in st.session_state.workers if w.get('사번') == emp_id), None)
                        
                        if existing:
                            # 사진 처리 (3:5 비율로 리사이징)
                            img = Image.open(photo_file)
                            img_resized = img.resize((300, 500), Image.Resampling.LANCZOS)
                            st.session_state.worker_photos[emp_id] = img_resized
                            success_count += 1
                        else:
                            st.warning(f"⚠️ 사번 {emp_id}는 등록되지 않았습니다. (파일: {file_name})")
                            error_count += 1
                    except Exception as e:
                        st.warning(f"사진 처리 중 오류 - {photo_file.name}: {e}")
                        error_count += 1
                
                save_data()
                st.success(f"✅ 사진 저장 완료: 성공 {success_count}개" + (f", 실패 {error_count}개" if error_count > 0 else ""))
    
    with tab2:
        st.subheader("👥 등록된 인력 정보 조회")
        
        if st.session_state.workers:
            # 필터 폼 - 부서/반 선택
            with st.form("filter_form"):
                st.write("**필터 조건 선택**")
                
                filter_col1, filter_col2 = st.columns(2)
                
                with filter_col1:
                    dept_list = list(dept_structure.keys())
                    selected_filter_dept = st.selectbox(
                        "부서 선택",
                        ["전체"] + dept_list,
                        key="filter_dept"
                    )
                
                with filter_col2:
                    # 선택된 부서에 따라 반 옵션 변경
                    if selected_filter_dept == "전체":
                        available_units = []
                        for dept in dept_list:
                            available_units.extend(dept_structure[dept]["반"])
                        available_units = sorted(list(set(available_units)))
                    else:
                        available_units = dept_structure[selected_filter_dept]["반"]
                    
                    selected_filter_unit = st.selectbox(
                        "반 선택",
                        ["전체"] + available_units,
                        key="filter_unit"
                    )
                
                # 검색 창
                search_term = st.text_input(
                    "이름 또는 사번으로 검색 (선택 사항)",
                    placeholder="이름이나 사번 입력...",
                    key="filter_search"
                )
                
                submitted = st.form_submit_button("🔍 조회", use_container_width=True)
            
            st.write("---")
            
            # 필터링 로직
            filtered_workers = st.session_state.workers.copy()
            
            # 부서 필터
            if selected_filter_dept != "전체":
                filtered_workers = [w for w in filtered_workers if w.get('부서') == selected_filter_dept]
            
            # 반 필터
            if selected_filter_unit != "전체":
                filtered_workers = [w for w in filtered_workers if w.get('반') == selected_filter_unit]
            
            # 검색 필터
            if search_term:
                filtered_workers = [w for w in filtered_workers 
                                   if search_term in str(w.get('사번', '')) or 
                                      search_term in w.get('이름', '')]
            
            # 결과 표시
            st.write(f"**조회 결과: {len(filtered_workers)}명** (전체: {len(st.session_state.workers)}명)")
            st.write("---")
            
            if len(filtered_workers) > 0:
                # 인력 개별 표시
                for idx, worker in enumerate(filtered_workers):
                    with st.container():
                        col_photo, col_info = st.columns([1, 4])
                        
                        with col_photo:
                            emp_id = worker.get("사번")
                            if emp_id in st.session_state.worker_photos:
                                st.image(st.session_state.worker_photos[emp_id], use_container_width=True, caption=f"사번: {emp_id}")
                            else:
                                st.info("📷\n사진\n없음")
                        
                        with col_info:
                            st.markdown(f"### 🆔 사번: {worker.get('사번')}")
                            
                            col_l, col_r = st.columns(2)
                            with col_l:
                                st.write(f"**📛 이름**: {worker.get('이름')}")
                                st.write(f"**🌐 영어이름**: {worker.get('영어이름')}")
                                st.write(f"**🌍 국적**: {worker.get('국적')}")
                                st.write(f"**📅 입국일**: {worker.get('입국일')}")
                                st.write(f"**⏱️ 근속개월**: {worker.get('근속개월')}개월")
                            
                            with col_r:
                                st.write(f"**🏢 부서**: {worker.get('부서')}")
                                st.write(f"**👥 반**: {worker.get('반')}")
                                st.write(f"**💼 직종**: {worker.get('직종')}")
                                st.write(f"**🏠 숙소**: {worker.get('숙소구분')}")
                                st.write(f"**👨‍👩‍👧 가족동반**: {worker.get('가족동반')}")
                            
                            with st.expander("📋 상세 정보"):
                                st.write(f"**주소**: {worker.get('주소', 'N/A')}")
                                st.write(f"**계약 구분**: {worker.get('계약', 'N/A')}")
                                st.write(f"**비고**: {worker.get('비고', 'N/A')}")
                        
                        st.divider()
            else:
                st.info("선택한 조건에 해당하는 인력이 없습니다.")
        else:
            st.info("등록된 인력이 없습니다. 먼저 엑셀 파일을 업로드해주세요.")


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