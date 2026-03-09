import streamlit as st
import pandas as pd
from datetime import datetime
import io

# 1. 페이지 설정 및 제목
st.set_page_config(page_title="Global 외국인 인력 관리 시스템", layout="wide")

# 2. 데이터 유지용 세션 상태 초기화 (브라우저 열려있는 동안 유지)
if 'workers' not in st.session_state:
    st.session_state.workers = []
if 'history' not in st.session_state:
    st.session_state.history = []
if 'korean_certs' not in st.session_state:
    st.session_state.korean_certs = []
if 'korean_classes' not in st.session_state:
    st.session_state.korean_classes = []

# 3. 조직 데이터 정의 (부서-반-직종 연동 구조)
dept_structure = {
    "공사1부5과": {
        "반": ["1직1반", "1직2반", "1직3반", "2직1반", "2직2반", "2직3반"],
        "직종": ["수동본딩", "ABM", "TBP"]
    },
    "공사2부3과": {
        "반": ["설치직1반", "설치직2반", "설치직3반", "설치직4반", "용접직1반", "용접직2반", "용접직3반"],
        "직종": ["MB설치", "MB수동용접", "MB자동용접", "MB리웰딩"]
    },
    "공사3부": {
        "반": ["2직1반", "2직2반", "2직3반"],
        "직종": ["의장", "LNGTIG"]
    }
}

# 4. 사이드바 메뉴 구성
st.sidebar.title("🌍 Global HRM System")
menu = st.sidebar.radio("메뉴 선택", 
    ["📊 대시보드 및 업로드", "👤 인력 신규 등록", "📚 한국어 교육/자격 관리"])

# --- 메뉴 1: 대시보드 및 업로드 (171명 대량 등록 및 현황) ---
if menu == "📊 대시보드 및 업로드":
    st.header("📊 인력 관리 현황 및 데이터 업로드")
    
    # 엑셀 업로드 섹션
    with st.expander("📥 기존 인력 엑셀 파일 일괄 등록", expanded=False):
        uploaded_file = st.file_uploader("엑셀 파일(.xlsx)을 업로드하세요", type=["xlsx"])
        if uploaded_file:
            try:
                df_upload = pd.read_excel(uploaded_file)
                if st.button("시스템에 데이터 반영하기"):
                    uploaded_data = df_upload.to_dict('records')
                    st.session_state.workers.extend(uploaded_data)
                    st.success(f"✅ {len(df_upload)}명의 데이터가 추가되었습니다.")
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")

    st.write("---")
    
    # 현황 테이블
    if st.session_state.workers:
        df = pd.DataFrame(st.session_state.workers)
        
        # 상단 요약 지표
        c1, c2, c3 = st.columns(3)
        c1.metric("총 인원", f"{len(df)}명")
        c2.metric("방글라데시", f"{len(df[df['국적']=='방글라데시'])}명")
        c3.metric("파키스탄", f"{len(df[df['국적']=='파키스탄'])}명")
        
        st.subheader("📋 전체 인력 명부")
        st.dataframe(df, use_container_width=True)
        
        # 다운로드 버튼
        csv = df.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 현재 명부 다운로드(CSV)", data=csv, file_name=f"global_list_{datetime.now().strftime('%Y%m%d')}.csv")
        
        st.write("---")
        st.subheader("🔄 조직 변경 이력")
        if st.session_state.history:
            st.table(pd.DataFrame(st.session_state.history))
    else:
        st.info("등록된 인력이 없습니다. 엑셀 업로드나 개별 등록을 이용해 주세요.")

# --- 메뉴 2: 인력 신규 등록 (개별 등록) ---
elif menu == "👤 인력 신규 등록":
    st.header("👤 신규 인력 정보 입력")
    
    with st.form("individual_reg"):
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("기본 정보")
            emp_id = st.text_input("사번 (6자리 고유번호)", max_chars=6)
            nation = st.selectbox("국적", ["방글라데시", "파키스탄"])
            dept = st.selectbox("소속 부서", list(dept_structure.keys()))
            unit = st.selectbox("상세 조직(반)", dept_structure[dept]["반"])
            job = st.selectbox("직종", dept_structure[dept]["직종"])
            entry_date = st.date_input("입국일자", datetime.now())
        
        with col2:
            st.subheader("숙소 및 가족")
            h_type = st.radio("숙소 구분", ["기숙사", "사외"], horizontal=True)
            h_addr = st.text_input("숙소 상세 주소")
            c_type = st.selectbox("계약 구분", ["부동산", "개인"])
            st.write("---")
            has_fam = st.radio("가족 동반 여부", ["X", "O"], horizontal=True)
            fam_note = st.text_area("비고 (가족 상세 및 기타 특이사항)")
            
        if st.form_submit_button("인력 등록"):
            if len(emp_id) == 6:
                new_worker = {
                    "사번": emp_id, "국적": nation, "부서": dept, "반": unit, "직종": job,
                    "입국일": entry_date.strftime("%Y-%m-%d"), "숙소구분": h_type, 
                    "주소": h_addr, "계약": c_type, "가족동반": has_fam, "비고": fam_note
                }
                st.session_state.workers.append(new_worker)
                st.session_state.history.append({
                    "사번": emp_id, "변경일": datetime.now().strftime("%Y-%m-%d"),
                    "내용": f"{dept} {unit} 신규 배정"
                })
                st.success(f"사번 {emp_id} 등록 완료!")
            else:
                st.error("사번은 반드시 6자리여야 합니다.")

# --- 메뉴 3: 한국어 교육/자격 관리 ---
elif menu == "📚 한국어 교육/자격 관리":
    st.header("📚 한국어 역량 관리")
    
    tab1, tab2 = st.tabs(["🎓 자격증 취득 이력", "🏫 교육 수업 참여"])
    
    with tab1:
        with st.form("cert_input"):
            c_id = st.text_input("사번 입력")
            c_level = st.selectbox("자격 종류", ["사통1급", "사통2급", "사통3급", "사통4급", "사통5급"])
            c_date = st.date_input("취득 날짜")
            if st.form_submit_button("자격 저장"):
                st.session_state.korean_certs.append({"사번": c_id, "자격": c_level, "취득일": c_date})
                st.success("자격 정보 저장 완료")
        if st.session_state.korean_certs:
            st.table(pd.DataFrame(st.session_state.korean_certs))

    with tab2:
        with st.form("class_input"):
            cl_id = st.text_input("사번 입력")
            cl_term = st.selectbox("구분", ["2026년 상반기", "2026년 하반기"])
            cl_name = st.selectbox("수업 명칭", ["사내 사통교육", "사외 사통교육", "2급이상 수업", "통역사주관수업"])
            if st.form_submit_button("수업 이력 저장"):
                st.session_state.korean_classes.append({"사번": cl_id, "시기": cl_term, "수업명": cl_name})
                st.success("수업 이력 저장 완료")
        if st.session_state.korean_classes:
            st.table(pd.DataFrame(st.session_state.korean_classes))