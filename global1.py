import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime, timedelta
import io

# 1. 페이지 설정
st.set_page_config(page_title="LNG선공사팀 글로벌인력관리", layout="wide")

# ---------------------------------------------------------
# [데이터 초기화] Session State 설정
# ---------------------------------------------------------
if 'workers' not in st.session_state:
    st.session_state.workers = []
if 'korean_certs' not in st.session_state:
    st.session_state.korean_certs = []
if 'korean_classes' not in st.session_state:
    st.session_state.korean_classes = []

# 엑셀 다운로드 함수
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    return output.getvalue()

# 조직 구조 정의
dept_structure = {
    "공사1부5과": {"반": ["1직1반", "1직2반", "1직3반", "2직1반", "2직2반", "2직3반"], "직종": ["수동본딩", "ABM", "TBP"]},
    "공사2부3과": {"반": ["설치직1반", "설치직2반", "설치직3반", "설치직4반", "용접직1반", "용접직2반", "용접직3반"], "직종": ["MB설치", "MB수동용접", "MB자동용접", "MB리웰딩"]},
    "공사3부": {"반": ["2직1반", "2직2반", "2직3반"], "직종": ["의장", "LNGTIG"]}
}

# ---------------------------------------------------------
# [사이드바] 메뉴 구성
# ---------------------------------------------------------
st.sidebar.title("🌍 LNG선공사팀 글로벌인력관리")
menu = st.sidebar.radio("메뉴 선택", ["📊 통합 대시보드", "👤 인력 등록/업로드", "📚 한국어 교육/자격 관리"])

# ---------------------------------------------------------
# 메뉴 1: 📊 통합 대시보드 (분석 및 시각화)
# ---------------------------------------------------------
if menu == "📊 통합 대시보드":
    st.header("📊 글로벌 인력 현황 및 데이터 업로드")
    
    # 데이터 병합 (인력 + 자격증)
    if st.session_state.workers and st.session_state.korean_certs:
        df_w = pd.DataFrame(st.session_state.workers)
        df_c = pd.DataFrame(st.session_state.korean_certs)
        # 사번을 기준으로 병합
        df = pd.merge(df_w, df_c, on="사번", how="left")
        df['고등급여부'] = df['자격'].apply(lambda x: int(str(x)[-2]) >= 3 if pd.notnull(x) else False)
        
        # 상단 KPI 섹션
        total_emp = 171 # 관리 대상 총 인원
        current_reg = len(df_w)
        cert_holders = df[df['자격'].notnull()]['사번'].nunique()
        high_grade_holders = df[df['고등급여부'] == True]['사번'].nunique()

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("전체 관리대상", f"{total_emp}명")
        col2.metric("현재 등록인원", f"{current_reg}명")
        col3.metric("자격 보유율", f"{(cert_holders/total_emp)*100:.1f}%", delta="목표 91%")
        col4.metric("고등급(3급↑) 비율", f"{(high_grade_holders/total_emp)*100:.1f}%")

        st.divider()

        # 시각화 차트
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("📊 자격증 종류별 분포")
            fig1 = px.pie(df[df['자격'].notnull()], names='자격', hole=0.4)
            st.plotly_chart(fig1, use_container_width=True)
            
        with c2:
            st.subheader("🏫 반별 자격 및 고등급 현황")
            class_stats = df.groupby('반').agg(
                취득인원=('사번', 'nunique'),
                고등급인원=('고등급여부', 'sum')
            ).reset_index()
            fig2 = px.bar(class_stats, x='반', y=['취득인원', '고등급인원'], barmode='group')
            st.plotly_chart(fig2, use_container_width=True)

        # 사원 상세 검색
        st.divider()
        st.subheader("🔍 사원별 상세 이력 조회")
        search_id = st.selectbox("조회할 사번을 선택하세요", options=[""] + sorted(df_w['사번'].unique().tolist()))
        if search_id:
            p_info = df_w[df_w['사번'] == search_id]
            p_certs = df_c[df_c['사번'] == search_id]
            p_classes = pd.DataFrame(st.session_state.korean_classes)
            p_classes = p_classes[p_classes['사번'] == search_id] if not p_classes.empty else p_classes

            col_a, col_b = st.columns([1, 2])
            with col_a:
                st.info(f"**신상 정보 ({search_id})**")
                st.write(p_info[['국적', '부서', '반', '직종']].iloc[0])
            with col_b:
                st.write("📋 **교육 및 자격 이력**")
                st.write("- 자격증:", p_certs[['자격', '취득일']] if not p_certs.empty else "내역 없음")
                st.write("- 수강이력:", p_classes[['시기', '수업명']] if not p_classes.empty else "내역 없음")
    else:
        st.info("대시보드 분석을 위해 인력 정보와 자격증 정보를 먼저 등록해주세요.")

# ---------------------------------------------------------
# 메뉴 2: 👤 인력 등록/업로드
# ---------------------------------------------------------
elif menu == "👤 인력 등록/업로드":
    st.header("👤 인력 정보 관리")
    
    with st.expander("📥 엑셀 파일로 일괄 등록", expanded=True):
        uploaded_file = st.file_uploader("사원 명부 엑셀 업로드", type=["xlsx"])
        if uploaded_file:
            df_upload = pd.read_excel(uploaded_file)
            if st.button("시스템에 데이터 반영"):
                st.session_state.workers.extend(df_upload.to_dict('records'))
                st.success(f"{len(df_upload)}명 등록 완료")

    st.write("---")
    with st.form("individual_reg"):
        st.subheader("기본 정보 등록")
        c1, c2 = st.columns(2)
        with c1:
            emp_id = st.text_input("사번 (6자리)", max_chars=6)
            nation = st.selectbox("국적", ["방글라데시", "파키스탄"])
            dept = st.selectbox("부서", list(dept_structure.keys()))
        with c2:
            unit = st.selectbox("반", dept_structure[dept]["반"])
            job = st.selectbox("직종", dept_structure[dept]["직종"])
            entry_date = st.date_input("입국일", datetime.now())
        
        if st.form_submit_button("등록하기"):
            new_worker = {"사번": emp_id, "국적": nation, "부서": dept, "반": unit, "직종": job, "입국일": entry_date.strftime("%Y-%m-%d")}
            st.session_state.workers.append(new_worker)
            st.success(f"사번 {emp_id} 등록 성공")

# ---------------------------------------------------------
# 메뉴 3: 📚 한국어 교육/자격 관리
# ---------------------------------------------------------
elif menu == "📚 한국어 교육/자격 관리":
    st.header("📚 한국어 역량")
    
    t1, t2 = st.tabs(["🎓 자격증 취득 정보 등록", "🏫 교육 수업 참여"])
    
    with t1:
        st.subheader("🎓 자격증 취득 정보 등록")
        with st.form("cert_form"):
            c_id = st.text_input("사번 입력")
            
            # 1. 시험 종류 선택
            exam_type = st.selectbox("시험 종류", ["사회통합프로그램", "사전평가", "TOPIK"])
            
            # 2. 시험 종류에 따른 동적 등급/단계 설정
            if exam_type == "사회통합프로그램":
                # 사회통합프로그램은 1단계 ~ 6단계
                level_options = [f"{i}단계" for i in range(1, 7)]
                label_text = "단계 선택"
            elif exam_type == "사전평가":
                # 사전평가는 1급 ~ 5급
                level_options = [f"{i}급" for i in range(1, 6)]
                label_text = "급수 선택"
            else:
                # TOPIK은 1급 ~ 5급 (6급까지 필요하시면 range(1, 7)로 수정)
                level_options = [f"{i}급" for i in range(1, 6)]
                label_text = "급수 선택"
                
            c_level = st.selectbox(label_text, level_options)
            c_date = st.date_input("취득일")
            
            if st.form_submit_button("자격증 저장"):
                if c_id:
                    st.session_state.korean_certs.append({
                        "사번": c_id, 
                        "시험종류": exam_type,
                        "자격": c_level, 
                        "취득일": c_date.strftime("%Y-%m-%d")
                    })
                    st.success(f"✅ 저장 완료: {exam_type} {c_level}")
                else:
                    st.error("사번을 입력해 주세요.")

    with t2:
        with st.form("class_form"):
            cl_id = st.text_input("사번 입력")
            cl_name = st.selectbox("수업명", ["사내 사통교육", "사외 사통교육", "통역사 주관 수업"])
            cl_date = st.selectbox("시기", ["2026년 상반기", "2026년 하반기"])
            if st.form_submit_button("수업 이력 저장"):
                st.session_state.korean_classes.append({"사번": cl_id, "수업명": cl_name, "시기": cl_date})
                st.success("저장되었습니다.")