import streamlit as st
import pandas as pd
import numpy as np
#import plotly.express as px
from datetime import datetime, timedelta
import io
from PIL import Image
import json
import os
import matplotlib.pyplot as plt
import matplotlib as mpl

# Matplotlib 한글 폰트 설정
mpl.rcParams['font.family'] = 'sans-serif'
try:
    # Windows에서 사용 가능한 한글 폰트
    mpl.rcParams['font.sans-serif'] = ['Malgun Gothic', 'DejaVu Sans']
except:
    pass
mpl.rcParams['axes.unicode_minus'] = False

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
    """모든 데이터를 JSON 파일로 저장하고 사진을 폴더에 저장"""
    data_dir = "data"
    photos_dir = f"{data_dir}/photos"
    
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)
    
    if not os.path.exists(photos_dir):
        os.makedirs(photos_dir)
    
    # JSON 데이터 저장
    with open(f"{data_dir}/workers.json", "w", encoding="utf-8") as f:
        json.dump(st.session_state.workers, f, ensure_ascii=False, indent=2)
    
    with open(f"{data_dir}/history.json", "w", encoding="utf-8") as f:
        json.dump(st.session_state.history, f, ensure_ascii=False, indent=2)
    
    with open(f"{data_dir}/korean_certs.json", "w", encoding="utf-8") as f:
        json.dump(st.session_state.korean_certs, f, ensure_ascii=False, indent=2)
    
    with open(f"{data_dir}/korean_classes.json", "w", encoding="utf-8") as f:
        json.dump(st.session_state.korean_classes, f, ensure_ascii=False, indent=2)
    
    # 메타데이터 저장 (파일명 등)
    metadata = {
        "uploaded_worker_file_name": st.session_state.uploaded_worker_file_name,
        "uploaded_cert_file_name": st.session_state.uploaded_cert_file_name,
        "uploaded_class_file_name": st.session_state.uploaded_class_file_name
    }
    with open(f"{data_dir}/metadata.json", "w", encoding="utf-8") as f:
        json.dump(metadata, f, ensure_ascii=False, indent=2)
    
    # 사진 저장 (PNG 형식으로 사번.png)
    for emp_id, photo in st.session_state.worker_photos.items():
        try:
            photo_path = f"{photos_dir}/{emp_id}.png"
            if isinstance(photo, Image.Image):
                photo.save(photo_path)
        except Exception as e:
            print(f"사진 저장 오류 - 사번 {emp_id}: {e}")

# 데이터 로드 함수
def load_data():
    """JSON 파일에서 데이터 로드하고 사진 폴더에서 이미지 불러오기"""
    data_dir = "data"
    photos_dir = f"{data_dir}/photos"
    
    # JSON 데이터 로드
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
    
    # 사진 로드 (data/photos 폴더에서 사번.png 파일 불러오기)
    if os.path.exists(photos_dir):
        for filename in os.listdir(photos_dir):
            if filename.endswith(".png"):
                try:
                    emp_id = filename[:-4]  # 확장자 제거
                    photo_path = os.path.join(photos_dir, filename)
                    photo = Image.open(photo_path)
                    st.session_state.worker_photos[emp_id] = photo
                except Exception as e:
                    print(f"사진 로드 오류 - {filename}: {e}")
    
    # 메타데이터 로드 (파일명 등)
    try:
        if os.path.exists(f"{data_dir}/metadata.json"):
            with open(f"{data_dir}/metadata.json", "r", encoding="utf-8") as f:
                metadata = json.load(f)
                st.session_state.uploaded_worker_file_name = metadata.get("uploaded_worker_file_name", "미등록")
                st.session_state.uploaded_cert_file_name = metadata.get("uploaded_cert_file_name", "미등록")
                st.session_state.uploaded_class_file_name = metadata.get("uploaded_class_file_name", "미등록")
    except (json.JSONDecodeError, FileNotFoundError):
        pass  # 메타데이터 파일이 없으면 기본값 사용

# 엑셀 데이터 정제 함수
def clean_excel_data(df):
    """엑셀 데이터에서 시간정보 제거 및 날짜 포맷팅, 사번을 문자열로 변환"""
    df_clean = df.copy()
    
    # 사번을 문자열로 변환 (형식 통일)
    if '사번' in df_clean.columns:
        df_clean['사번'] = df_clean['사번'].astype(str).str.strip()
    
    # 모든 컬럼의 Timestamp/datetime 타입을 문자열로 강제 변환
    for col in df_clean.columns:
        if col == '사번':  # 사번은 이미 처리됨
            continue
        
        # datetime64, Timestamp 타입 체크
        if pd.api.types.is_datetime64_any_dtype(df_clean[col]):
            df_clean[col] = df_clean[col].dt.strftime('%Y-%m-%d')
        elif df_clean[col].dtype == 'object':
            # object 타입에서 Timestamp 객체가 있는지 확인
            try:
                # 첫 번째 non-null 값을 확인
                first_valid = df_clean[col].dropna().iloc[0] if len(df_clean[col].dropna()) > 0 else None
                if isinstance(first_valid, pd.Timestamp):
                    df_clean[col] = df_clean[col].dt.strftime('%Y-%m-%d')
            except (AttributeError, IndexError):
                pass  # Timestamp가 아닌 경우
    
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
if 'uploaded_worker_file_name' not in st.session_state:
    st.session_state.uploaded_worker_file_name = "미등록"
if 'uploaded_cert_file_name' not in st.session_state:
    st.session_state.uploaded_cert_file_name = "미등록"
if 'uploaded_class_file_name' not in st.session_state:
    st.session_state.uploaded_class_file_name = "미등록"

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
                
                # 파일 정보 표시
                st.info(f"📊 업로드된 파일에 **{len(df_upload)}명**의 데이터가 있습니다.")
                
                # 필수 컬럼 체크
                required_cols = ['사번', '이름', '영어이름', '국적', '부/과', '직/반', '입사일', '근속개월', '직종']
                missing_cols = [col for col in required_cols if col not in df_upload.columns]
                
                if missing_cols:
                    st.warning(f"⚠️ 다음 필수 컬럼이 누락되었습니다: {', '.join(missing_cols)}")
                
                if st.button("시스템에 데이터 반영하기"):
                    uploaded_data = df_upload.to_dict('records')
                    
                    # 중복 체크 및 추가
                    added_count = 0
                    duplicate_count = 0
                    skip_count = 0
                    failed_records = []
                    
                    for idx, new_worker in enumerate(uploaded_data, start=1):
                        # 필수 필드 체크
                        if not new_worker.get('사번') or pd.isna(new_worker.get('사번')):
                            skip_count += 1
                            failed_records.append(f"Row {idx}: 사번이 없습니다")
                            continue
                        
                        # 필수 컬럼 기본값 설정
                        if '부/과' not in new_worker or pd.isna(new_worker.get('부/과')):
                            new_worker['부/과'] = 'N/A'
                        if '직/반' not in new_worker or pd.isna(new_worker.get('직/반')):
                            new_worker['직/반'] = 'N/A'
                        if '입사일' not in new_worker or pd.isna(new_worker.get('입사일')):
                            new_worker['입사일'] = 'N/A'
                        if '근속개월' not in new_worker or pd.isna(new_worker.get('근속개월')):
                            new_worker['근속개월'] = 0
                        if '직종' not in new_worker or pd.isna(new_worker.get('직종')):
                            new_worker['직종'] = 'N/A'
                        
                        # 사번이 이미 존재하는지 확인
                        existing = next((w for w in st.session_state.workers if w.get('사번') == new_worker.get('사번')), None)
                        if existing:
                            duplicate_count += 1
                        else:
                            st.session_state.workers.append(new_worker)
                            added_count += 1
                    
                    save_data()  # 데이터 자동 저장
                    
                    # 결과 표시
                    col_result1, col_result2, col_result3, col_result4 = st.columns(4)
                    col_result1.metric("📥 업로드", f"{len(uploaded_data)}명")
                    col_result2.metric("➕ 신규", f"{added_count}명")
                    col_result3.metric("🔄 중복", f"{duplicate_count}명")
                    col_result4.metric("⏭️ 스킵", f"{skip_count}명")
                    
                    if added_count > 0:
                        st.success(f"✅ {added_count}명의 새로운 데이터가 추가되었습니다!\n현재 총 인원: {len(st.session_state.workers)}명")
                    else:
                        st.info(f"ℹ️ 신규 추가된 데이터가 없습니다. (중복: {duplicate_count}명, 스킵: {skip_count}명)")
                    
                    if failed_records:
                        with st.expander(f"❌ 처리 불가 항목 ({len(failed_records)}개)"):
                            for record in failed_records:
                                st.write(f"  - {record}")
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")

    st.write("---")
    
    if st.session_state.workers:
        df = pd.DataFrame(st.session_state.workers)
        # 사번 열을 문자열로 명시적 지정
        if '사번' in df.columns:
            df['사번'] = df['사번'].astype(str)
        
        # 입사일 열에서 날짜만 추출 (시간 정보 제거)
        if '입사일' in df.columns:
            df['입사일'] = pd.to_datetime(df['입사일'], errors='coerce').dt.date
        
        # 시스템 현황 정보
        st.info(f"✅ 시스템에 저장된 현황: **총 {len(df)}명** | 방글라데시: {len(df[df['국적']=='방글라데시'])}명 | 파키스탄: {len(df[df['국적']=='파키스탄'])}명")
        
        c1, c2, c3 = st.columns(3)
        c1.metric("총 인원", f"{len(df)}명")
        c2.metric("방글라데시", f"{len(df[df['국적']=='방글라데시'])}명")
        c3.metric("파키스탄", f"{len(df[df['국적']=='파키스탄'])}명")
        
        st.write("---")
        st.subheader("📚 한국어 자격증 보유 현황")
        
        # 자격증 보유자 집계 (사번 기준 - 1사원=1건)
        # 유효한 자격증: 사번 + 취득일이 모두 있어야 함
        certified_ids = set()
        for cert in st.session_state.korean_certs:
            # 사번과 취득일이 모두 있을 때만 유효한 자격증으로 인정
            if cert.get('사번') and cert.get('취득일') and str(cert.get('취득일')).strip() not in ['None', '', 'nan']:
                certified_ids.add(str(cert.get('사번')))
        
        total_employees = len(df)
        # 사번 기준으로만 카운트 (같은 사원이 여러 자격증을 가져도 1명)
        certified_employees = df[df['사번'].isin(certified_ids)]
        certified_count = len(certified_employees)
        
        # 전체 현황 메트릭
        col_cert1, col_cert2, col_cert3 = st.columns(3)
        col_cert1.metric("🎓 자격증 보유자", f"{certified_count}명")
        col_cert2.metric("📊 보유율", f"{(certified_count/total_employees*100):.1f}%")
        col_cert3.metric("미보유", f"{total_employees-certified_count}명")
        
        # 부/과별 자격증 현황
        st.write("---")
        col_dept1, col_dept2 = st.columns(2)
        
        with col_dept1:
            st.subheader("부/과별 자격증 보유율")
            
            # 부/과별 - 사번 기준 카운트 (중복 제거)
            dept_total = df.groupby('부/과')['사번'].count()
            dept_certified = df[df['사번'].isin(certified_ids)].groupby('부/과')['사번'].count()
            
            # 누락된 부/과 추가 (자격증 보유자가 없는 부/과)
            all_depts = set(df['부/과'].unique())
            for dept in all_depts:
                if dept not in dept_certified.index:
                    dept_certified[dept] = 0
            
            dept_data = pd.DataFrame({
                '전체': dept_total,
                '자격증': dept_certified.sort_index()
            }).sort_index()
            dept_data['보유율(%)'] = (dept_data['자격증'] / dept_data['전체'] * 100).round(1)
            
            # 부/과별 막대 그래프
            fig_dept, ax_dept = plt.subplots(figsize=(10, 5))
            depts = dept_data.index.tolist()
            certified = dept_data['자격증'].tolist()
            not_certified = (dept_data['전체'] - dept_data['자격증']).tolist()
            
            x = np.arange(len(depts))
            width = 0.6
            
            ax_dept.bar(x, certified, width, label='자격증 보유', color='#2ecc71')
            ax_dept.bar(x, not_certified, width, bottom=certified, label='미보유', color='#e74c3c')
            
            ax_dept.set_xlabel('부/과', fontsize=12, fontweight='bold')
            ax_dept.set_ylabel('인원(명)', fontsize=12, fontweight='bold')
            ax_dept.set_title('부/과별 자격증 보유율', fontsize=14, fontweight='bold')
            ax_dept.set_xticks(x)
            ax_dept.set_xticklabels(depts, rotation=45, ha='right')
            ax_dept.legend()
            ax_dept.grid(axis='y', alpha=0.3)
            
            plt.tight_layout()
            st.pyplot(fig_dept)
            
            # 부/과별 표
            st.dataframe(dept_data[['전체', '자격증', '보유율(%)']], use_container_width=True)
        
        with col_dept2:
            st.subheader("직/반별 자격증 보유율")
            
            # 직/반별 - 사번 기준 카운트 (중복 제거)
            unit_total = df.groupby('직/반')['사번'].count()
            unit_certified = df[df['사번'].isin(certified_ids)].groupby('직/반')['사번'].count()
            
            # 누락된 직/반 추가 (자격증 보유자가 없는 직/반)
            all_units = set(df['직/반'].unique())
            for unit in all_units:
                if unit not in unit_certified.index:
                    unit_certified[unit] = 0
            
            unit_data = pd.DataFrame({
                '전체': unit_total,
                '자격증': unit_certified.sort_index()
            }).sort_index()
            unit_data['보유율(%)'] = (unit_data['자격증'] / unit_data['전체'] * 100).round(1)
            
            # 직/반별 막대 그래프
            fig_unit, ax_unit = plt.subplots(figsize=(10, 5))
            units = unit_data.index.tolist()
            certified_unit = unit_data['자격증'].tolist()
            not_certified_unit = (unit_data['전체'] - unit_data['자격증']).tolist()
            
            x_unit = np.arange(len(units))
            
            ax_unit.bar(x_unit, certified_unit, width, label='자격증 보유', color='#3498db')
            ax_unit.bar(x_unit, not_certified_unit, width, bottom=certified_unit, label='미보유', color='#95a5a6')
            
            ax_unit.set_xlabel('직/반', fontsize=12, fontweight='bold')
            ax_unit.set_ylabel('인원(명)', fontsize=12, fontweight='bold')
            ax_unit.set_title('직/반별 자격증 보유율', fontsize=14, fontweight='bold')
            ax_unit.set_xticks(x_unit)
            ax_unit.set_xticklabels(units, rotation=45, ha='right')
            ax_unit.legend()
            ax_unit.grid(axis='y', alpha=0.3)
            
            plt.tight_layout()
            st.pyplot(fig_unit)
            
            # 직/반별 표
            st.dataframe(unit_data[['전체', '자격증', '보유율(%)']], use_container_width=True)
        
        # 국적별 자격증 현황
        st.write("---")
        st.subheader("국적별 자격증 보유율")
        
        # 국적별 - 사번 기준 카운트 (중복 제거)
        nationality_total = df.groupby('국적')['사번'].count()
        nationality_certified = df[df['사번'].isin(certified_ids)].groupby('국적')['사번'].count()
        
        # 누락된 국적 추가 (자격증 보유자가 없는 국적)
        all_nationalities = set(df['국적'].unique())
        for nat in all_nationalities:
            if nat not in nationality_certified.index:
                nationality_certified[nat] = 0
        
        nationality_data = pd.DataFrame({
            '전체': nationality_total,
            '자격증': nationality_certified.sort_index()
        }).sort_index()
        nationality_data['보유율(%)'] = (nationality_data['자격증'] / nationality_data['전체'] * 100).round(1)
        
        col_nat1, col_nat2 = st.columns(2)
        
        with col_nat1:
            # 국적별 파이 차트
            fig_nat, ax_nat = plt.subplots(figsize=(8, 6))
            colors = plt.cm.Set3(np.linspace(0, 1, len(nationality_data)))
            wedges, texts, autotexts = ax_nat.pie(
                nationality_data['자격증'],
                labels=nationality_data.index.tolist(),
                autopct='%1.1f%%',
                colors=colors,
                startangle=90
            )
            ax_nat.set_title('국적별 자격증 보유 분포', fontsize=14, fontweight='bold')
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontweight('bold')
            plt.tight_layout()
            st.pyplot(fig_nat)
        
        with col_nat2:
            # 국적별 막대 그래프 (보유율 비교)
            fig_nat_bar, ax_nat_bar = plt.subplots(figsize=(8, 6))
            nationalities = nationality_data.index.tolist()
            rates = nationality_data['보유율(%)'].tolist()
            colors_bar = ['#2ecc71' if r >= 50 else '#f39c12' if r >= 25 else '#e74c3c' for r in rates]
            
            bars = ax_nat_bar.barh(nationalities, rates, color=colors_bar)
            ax_nat_bar.set_xlabel('자격증 보유율(%)', fontsize=12, fontweight='bold')
            ax_nat_bar.set_title('국적별 자격증 보유율 비교', fontsize=14, fontweight='bold')
            ax_nat_bar.set_xlim(0, 100)
            
            # 비율 값 표시
            for i, (bar, rate) in enumerate(zip(bars, rates)):
                ax_nat_bar.text(rate + 2, i, f'{rate:.1f}%', va='center', fontweight='bold')
            
            ax_nat_bar.grid(axis='x', alpha=0.3)
            plt.tight_layout()
            st.pyplot(fig_nat_bar)
            
            # 국적별 상세 표
            st.dataframe(nationality_data[['전체', '자격증', '보유율(%)']], use_container_width=True)
        
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
#        st.info("등록된 인력이 없습니다. 엑셀 업로드나 개별 등록을 이용해 주세요.")

# ---------------------------------------------------------
# 메뉴 2: 👤 인력 정보 관리
# ---------------------------------------------------------
elif menu == "👤 인력 정보 관리":
    st.header("👤 인력 정보 관리")
    
    # 탭 구성
    tab1, tab2, tab3 = st.tabs(["👥 인력 정보 조회", "📥 데이터 및 사진 업로드", "📚 한국어 교육/자격 조회"])
    
    with tab1:
        st.subheader("👥 등록된 인력 정보 조회")
        
        if st.session_state.workers:
            # 필터 폼 - 부서/반 선택
            with st.form("filter_form"):
                st.write("**필터 조건 선택**")
                
                filter_col1, filter_col2 = st.columns(2)
                
                with filter_col1:
                    dept_list = list(dept_structure.keys())
                    selected_filter_dept = st.selectbox(
                        "부/과 선택",
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
                        "직/반 선택",
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
                filtered_workers = [w for w in filtered_workers if w.get('부/과') == selected_filter_dept]
            
            # 반 필터
            if selected_filter_unit != "전체":
                filtered_workers = [w for w in filtered_workers if w.get('직/반') == selected_filter_unit]
            
            # 검색 필터
            if search_term:
                filtered_workers = [w for w in filtered_workers 
                                   if search_term in str(w.get('사번', '')) or 
                                      search_term in w.get('이름', '')]
            
            # 결과 표시
            result_col1, result_col2, result_col3 = st.columns([2, 1.5, 1.5])
            with result_col1:
                st.write(f"**조회 결과: {len(filtered_workers)}명** (전체: {len(st.session_state.workers)}명)")
            with result_col2:
                st.write(f"**📋 데이터 제목**: {st.session_state.uploaded_worker_file_name}")
            with result_col3:
                st.write(f"**📅 기준일시**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
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
                                st.write(f"**📛 이름**: {worker.get('이름', 'N/A')}")
                                st.write(f"**🌐 영어이름**: {worker.get('영어이름', 'N/A')}")
                                st.write(f"**🌍 국적**: {worker.get('국적', 'N/A')}")
                                st.write(f"**📅 입사일**: {worker.get('입사일', 'N/A')}")                                
                            
                            with col_r:
                                st.write(f"**🏢 부/과**: {worker.get('부/과', 'N/A')}")
                                st.write(f"**👥 직/반**: {worker.get('직/반', 'N/A')}")
                                st.write(f"**💼 직종**: {worker.get('직종', 'N/A')}")
                                st.write(f"**⏱️ 근속개월**: {worker.get('근속개월', 'N/A')}개월")
                            
                            # 최신 한국어 자격/교육 정보 조회
                            emp_id = worker.get('사번')
                            
                            # 최신 한국어 자격 찾기
                            latest_cert = None
                            latest_cert_date = None
                            for cert in st.session_state.korean_certs:
                                if cert.get('사번') == emp_id:
                                    cert_date = str(cert.get('취득일', '')).strip()
                                    # 취득일이 유효한지 확인 (None, 빈값, 'nan' 제외)
                                    if cert_date and cert_date not in ['None', 'nan', '']:
                                        if latest_cert_date is None or cert_date > latest_cert_date:
                                            latest_cert = cert
                                            latest_cert_date = cert_date
                            
                            # 최신 한국어 교육 이력 찾기
                            latest_class = None
                            for cls in st.session_state.korean_classes:
                                if cls.get('사번') == emp_id:
                                    latest_class = cls
                            
                            st.write("---")
                            
                            info_col1, info_col2 = st.columns(2)
                            with info_col1:
                                if latest_cert:
                                    st.write(f"**🎓 최신 한국어 자격**: {latest_cert.get('시험종류')} {latest_cert.get('자격')}")
                                    st.write(f"&nbsp;&nbsp;&nbsp;&nbsp;(취득일: {latest_cert_date})")
                                else:
                                    st.write("**🎓 최신 한국어 자격**: -")
                            
                            with info_col2:
                                if latest_class:
                                    st.write(f"**🏫 최신 한국어 교육**: {latest_class.get('수업명')}")
                                    st.write(f"&nbsp;&nbsp;&nbsp;&nbsp;({latest_class.get('시기')})")
                                else:
                                    st.write("**🏫 최신 한국어 교육**: -")
                            
                            with st.expander("📋 상세 정보"):
                                st.write(f"**🏠 숙소**: {worker.get('숙소구분', 'N/A')}")
                                st.write(f"**👨‍👩‍👧 가족동반**: {worker.get('가족동반', 'N/A')}")
                                st.write(f"**주소**: {worker.get('주소', 'N/A')}")
                                st.write(f"**계약 구분**: {worker.get('계약', 'N/A')}")
                                st.write(f"**비고**: {worker.get('비고', 'N/A')}")
                        
                        st.divider()
            else:
                st.info("선택한 조건에 해당하는 인력이 없습니다.")
        else:
            st.info("등록된 인력이 없습니다. 먼저 엑셀 파일을 업로드해주세요.")
   
    
    with tab2:
        st.subheader("📥 엑셀 데이터 일괄 업로드")
        st.write("엑셀 파일(사번, 이름, 영어이름, ... 등)을 업로드합니다.")
        
        uploaded_excel = st.file_uploader("엑셀 파일(.xlsx)을 업로드하세요", type=["xlsx"], key="upload_excel_menu2")
        if uploaded_excel:
            st.session_state.uploaded_worker_file_name = uploaded_excel.name  # 파일명 저장
            try:
                df_excel = pd.read_excel(uploaded_excel)
                df_excel = clean_excel_data(df_excel)  # 날짜 데이터 정제
                df_excel.index = range(1, len(df_excel) + 1)  # 인덱스 1부터 시작
                
                # 필수 컬럼 체크
                required_cols = ['사번', '이름', '영어이름', '국적', '부/과', '직/반', '입사일', '근속개월', '직종']
                missing_cols = [col for col in required_cols if col not in df_excel.columns]
                
                if missing_cols:
                    st.warning(f"⚠️ 다음 필수 컬럼이 누락되었습니다: {', '.join(missing_cols)}")
                
                st.dataframe(df_excel, width='stretch')
                
                if st.button("✅ 데이터 저장하기", key="save_excel_menu2"):
                    excel_data = df_excel.to_dict('records')
                    
                    # 중복 체크 및 추가
                    added_count = 0
                    updated_count = 0
                    duplicate_count = 0
                    
                    for new_worker in excel_data:
                        # 필수 컬럼 기본값 설정
                        if '부/과' not in new_worker or pd.isna(new_worker.get('부/과')):
                            new_worker['부/과'] = 'N/A'
                        if '직/반' not in new_worker or pd.isna(new_worker.get('직/반')):
                            new_worker['직/반'] = 'N/A'
                        if '입사일' not in new_worker or pd.isna(new_worker.get('입사일')):
                            new_worker['입사일'] = 'N/A'
                        if '근속개월' not in new_worker or pd.isna(new_worker.get('근속개월')):
                            new_worker['근속개월'] = 0
                        if '직종' not in new_worker or pd.isna(new_worker.get('직종')):
                            new_worker['직종'] = 'N/A'
                        
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

    with tab3:
        st.subheader("📚 한국어 교육/자격 조회")
        
        if st.session_state.korean_certs or st.session_state.korean_classes:
            with st.form("korean_filter_form"):
                st.write("**필터 조건 선택**")
                
                # 부서/반 선택
                filter_col1, filter_col2 = st.columns(2)
                with filter_col1:
                    dept_list = list(dept_structure.keys())
                    multi_dept = st.multiselect("📍 부/과 (복수선택 가능)", ["전체"] + dept_list, default=["전체"], key="korean_dept")
                
                with filter_col2:
                    available_units = []
                    if "전체" not in multi_dept:
                        for dept in multi_dept:
                            available_units.extend(dept_structure[dept]["반"])
                        available_units = sorted(list(set(available_units)))
                    else:
                        for dept in dept_list:
                            available_units.extend(dept_structure[dept]["반"])
                        available_units = sorted(list(set(available_units)))
                    
                    multi_unit = st.multiselect("👥 직/반 (복수선택 가능)", ["전체"] + available_units, default=["전체"], key="korean_unit")
                
                # 검색 창
                filter_col3, filter_col4 = st.columns(2)
                with filter_col3:
                    search_name = st.text_input("👤 이름 검색", placeholder="이름 입력...", key="korean_name")
                
                with filter_col4:
                    search_id = st.text_input("🆔 사번 검색", placeholder="사번 입력...", key="korean_id")
                
                # 시험종류/단계 필터
                filter_col5, filter_col6 = st.columns(2)
                with filter_col5:
                    multi_exam = st.multiselect("📋 시험 종류", ["사통", "사전평가", "TOPIK"], key="korean_exam")
                
                with filter_col6:
                    multi_level = st.multiselect("📊 단계/급수", [f"{i}단계" for i in range(1, 7)] + [f"{i}급" for i in range(1, 7)], key="korean_level")
                
                # 취득연도 필터
                # 현재 데이터에서 연도 추출
                available_years = set()
                for cert in st.session_state.korean_certs:
                    cert_date = str(cert.get('취득일', '')).strip()
                    if cert_date and cert_date != 'nan' and len(cert_date) >= 4:
                        try:
                            year = cert_date[:4]
                            if year.isdigit():
                                available_years.add(year)
                        except:
                            pass
                
                available_years = sorted(list(available_years), reverse=True)
                selected_year = st.selectbox("📅 취득연도", ["전체"] + available_years, key="korean_year")
                
                submitted_korean = st.form_submit_button("🔍 조회", use_container_width=True)
            
            st.write("---")
            
            # 필터링 로직
            filtered_certs = st.session_state.korean_certs.copy()
            
            # 워커 정보 매핑
            worker_map = {w.get('사번'): w for w in st.session_state.workers}
            
            # 부/과 필터 (자격증 데이터에서 직접)
            if "전체" not in multi_dept:
                filtered_certs = [c for c in filtered_certs if c.get('부/과') in multi_dept]
            
            # 직/반 필터 (자격증 데이터에서 직접)
            if "전체" not in multi_unit:
                filtered_certs = [c for c in filtered_certs if c.get('직/반') in multi_unit]
            
            # 이름 필터
            if search_name:
                filtered_certs = [c for c in filtered_certs if search_name in worker_map.get(c.get('사번'), {}).get('이름', '')]
            
            # 사번 필터
            if search_id:
                filtered_certs = [c for c in filtered_certs if search_id in str(c.get('사번', ''))]
            
            # 시험종류 필터
            if multi_exam:
                filtered_certs = [c for c in filtered_certs if c.get('시험') in multi_exam]
            
            # 단계/급수 필터
            if multi_level:
                filtered_certs = [c for c in filtered_certs if c.get('단계/급수') in multi_level]
            
            # 취득연도 필터
            if selected_year != "전체":
                filtered_certs = [c for c in filtered_certs if str(c.get('취득일', '')).startswith(selected_year)]
            
            # 결과 표시
            result_col1, result_col2, result_col3 = st.columns([2, 1.5, 1.5])
            with result_col1:
                st.write(f"**📊 조회 결과: {len(filtered_certs)}건** (전체: {len(st.session_state.korean_certs)}건)")
            with result_col2:
                st.write(f"**📋 데이터 제목**: {st.session_state.uploaded_cert_file_name}")
            with result_col3:
                st.write(f"**📅 기준일시**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            st.write("---")
            
            if len(filtered_certs) > 0:
                # 테이블로 표시
                cert_display = []
                for cert in filtered_certs:
                    # 단계/급수에서 첫 번째 숫자만 추출
                    level = cert.get('단계/급수', 'N/A')
                    # float(NaN) 또는 None 값 처리
                    if pd.isna(level) or level == 'N/A':
                        level = ''
                    else:
                        # 문자열로 변환 후 첫 번째 숫자만 추출
                        level_str = str(level).strip()
                        if level_str and level_str != 'N/A':
                            level_num = next((c for c in level_str if c.isdigit()), '')
                            level = level_num if level_num else level_str
                        else:
                            level = ''
                    
                    # 시험이 N/A면 공백으로
                    exam = cert.get('시험', 'N/A')
                    exam = '' if exam == 'N/A' else str(exam).strip()
                    
                    cert_display.append({
                        "사번": cert.get('사번'),
                        "이름": cert.get('이름', ''),
                        "영어이름": cert.get('영어이름', ''),
                        "국적": cert.get('국적', ''),
                        "부/과": cert.get('부/과', ''),
                        "직/반": cert.get('직/반', ''),
                        "직종": cert.get('직종', ''),
                        "시험": exam,
                        "단계/급수": level,
                        "취득일": cert.get('취득일')
                    })
                
                cert_df = pd.DataFrame(cert_display)
                st.dataframe(cert_df, use_container_width=True)
                
                # 다운로드 버튼
                csv_data = cert_df.to_csv(index=False).encode('utf-8-sig')
                excel_data = to_excel(cert_df)
                
                col_dlk1, col_dlk2 = st.columns(2)
                with col_dlk1:
                    st.download_button(
                        "📥 조회결과 다운로드(CSV)",
                        data=csv_data,
                        file_name=f"korean_certs_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv"
                    )
                with col_dlk2:
                    st.download_button(
                        "📥 조회결과 다운로드(Excel)",
                        data=excel_data,
                        file_name=f"korean_certs_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.info("선택한 조건에 해당하는 자격정보가 없습니다.")
        else:
            st.info("등록된 한국어 자격 정보가 없습니다. 메뉴 3에서 데이터를 추가해주세요.")
elif menu == "📚 한국어 교육/자격 관리":
    st.header("📚 한국어 역량")
    
    t1, t2, t3 = st.tabs(["📥 자격증 데이터 업로드", "📥 교육 데이터 업로드", "📊 데이터 다운로드"])
    
    with t1:
        st.subheader("📥 자격증 취득 정보 업로드")
        st.write("엑셀 파일 형식: 사번, 이름, 영어이름, 국적, 부/과, 직/반, 직종, 시험, 단계/급수, 취득일")
        
        uploaded_cert_file = st.file_uploader("자격증 엑셀 파일을 업로드하세요", type=["xlsx"], key="upload_cert_excel")
        if uploaded_cert_file:
            st.session_state.uploaded_cert_file_name = uploaded_cert_file.name  # 파일명 저장
            try:
                df_cert = pd.read_excel(uploaded_cert_file)
                df_cert = clean_excel_data(df_cert)  # 날짜 데이터 정제
                
                # 필요한 컬럼 확인
                required_cols = ['사번', '이름', '영어이름', '국적', '부/과', '직/반', '직종', '시험', '단계/급수', '취득일']
                missing_cols = [col for col in required_cols if col not in df_cert.columns]
                
                if missing_cols:
                    st.warning(f"⚠️ 다음 필수 컬럼이 누락되었습니다: {', '.join(missing_cols)}")
                
                df_cert.index = range(1, len(df_cert) + 1)  # 인덱스 1부터 시작
                st.dataframe(df_cert, use_container_width=True)
                
                # 두 가지 저장 옵션
                col_btn1, col_btn2 = st.columns(2)
                
                with col_btn1:
                    if st.button("🔄 기존 데이터 삭제 후 교체하기", key="replace_cert_excel", help="기존 자격증 데이터를 모두 삭제하고 새 데이터로 교체합니다"):
                        cert_data = df_cert.to_dict('records')
                        
                        # 필수 컬럼 기본값 설정
                        for new_cert in cert_data:
                            if '이름' not in new_cert or pd.isna(new_cert.get('이름')):
                                new_cert['이름'] = 'N/A'
                            if '영어이름' not in new_cert or pd.isna(new_cert.get('영어이름')):
                                new_cert['영어이름'] = 'N/A'
                            if '국적' not in new_cert or pd.isna(new_cert.get('국적')):
                                new_cert['국적'] = 'N/A'
                            if '부/과' not in new_cert or pd.isna(new_cert.get('부/과')):
                                new_cert['부/과'] = 'N/A'
                            if '직/반' not in new_cert or pd.isna(new_cert.get('직/반')):
                                new_cert['직/반'] = 'N/A'
                            if '직종' not in new_cert or pd.isna(new_cert.get('직종')):
                                new_cert['직종'] = 'N/A'
                            if '시험' not in new_cert or pd.isna(new_cert.get('시험')):
                                new_cert['시험'] = 'N/A'
                            if '단계/급수' not in new_cert or pd.isna(new_cert.get('단계/급수')):
                                new_cert['단계/급수'] = 'N/A'
                        
                        # 기존 데이터 모두 삭제
                        st.session_state.korean_certs = cert_data
                        save_data()
                        st.success(f"✅ 저장 완료: 기존 데이터 삭제 후 {len(cert_data)}건 새로 저장되었습니다")
                        st.rerun()
                
                with col_btn2:
                    if st.button("✅ 기존 데이터에 추가하기", key="save_cert_excel", help="기존 자격증 데이터에 새 데이터를 추가합니다"):
                        cert_data = df_cert.to_dict('records')
                        
                        added_count = 0
                        updated_count = 0
                        
                        for new_cert in cert_data:
                            # 필수 컬럼 기본값 설정
                            if '이름' not in new_cert or pd.isna(new_cert.get('이름')):
                                new_cert['이름'] = 'N/A'
                            if '영어이름' not in new_cert or pd.isna(new_cert.get('영어이름')):
                                new_cert['영어이름'] = 'N/A'
                            if '국적' not in new_cert or pd.isna(new_cert.get('국적')):
                                new_cert['국적'] = 'N/A'
                            if '부/과' not in new_cert or pd.isna(new_cert.get('부/과')):
                                new_cert['부/과'] = 'N/A'
                            if '직/반' not in new_cert or pd.isna(new_cert.get('직/반')):
                                new_cert['직/반'] = 'N/A'
                            if '직종' not in new_cert or pd.isna(new_cert.get('직종')):
                                new_cert['직종'] = 'N/A'
                            if '시험' not in new_cert or pd.isna(new_cert.get('시험')):
                                new_cert['시험'] = 'N/A'
                            if '단계/급수' not in new_cert or pd.isna(new_cert.get('단계/급수')):
                                new_cert['단계/급수'] = 'N/A'
                            
                            # 중복 확인 (사번, 시험, 단계/급수, 취득일 기준)
                            existing = next((c for c in st.session_state.korean_certs 
                                           if c.get('사번') == new_cert.get('사번') and 
                                              c.get('시험') == new_cert.get('시험') and
                                              c.get('단계/급수') == new_cert.get('단계/급수') and
                                              c.get('취득일') == new_cert.get('취득일')), None)
                            
                            if existing:
                                # 기존 데이터 업데이트
                                for i, cert in enumerate(st.session_state.korean_certs):
                                    if (cert.get('사번') == new_cert.get('사번') and 
                                        cert.get('시험') == new_cert.get('시험') and
                                        cert.get('단계/급수') == new_cert.get('단계/급수') and
                                        cert.get('취득일') == new_cert.get('취득일')):
                                        st.session_state.korean_certs[i] = new_cert
                                        updated_count += 1
                                        break
                            else:
                                # 신규 등록
                                st.session_state.korean_certs.append(new_cert)
                                added_count += 1
                        
                        save_data()
                        st.success(f"✅ 저장 완료: 신규 {added_count}건, 업데이트 {updated_count}건")
                        st.rerun()
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")
    
    with t2:
        st.subheader("📥 교육 수업 참여 정보 업로드")
        st.write("엑셀 파일 형식: 사번, 수업명(사내 사통교육/사외 사통교육/통역사 주관 수업), 시기(2026년 상반기/2026년 하반기)")
        
        uploaded_class_file = st.file_uploader("교육 엑셀 파일을 업로드하세요", type=["xlsx"], key="upload_class_excel")
        if uploaded_class_file:
            st.session_state.uploaded_class_file_name = uploaded_class_file.name  # 파일명 저장
            try:
                df_class = pd.read_excel(uploaded_class_file)
                # 필요한 컬럼 확인
                required_cols = ['사번', '수업명', '시기']
                if all(col in df_class.columns for col in required_cols):
                    df_class.index = range(1, len(df_class) + 1)  # 인덱스 1부터 시작
                    st.dataframe(df_class, use_container_width=True)
                    
                    if st.button("✅ 교육 데이터 저장하기", key="save_class_excel"):
                        class_data = df_class.to_dict('records')
                        
                        added_count = 0
                        updated_count = 0
                        
                        for new_class in class_data:
                            # 중복 확인
                            existing = next((c for c in st.session_state.korean_classes 
                                           if c.get('사번') == new_class.get('사번') and 
                                              c.get('수업명') == new_class.get('수업명') and
                                              c.get('시기') == new_class.get('시기')), None)
                            
                            if existing:
                                # 기존 데이터 업데이트
                                for i, cls in enumerate(st.session_state.korean_classes):
                                    if (cls.get('사번') == new_class.get('사번') and 
                                        cls.get('수업명') == new_class.get('수업명') and
                                        cls.get('시기') == new_class.get('시기')):
                                        st.session_state.korean_classes[i] = new_class
                                        updated_count += 1
                                        break
                            else:
                                # 신규 등록
                                st.session_state.korean_classes.append(new_class)
                                added_count += 1
                        
                        save_data()
                        st.success(f"✅ 저장 완료: 신규 {added_count}건, 업데이트 {updated_count}건")
                else:
                    st.error(f"❌ 필수 컬럼이 없습니다. 필요한 컬럼: {', '.join(required_cols)}")
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")
    
    with t3:
        st.subheader("📊 데이터 다운로드")
        
        col_d1, col_d2 = st.columns(2)
        
        with col_d1:
            if st.session_state.korean_certs:
                st.write("**🎓 자격증 취득 정보**")
                st.caption(f"📋 {st.session_state.uploaded_cert_file_name} | 📅 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                
                # 표시용 데이터 처리
                cert_display = []
                for cert in st.session_state.korean_certs:
                    # 단계/급수에서 첫 번째 숫자만 추출
                    level = cert.get('단계/급수', 'N/A')
                    # float(NaN) 또는 None 값 처리
                    if pd.isna(level) or level == 'N/A':
                        level = ''
                    else:
                        # 문자열로 변환 후 첫 번째 숫자만 추출
                        level_str = str(level).strip()
                        if level_str and level_str != 'N/A':
                            level_num = next((c for c in level_str if c.isdigit()), '')
                            level = level_num if level_num else level_str
                        else:
                            level = ''
                    
                    # 시험이 N/A면 공백으로
                    exam = cert.get('시험', 'N/A')
                    exam = '' if exam == 'N/A' else str(exam).strip()
                    
                    cert_display.append({
                        "사번": cert.get('사번'),
                        "이름": cert.get('이름', ''),
                        "영어이름": cert.get('영어이름', ''),
                        "국적": cert.get('국적', ''),
                        "부/과": cert.get('부/과', ''),
                        "직/반": cert.get('직/반', ''),
                        "직종": cert.get('직종', ''),
                        "시험": exam,
                        "단계/급수": level,
                        "취득일": cert.get('취득일')
                    })
                
                cert_df_display = pd.DataFrame(cert_display)
                st.dataframe(cert_df_display, use_container_width=True)
                
                # 다운로드용 원본 데이터
                cert_df = pd.DataFrame(st.session_state.korean_certs)
                cert_csv = cert_df.to_csv(index=False).encode('utf-8-sig')
                cert_excel = to_excel(cert_df)
                
                st.download_button(
                    "📥 자격증 데이터 다운로드(CSV)",
                    data=cert_csv,
                    file_name=f"korean_certs_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    key="dl_cert_csv"
                )
                st.download_button(
                    "📥 자격증 데이터 다운로드(Excel)",
                    data=cert_excel,
                    file_name=f"korean_certs_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_cert_excel"
                )
            else:
                st.info("등록된 자격증 데이터가 없습니다.")
        
        with col_d2:
            if st.session_state.korean_classes:
                st.write("**🏫 교육 수업 참여 정보**")
                st.caption(f"📋 {st.session_state.uploaded_class_file_name} | 📅 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                class_df = pd.DataFrame(st.session_state.korean_classes)
                st.dataframe(class_df, use_container_width=True)
                
                class_csv = class_df.to_csv(index=False).encode('utf-8-sig')
                class_excel = to_excel(class_df)
                
                st.download_button(
                    "📥 교육 데이터 다운로드(CSV)",
                    data=class_csv,
                    file_name=f"korean_classes_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    key="dl_class_csv"
                )
                st.download_button(
                    "📥 교육 데이터 다운로드(Excel)",
                    data=class_excel,
                    file_name=f"korean_classes_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_class_excel"
                )
            else:
                st.info("등록된 교육 데이터가 없습니다.")