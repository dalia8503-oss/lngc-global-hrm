import streamlit as st          # 웹 앱 UI 프레임워크 (브라우저에서 실행되는 화면 구성)
import streamlit.components.v1 as components  # 사용자 정의 HTML/JS 삽입
import pandas as pd              # 표(DataFrame) 형태의 데이터 처리 라이브러리
import numpy as np               # 수치 계산 및 배열 처리 라이브러리
import re                        # 문자열 패턴(숫자 토큰) 추출
#import plotly.express as px     # (사용 안 함) 인터랙티브 차트 라이브러리
from datetime import datetime, timedelta  # 날짜/시간 처리 (현재 시각, 날짜 차이 계산 등)
import io                        # 메모리 내 파일 스트림 처리 (엑셀 다운로드 시 사용)
from PIL import Image            # 이미지 파일 열기, 리사이징, 저장 (사진 처리)
import json                      # JSON 파일 읽기/쓰기 (데이터 영구 저장)
import os                        # 폴더·파일 존재 확인 및 생성/삭제
import webbrowser                # 로컬 기본 브라우저 열기
import matplotlib.pyplot as plt  # 차트(막대, 파이 등) 그리기
import matplotlib as mpl         # Matplotlib 전역 설정 (폰트 등)
try:
    import win32com.client as win32  # DRM 엑셀 파일을 실제 Excel 앱으로 읽기
except ImportError:
    win32 = None

# ──────────────────────────────────────────────────────────
# Matplotlib 한글 폰트 설정
# ── 차트 레이블에 한글이 깨지지 않도록 폰트를 미리 지정
# ──────────────────────────────────────────────────────────
mpl.rcParams['font.family'] = 'sans-serif'  # 기본 폰트 계열을 sans-serif로 설정
try:
    # Windows 환경에 설치된 한글 폰트 목록 지정 (앞에 있는 것 우선 사용)
    mpl.rcParams['font.sans-serif'] = ['Malgun Gothic', 'DejaVu Sans']
except:
    pass  # 폰트 설정에 실패해도 앱이 중단되지 않도록 예외 무시
mpl.rcParams['axes.unicode_minus'] = False  # 음수 기호(-)가 깨지는 문제 방지

# ──────────────────────────────────────────────────────────
# 1. 페이지 설정
# ── 브라우저 탭 제목을 설정하고, 화면을 wide 레이아웃으로 지정
# ──────────────────────────────────────────────────────────
st.set_page_config(page_title="LNG선공사팀 글로벌인력관리", layout="wide")

# ──────────────────────────────────────────────────────────
# 사진 아래 사번 표시 글자 크기 설정
# ── 숫자만 바꾸면 사진 밑 "사번" 텍스트 크기를 바로 조정할 수 있음
# ──────────────────────────────────────────────────────────
PHOTO_EMP_ID_FONT_SIZE_PX = 12

# ──────────────────────────────────────────────────────────
# 인력 정보 영역 사번 표시 글자 크기 설정
# ── 이름/영어이름보다 "약간만" 크게 보이도록 기본값 19px 사용
# ──────────────────────────────────────────────────────────
INFO_EMP_ID_FONT_SIZE_PX = 19

# ──────────────────────────────────────────────────────────
# 전역 배경색 설정
# ── 화면 전체 배경을 아주 연한 핑크색으로 지정
# ──────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
    .stApp {
        background-color: #fff7fb;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ──────────────────────────────────────────────────────────
# 스크롤 위치 유지 설정
# ── 브라우저 새로고침(F5) 후에도 마지막 스크롤 위치를 복원
# ──────────────────────────────────────────────────────────
components.html(
        """
        <script>
            const parentWin = window.parent;
            const storageKey = `global-scroll-pos:${parentWin.location.pathname}`;

            if (!parentWin.__scrollPositionKeeperInitialized) {
                parentWin.__scrollPositionKeeperInitialized = true;

                const saveScroll = () => {
                    try {
                        parentWin.localStorage.setItem(storageKey, String(parentWin.scrollY || 0));
                    } catch (e) {}
                };

                parentWin.addEventListener('scroll', saveScroll, { passive: true });
                parentWin.addEventListener('beforeunload', saveScroll);
            }

            const restoreScroll = () => {
                try {
                    const saved = parentWin.localStorage.getItem(storageKey);
                    if (saved !== null) {
                        parentWin.scrollTo(0, parseInt(saved, 10) || 0);
                    }
                } catch (e) {}
            };

            setTimeout(restoreScroll, 50);
            setTimeout(restoreScroll, 300);
        </script>
        """,
        height=0,
        width=0,
)

# ──────────────────────────────────────────────────────────
# [유틸리티 함수] DataFrame → 엑셀 파일 변환
# ── 다운로드 버튼에서 사용할 엑셀 바이너리 데이터를 반환
# ──────────────────────────────────────────────────────────
def to_excel(df):
    output = io.BytesIO()  # 메모리 내 임시 바이트 버퍼 생성 (실제 파일 대신 메모리 사용)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:  # openpyxl 엔진으로 엑셀 작성기 생성
        df.to_excel(writer, index=False, sheet_name="data")    # 인덱스 제외하고 'data' 시트에 저장
    return output.getvalue()  # 작성된 엑셀 파일의 바이트 데이터를 반환

# ──────────────────────────────────────────────────────────
# [유틸리티 함수] 업로드된 Excel/CSV 파일을 DataFrame으로 읽기
# ── CSV는 UTF-8과 CP949를 순차 시도해서 한글 인코딩 차이를 흡수
# ──────────────────────────────────────────────────────────
def read_uploaded_table(uploaded_file):
    if uploaded_file.name.lower().endswith('.csv'):  # CSV 파일이면
        try:
            return pd.read_csv(uploaded_file, encoding='utf-8-sig')  # UTF-8 CSV 우선 시도
        except UnicodeDecodeError:
            uploaded_file.seek(0)  # 재시도를 위해 파일 포인터를 처음으로 이동
            return pd.read_csv(uploaded_file, encoding='cp949')      # 한글 Windows CSV 인코딩 대응
    return pd.read_excel(uploaded_file)  # CSV가 아니면 Excel 파일로 읽기


def load_data_from_excel(excel_file=None):
    """보안(DRM)이 걸린 엑셀 파일을 실제 엑셀 앱을 통해 읽어옵니다."""
    target_file = excel_file if excel_file is not None else globals().get("EXCEL_FILE")
    if not target_file:
        print("EXCEL_FILE이 정의되지 않았습니다.")
        return []

    abs_path = os.path.abspath(target_file)

    if not os.path.exists(abs_path):
        print(f"파일 없음: {abs_path}")
        return []

    if win32 is None:
        print("pywin32(win32com)가 설치되지 않았습니다. `pip install pywin32` 후 다시 시도하세요.")
        return []

    excel = None
    wb = None

    try:
        # 1. 실제 엑셀 프로그램 실행
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False  # 엑셀 창은 띄우지 않음

        # 2. 파일 열기 (보안 프로그램이 설치된 PC라면 여기서 자동으로 풀려서 열림)
        wb = excel.Workbooks.Open(abs_path)
        sheet = wb.ActiveSheet

        # 3. 엑셀의 데이터 영역을 읽어서 리스트로 변환
        used_range_value = sheet.UsedRange.Value
        if not used_range_value:
            return []

        data_rows = list(used_range_value)
        if not data_rows:
            return []

        # 4. pandas 데이터프레임으로 변환
        df = pd.DataFrame(data_rows)

        # 첫 번째 행을 컬럼명으로 설정
        df.columns = df.iloc[0]
        df = df[1:]

        # 5. 기존과 동일한 데이터 재구조화 로직
        df.columns = [str(col).strip() for col in df.columns]

        id_cols = ['과', '항목']
        date_cols = [col for col in df.columns if col not in id_cols]

        df_melted = df.melt(
            id_vars=id_cols,
            value_vars=date_cols,
            var_name='raw_date',
            value_name='val'
        )

        return df_melted

    finally:
        if wb:
            wb.Close(False)
        if excel:
            excel.Quit()

# ──────────────────────────────────────────────────────────
# [유틸리티 함수] 시험 종류별 단계/급수 표시 문자열 정규화
# ── 사전평가는 N급, 사통은 N단계로 일관되게 표시
# ──────────────────────────────────────────────────────────
def format_cert_level(exam, level):
    exam_text = '' if pd.isna(exam) else str(exam).strip()
    if pd.isna(level):
        return ''

    level_text = str(level).strip()
    if level_text in ['', 'N/A', 'nan', 'None']:
        return ''

    # 숫자형 값(예: 1, 1.0, "1.0")은 보기 좋게 정규화 (1.0 -> 1)
    number_match = re.search(r"\d+(?:\.\d+)?", level_text)
    if number_match:
        number_token = number_match.group(0)
        if '.' in number_token:
            number_value = float(number_token)
            base_value = str(int(number_value)) if number_value.is_integer() else str(number_value).rstrip('0').rstrip('.')
        else:
            base_value = number_token
    else:
        base_value = level_text

    if exam_text == '사전평가':
        return f"{base_value}급"
    if exam_text == '사통':
        return f"{base_value}단계"
    if exam_text == 'TOPIK':
        return f"{base_value}급"
    return base_value


def normalize_emp_id(emp_id):
    """사번 정규화: 공백 제거, .0 제거"""
    emp_id_text = str(emp_id).strip()
    if emp_id_text.endswith('.0'):
        emp_id_text = emp_id_text[:-2]
    return emp_id_text


def normalize_cert_level_bucket(exam, level):
    """단계/급수를 1~6급/단계, 없음으로 표준화"""
    formatted_level = format_cert_level(exam, level)
    number_match = re.search(r"\d+", str(formatted_level))
    if number_match:
        level_number = int(number_match.group(0))
        if 1 <= level_number <= 6:
            return f"{level_number}급/{level_number}단계"
    return "없음"

# ──────────────────────────────────────────────────────────
# [저장 함수] 모든 세션 데이터를 JSON 파일과 사진으로 디스크에 보존
# ──────────────────────────────────────────────────────────
def save_data():
    """모든 데이터를 JSON 파일로 저장하고 사진을 폴더에 저장"""
    data_dir = "data"            # 데이터 디렉토리 이름 지정
    photos_dir = f"{data_dir}/photos"  # 사진 저장 하위 폴더 경로 지정
    
    if not os.path.exists(data_dir):  # data 폴더가 없으면
        os.makedirs(data_dir)          # data 폴더를 새로 생성
    
    if not os.path.exists(photos_dir):  # photos 폴더가 없으면
        os.makedirs(photos_dir)          # photos 폴더를 새로 생성
    
    # 세션 데이터를 JSON 파일로 저장
    with open(f"{data_dir}/workers.json", "w", encoding="utf-8") as f:
        # 인력 리스트를 workers.json에 한글 그대로 저장하고, 2칸 들여쓰기 적용
        json.dump(st.session_state.workers, f, ensure_ascii=False, indent=2)
    
    with open(f"{data_dir}/history.json", "w", encoding="utf-8") as f:
        # 변경 이력 리스트를 history.json에 저장
        json.dump(st.session_state.history, f, ensure_ascii=False, indent=2)
    
    with open(f"{data_dir}/korean_certs.json", "w", encoding="utf-8") as f:
        # 한국어 자격증 데이터를 korean_certs.json에 저장
        json.dump(st.session_state.korean_certs, f, ensure_ascii=False, indent=2)
    
    with open(f"{data_dir}/korean_classes.json", "w", encoding="utf-8") as f:
        # 한국어 교육 데이터를 korean_classes.json에 저장
        json.dump(st.session_state.korean_classes, f, ensure_ascii=False, indent=2)
    
    # 메타데이터 사전 구성 (업로드된 파일명 3가지)
    metadata = {
        "uploaded_worker_file_name": st.session_state.uploaded_worker_file_name,  # 인력 파일명
        "uploaded_cert_file_name": st.session_state.uploaded_cert_file_name,      # 자격증 파일명
        "uploaded_class_file_name": st.session_state.uploaded_class_file_name     # 교육 파일명
    }
    with open(f"{data_dir}/metadata.json", "w", encoding="utf-8") as f:
        json.dump(metadata, f, ensure_ascii=False, indent=2)  # 메타데이터를 metadata.json에 저장
    
    # 세션에 있는 사원 사진을 하나씩 PNG로 저장 (파일명: 사번.png)
    for emp_id, photo in st.session_state.worker_photos.items():  # 사번(emp_id)과 사진(photo)을 하나씩 꺼내 반복
        try:
            photo_path = f"{photos_dir}/{emp_id}.png"   # 저장할 전체 경로 구성
            if isinstance(photo, Image.Image):           # PIL Image 객체인지 확인
                photo.save(photo_path)                   # PNG 형식으로 디스크에 저장
        except Exception as e:
            print(f"사진 저장 오류 - 사번 {emp_id}: {e}")  # 저장 중 오류만 출력하고 계속 실행

# ──────────────────────────────────────────────────────────
# [로드 함수] 분기 기량평가 엑셀 데이터 캐시 로드
# ──────────────────────────────────────────────────────────
@st.cache_data
def load_aqe_data():
    """data/test_qae_list.xlsx 에서 분기 기량평가 이력을 읽어 DataFrame 반환"""
    aqe_path = "data/test_qae_list.xlsx"
    if os.path.exists(aqe_path):
        df = pd.read_excel(aqe_path, dtype={'사번': str})
        return df
    return None

# ──────────────────────────────────────────────────────────
# [로드 함수] 앱 시작 시 JSON 파일에서 세션으로 데이터 복원
# ──────────────────────────────────────────────────────────
def load_data():
    """JSON 파일에서 데이터 로드하고 사진 폴더에서 이미지 불러오기"""
    data_dir = "data"               # 데이터 폴더 이름
    photos_dir = f"{data_dir}/photos"  # 사진 폴더 경로
    
    # 인력 데이터 로드
    try:
        if os.path.exists(f"{data_dir}/workers.json"):  # 파일이 있으면
            with open(f"{data_dir}/workers.json", "r", encoding="utf-8") as f:
                st.session_state.workers = json.load(f)  # JSON을 리스트로 읽어 세션에 저장
    except (json.JSONDecodeError, FileNotFoundError):
        st.session_state.workers = []  # 파일 없거나 파싱 실패 시 빈 리스트로 초기화
    
    # 변경 이력 데이터 로드
    try:
        if os.path.exists(f"{data_dir}/history.json"):
            with open(f"{data_dir}/history.json", "r", encoding="utf-8") as f:
                st.session_state.history = json.load(f)  # JSON을 읽어 세션에 저장
    except (json.JSONDecodeError, FileNotFoundError):
        st.session_state.history = []  # 오류 시 빈 리스트
    
    # 한국어 자격증 데이터 로드
    try:
        if os.path.exists(f"{data_dir}/korean_certs.json"):
            with open(f"{data_dir}/korean_certs.json", "r", encoding="utf-8") as f:
                st.session_state.korean_certs = json.load(f)
    except (json.JSONDecodeError, FileNotFoundError):
        st.session_state.korean_certs = []  # 오류 시 빈 리스트
    
    # 한국어 교육 데이터 로드
    try:
        if os.path.exists(f"{data_dir}/korean_classes.json"):
            with open(f"{data_dir}/korean_classes.json", "r", encoding="utf-8") as f:
                st.session_state.korean_classes = json.load(f)
    except (json.JSONDecodeError, FileNotFoundError):
        st.session_state.korean_classes = []  # 오류 시 빈 리스트
    
    # 사진 로드: data/photos 폴더의 사번.png 파일들을 세션으로 읽기
    if os.path.exists(photos_dir):              # photos 폴더가 존재할 때만
        for filename in os.listdir(photos_dir):  # 폴더 내 모든 파일명 순회
            if filename.endswith(".png"):         # .png 파일만 실제 처리
                try:
                    emp_id = filename[:-4]  # 파일명에서 .png 확장자(4자) 제거 → 사번 추출
                    photo_path = os.path.join(photos_dir, filename)  # 전체 파일 경로 조합
                    photo = Image.open(photo_path)                    # PIL로 이미지 열기
                    st.session_state.worker_photos[emp_id] = photo   # 사번을 키로 세션에 저장
                except Exception as e:
                    print(f"사진 로드 오류 - {filename}: {e}")  # 로드 실패해도 계속 진행
    
    # 메타데이터 로드: 마지막에 업로드한 파일명들을 복원
    try:
        if os.path.exists(f"{data_dir}/metadata.json"):
            with open(f"{data_dir}/metadata.json", "r", encoding="utf-8") as f:
                metadata = json.load(f)  # metadata.json 파일 로드
                # 각 파일명을 세션에 복원 (없으면 기본값 '미등록' 사용)
                st.session_state.uploaded_worker_file_name = metadata.get("uploaded_worker_file_name", "미등록")
                st.session_state.uploaded_cert_file_name = metadata.get("uploaded_cert_file_name", "미등록")
                st.session_state.uploaded_class_file_name = metadata.get("uploaded_class_file_name", "미등록")
    except (json.JSONDecodeError, FileNotFoundError):
        pass  # 메타데이터 파일이 없으면 기본값 유지 (세션 초기값 그대로 사용)

# ──────────────────────────────────────────────────────────
# [유틸리티 함수] 엑셀에서 읽은 DataFrame 날짜/사번 데이터 정제
# ──────────────────────────────────────────────────────────
def clean_excel_data(df):
    """엑셀 데이터에서 시간정보 제거 및 날짜 포맷팅, 사번을 문자열로 변환"""
    df_clean = df.copy()  # 원본 DataFrame을 보호하기 위해 복사본으로 작업
    
    # 사번 열을 문자열로 변환 (숫자로 나오는 것 방지, 동일 형식 통일)
    if '사번' in df_clean.columns:  # '사번' 열이 존재할 때만 처리
        df_clean['사번'] = df_clean['사번'].astype(str).str.strip()  # str로 정수를 문자로, .strip()으로 양끝 공백 제거
    
    # 모든 열의 Timestamp/datetime 타입을 '년-월-일' 문자열로 강제 변환
    for col in df_clean.columns:   # DataFrame의 모든 열 이름을 하나씩 확인
        if col == '사번':  # 사번 열은 이미 위에서 처리했으므로 건너뜀
            continue
        
        # pandas가 인식한 datetime64 또는 대부분의 날짜형 열일 경우
        if pd.api.types.is_datetime64_any_dtype(df_clean[col]):
            df_clean[col] = df_clean[col].dt.strftime('%Y-%m-%d')  # 날짜를 '년-월-일' 문자열로
        elif df_clean[col].dtype == 'object':  # object 타입(일반 문자열 or 혼합 타입)
            # 세세하게 Timestamp 객체가 섞여 있는지 첫번째 유효값으로 확인
            try:
                # NaN 제외한 첫 번째 값을 가져와 타입 확인
                first_valid = df_clean[col].dropna().iloc[0] if len(df_clean[col].dropna()) > 0 else None
                if isinstance(first_valid, pd.Timestamp):  # pd.Timestamp 객체이면
                    df_clean[col] = df_clean[col].dt.strftime('%Y-%m-%d')  # 날짜 포맷 적용
            except (AttributeError, IndexError):
                pass  # Timestamp가 아니거나 빈 열이면 무시
    
    return df_clean  # 정제된 DataFrame 반환

# ---------------------------------------------------------
# [데이터 초기화] Session State 설정
# ---------------------------------------------------------
# session_state는 페이지가 다시 로드되어도 값이 유지되는 Streamlit의 전역저장소
# 파이썬의 'if key not in dict' 패턴: 키가 없을 때만 초기값 설정
if 'workers' not in st.session_state:
    st.session_state.workers = []          # 인력 리스트 (각 사원 = dict 하나)
if 'history' not in st.session_state:
    st.session_state.history = []          # 조직 변경 이력 리스트
if 'korean_certs' not in st.session_state:
    st.session_state.korean_certs = []     # 한국어 자격증 데이터 리스트
if 'korean_classes' not in st.session_state:
    st.session_state.korean_classes = []   # 한국어 교육 수업 데이터 리스트
if 'show_update_confirm' not in st.session_state:
    st.session_state.show_update_confirm = False  # 업데이트 확인 UI 표시 여부
if 'pending_worker_data' not in st.session_state:
    st.session_state.pending_worker_data = None   # 대기 중인 승인대기 작업자 데이터
if 'worker_photos' not in st.session_state:
    st.session_state.worker_photos = {}    # {사번: PIL.Image} 형태의 사진 사전
if 'confirm_reset_mode' not in st.session_state:
    st.session_state.confirm_reset_mode = False   # 데이터 전체 초기화 확인 상태
if 'selected_employee_data' not in st.session_state:
    st.session_state.selected_employee_data = None  # 현재 선택된 사원 데이터
if 'confirm_delete_mode' not in st.session_state:
    st.session_state.confirm_delete_mode = False    # 삭제 확인 다이얼로그 표시 여부
if 'pending_delete_emp_id' not in st.session_state:
    st.session_state.pending_delete_emp_id = None   # 삭제 대기 중인 사원의 사번
if 'pending_delete_name' not in st.session_state:
    st.session_state.pending_delete_name = None     # 삭제 대기 중인 사원의 이름
if 'edit_selected_emp_id' not in st.session_state:
    st.session_state.edit_selected_emp_id = None    # 현재 편집 중인 사원의 사번
if 'uploaded_worker_file_name' not in st.session_state:
    st.session_state.uploaded_worker_file_name = "미등록"  # 마지막에 업로드한 인력 파일명
if 'uploaded_cert_file_name' not in st.session_state:
    st.session_state.uploaded_cert_file_name = "미등록"    # 마지막에 업로드한 자격증 파일명
if 'uploaded_class_file_name' not in st.session_state:
    st.session_state.uploaded_class_file_name = "미등록"   # 마지막에 업로드한 교육 파일명

# 앱 시작 시 디스크에 저장된 JSON 파일에서 세션으로 데이터 복원
load_data()

# ──────────────────────────────────────────────────────────
# 조직 구조 정의: 부/과 → 반 → 직종 계층 구조
# ── 선택상자 옵션 종류가 어떻게 연결되는지 정의
# ──────────────────────────────────────────────────────────
dept_structure = {
    "공사1부5과": {                              # 제1 부/과 코드
        "반": ["1직1반", "1직2반", "1직3반", "2직1반", "2직2반", "2직3반"],  # 소속 반 목록
        "직종": ["수동본딩", "ABM", "TBP"]               # 해당 부/과의 직종 목록
    },
    "공사2부3과": {                              # 제2 부/과 코드
        "반": ["설치직1반", "설치직2반", "설치직3반", "설치직4반", "용접직1반", "용접직2반", "용접직3반"],
        "직종": ["MB설치", "MB수동용접", "MB자동용접", "MB리웰딩"]
    },
    "공사3부의장과": {                               # 제3 부/과 코드
        "반": ["2직1반", "2직2반", "2직3반"],
        "직종": ["의장", "LNGTIG"]
    }
}

# ---------------------------------------------------------
# [사이드바] 앱 좌측 메뉴 및 데이터 관리 버튼 구성
# ---------------------------------------------------------
st.sidebar.title("🌍 LNG선공사팀\n글로벌 인력관리")  # 사이드바 제목 표시
menu_map = {
    "dashboard": "📊 통합 대시보드",
    "workers": "👤 인력 정보 관리",
    "korean": "📚 한국어 교육/자격 관리",
    "evaluation": "📋 평가 이력",
    "safety_quiz": "🛡️ 안전 퀴즈",
}
menu_keys = list(menu_map.keys())
menu_options = [menu_map[key] for key in menu_keys]

# 새로고침(F5) 후에도 마지막 메뉴가 유지되도록, 라벨 대신 안정적인 키를 URL에 저장
saved_menu_key = st.query_params.get("menu", "dashboard")
if isinstance(saved_menu_key, list):
    saved_menu_key = saved_menu_key[0] if saved_menu_key else "dashboard"
if saved_menu_key not in menu_map:
    saved_menu_key = "dashboard"

default_index = menu_keys.index(saved_menu_key)
menu = st.sidebar.radio("메뉴 선택", menu_options, index=default_index)  # 라디오 버튼에서 선택된 메뉴 라벨 저장

selected_menu_key = menu_keys[menu_options.index(menu)]
current_menu_param = st.query_params.get("menu", "")
if isinstance(current_menu_param, list):
    current_menu_param = current_menu_param[0] if current_menu_param else ""
if current_menu_param != selected_menu_key:
    st.query_params["menu"] = selected_menu_key

# 메뉴가 바뀌면 화면 스크롤을 최상단으로 이동
previous_menu = st.session_state.get("last_selected_menu")
menu_changed = previous_menu is not None and previous_menu != menu
st.session_state.last_selected_menu = menu

if menu_changed:
        components.html(
                """
                <script>
                    const parentWin = window.parent;
                    const storageKey = `global-scroll-pos:${parentWin.location.pathname}`;
                    parentWin.scrollTo(0, 0);
                    try {
                        parentWin.localStorage.setItem(storageKey, '0');
                    } catch (e) {}
                </script>
                """,
                height=0,
                width=0,
        )

# 사이드바 - 데이터 관리 섹션
st.sidebar.write("---")          # 가로 구분선 출력
st.sidebar.subheader("⚙️ 데이터 관리")  # '데이터 관리' 소제목 출력

if st.sidebar.button("💾 데이터 저장", use_container_width=True):  # '데이터 저장' 버튼 클릭 시
    save_data()                                                        # save_data() 호출
    st.sidebar.success("✅ 데이터가 저장되었습니다!")               # 성공 메시지 표시

if st.sidebar.button("🔄 데이터 초기화", use_container_width=True):  # '데이터 초기화' 버튼 클릭 시
    st.session_state.confirm_reset_mode = True  # 초기화 확인 모드 활성화

# 초기화 확인 UI: confirm_reset_mode가 True일 때만 표시
if st.session_state.confirm_reset_mode:
    st.sidebar.warning("⚠️ **정말로 모든 데이터를 삭제하시겠습니까?**")  # 경고 메시지
    col_yn1, col_yn2 = st.sidebar.columns(2)  # 예/아니오 버튼을 2칸으로 나눔
    
    with col_yn1:  # 왼쪽 칸: 확인 버튼
        if st.button("✅ 예, 삭제합니다", use_container_width=True):
            # 세션 데이터 모두 빈 리스트로 지우기
            st.session_state.workers = []
            st.session_state.history = []
            st.session_state.korean_certs = []
            st.session_state.korean_classes = []
            st.session_state.worker_photos = {}
            
            # 디스크에 저장된 JSON 파일도 삭제
            data_dir = "data"  # 데이터 폴더 이름
            if os.path.exists(f"{data_dir}/workers.json"):         # 파일이 있으면
                os.remove(f"{data_dir}/workers.json")              # 인력 파일 삭제
            if os.path.exists(f"{data_dir}/history.json"):
                os.remove(f"{data_dir}/history.json")              # 이력 파일 삭제
            if os.path.exists(f"{data_dir}/korean_certs.json"):
                os.remove(f"{data_dir}/korean_certs.json")         # 자격증 파일 삭제
            if os.path.exists(f"{data_dir}/korean_classes.json"):
                os.remove(f"{data_dir}/korean_classes.json")       # 교육 파일 삭제
            
            st.session_state.confirm_reset_mode = False  # 확인 모드 해제
            st.sidebar.success("✅ 모든 데이터가 초기화되었습니다!")  # 성공 메시지
            st.rerun()  # 페이지 새로고침(삭제된 상태 반영)
    
    with col_yn2:  # 오른쪽 칸: 취소 버튼
        if st.button("❌ 아니오, 취소", use_container_width=True):
            st.session_state.confirm_reset_mode = False  # 확인 모드 해제
            st.rerun()  # 페이지 새로고침(취소 상태 반영)

# ---------------------------------------------------------
# 메뉴 1: 📊 통합 대시보드 (분석 및 시각화)
# ── menu 변수가 '통합 대시보드'일 때 이 섹션 표시
# ---------------------------------------------------------
if menu == "📊 통합 대시보드":
    st.header("📊 글로벌 인력 대시보드")  # 페이지 제목
    
    st.write("---")  # 구분선
    
    if st.session_state.workers:  # 등록된 인력이 있을 때만 아래 내용 표시
        df = pd.DataFrame(st.session_state.workers)  # 인력 리스트를 DataFrame으로 변환

        # 사번 열이 숫자로 읽혔을 경우 문자열로 명시 변환
        if '사번' in df.columns:
            df['사번'] = df['사번'].apply(normalize_emp_id)
        
        # 입사일에 시간 정보(HH:MM:SS)가 붙어있으면 날짜만 추출
        if '입사일' in df.columns:
            df['입사일'] = pd.to_datetime(df['입사일'], errors='coerce').dt.date  # 변환 실패는 NaT로 처리
        
        # 상단 인포 배너: 전체/국적별 현황 한눈에 표시
        st.info(f"✅ 시스템에 저장된 현황: **총 {len(df)}명** | 방글라데시: {len(df[df['국적']=='방글라데시'])}명 | 파키스탄: {len(df[df['국적']=='파키스탄'])}명")
        
        c1, c2, c3 = st.columns(3)  # 3칸 레이아웃
        c1.metric("총 인원", f"{len(df)}명")                                   # 전체 인원 수
        c2.metric("방글라데시", f"{len(df[df['국적']=='방글라데시'])}명")        # 방글라데시 인원 수
        c3.metric("파키스탄", f"{len(df[df['국적']=='파키스탄'])}명")            # 파키스탄 인원 수
        
        st.write("---")
        st.subheader("📚 한국어 자격증 보유 현황")
        
        # 자격증 보유자 집계 (사번 기준 최신 자격 1건만 인정)
        latest_cert_by_emp_id = {}
        for cert in st.session_state.korean_certs:
            cert_emp_id = normalize_emp_id(cert.get('사번', ''))
            cert_date = str(cert.get('취득일', '')).strip()

            if not cert_emp_id or cert_date in ['None', '', 'nan']:
                continue

            cert_sort_key = pd.to_datetime(cert_date, errors='coerce')
            if pd.isna(cert_sort_key):
                continue

            current_latest = latest_cert_by_emp_id.get(cert_emp_id)
            if current_latest is None or cert_sort_key > current_latest['_sort_key']:
                latest_cert_by_emp_id[cert_emp_id] = {
                    **cert,
                    '_sort_key': cert_sort_key
                }

        certified_ids = set(latest_cert_by_emp_id.keys())

        # 보유율 집계는 사번 기준으로 중복 제거 후 계산
        unique_workers_df = df.drop_duplicates(subset=['사번']).copy()
        
        total_employees = len(unique_workers_df)  # 전체 인원 수 (사번 중복 제거)
        # 인력 DataFrame에서 자격증 보유 사번에 해당하는 행만 추리기
        certified_employees = unique_workers_df[unique_workers_df['사번'].isin(certified_ids)]
        certified_count = len(certified_employees)  # 자격증 보유 인원 수
        
        # 3칸 메트릭으로 자격증 현황 요약
        col_cert1, col_cert2, col_cert3 = st.columns(3)
        col_cert1.metric("🎓 자격증 보유자", f"{certified_count}명")
        col_cert2.metric("📊 보유율", f"{(certified_count/total_employees*100):.1f}%")  # 소수점 1자리
        col_cert3.metric("미보유", f"{total_employees-certified_count}명")
        
        # 부/과별 자격증 현황 (왼쪽), 직/반별 (오른쪽) 2개 차트
        st.write("---")
        col_dept1, col_dept2 = st.columns(2)  # 2칸 레이아웃
        
        with col_dept1:
            st.subheader("부/과별 자격증 보유율")
            
            # 부/과별 전체 인원: groupby로 부/과 이름을 기준으로 사번 개수 집계
            dept_total = unique_workers_df.groupby('부/과')['사번'].count()
            # 자격증 보유자만 필터링 후 부/과별 집계
            dept_certified = unique_workers_df[unique_workers_df['사번'].isin(certified_ids)].groupby('부/과')['사번'].count()
            
            # 누락된 부/과 추가: 자격증 보유자가 없는 부/과는 집계에서 빠지므로 0으로 채워 넣기
            all_depts = set(unique_workers_df['부/과'].unique())  # 인력 데이터의 모든 부/과 이름 집합
            for dept in all_depts:
                if dept not in dept_certified.index:  # 자격증 집계에 없는 부/과면
                    dept_certified[dept] = 0           # 0으로 추가
            
            # 전체/자격증 칼럼을 합친 DataFrame 생성, 인덱스 오름차순 정렬
            dept_data = pd.DataFrame({
                '전체': dept_total,
                '자격증': dept_certified.sort_index()
            }).sort_index()
            dept_data['보유율(%)'] = (dept_data['자격증'] / dept_data['전체'] * 100).round(1)  # 보유율 계산 후 소수점 1자리 반올림
            
            # 부/과별 누적 막대 그래프 생성
            fig_dept, ax_dept = plt.subplots(figsize=(6, 4))  # 6x4인치 캔버스
            depts = dept_data.index.tolist()           # x축 레이블: 부/과 이름 목록
            certified = dept_data['자격증'].tolist()    # 자격증 보유 인원 목록
            not_certified = (dept_data['전체'] - dept_data['자격증']).tolist()  # 미보유 인원 목록
            
            x = np.arange(len(depts))  # x축 위치를 0,1,2,... 정수 배열로 생성
            width = 0.6                # 막대 너비 설정
            
            # 아래부터: 자격증 보유(초록), 위에 미보유(빨강) 누적
            ax_dept.bar(x, certified, width, label='자격증 보유', color='#2ecc71')               # 보유 막대 (초록)
            ax_dept.bar(x, not_certified, width, bottom=certified, label='미보유', color='#e74c3c')  # 미보유 막대 (빨강, 위에 쌓임)
            
            ax_dept.set_xlabel('부/과', fontsize=9, fontweight='bold')     # x축 레이블
            ax_dept.set_ylabel('인원(명)', fontsize=9, fontweight='bold')   # y축 레이블
            ax_dept.set_title('부/과별 자격증 보유율', fontsize=12, fontweight='bold')  # 차트 제목
            ax_dept.set_xticks(x)                                            # x 눈금 위치
            ax_dept.set_xticklabels(depts, rotation=45, ha='right', fontsize=8)          # 레이블 45도 회전
            ax_dept.legend(fontsize=8)                                                  # 범례 표시
            ax_dept.grid(axis='y', alpha=0.3)                                # y축 격자선 (반투명)
            
            plt.tight_layout()          # 레이블이 잘리지 않게 여백 자동 조정
            st.pyplot(fig_dept)          # Streamlit에 차트 렌더링
            
            # 부/과별 수치 테이블 표시
            with st.expander("부/과별 상세 통계", expanded=True):
                st.dataframe(dept_data[['전체', '자격증', '보유율(%)']], use_container_width=True)
        
        with col_dept2:
            st.subheader("직/반별 자격증 보유율")
            
            # 직/반별 전체 인원 집계
            unit_total = unique_workers_df.groupby('직/반')['사번'].count()
            # 자격증 보유자만 필터링 후 직/반별 집계
            unit_certified = unique_workers_df[unique_workers_df['사번'].isin(certified_ids)].groupby('직/반')['사번'].count()
            
            # 자격증 보유자가 없는 직/반은 집계에서 빠지므로 0으로 채워주기
            all_units = set(unique_workers_df['직/반'].unique())
            for unit in all_units:
                if unit not in unit_certified.index:
                    unit_certified[unit] = 0
            
            unit_data = pd.DataFrame({
                '전체': unit_total,
                '자격증': unit_certified.sort_index()
            }).sort_index()
            unit_data['보유율(%)'] = (unit_data['자격증'] / unit_data['전체'] * 100).round(1)
            
            # 직/반별 부/과 매핑 - 각 직/반이 어느 부/과에 속하는지 추출
            unit_to_dept = df.drop_duplicates('직/반')[['직/반', '부/과']].set_index('직/반')['부/과'].to_dict()
            unit_data['부/과'] = unit_data.index.map(unit_to_dept)
            
            # 직/반별 누적 막대 그래프
            fig_unit, ax_unit = plt.subplots(figsize=(6, 4))
            units = unit_data.index.tolist()
            certified_unit = unit_data['자격증'].tolist()
            not_certified_unit = (unit_data['전체'] - unit_data['자격증']).tolist()
            
            x_unit = np.arange(len(units))  # x축 위치 배열
            
            ax_unit.bar(x_unit, certified_unit, width, label='자격증 보유', color='#3498db')                           # 보유 막대 (파랑)
            ax_unit.bar(x_unit, not_certified_unit, width, bottom=certified_unit, label='미보유', color='#95a5a6')  # 미보유 막대 (회색)
            
            ax_unit.set_xlabel('직/반', fontsize=9, fontweight='bold')
            ax_unit.set_ylabel('인원(명)', fontsize=9, fontweight='bold')
            ax_unit.set_title('직/반별 자격증 보유율', fontsize=12, fontweight='bold')
            ax_unit.set_xticks(x_unit)
            ax_unit.set_xticklabels(units, rotation=45, ha='right', fontsize=8)
            ax_unit.legend(fontsize=8)
            ax_unit.grid(axis='y', alpha=0.3)
            
            plt.tight_layout()
            st.pyplot(fig_unit)
            
            # 직/반별 수치 테이블 (부/과 → 직/반 순서로 표시, 스크롤 없이 전체 표시)
            unit_data_display = unit_data.reset_index()
            unit_data_display['보유율(%)'] = unit_data_display['보유율(%)'].apply(lambda x: f"{x:.1f}%")
            st.table(unit_data_display[['부/과', '직/반', '전체', '자격증', '보유율(%)']])
        
        st.write("---")
        st.subheader("📋 전체 인력 명부")

        # 전체 인력 명부(기본정보 + 최신 자격증 정보) 생성
        roster_df = df[['사번', '이름', '영어이름', '국적', '부/과', '직/반', '직종', '입사일', '근속개월']].copy()

        latest_cert_map = {}
        for cert_emp_id, cert in latest_cert_by_emp_id.items():
            cert_date = str(cert.get('취득일', '')).strip()
            cert_exam = str(cert.get('시험', cert.get('시험종류', ''))).strip()
            cert_level = format_cert_level(cert_exam, cert.get('단계/급수', cert.get('자격', '')))

            latest_cert_map[cert_emp_id] = {
                '시험': cert_exam if cert_exam else '-',
                '단계/급수': cert_level if cert_level else '-',
                '취득일': cert_date if cert_date not in ['', 'None', 'nan'] else '-',
                '_sort_key': cert.get('_sort_key', pd.Timestamp.min)
            }

        roster_df['시험'] = roster_df['사번'].astype(str).map(lambda emp_id: latest_cert_map.get(emp_id, {}).get('시험', '-'))
        roster_df['단계/급수'] = roster_df['사번'].astype(str).map(lambda emp_id: latest_cert_map.get(emp_id, {}).get('단계/급수', '-'))
        roster_df['취득일'] = roster_df['사번'].astype(str).map(lambda emp_id: latest_cert_map.get(emp_id, {}).get('취득일', '-'))

        roster_df = roster_df[['사번', '이름', '영어이름', '국적', '부/과', '직/반', '직종', '입사일', '근속개월', '시험', '단계/급수', '취득일']]
        join_year_series = pd.to_datetime(roster_df['입사일'], errors='coerce').dt.year
        available_join_years = sorted(
            join_year_series.dropna().astype(int).astype(str).unique().tolist(),
            reverse=True
        )
        
        # 👇 필터 기능 추가
        st.write("**필터 옵션**")
        filter_col1, filter_col2, filter_col3, filter_col4, filter_col5 = st.columns(5)
        
        with filter_col1:
            dept_list = list(dept_structure.keys())
            filter_dept = st.selectbox(
                "부/과 선택",
                ["전체"] + dept_list,
                key="roster_filter_dept"
            )
        
        with filter_col2:
            if filter_dept == "전체":
                available_units = []
                for d in dept_list:
                    available_units.extend(dept_structure[d]["반"])
                available_units = sorted(list(set(available_units)))
            else:
                available_units = dept_structure[filter_dept]["반"]
            
            filter_unit = st.selectbox(
                "직/반 선택",
                ["전체"] + available_units,
                key="roster_filter_unit"
            )
        
        with filter_col3:
            filter_search = st.text_input(
                "이름/사번 검색",
                placeholder="검색...",
                key="roster_filter_search"
            )
        
        with filter_col4:
            filter_cert = st.selectbox(
                "자격증 소유",
                ["전체", "O", "X"],
                key="roster_filter_cert"
            )

        with filter_col5:
            filter_join_year = st.selectbox(
                "입사연도",
                ["전체"] + available_join_years,
                key="roster_filter_join_year"
            )
        
        # 필터 적용
        filtered_roster = roster_df.copy()
        
        # 부/과 필터
        if filter_dept != "전체":
            filtered_roster = filtered_roster[filtered_roster['부/과'] == filter_dept]
        
        # 직/반 필터
        if filter_unit != "전체":
            filtered_roster = filtered_roster[filtered_roster['직/반'] == filter_unit]
        
        # 이름/사번 검색
        if filter_search:
            filtered_roster = filtered_roster[
                (filtered_roster['이름'].str.contains(filter_search, case=False, na=False)) |
                (filtered_roster['사번'].astype(str).str.contains(filter_search, na=False))
            ]
        
        # 자격증 소유 필터
        if filter_cert == "O":
            filtered_roster = filtered_roster[filtered_roster['시험'] != '-']
        elif filter_cert == "X":
            filtered_roster = filtered_roster[filtered_roster['시험'] == '-']

        # 입사연도 필터
        if filter_join_year != "전체":
            filtered_roster = filtered_roster[
                pd.to_datetime(filtered_roster['입사일'], errors='coerce').dt.year == int(filter_join_year)
            ]
        
        st.caption(f"📋 조회 결과: **{len(filtered_roster)}명** (전체: {len(roster_df)}명)")
        st.dataframe(filtered_roster, width='stretch')
                    
        st.write("---")
        csv = filtered_roster.to_csv(index=False).encode('utf-8-sig')      # CSV 변환 (한글 깨짐 방지 utf-8-sig)
        excel_data = to_excel(filtered_roster)                              # Excel 변환 (to_excel 유틸리티 함수)

        col_dl1, col_dl2 = st.columns(2)  # 다운로드 버튼 2칸
        with col_dl1:
            st.download_button(
                "📥 현재 인력 명부 다운로드(CSV)",
                data=csv,
                file_name=f"global_list_{datetime.now().strftime('%Y%m%d')}.csv",  # 오늘 날짜로 파일명
                mime="text/csv"
            )
        with col_dl2:
            st.download_button(
                "📥 현재 인력 명부 다운로드(Excel)",
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
# ── menu 변수가 '인력 정보 관리'일 때 이 섹션 표시
# ---------------------------------------------------------
elif menu == "👤 인력 정보 관리":
    st.header("👤 인력 정보 관리")
    
    # 2개의 탭 생성: 조회 / 업로드
    tab1, tab2 = st.tabs(["👥 인력 정보 조회", "📥 데이터 및 사진 업로드/다운로드"])
    
    with tab1:  # ── 탭 1: 인력 정보 조회
        st.subheader("👥 등록된 인력 정보 조회")
        
        if st.session_state.workers:  # 등록된 인력이 있을 때만 표시
            # 자격 필터 드롭다운 옵션 생성 (등록된 자격 데이터 기준)
            cert_exam_options = sorted({
                str(cert.get('시험', '')).strip()
                for cert in st.session_state.korean_certs
                if str(cert.get('시험', '')).strip() not in ['', 'N/A', 'nan', 'None']
            })
            cert_level_options = [f"{i}급/{i}단계" for i in range(1, 7)] + ["없음"]

            # st.form: 위젯을 바꿔도 즉시 rerun되지 않도록 입력 요소를 묶음
            with st.form("filter_form"):
                st.write("**필터 조건 선택**")
                
                filter_col1, filter_col2 = st.columns(2)  # 2칸 레이아웃
                
                with filter_col1:
                    dept_list = list(dept_structure.keys())  # 부/과 이름 목록 추출
                    selected_filter_dept = st.selectbox(
                        "부/과 선택",
                        ["전체"] + dept_list,   # '전체' + 실제 부/과 목록
                        key="filter_dept"
                    )
                
                with filter_col2:
                    # 선택된 부/과에 따라 반 목록을 동적으로 변경
                    if selected_filter_dept == "전체":
                        available_units = []
                        for dept in dept_list:
                            available_units.extend(dept_structure[dept]["반"])  # 모든 부/과의 반을 합침
                        available_units = sorted(list(set(available_units)))    # 중복 제거 후 정렬
                    else:
                        available_units = dept_structure[selected_filter_dept]["반"]  # 선택된 부/과의 반만
                    
                    selected_filter_unit = st.selectbox(
                        "직/반 선택",
                        ["전체"] + available_units,
                        key="filter_unit"
                    )
                
                # 이름 또는 사번 텍스트 검색
                search_term = st.text_input(
                    "이름 또는 사번으로 검색 (선택 사항)",
                    placeholder="이름이나 사번 입력...",
                    key="filter_search"
                )

                # 자격 정보 필터 (시험 종류 / 단계·급수)
                filter_col3, filter_col4 = st.columns(2)
                with filter_col3:
                    selected_cert_exam = st.selectbox(
                        "자격 시험 종류",
                        ["전체"] + cert_exam_options,
                        key="filter_cert_exam"
                    )
                with filter_col4:
                    selected_cert_levels = st.multiselect(
                        "자격 단계/급수(복수 선택)",
                        cert_level_options,
                        key="filter_cert_level_multi"
                    )
                
                submitted = st.form_submit_button("🔍 조회", use_container_width=True)  # 조회 버튼 (폼 제출)
            
            st.write("---")
            
            # ── 필터링 로직: 세션에서 복사본을 만들어 순차적으로 필터 적용
            filtered_workers = st.session_state.workers.copy()  # 원본 훼손 방지를 위한 복사
            
            # 부/과 필터: '전체'가 아니면 해당 부/과 사원만 유지
            if selected_filter_dept != "전체":
                filtered_workers = [w for w in filtered_workers if w.get('부/과') == selected_filter_dept]
            
            # 직/반 필터: '전체'가 아니면 해당 직/반 사원만 유지
            if selected_filter_unit != "전체":
                filtered_workers = [w for w in filtered_workers if w.get('직/반') == selected_filter_unit]
            
            # 텍스트 검색 필터: 사번 또는 이름 어디에든 포함되면 유지
            if search_term:
                filtered_workers = [w for w in filtered_workers 
                                   if search_term in str(w.get('사번', '')) or 
                                      search_term in w.get('이름', '')]

            # 자격정보 필터: 사원의 자격 이력 중 하나라도 조건과 일치하면 유지
            if selected_cert_exam != "전체" or selected_cert_levels:
                filtered_workers_with_cert = []
                for worker in filtered_workers:
                    worker_emp_id = str(worker.get('사번', '')).strip()
                    has_matching_cert = False
                    selected_level_set = set(selected_cert_levels)
                    include_none_level = "없음" in selected_level_set
                    selected_numeric_levels = selected_level_set - {"없음"}

                    worker_certs = [
                        cert for cert in st.session_state.korean_certs
                        if str(cert.get('사번', '')).strip() == worker_emp_id
                    ]

                    for cert in worker_certs:
                        cert_exam = str(cert.get('시험', cert.get('시험종류', ''))).strip()
                        cert_level_bucket = normalize_cert_level_bucket(cert_exam, cert.get('단계/급수', cert.get('자격', '')))

                        exam_ok = selected_cert_exam == "전체" or cert_exam == selected_cert_exam
                        if not selected_cert_levels:
                            level_ok = True
                        else:
                            level_ok = cert_level_bucket in selected_numeric_levels

                        if exam_ok and level_ok:
                            has_matching_cert = True
                            break

                    if include_none_level and selected_cert_exam == "전체" and not worker_certs:
                        has_matching_cert = True

                    if has_matching_cert:
                        filtered_workers_with_cert.append(worker)

                filtered_workers = filtered_workers_with_cert
            
            # 조회 결과 헤더 출력 (건수, 파일명, 기준일시)
            result_col1, result_col2, result_col3 = st.columns([2, 1.5, 1.5])
            with result_col1:
                st.write(f"**조회 결과: {len(filtered_workers)}명** (전체: {len(st.session_state.workers)}명)")
            with result_col2:
                st.write(f"**📋 데이터 제목**: {st.session_state.uploaded_worker_file_name}")
            with result_col3:
                st.write(f"**📅 기준일시**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            st.write("---")
            
            if len(filtered_workers) > 0:  # 조회 결과가 1명 이상이면
                # 조회 결과 영역은 고정 높이 컨테이너로 표시하여 내부 스크롤 제공
                with st.container(height=900):
                    # 인력별 카드 형태로 표시 (사진 1칸 + 정보 4칸)
                    for idx, worker in enumerate(filtered_workers):
                        col_photo, col_info = st.columns([1, 4])  # 사진:정보 = 1:4 비율
                        
                        with col_photo:
                            emp_id = worker.get("사번")  # 현재 사원의 사번 추출
                            if emp_id in st.session_state.worker_photos:  # 사진이 등록되어 있으면
                                display_photo = st.session_state.worker_photos[emp_id].resize((180, 240), Image.Resampling.LANCZOS)  # 표시용 크기(가로 180, 세로 240, 비율 3:4)로 조정
                                st.image(display_photo, width=180)
                                st.markdown(
                                    f"<p style='text-align: center; margin-top: 0.25rem; font-size: {PHOTO_EMP_ID_FONT_SIZE_PX}px;'>사번: {emp_id}</p>",
                                    unsafe_allow_html=True
                                )  # 상단 글자 크기 설정값을 사용해 사진 아래 사번 텍스트 크기를 쉽게 조정
                            else:
                                st.info("📷\n사진\n없음")  # 사진 없을 때 안내 표시
                        
                        with col_info:
                            st.markdown(
                                f"<p style='font-size: {INFO_EMP_ID_FONT_SIZE_PX}px; font-weight: 700; margin: 0.1rem 0 0.4rem 0;'>🆔 사번: {worker.get('사번')}</p>",
                                unsafe_allow_html=True
                            )
                            worker_cert_history_rows = []  # 카드 단위 한국어 자격 이력 기본값
                            worker_class_history_rows = []  # 카드 단위 한국어 교육 이력 기본값
                            
                            col_l, col_r = st.columns(2)  # 기본 정보 2열 분할
                            with col_l:  # 왼쪽: 이름/국적 등
                                st.write(f"**📛 이름**: {worker.get('이름', 'N/A')}")
                                st.write(f"**🌐 영어이름**: {worker.get('영어이름', 'N/A')}")
                                st.write(f"**🌍 국적**: {worker.get('국적', 'N/A')}")
                                st.write(f"**📅 입사일**: {worker.get('입사일', 'N/A')}")
                            
                            with col_r:  # 오른쪽: 부서/직종 등
                                st.write(f"**🏢 부/과**: {worker.get('부/과', 'N/A')}")
                                st.write(f"**👥 직/반**: {worker.get('직/반', 'N/A')}")
                                st.write(f"**💼 직종**: {worker.get('직종', 'N/A')}")
                                st.write(f"**⏱️ 근속개월**: {worker.get('근속개월', 'N/A')}개월")
                            
                            # 해당 사원의 가장 최신 자격/교육 정보 조회
                            emp_id = worker.get('사번')
                            
                            # 자격증 목록에서 같은 사번의 항목을 모두 수집하고, 최신 1건을 별도로 선택
                            latest_cert = None       # 최신 자격증 (초기값 없음)
                            latest_cert_date = None  # 최신 취득일 (초기값 없음)
                            latest_cert_sort_key = None  # 최신 정렬 기준 날짜
                            worker_cert_history_rows = []   # 상세정보에 표시할 자격 이력 목록
                            for cert in st.session_state.korean_certs:
                                if str(cert.get('사번', '')).strip() == str(emp_id).strip():
                                    cert_exam = cert.get('시험', cert.get('시험종류', ''))
                                    cert_level = format_cert_level(cert_exam, cert.get('단계/급수', cert.get('자격', '')))
                                    cert_date = str(cert.get('취득일', '')).strip()

                                    # 상세 이력 테이블용 데이터 누적
                                    worker_cert_history_rows.append({
                                        "취득일": cert_date if cert_date not in ['None', 'nan', ''] else '-',
                                        "시험": cert_exam,
                                        "단계/급수": cert_level if cert_level else '-'
                                    })

                                    # 취득일이 유효(None·빈값·'nan' 제외)하고 기존보다 최신이면 갱신
                                    if cert_date and cert_date not in ['None', 'nan', '']:
                                        cert_sort_key = pd.to_datetime(cert_date, errors='coerce')
                                        if pd.isna(cert_sort_key):
                                            cert_sort_key = cert_date

                                        if latest_cert_sort_key is None or cert_sort_key > latest_cert_sort_key:
                                            latest_cert = cert
                                            latest_cert_date = cert_date
                                            latest_cert_sort_key = cert_sort_key
                            
                            # 교육 목록에서 같은 사번의 항목을 모두 수집하고, 최신 1건을 별도로 선택
                            latest_class = None
                            for cls in st.session_state.korean_classes:
                                if str(cls.get('사번', '')).strip() == str(emp_id).strip():
                                    latest_class = cls
                                    # 상세 이력 테이블용 데이터 누적
                                    class_result = cls.get('이수결과', '-')
                                    worker_class_history_rows.append({
                                        "수업시기": cls.get('수업시기', '-'),
                                        "수업명": cls.get('수업명', '-'),
                                        "이수결과": class_result if class_result else '-'
                                    })
                            
                            st.write("---")
                            
                            info_col1, info_col2 = st.columns(2)
                            with info_col1:
                                if latest_cert:
                                    cert_exam = latest_cert.get('시험', latest_cert.get('시험종류', ''))
                                    cert_level = format_cert_level(cert_exam, latest_cert.get('단계/급수', latest_cert.get('자격', '')))
                                    st.write(f"**🎓 한국어 자격**: {cert_exam} {cert_level}".strip())
                                    st.write(f"&nbsp;&nbsp;&nbsp;&nbsp;(취득일: {latest_cert_date})")
                                else:
                                    st.write("**🎓 한국어 자격**: -")
                            
                            with info_col2:
                                if latest_class:
                                    st.write(f"**🏫 한국어 교육**: {latest_class.get('수업명')}")
                                    st.write(f"&nbsp;&nbsp;&nbsp;&nbsp;({latest_class.get('수업시기')})")
                                else:
                                    st.write("**🏫 한국어 교육**: -")
                            # 접기/펼치기 가능한 상세 정보 영역
                            with st.expander("📋 상세 정보"):
                                col_l, col_r = st.columns(2)
                                with col_l:
                                    st.write(f"**🏠 숙소**: {worker.get('숙소구분', 'N/A')}")
                                    st.write(f"**👨‍👩‍👧 가족동반**: {worker.get('가족동반', 'N/A')}")
                                    st.write(f"**주소**: {worker.get('주소', 'N/A')}")
                                    st.write(f"**계약 구분**: {worker.get('계약', 'N/A')}")
                                    st.write(f"**비고**: {worker.get('비고', 'N/A')}")
                                    st.write("---")
                                with col_r:
                                    st.write(f"**🙏 종교**: {worker.get('종교', 'N/A')}")
                                    st.write(f"**⚠️ 안전사고 발생이력**: {worker.get('안전사고 발생이력', 'N/A')}")
                                    st.write(f"**🚲 전기자전거**: {worker.get('전기자전거', 'N/A')}")
                                    st.markdown("&nbsp;")
                                    st.markdown("&nbsp;")
                                    st.write("---")

                                st.write("**🎓 한국어 자격 이력**")
                                if worker_cert_history_rows:
                                    cert_history_df = pd.DataFrame(worker_cert_history_rows)
                                    cert_history_df['_취득일정렬용'] = pd.to_datetime(
                                        cert_history_df['취득일'].replace('-', pd.NA),
                                        errors='coerce'
                                    )
                                    cert_history_df = cert_history_df.sort_values(
                                        by='_취득일정렬용',
                                        ascending=False,
                                        na_position='last'
                                    ).drop(columns=['_취득일정렬용'])
                                    cert_history_df.index = range(1, len(cert_history_df) + 1)
                                    st.dataframe(cert_history_df, use_container_width=True)
                                else:
                                    st.write("- 등록된 한국어 자격 이력이 없습니다.")
                                
                                st.write("---")
                                st.write("**🏫 한국어 교육 이력**")

                                class_history_rows = worker_class_history_rows if 'worker_class_history_rows' in locals() else []
                                if class_history_rows:
                                    class_history_df = pd.DataFrame(class_history_rows)
                                    class_history_df.index = range(1, len(class_history_df) + 1)
                                    st.dataframe(class_history_df, use_container_width=True)
                                else:
                                    st.write("- 등록된 한국어 교육 이력이 없습니다.")

                                st.write("---")
                                st.write("**📊 분기 기량평가 이력**")
                                aqe_df = load_aqe_data()
                                if aqe_df is not None:
                                    emp_aqe = aqe_df[aqe_df['사번'].astype(str) == str(emp_id)][
                                        ['평가시기', '평가일', '점수합계', 'EE/ME/BE', '기량등급']
                                    ].copy()
                                    if not emp_aqe.empty:
                                        emp_aqe = emp_aqe.sort_values('평가시기', ascending=False).reset_index(drop=True)
                                        emp_aqe.index = range(1, len(emp_aqe) + 1)
                                        st.dataframe(emp_aqe, use_container_width=True)
                                    else:
                                        st.write("- 등록된 분기 기량평가 이력이 없습니다.")
                                else:
                                    st.write("- 분기평가 데이터 파일(data/test_qae_list.xlsx)을 찾을 수 없습니다.")
                        
                        st.divider()  # 사원 카드 간 구분선
            else:
                st.info("선택한 조건에 해당하는 인력이 없습니다.")
        else:
            st.info("등록된 인력이 없습니다. 먼저 엑셀 파일을 업로드해주세요.")
   
    with tab2:  # ── 탭 2: 데이터 및 사진 업로드
        st.subheader("📥 엑셀/CSV 데이터 일괄 업로드")
        st.write("엑셀 또는 CSV 파일(사번, 이름, 영어이름, 국적, 부/과, 직/반, 직종, 입사일, 근속개월, 안전사고 발생이력, 전기자전거, 종교)을 업로드합니다.")
        
        # tab2 전용 key 사용 (같은 페이지에서 file_uploader 중복 방지)
        uploaded_excel = st.file_uploader("엑셀 또는 CSV 파일(.xlsx, .csv)을 업로드하세요", type=["xlsx", "csv"], key="upload_excel_menu2")
        if uploaded_excel:
            st.session_state.uploaded_worker_file_name = uploaded_excel.name  # 업로드된 파일명 기록
            try:
                df_excel = read_uploaded_table(uploaded_excel)       # 업로드 파일 형식에 맞춰 읽기
                df_excel = clean_excel_data(df_excel)                # 날짜/사번 정제
                df_excel = df_excel.rename(columns={'전기 자전거': '전기자전거'})  # 구 형식 컬럼명 자동 변환
                df_excel.index = range(1, len(df_excel) + 1)        # 인덱스를 0이 아닌 1부터 시작
                
                # 필수 열 누락 여부 확인
                required_cols = ['사번', '이름', '영어이름', '국적', '부/과', '직/반', '직종', '입사일', '근속개월', '안전사고 발생이력', '전기자전거', '종교']
                missing_cols = [col for col in required_cols if col not in df_excel.columns]
                
                if missing_cols:
                    st.warning(f"⚠️ 다음 필수 컬럼이 누락되었습니다: {', '.join(missing_cols)}")
                
                st.dataframe(df_excel, width='stretch')  # 업로드된 데이터 미리보기
                
                if st.button("✅ 데이터 저장하기", key="save_excel_menu2"):
                    excel_data = df_excel.to_dict('records')  # dict 리스트로 변환
                    
                    # 카운터 초기화
                    added_count = 0    # 신규 추가 수
                    updated_count = 0  # 기존 데이터 업데이트 수
                    duplicate_count = 0
                    
                    for new_worker in excel_data:
                        # 빈 필드는 기본값으로 채워서 KeyError 방지
                        if '부/과' not in new_worker or pd.isna(new_worker.get('부/과')):
                            new_worker['부/과'] = 'N/A'
                        if '직/반' not in new_worker or pd.isna(new_worker.get('직/반')):
                            new_worker['직/반'] = 'N/A'
                        if '직종' not in new_worker or pd.isna(new_worker.get('직종')):
                            new_worker['직종'] = 'N/A'
                        if '입사일' not in new_worker or pd.isna(new_worker.get('입사일')):
                            new_worker['입사일'] = 'N/A'
                        if '근속개월' not in new_worker or pd.isna(new_worker.get('근속개월')):
                            new_worker['근속개월'] = 0
                        if '안전사고 발생이력' not in new_worker or pd.isna(new_worker.get('안전사고 발생이력')):
                            new_worker['안전사고 발생이력'] = 'N/A'
                        if '전기자전거' not in new_worker or pd.isna(new_worker.get('전기자전거')):
                            new_worker['전기자전거'] = 'N/A'
                        if '종교' not in new_worker or pd.isna(new_worker.get('종교')):
                            new_worker['종교'] = 'N/A'
                        
                        emp_id = new_worker.get('사번')  # 현재 처리 중인 사번 추출
                        # 같은 사번이 이미 있는지 확인
                        existing = next((w for w in st.session_state.workers if w.get('사번') == emp_id), None)
                        
                        if existing:  # 이미 존재하는 사번이면
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
                    
                    save_data()  # 모든 변경 완료 후 저장
                    
                    st.success(f"✅ 저장 완료: 신규 {added_count}명, 업데이트 {updated_count}명")
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")
        
        st.write("---")
        st.subheader("📷 사진 일괄 업로드")
        st.write("사번.jpg 형식으로 이름이 지정된 사진들을 업로드합니다. (예: 111111.jpg, 222222.jpg, ...)")
        
        uploaded_photos = st.file_uploader(
            "사진 파일을 선택하세요 (JPG, PNG)", 
            type=["jpg", "jpeg", "png"],   # 허용 확장자
            accept_multiple_files=True,    # 여러 파일 동시 선택 허용
            key="upload_photos_menu2"
        )
        
        if uploaded_photos:  # 하나 이상 선택됐으면
            photo_preview = []
            for photo_file in uploaded_photos:
                file_name = photo_file.name                          # 파일명 (예: 123456.jpg)
                emp_id_from_file = file_name.split('.')[0]           # '.' 기준으로 나눠 첫 부분 = 사번
                photo_preview.append({
                    "파일명": file_name,
                    "사번": emp_id_from_file
                })
            
            # 업로드 목록 미리보기 테이블
            preview_df = pd.DataFrame(photo_preview)
            st.dataframe(preview_df, width='stretch', use_container_width=True)
            
            if st.button("✅ 사진 저장하기", key="save_photos_menu2"):
                success_count = 0  # 저장 성공 수
                error_count = 0    # 오류 발생 수
                
                for photo_file in uploaded_photos:
                    try:
                        file_name = photo_file.name
                        emp_id = file_name.split('.')[0]  # 파일명에서 사번 추출
                        
                        # 해당 사번이 인력 리스트에 있는지 확인
                        existing = next((w for w in st.session_state.workers if w.get('사번') == emp_id), None)
                        
                        if existing:
                            # 사진을 3:5 비율(300x500)로 리사이징 후 세션에 저장
                            img = Image.open(photo_file)                           # PIL로 이미지 열기
                            img_resized = img.resize((300, 500), Image.Resampling.LANCZOS)  # 고품질 리사이징
                            st.session_state.worker_photos[emp_id] = img_resized   # 사번을 키로 세션에 저장
                            success_count += 1
                        else:
                            st.warning(f"⚠️ 사번 {emp_id}는 등록되지 않았습니다. (파일: {file_name})")
                            error_count += 1
                    except Exception as e:
                        st.warning(f"사진 처리 중 오류 - {photo_file.name}: {e}")
                        error_count += 1
                
                save_data()  # 메모리의 사진을 PNG 파일로 디스크에 저장
                st.success(f"✅ 사진 저장 완료: 성공 {success_count}개" + (f", 실패 {error_count}개" if error_count > 0 else ""))
        
        st.write("---")
        st.subheader("📥 인력 정보 다운로드")
        
        if st.session_state.workers:
            st.write(f"**저장된 인력 정보**: {len(st.session_state.workers)}명")
            
            # 인력 정보 DataFrame 생성
            download_df = pd.DataFrame(st.session_state.workers)
            
            # 사번 문자열 변환
            if '사번' in download_df.columns:
                download_df['사번'] = download_df['사번'].astype(str)
            
            # 입사일 날짜 형식 처리
            if '입사일' in download_df.columns:
                download_df['입사일'] = pd.to_datetime(download_df['입사일'], errors='coerce').dt.date
            
            # 다운로드용 CSV, Excel 변환
            csv_data = download_df.to_csv(index=False).encode('utf-8-sig')
            excel_data = to_excel(download_df)
            
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                st.download_button(
                    "📥 인력 정보 다운로드(CSV)",
                    data=csv_data,
                    file_name=f"workers_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv"
                )
            with col_dl2:
                st.download_button(
                    "📥 인력 정보 다운로드(Excel)",
                    data=excel_data,
                    file_name=f"workers_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.info("저장된 인력 정보가 없습니다.")

# ══════════════════════════════════════════════════════════════════
# 메뉴 3: 한국어 교육/자격 관리
# - t1: 자격증 취득 정보 업로드 및 조회
# - t2: 교육 수업 참여 정보 업로드 및 조회
# ══════════════════════════════════════════════════════════════════
elif menu == "📚 한국어 교육/자격 관리":
    st.header("📚 한국어 역량")          # 페이지 제목 표시
    
    # 2개 탭 생성: 자격증 데이터, 교육 데이터
    t1, t2 = st.tabs(["📥 자격증 데이터", "📥 교육 데이터"])
    
    with t1:
        st.subheader("📥 자격증 취득 정보 업로드")
        st.write("엑셀 또는 CSV 파일 형식: 사번, 이름, 영어이름, 국적, 부/과, 직/반, 직종, 시험, 단계/급수, 취득일")  # 업로드 파일의 필수 열 안내
        
        uploaded_cert_file = st.file_uploader("자격증 엑셀 또는 CSV 파일을 업로드하세요", type=["xlsx", "csv"], key="upload_cert_excel")
        if uploaded_cert_file:
            st.session_state.uploaded_cert_file_name = uploaded_cert_file.name  # 업로드한 파일명을 세션에 저장 (메뉴 2 조회 탭에 표시할 때 사용)
            try:
                df_cert = read_uploaded_table(uploaded_cert_file)    # 업로드 파일 형식에 맞춰 읽기
                df_cert = clean_excel_data(df_cert)                  # 날짜 형식 등 데이터 정제
                
                # 자격증 업로드 파일에 필요한 열이 모두 있는지 확인
                required_cols = ['사번', '이름', '영어이름', '국적', '부/과', '직/반', '직종', '시험', '단계/급수', '취득일']
                missing_cols = [col for col in required_cols if col not in df_cert.columns]  # 없는 컬럼 목록
                
                if missing_cols:
                    st.warning(f"⚠️ 다음 필수 컬럼이 누락되었습니다: {', '.join(missing_cols)}")
                
                df_cert.index = range(1, len(df_cert) + 1)  # 인덱스를 1부터 시작하도록 리셋
                
                st.write("**📋 업로드 데이터 미리보기**")
                st.dataframe(df_cert, use_container_width=True)  # 업로드 원본 그대로 표시
                
                # 저장 방식을 선택하는 버튼 2개를 좌우로 배치
                col_btn1, col_btn2 = st.columns(2)
                
                with col_btn1:
                    # 교체 모드: 기존 데이터를 전부 지우고 새 파일 데이터로 덮어쓰기
                    if st.button("🔄 기존 데이터 삭제 후 교체하기", key="replace_cert_excel", help="기존 자격증 데이터를 모두 삭제하고 새 데이터로 교체합니다"):
                        cert_data = df_cert.to_dict('records')  # DataFrame → dict 리스트 변환
                        
                        # 비어 있거나 누락된 값을 'N/A' 기본값으로 보정
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
                        
                        # 기존 자격증 데이터를 새 파일 내용으로 완전히 교체
                        st.session_state.korean_certs = cert_data
                        save_data()                                   # 변경된 데이터를 JSON 파일로 저장
                        st.success(f"✅ 저장 완료: 기존 데이터 삭제 후 {len(cert_data)}건 새로 저장되었습니다")
                        st.rerun()                                    # 저장 결과가 바로 보이도록 다시 실행
                
                with col_btn2:
                    # 추가 모드: 기존 데이터에 새 데이터를 병합 (중복 시 업데이트)
                    if st.button("✅ 기존 데이터에 추가하기", key="save_cert_excel", help="기존 자격증 데이터에 새 데이터를 추가합니다"):
                        cert_data = df_cert.to_dict('records')
                        
                        added_count = 0    # 신규 추가된 건수
                        updated_count = 0  # 기존 데이터 업데이트된 건수
                        
                        for new_cert in cert_data:
                            # 비어 있거나 누락된 값을 기본값으로 보정
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
                            
                            # 사번 + 시험 + 단계/급수 + 취득일이 모두 같으면 같은 자격 기록으로 간주
                            existing = next((c for c in st.session_state.korean_certs 
                                           if c.get('사번') == new_cert.get('사번') and 
                                              c.get('시험') == new_cert.get('시험') and
                                              c.get('단계/급수') == new_cert.get('단계/급수') and
                                              c.get('취득일') == new_cert.get('취득일')), None)
                            
                            if existing:
                                # 기존 데이터가 있으면 해당 항목을 새 값으로 덮어씀
                                for i, cert in enumerate(st.session_state.korean_certs):
                                    if (cert.get('사번') == new_cert.get('사번') and 
                                        cert.get('시험') == new_cert.get('시험') and
                                        cert.get('단계/급수') == new_cert.get('단계/급수') and
                                        cert.get('취득일') == new_cert.get('취득일')):
                                        st.session_state.korean_certs[i] = new_cert  # 찾은 위치의 데이터를 새 값으로 교체
                                        updated_count += 1
                                        break
                            else:
                                # 기존 데이터가 없으면 새 항목으로 추가
                                st.session_state.korean_certs.append(new_cert)
                                added_count += 1
                        
                        save_data()
                        st.success(f"✅ 저장 완료: 신규 {added_count}건, 업데이트 {updated_count}건")
                        st.rerun()
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")  # 파일 파싱 실패 시 오류 메시지

        st.write("---")
        st.subheader("📚 한국어 교육/자격 조회")

        # 자격증 또는 교육 데이터가 하나라도 있으면 조회 폼 표시
        if st.session_state.korean_certs or st.session_state.korean_classes:
            st.write("**필터 조건 선택**")

            filter_col1, filter_col2 = st.columns(2)
            with filter_col1:
                dept_list = list(dept_structure.keys())
                selected_dept = st.selectbox("📍 부/과", ["전체"] + dept_list, key="korean_m3_dept_single")

            with filter_col2:
                if selected_dept != "전체":
                    available_units = sorted(list(set(dept_structure[selected_dept]["반"])))
                else:
                    available_units = []
                    for dept in dept_list:
                        available_units.extend(dept_structure[dept]["반"])
                    available_units = sorted(list(set(available_units)))

                selected_unit = st.selectbox("👥 직/반", ["전체"] + available_units, key="korean_m3_unit_single")

            filter_col3, filter_col4 = st.columns(2)
            with filter_col3:
                search_name = st.text_input("👤 이름 검색", placeholder="이름 입력...", key="korean_m3_name")

            with filter_col4:
                search_id = st.text_input("🆔 사번 검색", placeholder="사번 입력...", key="korean_m3_id")

            filter_col5, filter_col6 = st.columns(2)
            with filter_col5:
                selected_exam = st.selectbox("📋 시험 종류", ["전체", "사통", "사전평가", "TOPIK"], key="korean_m3_exam_single")

            with filter_col6:
                selected_level = st.selectbox("📊 단계/급수", ["전체"] + [f"{i}단계" for i in range(1, 7)] + [f"{i}급" for i in range(1, 7)], key="korean_m3_level_single")

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
            selected_year = st.selectbox("📅 취득연도", ["전체"] + available_years, key="korean_m3_year")

            available_join_years = set()
            for worker in st.session_state.workers:
                join_date = str(worker.get('입사일', '')).strip()
                join_year = pd.to_datetime(join_date, errors='coerce').year
                if pd.notna(join_year):
                    available_join_years.add(str(int(join_year)))

            available_join_years = sorted(list(available_join_years), reverse=True)
            selected_join_year = st.selectbox("📅 입사연도", ["전체"] + available_join_years, key="korean_m3_join_year")

            st.write("---")

            filtered_certs = st.session_state.korean_certs.copy()
            worker_map = {w.get('사번'): w for w in st.session_state.workers}

            if selected_dept != "전체":
                filtered_certs = [c for c in filtered_certs if c.get('부/과') == selected_dept]

            if selected_unit != "전체":
                filtered_certs = [c for c in filtered_certs if c.get('직/반') == selected_unit]

            if search_name:
                filtered_certs = [c for c in filtered_certs if search_name in worker_map.get(c.get('사번'), {}).get('이름', '')]

            if search_id:
                filtered_certs = [c for c in filtered_certs if search_id in str(c.get('사번', ''))]

            if selected_exam != "전체":
                filtered_certs = [c for c in filtered_certs if c.get('시험') == selected_exam]

            if selected_level != "전체":
                filtered_certs = [
                    c for c in filtered_certs
                    if format_cert_level(c.get('시험', ''), c.get('단계/급수', '')) == selected_level
                ]

            if selected_year != "전체":
                filtered_certs = [c for c in filtered_certs if str(c.get('취득일', '')).startswith(selected_year)]

            if selected_join_year != "전체":
                filtered_certs = [
                    c for c in filtered_certs
                    if str(pd.to_datetime(worker_map.get(c.get('사번'), {}).get('입사일', ''), errors='coerce').year) == selected_join_year
                ]

            result_col1, result_col2, result_col3 = st.columns([2, 1.5, 1.5])
            with result_col1:
                st.write(f"**📊 조회 결과: {len(filtered_certs)}건** (전체: {len(st.session_state.korean_certs)}건)")
            with result_col2:
                st.write(f"**📋 데이터 제목**: {st.session_state.uploaded_cert_file_name}")
            with result_col3:
                st.write(f"**📅 기준일시**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            st.write("---")

            if len(filtered_certs) > 0:
                cert_display = []
                for cert in filtered_certs:
                    exam = cert.get('시험', 'N/A')
                    exam = '' if exam == 'N/A' else str(exam).strip()
                    level = format_cert_level(exam, cert.get('단계/급수', 'N/A'))

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

                csv_data = cert_df.to_csv(index=False).encode('utf-8-sig')
                excel_data = to_excel(cert_df)

                col_dlk1, col_dlk2 = st.columns(2)
                with col_dlk1:
                    st.download_button(
                        "📥 조회결과 다운로드(CSV)",
                        data=csv_data,
                        file_name=f"korean_certs_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv",
                        key="korean_m3_dl_csv"
                    )
                with col_dlk2:
                    st.download_button(
                        "📥 조회결과 다운로드(Excel)",
                        data=excel_data,
                        file_name=f"korean_certs_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="korean_m3_dl_excel"
                    )
            else:
                st.info("선택한 조건에 해당하는 자격정보가 없습니다.")
        else:
            st.info("등록된 한국어 자격 정보가 없습니다. 메뉴 3에서 데이터를 추가해주세요.")
    
    with t2:
        st.subheader("📥 교육 참여 정보 업로드")
        st.write("엑셀 또는 CSV 파일 형식: 사번, 이름, 영어이름, 국적, 부/과, 직/반, 직종, 수업명, 수업시기, 이수결과")  # 업로드 파일 형식 안내
        
        uploaded_class_file = st.file_uploader("교육 엑셀 또는 CSV 파일을 업로드하세요", type=["xlsx", "csv"], key="upload_class_excel")
        if uploaded_class_file:
            st.session_state.uploaded_class_file_name = uploaded_class_file.name  # 파일명 세션에 저장
            try:
                df_class = read_uploaded_table(uploaded_class_file)  # 업로드 파일 형식에 맞춰 읽기
                # 교육 업로드 파일에 필요한 열(사번·이름·수업명·수업시기·이수결과)이 모두 있는지 확인
                required_cols = ['사번', '이름', '영어이름', '국적', '부/과', '직/반', '직종', '수업명', '수업시기', '이수결과']
                missing_cols = [col for col in required_cols if col not in df_class.columns]
                
                if missing_cols:
                    st.warning(f"⚠️ 다음 필수 컬럼이 누락되었습니다: {', '.join(missing_cols)}")
                
                df_class.index = range(1, len(df_class) + 1)  # 인덱스 1부터 시작
                st.dataframe(df_class, use_container_width=True)  # 업로드 미리보기
                
                col_btn1, col_btn2 = st.columns(2)
                
                with col_btn1:
                    # 교체 모드: 기존 데이터를 전부 지우고 새 파일 데이터로 덮어쓰기
                    if st.button("🔄 기존 데이터 삭제 후 교체하기", key="replace_class_excel", help="기존 교육 데이터를 모두 삭제하고 새 데이터로 교체합니다"):
                        class_data = df_class.to_dict('records')  # DataFrame → dict 리스트 변환
                        st.session_state.korean_classes = class_data
                        save_data()
                        st.success(f"✅ 저장 완료: 기존 데이터 삭제 후 {len(class_data)}건 새로 저장되었습니다")
                        st.rerun()
                
                with col_btn2:
                    # 추가 모드: 기존 데이터에 새 데이터를 병합 (중복 시 업데이트)
                    if st.button("✅ 기존 데이터에 추가하기", key="save_class_excel", help="기존 교육 데이터에 새 데이터를 추가합니다"):
                        class_data = df_class.to_dict('records')  # dict 리스트로 변환
                        
                        added_count = 0    # 신규 추가 건수
                        updated_count = 0  # 업데이트 건수
                        
                        for new_class in class_data:
                            # 사번 + 수업명 + 시기가 모두 같으면 같은 교육 기록으로 간주
                            existing = next((c for c in st.session_state.korean_classes 
                                           if c.get('사번') == new_class.get('사번') and 
                                              c.get('수업명') == new_class.get('수업명') and
                                              c.get('수업시기') == new_class.get('수업시기')), None)
                            
                            if existing:
                                # 기존 데이터가 있으면 새 값으로 업데이트
                                for i, cls in enumerate(st.session_state.korean_classes):
                                    if (cls.get('사번') == new_class.get('사번') and 
                                        cls.get('수업명') == new_class.get('수업명') and
                                        cls.get('수업시기') == new_class.get('수업시기')):
                                        st.session_state.korean_classes[i] = new_class  # 찾은 위치의 데이터를 새 값으로 교체
                                        updated_count += 1
                                        break
                            else:
                                # 기존 데이터가 없으면 새 항목으로 추가
                                st.session_state.korean_classes.append(new_class)
                                added_count += 1
                        
                        save_data()  # 변경된 교육 데이터를 JSON에 저장
                        st.success(f"✅ 저장 완료: 신규 {added_count}건, 업데이트 {updated_count}건")
                        st.rerun()
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")  # 파싱 실패 시 오류 표시

        st.write("---")
        st.write("**📋 교육 데이터 표 (필터 적용)**")

        class_table_columns = ['사번', '이름', '영어이름', '국적', '부/과', '직/반', '직종', '수업명', '수업시기', '이수결과']
        class_table_df = pd.DataFrame(st.session_state.korean_classes)

        if not class_table_df.empty:
            for column_name in class_table_columns:
                if column_name not in class_table_df.columns:
                    class_table_df[column_name] = ''
            class_table_df = class_table_df[class_table_columns].copy()

            class_filter_col1, class_filter_col2, class_filter_col3, class_filter_col4 = st.columns(4)

            with class_filter_col1:
                class_filter_dept = st.selectbox(
                    "부/과 선택",
                    ["전체"] + sorted(class_table_df['부/과'].dropna().astype(str).unique().tolist()),
                    key="korean_class_filter_dept_t2"
                )

            with class_filter_col2:
                if class_filter_dept != "전체":
                    class_available_units = sorted(
                        class_table_df[class_table_df['부/과'].astype(str) == class_filter_dept]['직/반']
                        .dropna().astype(str).unique().tolist()
                    )
                else:
                    class_available_units = sorted(class_table_df['직/반'].dropna().astype(str).unique().tolist())

                class_filter_unit = st.selectbox(
                    "직/반 선택",
                    ["전체"] + class_available_units,
                    key="korean_class_filter_unit_t2"
                )

            with class_filter_col3:
                class_filter_search = st.text_input(
                    "이름/사번 검색",
                    placeholder="검색...",
                    key="korean_class_filter_search_t2"
                )

            with class_filter_col4:
                class_filter_course = st.selectbox(
                    "수업명",
                    ["전체"] + sorted(class_table_df['수업명'].dropna().astype(str).unique().tolist()),
                    key="korean_class_filter_course_t2"
                )

            filtered_class_table_df = class_table_df.copy()

            if class_filter_dept != "전체":
                filtered_class_table_df = filtered_class_table_df[
                    filtered_class_table_df['부/과'].astype(str) == class_filter_dept
                ]

            if class_filter_unit != "전체":
                filtered_class_table_df = filtered_class_table_df[
                    filtered_class_table_df['직/반'].astype(str) == class_filter_unit
                ]

            if class_filter_search:
                filtered_class_table_df = filtered_class_table_df[
                    filtered_class_table_df['이름'].astype(str).str.contains(class_filter_search, case=False, na=False)
                    | filtered_class_table_df['사번'].astype(str).str.contains(class_filter_search, na=False)
                ]

            if class_filter_course != "전체":
                filtered_class_table_df = filtered_class_table_df[
                    filtered_class_table_df['수업명'].astype(str) == class_filter_course
                ]

            # 최신 자격정보 추가
            latest_cert_by_emp_id = {}
            for cert in st.session_state.korean_certs:
                cert_emp_id = normalize_emp_id(cert.get('사번', ''))
                cert_date = str(cert.get('취득일', '')).strip()

                if not cert_emp_id or cert_date in ['None', '', 'nan']:
                    continue

                cert_sort_key = pd.to_datetime(cert_date, errors='coerce')
                if pd.isna(cert_sort_key):
                    continue
                current_latest = latest_cert_by_emp_id.get(cert_emp_id)
                if current_latest is None or cert_sort_key > current_latest['_sort_key']:
                    latest_cert_by_emp_id[cert_emp_id] = {
                        **cert,
                        '_sort_key': cert_sort_key
                    }

            filtered_class_table_df['최신자격시험'] = filtered_class_table_df['사번'].astype(str).apply(
                lambda emp_id: latest_cert_by_emp_id.get(normalize_emp_id(emp_id), {}).get('시험', '-')
            )
            filtered_class_table_df['단계.급수'] = filtered_class_table_df['사번'].astype(str).apply(
                lambda emp_id: format_cert_level(latest_cert_by_emp_id.get(normalize_emp_id(emp_id), {}).get('시험', ''),
                                                 latest_cert_by_emp_id.get(normalize_emp_id(emp_id), {}).get('단계/급수', ''))
            )
            filtered_class_table_df['취득일'] = filtered_class_table_df['사번'].astype(str).apply(
                lambda emp_id: latest_cert_by_emp_id.get(normalize_emp_id(emp_id), {}).get('취득일', '-')
            )

            filtered_class_table_df.index = range(1, len(filtered_class_table_df) + 1)
            st.caption(f"📋 조회 결과: **{len(filtered_class_table_df)}건** (전체: {len(class_table_df)}건)")
            st.dataframe(filtered_class_table_df, use_container_width=True)

            class_csv_data = filtered_class_table_df.to_csv(index=False).encode('utf-8-sig')
            class_excel_data = to_excel(filtered_class_table_df)

            class_dl_col1, class_dl_col2 = st.columns(2)
            with class_dl_col1:
                st.download_button(
                    "📥 조회결과 다운로드(CSV)",
                    data=class_csv_data,
                    file_name=f"korean_classes_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    key="korean_class_dl_csv_t2"
                )
            with class_dl_col2:
                st.download_button(
                    "📥 조회결과 다운로드(Excel)",
                    data=class_excel_data,
                    file_name=f"korean_classes_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="korean_class_dl_excel_t2"
                )
        else:
            st.info("등록된 교육 데이터가 없습니다.")
    
# ══════════════════════════════════════════════════════════════════
# 메뉴 4: 평가 이력 관리
# - t1: 분기 평가 데이터 업로드 및 저장
# - t2: 월별 평가 데이터 업로드 및 저장
# ══════════════════════════════════════════════════════════════════
elif menu == "📋 평가 이력":
    st.header("📋 평가 이력")  # 페이지 제목 표시
    
    # 2개 탭 생성: 분기평가 업로드, 월별평가 업로드
    t1, t2 = st.tabs(["📥 분기 평가 데이터 업로드", "📥 월별 평가 데이터 업로드"])
    
    with t1:
        st.subheader("📥 분기 평가 정보 업로드")
        st.write("엑셀 또는 CSV 파일 형식: 사번, 이름, 영어이름, 국적, 부/과, 직/반, 직종, 입사일, 평가시기, 평가일, 반장 기능(70), 반장 역량(30), 직장 기능(70), 직장 역량(30), 과장 기능(70), 과장 역량(30), 부장 기능(70), 부장 역량(30), EE/ME/BE, 점수합계, 기량등급")
        
        uploaded_aqe_file = st.file_uploader("분기 평가 엑셀 또는 CSV 파일을 업로드하세요", type=["xlsx", "csv"], key="upload_aqe_excel")
        if uploaded_aqe_file:
            try:
                df_aqe = read_uploaded_table(uploaded_aqe_file)  # 업로드 파일 형식에 맞춰 읽기
                df_aqe = clean_excel_data(df_aqe)  # 날짜 형식 등 데이터 정제
                
                # 분기평가 업로드 파일에 필요한 열이 모두 있는지 확인
                required_cols = ['사번', '이름', '영어이름', '국적', '부/과', '직/반', '직종', '입사일', '평가시기', '평가일', 
                               '반장 기능(70)', '반장 역량(30)', '직장 기능(70)', '직장 역량(30)', 
                               '과장 기능(70)', '과장 역량(30)', '부장 기능(70)', '부장 역량(30)', 
                               'EE/ME/BE', '점수합계', '기량등급']
                missing_cols = [col for col in required_cols if col not in df_aqe.columns]
                
                if missing_cols:
                    st.warning(f"⚠️ 다음 필수 컬럼이 누락되었습니다: {', '.join(missing_cols)}")
                
                df_aqe.index = range(1, len(df_aqe) + 1)  # 인덱스를 1부터 시작하도록 리셋
                
                st.write("**📋 업로드 데이터 미리보기 (필터 적용)**")
                aqe_filter_col1, aqe_filter_col2, aqe_filter_col3, aqe_filter_col4 = st.columns(4)

                with aqe_filter_col1:
                    aqe_filter_dept = st.selectbox(
                        "부/과 선택",
                        ["전체"] + sorted(df_aqe['부/과'].dropna().astype(str).unique().tolist()) if '부/과' in df_aqe.columns else ["전체"],
                        key="eval_upload_aqe_filter_dept"
                    )

                with aqe_filter_col2:
                    if '직/반' in df_aqe.columns:
                        if aqe_filter_dept != "전체" and '부/과' in df_aqe.columns:
                            aqe_units = sorted(
                                df_aqe[df_aqe['부/과'].astype(str) == aqe_filter_dept]['직/반']
                                .dropna().astype(str).unique().tolist()
                            )
                        else:
                            aqe_units = sorted(df_aqe['직/반'].dropna().astype(str).unique().tolist())
                    else:
                        aqe_units = []

                    aqe_filter_unit = st.selectbox(
                        "직/반 선택",
                        ["전체"] + aqe_units,
                        key="eval_upload_aqe_filter_unit"
                    )

                with aqe_filter_col3:
                    aqe_filter_search = st.text_input(
                        "이름/사번 검색",
                        placeholder="검색...",
                        key="eval_upload_aqe_filter_search"
                    )

                with aqe_filter_col4:
                    aqe_filter_grade = st.selectbox(
                        "EE/ME/BE",
                        ["전체"] + sorted(df_aqe['EE/ME/BE'].dropna().astype(str).unique().tolist()) if 'EE/ME/BE' in df_aqe.columns else ["전체"],
                        key="eval_upload_aqe_filter_grade"
                    )

                filtered_df_aqe = df_aqe.copy()

                if aqe_filter_dept != "전체" and '부/과' in filtered_df_aqe.columns:
                    filtered_df_aqe = filtered_df_aqe[filtered_df_aqe['부/과'].astype(str) == aqe_filter_dept]

                if aqe_filter_unit != "전체" and '직/반' in filtered_df_aqe.columns:
                    filtered_df_aqe = filtered_df_aqe[filtered_df_aqe['직/반'].astype(str) == aqe_filter_unit]

                if aqe_filter_search:
                    aqe_search_mask = pd.Series(False, index=filtered_df_aqe.index)
                    if '이름' in filtered_df_aqe.columns:
                        aqe_search_mask = aqe_search_mask | filtered_df_aqe['이름'].astype(str).str.contains(aqe_filter_search, case=False, na=False)
                    if '사번' in filtered_df_aqe.columns:
                        aqe_search_mask = aqe_search_mask | filtered_df_aqe['사번'].astype(str).str.contains(aqe_filter_search, na=False)
                    filtered_df_aqe = filtered_df_aqe[aqe_search_mask]

                if aqe_filter_grade != "전체" and 'EE/ME/BE' in filtered_df_aqe.columns:
                    filtered_df_aqe = filtered_df_aqe[filtered_df_aqe['EE/ME/BE'].astype(str) == aqe_filter_grade]

                st.caption(f"📋 조회 결과: **{len(filtered_df_aqe)}건** (원본 업로드: {len(df_aqe)}건)")
                st.dataframe(filtered_df_aqe, use_container_width=True)  # 업로드 원본 기반 필터 미리보기
                
                st.write("---")
                st.write("**저장 방식 선택**")
                col_save1, col_save2 = st.columns(2)
                
                with col_save1:
                    if st.button("💾 분기평가 데이터 저장 (test_qae_list.xlsx)", key="save_aqe_file"):
                        try:
                            # DataFrame을 test_qae_list.xlsx로 저장
                            aqe_output_path = "data/test_qae_list.xlsx"
                            with pd.ExcelWriter(aqe_output_path, engine='openpyxl') as writer:
                                df_aqe.to_excel(writer, index=False, sheet_name='분기평가')
                            st.success(f"✅ 저장 완료: {len(df_aqe)}건의 분기평가 데이터가 저장되었습니다")
                            # 캐시 초기화 (새로운 파일을 다시 로드하기 위해)
                            load_aqe_data.clear()
                        except Exception as e:
                            st.error(f"❌ 저장 실패: {e}")
                
                with col_save2:
                    csv_data = df_aqe.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        "📥 CSV로 다운로드",
                        data=csv_data,
                        file_name=f"aqe_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv",
                        key="download_aqe_csv"
                    )
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")

        st.write("---")
        st.write("**📋 분기 평가 이력 표 (필터 적용)**")

        aqe_history_columns = [
            '사번', '이름', '영어이름', '국적', '부/과', '직/반', '직종', '입사일',
            '평가시기', '평가일', '반장 기능(70)', '반장 역량(30)', '직장 기능(70)', '직장 역량(30)',
            '과장 기능(70)', '과장 역량(30)', '부장 기능(70)', '부장 역량(30)',
            'EE/ME/BE', '점수합계', '기량등급'
        ]
        aqe_history_df = load_aqe_data()

        if aqe_history_df is not None and not aqe_history_df.empty:
            aqe_history_df = aqe_history_df.copy()
            for column_name in aqe_history_columns:
                if column_name not in aqe_history_df.columns:
                    aqe_history_df[column_name] = ''
            aqe_history_df = aqe_history_df[aqe_history_columns]

            history_filter_col1, history_filter_col2, history_filter_col3, history_filter_col4, history_filter_col5 = st.columns(5)

            with history_filter_col1:
                history_filter_dept = st.selectbox(
                    "부/과 선택",
                    ["전체"] + sorted(aqe_history_df['부/과'].dropna().astype(str).unique().tolist()),
                    key="eval_history_filter_dept_t1"
                )

            with history_filter_col2:
                if history_filter_dept != "전체":
                    history_available_units = sorted(
                        aqe_history_df[aqe_history_df['부/과'].astype(str) == history_filter_dept]['직/반']
                        .dropna().astype(str).unique().tolist()
                    )
                else:
                    history_available_units = sorted(aqe_history_df['직/반'].dropna().astype(str).unique().tolist())

                history_filter_unit = st.selectbox(
                    "직/반 선택",
                    ["전체"] + history_available_units,
                    key="eval_history_filter_unit_t1"
                )

            with history_filter_col3:
                history_filter_search = st.text_input(
                    "이름/사번 검색",
                    placeholder="검색...",
                    key="eval_history_filter_search_t1"
                )

            with history_filter_col4:
                history_filter_period = st.selectbox(
                    "평가시기",
                    ["전체"] + sorted(aqe_history_df['평가시기'].dropna().astype(str).unique().tolist(), reverse=True),
                    key="eval_history_filter_period_t1"
                )

            with history_filter_col5:
                history_filter_grade = st.selectbox(
                    "EE/ME/BE",
                    ["전체"] + sorted(aqe_history_df['EE/ME/BE'].dropna().astype(str).unique().tolist()),
                    key="eval_history_filter_grade_t1"
                )

            filtered_aqe_history_df = aqe_history_df.copy()

            if history_filter_dept != "전체":
                filtered_aqe_history_df = filtered_aqe_history_df[
                    filtered_aqe_history_df['부/과'].astype(str) == history_filter_dept
                ]

            if history_filter_unit != "전체":
                filtered_aqe_history_df = filtered_aqe_history_df[
                    filtered_aqe_history_df['직/반'].astype(str) == history_filter_unit
                ]

            if history_filter_search:
                filtered_aqe_history_df = filtered_aqe_history_df[
                    filtered_aqe_history_df['이름'].astype(str).str.contains(history_filter_search, case=False, na=False)
                    | filtered_aqe_history_df['사번'].astype(str).str.contains(history_filter_search, na=False)
                ]

            if history_filter_period != "전체":
                filtered_aqe_history_df = filtered_aqe_history_df[
                    filtered_aqe_history_df['평가시기'].astype(str) == history_filter_period
                ]

            if history_filter_grade != "전체":
                filtered_aqe_history_df = filtered_aqe_history_df[
                    filtered_aqe_history_df['EE/ME/BE'].astype(str) == history_filter_grade
                ]

            filtered_aqe_history_df.index = range(1, len(filtered_aqe_history_df) + 1)
            st.caption(f"📋 조회 결과: **{len(filtered_aqe_history_df)}건** (전체: {len(aqe_history_df)}건)")
            st.dataframe(filtered_aqe_history_df, use_container_width=True)

            aqe_history_csv_data = filtered_aqe_history_df.to_csv(index=False).encode('utf-8-sig')
            aqe_history_excel_data = to_excel(filtered_aqe_history_df)

            history_dl_col1, history_dl_col2 = st.columns(2)
            with history_dl_col1:
                st.download_button(
                    "📥 조회결과 다운로드(CSV)",
                    data=aqe_history_csv_data,
                    file_name=f"aqe_history_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    key="eval_history_dl_csv_t1"
                )
            with history_dl_col2:
                st.download_button(
                    "📥 조회결과 다운로드(Excel)",
                    data=aqe_history_excel_data,
                    file_name=f"aqe_history_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="eval_history_dl_excel_t1"
                )
        else:
            st.info("저장된 분기 평가 이력이 없습니다.")
    
    with t2:
        st.subheader("📥 월별 평가 정보 업로드")
        st.write("엑셀 또는 CSV 파일 형식: 사번, 이름, 평가월, 평가점수, 평가등급, 비고")
        
        uploaded_monthly_file = st.file_uploader("월별 평가 엑셀 또는 CSV 파일을 업로드하세요", type=["xlsx", "csv"], key="upload_monthly_excel")
        if uploaded_monthly_file:
            try:
                df_monthly = read_uploaded_table(uploaded_monthly_file)
                df_monthly = clean_excel_data(df_monthly)
                
                required_cols = ['사번', '이름', '평가월', '평가점수', '평가등급']
                missing_cols = [col for col in required_cols if col not in df_monthly.columns]
                
                if missing_cols:
                    st.warning(f"⚠️ 다음 필수 컬럼이 누락되었습니다: {', '.join(missing_cols)}")
                
                df_monthly.index = range(1, len(df_monthly) + 1)
                
                st.write("**📋 업로드 데이터 미리보기 (필터 적용)**")
                monthly_filter_col1, monthly_filter_col2, monthly_filter_col3 = st.columns(3)

                with monthly_filter_col1:
                    monthly_filter_grade = st.selectbox(
                        "평가등급",
                        ["전체"] + sorted(df_monthly['평가등급'].dropna().astype(str).unique().tolist()) if '평가등급' in df_monthly.columns else ["전체"],
                        key="eval_upload_monthly_filter_grade"
                    )

                with monthly_filter_col2:
                    monthly_filter_month = st.selectbox(
                        "평가월",
                        ["전체"] + sorted(df_monthly['평가월'].dropna().astype(str).unique().tolist()) if '평가월' in df_monthly.columns else ["전체"],
                        key="eval_upload_monthly_filter_month"
                    )

                with monthly_filter_col3:
                    monthly_filter_search = st.text_input(
                        "이름/사번 검색",
                        placeholder="검색...",
                        key="eval_upload_monthly_filter_search"
                    )

                filtered_df_monthly = df_monthly.copy()

                if monthly_filter_grade != "전체" and '평가등급' in filtered_df_monthly.columns:
                    filtered_df_monthly = filtered_df_monthly[filtered_df_monthly['평가등급'].astype(str) == monthly_filter_grade]

                if monthly_filter_month != "전체" and '평가월' in filtered_df_monthly.columns:
                    filtered_df_monthly = filtered_df_monthly[filtered_df_monthly['평가월'].astype(str) == monthly_filter_month]

                if monthly_filter_search:
                    monthly_search_mask = pd.Series(False, index=filtered_df_monthly.index)
                    if '이름' in filtered_df_monthly.columns:
                        monthly_search_mask = monthly_search_mask | filtered_df_monthly['이름'].astype(str).str.contains(monthly_filter_search, case=False, na=False)
                    if '사번' in filtered_df_monthly.columns:
                        monthly_search_mask = monthly_search_mask | filtered_df_monthly['사번'].astype(str).str.contains(monthly_filter_search, na=False)
                    filtered_df_monthly = filtered_df_monthly[monthly_search_mask]

                st.caption(f"📋 조회 결과: **{len(filtered_df_monthly)}건** (원본 업로드: {len(df_monthly)}건)")
                st.dataframe(filtered_df_monthly, use_container_width=True)
                
                st.write("---")
                col_m1, col_m2 = st.columns(2)
                
                with col_m1:
                    if st.button("💾 월별평가 데이터 저장", key="save_monthly_file"):
                        try:
                            monthly_output_path = "data/monthly_evaluation.xlsx"
                            with pd.ExcelWriter(monthly_output_path, engine='openpyxl') as writer:
                                df_monthly.to_excel(writer, index=False, sheet_name='월별평가')
                            st.success(f"✅ 저장 완료: {len(df_monthly)}건의 월별평가 데이터가 저장되었습니다")
                        except Exception as e:
                            st.error(f"❌ 저장 실패: {e}")
                
                with col_m2:
                    csv_data_m = df_monthly.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        "📥 CSV로 다운로드",
                        data=csv_data_m,
                        file_name=f"monthly_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv",
                        key="download_monthly_csv"
                    )
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")
    


# ══════════════════════════════════════════════════════════════════
# 메뉴 5: 안전 퀴즈
# ══════════════════════════════════════════════════════════════════
elif menu == "🛡️ 안전 퀴즈":
    st.header("🛡️ 안전 퀴즈")
    st.write("퀴즈 파일을 실행하거나 배포용 zip 파일을 다운로드할 수 있습니다.")

    quiz_html_path = os.path.join("quiz_web", "index.html")
    quiz_zip_path = "안전용어퀴즈.zip"

    if os.path.exists(quiz_html_path):
        html_abs_path = os.path.abspath(quiz_html_path).replace('\\', '/')
        quiz_file_url = f"file:///{html_abs_path}"
        if st.button("▶ 안전 퀴즈 열기", use_container_width=True, key="open_safety_quiz_local"):
            opened = webbrowser.open_new_tab(quiz_file_url)
            if opened:
                st.success("브라우저에서 안전 퀴즈를 열었습니다.")
            else:
                st.warning("브라우저 자동 실행에 실패했습니다. 아래 경로를 복사해 직접 열어주세요.")

        st.caption("버튼이 동작하지 않으면 아래 경로를 파일 탐색기에 붙여넣어 실행하세요.")
        st.text(os.path.abspath(quiz_html_path))
    else:
        st.warning("quiz_web/index.html 파일이 없습니다. 먼저 global_kw.py를 실행해 퀴즈를 생성해주세요.")

    if os.path.exists(quiz_zip_path):
        with open(quiz_zip_path, "rb") as f:
            st.download_button(
                "📦 안전 퀴즈 zip 다운로드",
                data=f,
                file_name="안전용어퀴즈.zip",
                mime="application/zip",
                use_container_width=True,
                key="download_safety_quiz_zip"
            )
    else:
        st.info("안전용어퀴즈.zip 파일이 없습니다. 터미널에서 압축 생성 후 다운로드할 수 있습니다.")



