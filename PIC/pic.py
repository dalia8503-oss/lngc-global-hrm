# =============================================
#  호선별 담당정보 입력 폼 - 설정 커스터마이징
#  이 파일과 hoseon_input.html을 같은 폴더에 두고 실행하세요
# =============================================

# ✅ 여기만 수정하세요 -----------------------

# 앱 제목 / 부제목
TITLE    = "호선별 담당정보 입력"
SUBTITLE = "직종 선택 후 TK별 담당자를 입력하세요"

# 직종 목록 (개수 자유롭게 추가/삭제 가능)
JOBS = [
    'IP',
    'FSB(ABM/Manual)',
    'TBP',
    'MB(설치/배재/마킹)',
    'MB(수동)',
    'MB(자동)',
    'MB(리웰딩)',
    'L/QC',
    '족장',
]

# TK 목록 (4→3→2→1 순서로 입력)
TKS = ['4TK', '3TK', '2TK', '1TK']

# 담당자 입력 수
MAX_PERSONS = 1

# 직종별 담당자 버튼 목록 (기본값)
NAMES_BY_JOB = {
    'IP':              ['구일', '송안', '한국', '화진'],
    'FSB(ABM/Manual)': ['구일', '송안', '한국', '화진', '오셔나즈', '1부2과', '1부4과', '1부5과(E7)'],
    'TBP':             ['구일', '송안', '한국', '화진', '오셔나즈', '1부5과(E7)'],
    'MB(설치/배재/마킹)':  ['구일', '다온', '엠알', '한국', '2부1과', '2부2과', '2부3과(E7)'],
    'MB(수동)':         ['구일', '다온', '엠알', '한국', '2부1과', '2부2과', '2부3과(E7)'],
    'MB(자동)':         ['구일', '다온', '엠알', '한국', '2부3과(E7)'],
    'MB(리웰딩)':        ['구일', '한국', '2부1과', '2부2과', '2부3과(E7)'],
    'L/QC':            [],   # 호선별로 지정 (아래 NAMES_BY_HOSEON_JOB 참고)
    '족장':             [],   # 호선별로 지정 (아래 NAMES_BY_HOSEON_JOB 참고)
}

# 호선별 직종별 담당자 버튼 목록 (여기 지정된 호선+직종은 기본값 대신 이 목록 사용)
NAMES_BY_HOSEON_JOB = {
    '2665': {
        'L/QC': ['김남균', '신효진', '서정민', '강동국', '김택준', '원호정', '문형진', '최세훈'],
        '족장': ['광진', '하성'],
    },
}

# 1회 선택으로 전체 TK에 동일 적용되는 직종
SINGLE_ENTRY_JOBS = ['L/QC', '족장']

# 출력 파일 이름
OUTPUT_FILE = "hoseon_input_custom.html"

# --------------------------------------------


import re, os, json

INPUT_FILE = "hoseon_input.html"

def build_job_tabs(jobs):
    return "\n".join(
        f'    <button class="job-tab" onclick="selectJob(\'{j}\')">{j}</button>'
        for j in jobs
    )

def build_person_rows(max_persons):
    rows = []
    for i in range(max_persons):
        rows.append(
            f'        <div class="name-row">'
            f'<input type="text" list="names" data-tk="${{tk}}" data-idx="{i}" placeholder="담당자" value="${{escHtml(vals[{i}])}}"></div>'
        )
    return "\n".join(rows)

def main():
    if not os.path.exists(INPUT_FILE):
        print(f"❌ '{INPUT_FILE}' 파일을 찾을 수 없어요.")
        print(f"   이 스크립트와 같은 폴더에 '{INPUT_FILE}'을 넣어주세요.")
        return

    with open(INPUT_FILE, encoding="utf-8") as f:
        html = f.read()

    # 1) 제목 / 부제목
    html = re.sub(
        r'(<h1>).*?(</h1>)',
        rf'\g<1>{TITLE}\g<2>',
        html
    )
    html = re.sub(
        r'(<title>).*?(</title>)',
        rf'\g<1>{TITLE}\g<2>',
        html
    )
    html = re.sub(
        r'(<p>)(직종 선택 후.*?)(</p>)',
        rf'\g<1>{SUBTITLE}\g<3>',
        html
    )

    # 2) 직종 탭 버튼 교체
    tab_html = build_job_tabs(JOBS)
    html = re.sub(
        r'(<div class="job-tabs" id="jobTabs">).*?(</div>)',
        rf'\1\n{tab_html}\n  \2',
        html,
        flags=re.DOTALL
    )

    # 3) JS: jobs 배열
    jobs_js = ", ".join(f"'{j}'" for j in JOBS)
    html = re.sub(
        r"const jobs = \[.*?\];",
        f"const jobs = [{jobs_js}];",
        html
    )

    # 4) JS: tks 배열
    tks_js = ", ".join(f"'{t}'" for t in TKS)
    html = re.sub(
        r"const tks = \[.*?\];",
        f"const tks = [{tks_js}];",
        html
    )

    # 5) 담당자 수: data 초기화 배열 크기
    empty = ", ".join(["''"] * MAX_PERSONS)
    html = re.sub(
        r"data\[j\]\[tk\] = \[.*?\];",
        f"data[j][tk] = [{empty}];",
        html
    )

    # 6) JS namesByJob 객체 주입
    html = html.replace('/* NAMES_BY_JOB */', json.dumps(NAMES_BY_JOB, ensure_ascii=False))

    # 6c) JS singleEntryJobs 배열 주입
    html = html.replace('/* SINGLE_ENTRY_JOBS */', json.dumps(SINGLE_ENTRY_JOBS, ensure_ascii=False))

    # 6b) JS namesByHoseonJob 객체 주입
    html = html.replace('/* NAMES_BY_HOSEON_JOB */', json.dumps(NAMES_BY_HOSEON_JOB, ensure_ascii=False))

    # 7) 자동완성 datalist 주입 (전체 이름 합집합)
    all_names_hoseon = {n for h in NAMES_BY_HOSEON_JOB.values() for names in h.values() for n in names}
    all_names = sorted({n for names in NAMES_BY_JOB.values() for n in names} | all_names_hoseon)
    options = "\n".join(f'  <option value="{n}">' for n in all_names)
    html = re.sub(
        r'<datalist id="names">.*?</datalist>',
        f'<datalist id="names">\n{options}\n</datalist>',
        html,
        flags=re.DOTALL
    )

    # 8) 담당자 입력 rows (renderTKs 함수 내부 name-inputs 블록 — 호환용)
    person_rows = build_person_rows(MAX_PERSONS)
    html = re.sub(
        r'(<div class="name-inputs">).*?(</div>\s*</div>\s*`;)',
        rf'\1\n{person_rows}\n      \2',
        html,
        flags=re.DOTALL
    )

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅ '{OUTPUT_FILE}' 생성 완료!")
    print(f"   직종: {JOBS}")
    print(f"   TK  : {TKS}")
    print(f"   담당자 최대 {MAX_PERSONS}명")
    print(f"   직종별 버튼: { {j: len(v) for j, v in NAMES_BY_JOB.items()} }")

if __name__ == "__main__":
    main()
