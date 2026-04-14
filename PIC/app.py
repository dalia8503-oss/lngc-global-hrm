from flask import (
    Flask, request, jsonify, send_from_directory,
    redirect, url_for, session, Response
)
import os, csv, io, psycopg2, psycopg2.extras
from datetime import datetime
from functools import wraps

app = Flask(__name__, static_folder='.')
app.secret_key = os.environ.get('SECRET_KEY', 'change-this-in-production')

ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin1234')
DATABASE_URL    = os.environ.get('DATABASE_URL')


# ── DB 연결 ────────────────────────────────────────────────
def get_conn():
    return psycopg2.connect(DATABASE_URL)


# ── DB 초기화 ──────────────────────────────────────────────
def init_db():
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('''
                CREATE TABLE IF NOT EXISTS submissions (
                    id           SERIAL PRIMARY KEY,
                    submitted_at TEXT,
                    hoseon       TEXT,
                    job          TEXT,
                    tk           TEXT,
                    name         TEXT
                )
            ''')

init_db()


# ── 관리자 인증 데코레이터 ─────────────────────────────────
def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('admin'):
            return redirect(url_for('admin_login'))
        return f(*args, **kwargs)
    return decorated


# ── 폼 서빙 ────────────────────────────────────────────────
@app.route('/')
def index():
    return send_from_directory('.', 'hoseon_input_custom.html')


# ── 폼 제출 → DB 저장 ─────────────────────────────────────
@app.route('/submit', methods=['POST'])
def submit():
    body   = request.get_json(force=True)
    hoseon = body.get('hoseon', '').strip()
    jobs   = body.get('jobs', {})
    now    = datetime.now().strftime('%Y-%m-%d %H:%M')

    rows = [
        (now, hoseon, job, tk, name)
        for job, tks in jobs.items()
        for tk, name in tks.items()
        if name
    ]

    if not rows:
        return jsonify({'ok': False, 'error': '입력 데이터 없음'}), 400

    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('DELETE FROM submissions WHERE hoseon = %s', (hoseon,))
            psycopg2.extras.execute_values(
                cur,
                'INSERT INTO submissions (submitted_at, hoseon, job, tk, name) VALUES %s',
                rows
            )

    return jsonify({'ok': True})


# ── 관리자 로그인 ──────────────────────────────────────────
@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    error = ''
    if request.method == 'POST':
        if request.form.get('password') == ADMIN_PASSWORD:
            session['admin'] = True
            return redirect(url_for('admin'))
        error = '비밀번호가 틀렸습니다'

    return f'''<!DOCTYPE html><html lang="ko">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>관리자 로그인</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:-apple-system,sans-serif;background:#f5f4f0;display:flex;
        align-items:center;justify-content:center;min-height:100vh}}
  .box{{background:#fff;padding:2rem;border-radius:16px;width:300px;
        box-shadow:0 2px 12px rgba(0,0,0,.08)}}
  h2{{font-size:18px;margin-bottom:1.5rem}}
  input{{width:100%;height:44px;padding:0 12px;border:1.5px solid #ddd;
         border-radius:10px;font-size:16px;margin-bottom:10px}}
  button{{width:100%;height:44px;background:#1a1a1a;color:#fff;border:none;
          border-radius:10px;font-size:15px;font-weight:600;cursor:pointer}}
  .err{{color:#e53935;font-size:13px;margin-bottom:10px}}
</style></head>
<body><div class="box">
  <h2>관리자 로그인</h2>
  <form method="post">
    {'<p class="err">'+error+'</p>' if error else ''}
    <input type="password" name="password" placeholder="비밀번호" autofocus>
    <button type="submit">로그인</button>
  </form>
</div></body></html>'''


# ── 관리자 메인 페이지 ────────────────────────────────────
@app.route('/admin')
@admin_required
def admin():
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                'SELECT submitted_at, hoseon, job, tk, name FROM submissions ORDER BY submitted_at DESC, hoseon'
            )
            rows = cur.fetchall()

    trs = ''.join(
        f'<tr><td>{r[0]}</td><td>{r[1]}</td><td>{r[2]}</td><td>{r[3]}</td><td>{r[4]}</td></tr>'
        for r in rows
    ) or '<tr><td colspan="5" class="empty">제출된 데이터가 없습니다</td></tr>'

    return f'''<!DOCTYPE html><html lang="ko">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>담당정보 취합</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:-apple-system,sans-serif;background:#f5f4f0}}
  .header{{background:#1a1a1a;color:#fff;padding:1rem 1.5rem;
            display:flex;align-items:center;justify-content:space-between}}
  .header h1{{font-size:17px}}
  .header a{{color:#aaa;font-size:13px;text-decoration:none}}
  .actions{{padding:1rem 1.5rem;display:flex;gap:10px;flex-wrap:wrap;align-items:center}}
  .btn{{height:40px;padding:0 20px;border-radius:10px;font-size:14px;
         font-weight:600;cursor:pointer;border:none;text-decoration:none;
         display:inline-flex;align-items:center}}
  .btn-dark{{background:#1a1a1a;color:#fff}}
  .btn-danger{{background:transparent;color:#e53935;border:1.5px solid #e53935}}
  .count{{padding:0 1.5rem .75rem;font-size:13px;color:#888}}
  .wrap{{overflow-x:auto;padding:0 1.5rem 2rem}}
  table{{width:100%;border-collapse:collapse;background:#fff;
          border-radius:12px;overflow:hidden;font-size:14px}}
  th{{background:#f0f0f0;padding:10px 14px;text-align:left;
       font-size:12px;color:#666;white-space:nowrap}}
  td{{padding:10px 14px;border-top:1px solid #f0f0f0}}
  tr:hover td{{background:#fafafa}}
  .empty{{text-align:center;padding:3rem;color:#bbb}}
</style></head>
<body>
<div class="header">
  <h1>담당정보 취합 현황</h1>
  <a href="/admin/logout">로그아웃</a>
</div>
<div class="actions">
  <a href="/admin/download" class="btn btn-dark">CSV 다운로드</a>
  <form method="post" action="/admin/clear"
        onsubmit="return confirm('전체 데이터를 삭제할까요?')">
    <button class="btn btn-danger" type="submit">데이터 초기화</button>
  </form>
</div>
<p class="count">총 {len(rows)}건</p>
<div class="wrap">
  <table>
    <thead><tr><th>제출시각</th><th>호선</th><th>직종</th><th>TK</th><th>담당자</th></tr></thead>
    <tbody>{trs}</tbody>
  </table>
</div>
</body></html>'''


# ── CSV 다운로드 ───────────────────────────────────────────
@app.route('/admin/download')
@admin_required
def admin_download():
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                'SELECT submitted_at, hoseon, job, tk, name FROM submissions ORDER BY submitted_at, hoseon'
            )
            rows = cur.fetchall()

    buf = io.StringIO()
    w   = csv.writer(buf)
    w.writerow(['제출시각', '호선', '직종', 'TK', '담당자'])
    w.writerows(rows)

    today = datetime.now().strftime('%Y%m%d')
    data  = ('\uFEFF' + buf.getvalue()).encode('utf-8')

    return Response(
        data,
        mimetype='text/csv',
        headers={'Content-Disposition': f'attachment; filename=담당정보_{today}.csv'}
    )


# ── 데이터 초기화 ──────────────────────────────────────────
@app.route('/admin/clear', methods=['POST'])
@admin_required
def admin_clear():
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute('DELETE FROM submissions')
    return redirect(url_for('admin'))


# ── 로그아웃 ───────────────────────────────────────────────
@app.route('/admin/logout')
def admin_logout():
    session.pop('admin', None)
    return redirect(url_for('admin_login'))


if __name__ == '__main__':
    app.run(debug=True)
