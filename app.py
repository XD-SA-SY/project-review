import os
import re
import sqlite3
import uuid
from datetime import datetime
from flask import Flask, request, jsonify, render_template, send_from_directory
from werkzeug.utils import secure_filename
from openpyxl import Workbook

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024
app.secret_key = 'project_review_secret_key_2026'

# 自动创建目录
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('exports', exist_ok=True)

# 登录凭证存储
token_store = {}


# ===================== 数据库工具 =====================
def get_db():
    conn = sqlite3.connect('project_review.db')
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()

    # 用户表
    c.execute('''
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        name TEXT NOT NULL,
        role TEXT NOT NULL,
        class_name TEXT,
        student_no TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    # 项目作品表
    c.execute('''
    CREATE TABLE IF NOT EXISTS projects (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER NOT NULL,
        project_name TEXT NOT NULL,
        file_path TEXT NOT NULL,
        description TEXT,
        status INTEGER DEFAULT 1,
        submit_time TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (student_id) REFERENCES users(id)
    )
    ''')

    # 学生互评表
    c.execute('''
    CREATE TABLE IF NOT EXISTS student_reviews (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        reviewer_id INTEGER NOT NULL,
        project_id INTEGER NOT NULL,
        score REAL NOT NULL,
        comment TEXT,
        review_time TEXT DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(reviewer_id, project_id)
    )
    ''')

    # 教师评分表
    c.execute('''
    CREATE TABLE IF NOT EXISTS teacher_scores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        teacher_id INTEGER NOT NULL,
        project_id INTEGER NOT NULL,
        score REAL NOT NULL,
        comment TEXT,
        score_time TEXT DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(teacher_id, project_id)
    )
    ''')

    # 导师评分表
    c.execute('''
    CREATE TABLE IF NOT EXISTS mentor_scores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        mentor_id INTEGER NOT NULL,
        project_id INTEGER NOT NULL,
        score REAL NOT NULL,
        comment TEXT,
        score_time TEXT DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(mentor_id, project_id)
    )
    ''')

    # 最终成绩表
    c.execute('''
    CREATE TABLE IF NOT EXISTS final_scores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_id INTEGER NOT NULL UNIQUE,
        student_id INTEGER NOT NULL,
        student_avg_score REAL,
        teacher_score REAL,
        mentor_score REAL,
        total_score REAL,
        calculate_time TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (project_id) REFERENCES projects(id)
    )
    ''')

    # 初始化测试账号
    test_users = [
        ('stu001', '123456', '张三', 'student', '计科2班', '2026001'),
        ('stu002', '123456', '李四', 'student', '计科2班', '2026002'),
        ('tea001', '123456', '王老师', 'teacher', '计科2班', None),
        ('men001', '123456', '李导师', 'mentor', None, None),
        ('admin', '123456', '管理员', 'admin', None, None)
    ]
    for user in test_users:
        c.execute('SELECT id FROM users WHERE username = ?', (user[0],))
        if not c.fetchone():
            c.execute('INSERT INTO users (username, password, name, role, class_name, student_no) VALUES (?,?,?,?,?,?)', user)

    conn.commit()
    conn.close()


# ===================== 权限校验 =====================
def require_role(roles):
    def decorator(f):
        def wrapper(*args, **kwargs):
            token = request.headers.get('Authorization', '').replace('Bearer ', '')
            if not token:
                token = request.args.get('token', '')
            
            user = token_store.get(token)
            if not user:
                return jsonify({'msg': '未登录或登录失效'}), 401
            if roles and user['role'] not in roles:
                return jsonify({'msg': '权限不足'}), 403
            
            request.user = user
            return f(*args, **kwargs)
        wrapper.__name__ = f.__name__
        return wrapper
    return decorator


# ===================== 工具函数 =====================
def anonymize_student(row, is_self=False):
    data = dict(row)
    if not is_self and data.get('role') == 'student':
        data['student_name'] = f'匿名学生 #{data["id"]}'
        data['class_name'] = '未公开'
        data['student_no'] = '未公开'
        data.pop('name', None)
    return data


def calculate_final_score(project_id):
    conn = get_db()
    c = conn.cursor()

    c.execute('SELECT AVG(score) as avg FROM student_reviews WHERE project_id = ?', (project_id,))
    student_avg = c.fetchone()['avg'] or 0

    c.execute('SELECT score FROM teacher_scores WHERE project_id = ? LIMIT 1', (project_id,))
    teacher_row = c.fetchone()
    teacher_score = teacher_row['score'] if teacher_row else 0

    c.execute('SELECT score FROM mentor_scores WHERE project_id = ? LIMIT 1', (project_id,))
    mentor_row = c.fetchone()
    mentor_score = mentor_row['score'] if mentor_row else 0

    total = student_avg * 0.2 + teacher_score * 0.5 + mentor_score * 0.3

    c.execute('SELECT student_id FROM projects WHERE id = ?', (project_id,))
    student_id = c.fetchone()['student_id']

    c.execute('SELECT id FROM final_scores WHERE project_id = ?', (project_id,))
    if c.fetchone():
        c.execute('''
        UPDATE final_scores SET 
            student_avg_score = ?, teacher_score = ?, mentor_score = ?,
            total_score = ?, calculate_time = ?
        WHERE project_id = ?
        ''', (student_avg, teacher_score, mentor_score, total, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), project_id))
    else:
        c.execute('''
        INSERT INTO final_scores 
        (project_id, student_id, student_avg_score, teacher_score, mentor_score, total_score)
        VALUES (?,?,?,?,?,?)
        ''', (project_id, student_id, student_avg, teacher_score, mentor_score, total))

    if student_avg > 0 and teacher_score > 0 and mentor_score > 0:
        c.execute('UPDATE projects SET status = 2 WHERE id = ?', (project_id,))

    conn.commit()
    conn.close()


# ===================== 接口路由 =====================

# 首页
@app.route('/')
def index():
    return render_template('index.html')

# 1. 登录
@app.route('/api/login', methods=['POST'])
def login():
    data = request.get_json()
    username = data.get('username', '').strip()
    password = data.get('password', '').strip()

    conn = get_db()
    user = conn.execute('SELECT * FROM users WHERE username = ?', (username,)).fetchone()
    conn.close()

    if not user or user['password'] != password:
        return jsonify({'msg': '账号或密码错误'}), 400

    token = str(uuid.uuid4())
    user_info = {
        'id': user['id'],
        'name': user['name'],
        'role': user['role'],
        'className': user['class_name']
    }
    token_store[token] = user_info
    return jsonify({'token': token, 'user': user_info})

# 2. 学生注册
@app.route('/api/register', methods=['POST'])
def register():
    data = request.get_json()
    username = data.get('username', '').strip()
    password = data.get('password', '').strip()
    name = data.get('name', '').strip()
    class_name = data.get('className', '').strip()
    student_no = data.get('studentNo', '').strip()

    if not all([username, password, name, class_name, student_no]):
        return jsonify({'msg': '请填写完整信息'}), 400
    if len(password) < 6:
        return jsonify({'msg': '密码长度不能少于6位'}), 400

    conn = get_db()
    exist = conn.execute('SELECT id FROM users WHERE username = ?', (username,)).fetchone()
    if exist:
        conn.close()
        return jsonify({'msg': '该账号已被注册'}), 400

    try:
        conn.execute('''
        INSERT INTO users (username, password, name, role, class_name, student_no)
        VALUES (?,?,?,?,?,?)
        ''', (username, password, name, 'student', class_name, student_no))
        conn.commit()
        conn.close()
        return jsonify({'msg': '注册成功，请登录'})
    except Exception as e:
        conn.close()
        return jsonify({'msg': '注册失败：' + str(e)}), 500

# 3. 提交项目成果（支持Word+PDF）
@app.route('/api/project/upload', methods=['POST'])
@require_role(['student'])
def upload_project():
    project_name = request.form.get('projectName', '').strip()
    description = request.form.get('description', '')
    file = request.files.get('file')

    if not project_name or not file:
        return jsonify({'msg': '请填写项目名称并上传文档'}), 400

    allowed_exts = ['.doc', '.docx', '.pdf']
    file_ext = os.path.splitext(file.filename)[1].lower()
    if file_ext not in allowed_exts:
        return jsonify({'msg': '仅支持 .doc / .docx / .pdf 格式'}), 400

    filename = secure_filename(f"{int(datetime.now().timestamp())}_{file.filename}")
    save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(save_path)

    conn = get_db()
    conn.execute('''
    INSERT INTO projects (student_id, project_name, file_path, description)
    VALUES (?,?,?,?)
    ''', (request.user['id'], project_name, f'/uploads/{filename}', description))
    conn.commit()
    conn.close()
    return jsonify({'msg': '项目成果提交成功'})

# 4. 获取匿名互评列表
@app.route('/api/projects/review')
@require_role(['student'])
def get_review_projects():
    student_id = request.user['id']
    conn = get_db()
    projects = conn.execute('''
    SELECT p.id, p.project_name, p.description, p.submit_time
    FROM projects p
    WHERE p.student_id != ?
    AND p.id NOT IN (SELECT project_id FROM student_reviews WHERE reviewer_id = ?)
    ''', (student_id, student_id)).fetchall()
    conn.close()
    return jsonify([dict(p) for p in projects])

# 5. 提交互评
@app.route('/api/project/review', methods=['POST'])
@require_role(['student'])
def submit_review():
    data = request.get_json()
    project_id = data.get('projectId')
    score = data.get('score')
    comment = data.get('comment', '')

    if not project_id or not score:
        return jsonify({'msg': '参数不完整'}), 400

    conn = get_db()
    try:
        conn.execute('''
        INSERT INTO student_reviews (reviewer_id, project_id, score, comment)
        VALUES (?,?,?,?)
        ''', (request.user['id'], project_id, score, comment))
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({'msg': '你已评过该项目'}), 400

    conn.close()
    calculate_final_score(project_id)
    return jsonify({'msg': '互评提交成功'})

# 6. 教师/导师获取待评审项目
@app.route('/api/projects/teacher')
@require_role(['teacher', 'mentor'])
def get_teacher_projects():
    conn = get_db()
    projects = conn.execute('''
    SELECT p.id, p.project_name, p.description, p.submit_time, 
           u.id as student_id, u.name as student_name, u.class_name
    FROM projects p
    JOIN users u ON p.student_id = u.id
    WHERE p.status = 1
    ''').fetchall()
    conn.close()
    return jsonify([anonymize_student(p) for p in projects])

# 7. 获取项目详情（含文件信息）
@app.route('/api/project/<int:project_id>/detail')
@require_role(['teacher', 'mentor'])
def get_project_detail(project_id):
    conn = get_db()
    project = conn.execute('''
    SELECT p.id, p.project_name, p.description, p.file_path, p.submit_time,
           u.id as student_id, u.name as student_name, u.class_name
    FROM projects p
    JOIN users u ON p.student_id = u.id
    WHERE p.id = ?
    ''', (project_id,)).fetchone()
    conn.close()

    if not project:
        return jsonify({'msg': '项目不存在'}), 404

    result = anonymize_student(project)
    if result.get('file_path'):
        result['file_name'] = os.path.basename(result['file_path'])
    else:
        result['file_name'] = None
    return jsonify(result)

# 8. 下载项目文件
@app.route('/api/project/<int:project_id>/download')
@require_role(['teacher', 'mentor'])
def download_project_file(project_id):
    conn = get_db()
    project = conn.execute('SELECT file_path, project_name FROM projects WHERE id = ?', (project_id,)).fetchone()
    conn.close()

    if not project or not project['file_path']:
        return jsonify({'msg': '该项目未上传文件'}), 404

    file_path = project['file_path'].lstrip('/')
    full_path = os.path.join(os.getcwd(), file_path)

    if not os.path.exists(full_path):
        return jsonify({'msg': '文件不存在'}), 404

    filename = os.path.basename(full_path)
    return send_from_directory(os.path.dirname(full_path), filename, as_attachment=True)

# 9. 教师提交评分
@app.route('/api/score/teacher', methods=['POST'])
@require_role(['teacher'])
def teacher_score():
    data = request.get_json()
    project_id = data.get('projectId')
    score = data.get('score')
    comment = data.get('comment', '')

    conn = get_db()
    try:
        conn.execute('''
        INSERT INTO teacher_scores (teacher_id, project_id, score, comment)
        VALUES (?,?,?,?)
        ''', (request.user['id'], project_id, score, comment))
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({'msg': '已评过该项目'}), 400

    conn.close()
    calculate_final_score(project_id)
    return jsonify({'msg': '评分提交成功'})

# 10. 导师提交评分
@app.route('/api/score/mentor', methods=['POST'])
@require_role(['mentor'])
def mentor_score():
    data = request.get_json()
    project_id = data.get('projectId')
    score = data.get('score')
    comment = data.get('comment', '')

    conn = get_db()
    try:
        conn.execute('''
        INSERT INTO mentor_scores (mentor_id, project_id, score, comment)
        VALUES (?,?,?,?)
        ''', (request.user['id'], project_id, score, comment))
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({'msg': '已评过该项目'}), 400

    conn.close()
    calculate_final_score(project_id)
    return jsonify({'msg': '评分提交成功，综合成绩已自动生成'})

# 11. 成绩列表
@app.route('/api/scores/list')
@require_role(['teacher', 'admin'])
def get_scores_list():
    conn = get_db()
    scores = conn.execute('''
    SELECT f.id, u.id as student_id, u.name as student_name, u.class_name, 
           p.project_name, f.student_avg_score, f.teacher_score, 
           f.mentor_score, f.total_score, f.calculate_time
    FROM final_scores f
    JOIN users u ON f.student_id = u.id
    JOIN projects p ON f.project_id = p.id
    ORDER BY f.total_score DESC
    ''').fetchall()
    conn.close()
    
    result = []
    for s in scores:
        is_admin = request.user['role'] == 'admin'
        result.append(anonymize_student(s, is_self=is_admin))
    return jsonify(result)

# 12. 导出Excel成绩表
@app.route('/api/scores/export')
@require_role(['teacher', 'admin'])
def export_scores():
    conn = get_db()
    scores = conn.execute('''
    SELECT u.name as 学生姓名, u.student_no as 学号, u.class_name as 班级, 
           p.project_name as 项目名称, f.student_avg_score as 学生互评,
           f.teacher_score as 教师评分, f.mentor_score as 导师评分,
           f.total_score as 综合得分, f.calculate_time as 计算时间
    FROM final_scores f
    JOIN users u ON f.student_id = u.id
    JOIN projects p ON f.project_id = p.id
    ORDER BY f.total_score DESC
    ''').fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "班级综合成绩"

    headers = ["学生姓名", "学号", "班级", "项目名称", "学生互评(20%)", "教师评分(50%)", "导师评分(30%)", "综合得分", "计算时间"]
    ws.append(headers)

    for row in scores:
        ws.append(list(row))

    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        ws.column_dimensions[column].width = max_length + 2

    excel_path = 'exports/班级综合成绩.xlsx'
    wb.save(excel_path)
    return send_from_directory('exports', '班级综合成绩.xlsx', as_attachment=True)

# 13. 学生查看个人成绩
@app.route('/api/my/score')
@require_role(['student'])
def get_my_score():
    conn = get_db()
    scores = conn.execute('''
    SELECT p.project_name, f.student_avg_score, f.teacher_score, 
           f.mentor_score, f.total_score
    FROM final_scores f
    JOIN projects p ON f.project_id = p.id
    WHERE f.student_id = ?
    ''', (request.user['id'],)).fetchall()
    conn.close()
    return jsonify([dict(s) for s in scores])

# 14. 管理员：获取所有用户
@app.route('/api/admin/users')
@require_role(['admin'])
def get_all_users():
    conn = get_db()
    users = conn.execute('SELECT id, username, name, role, class_name, student_no, created_at FROM users ORDER BY id').fetchall()
    conn.close()
    return jsonify([dict(u) for u in users])

# 15. 管理员：创建账号
@app.route('/api/admin/create_user', methods=['POST'])
@require_role(['admin'])
def create_user():
    data = request.get_json()
    username = data.get('username', '').strip()
    password = data.get('password', '').strip()
    name = data.get('name', '').strip()
    role = data.get('role', '').strip()
    class_name = data.get('className', '').strip()

    if not all([username, password, name, role]):
        return jsonify({'msg': '请填写完整信息'}), 400
    if role not in ['teacher', 'mentor', 'admin']:
        return jsonify({'msg': '只能创建教师、导师、管理员账号'}), 400

    conn = get_db()
    exist = conn.execute('SELECT id FROM users WHERE username = ?', (username,)).fetchone()
    if exist:
        conn.close()
        return jsonify({'msg': '账号已存在'}), 400

    try:
        conn.execute('''
        INSERT INTO users (username, password, name, role, class_name)
        VALUES (?,?,?,?,?)
        ''', (username, password, name, role, class_name))
        conn.commit()
        conn.close()
        return jsonify({'msg': '账号创建成功'})
    except Exception as e:
        conn.close()
        return jsonify({'msg': '创建失败：' + str(e)}), 500

# 访问上传文件
@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)


if __name__ == '__main__':
    init_db()
    port = int(os.environ.get('PORT', 5000))
    print("="*50)
    print("课程项目互评平台 启动成功")
    print(f"本机访问：http://localhost:{port}")
    print("局域网访问：http://[你的IP]:5000")
    print("="*50)
    app.run(host='0.0.0.0', port=port, debug=False)
