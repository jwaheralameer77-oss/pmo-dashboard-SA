from flask import Flask, render_template, request, jsonify, session, redirect, url_for, flash
from functools import wraps
from models import db, Employee, Portfolio, Project, TaskDefinition, TaskResponse, Attendance, AnnualGoal, Initiative, CustomTask
from seed import seed_database
import os, calendar, datetime, logging

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'pmo-dashboard-secret-key-2026')

# Use /tmp for Render (writable directory), SQLite locally
database_url = os.environ.get('DATABASE_URL')
if database_url:
    # Render PostgreSQL
    app.config['SQLALCHEMY_DATABASE_URI'] = database_url.replace('postgres://', 'postgresql://')
else:
    # Local SQLite or Render with /tmp
    if os.environ.get('RENDER'):
        # Render with SQLite in /tmp
        app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:////tmp/pmo.db'
    else:
        # Local SQLite
        app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(BASE_DIR, 'pmo.db')

app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
logging.basicConfig(level=logging.DEBUG)
db.init_app(app)

# Login required decorator
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

PERFORMANCE_RATING = [
    {"scale": "0 - 70", "score": 1, "en": "Unsatisfactory", "ar": "غير مرضي: لم يحقق الأهداف الرئيسية أو معايير الأداء. يتطلب تحسين فوري."},
    {"scale": "71 - 85", "score": 2, "en": "Needs Improvement", "ar": "يحتاج تحسين: حقق بعض الأهداف لكن يوجد فجوات تتطلب تطوير إضافي."},
    {"scale": "86 - 100", "score": 3, "en": "Meets Expectations", "ar": "يلبي التوقعات: حقق جميع الأهداف المحددة وقدم نتائج متسقة مع الأداء المتوقع."},
    {"scale": "101 - 119", "score": 4, "en": "Exceeds Expectations", "ar": "يتجاوز التوقعات: حقق جميع الأهداف وقدم مبادرة ذات أثر قابل للقياس على مستوى القطاع."},
    {"scale": "120 - 130", "score": 5, "en": "Outstanding", "ar": "متميز: حقق نتائج استثنائية ذات أثر إيجابي على مستوى المركز. أظهر ابتكاراً وقيادة تتجاوز نطاق العمل."},
]

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '')
        password = request.form.get('password', '')
        if username == 'admin' and password == 'pmo2024':
            session['logged_in'] = True
            session['username'] = username
            return redirect(url_for('dashboard'))
        flash('بيانات الدخول غير صحيحة', 'error')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    session.pop('username', None)
    return redirect(url_for('login'))

@app.route('/')
@login_required
def dashboard():
    total_projects = Project.query.count()
    total_portfolios = Portfolio.query.count()
    total_initiatives = Initiative.query.count()
    completed_projects = Project.query.filter_by(status='مكتمل ومغلق').count()
    in_progress_projects = Project.query.filter_by(status='جاري العمل عليه').count()
    stopped_projects = Project.query.filter_by(status='متوقف').count()
    completion_rate = round((completed_projects / total_projects * 100) if total_projects else 0)

    completed_init = Initiative.query.filter_by(status='مكتملة').count()
    in_progress_init = Initiative.query.filter_by(status='جاري العمل').count()

    emp_data = []
    chart_labels = []
    chart_proj_values = []
    chart_init_values = []
    chart_tasks_values = []
    portfolio_counts = {}

    for emp in employees:
        proj_count = emp.projects.count()
        init_count = emp.initiatives.count()
        task_count = emp.custom_tasks.count()
        chart_labels.append(emp.name)
        chart_proj_values.append(proj_count)
        chart_init_values.append(init_count)
        chart_tasks_values.append(task_count)
        for p in emp.portfolios:
            portfolio_counts[p.name] = portfolio_counts.get(p.name, 0) + emp.projects.filter_by(portfolio_id=p.id).count()
        emp_data.append({
            'id': emp.id, 'name': emp.name,
            'portfolios': ', '.join([p.name for p in emp.portfolios]),
            'project_count': proj_count,
            'init_count': init_count,
            'task_count': task_count,
            'completed': emp.projects.filter_by(status='مكتمل ومغلق').count(),
        })

    port_labels = list(portfolio_counts.keys())[:10]
    port_values = [portfolio_counts[k] for k in port_labels]

    status_labels = ['مكتمل ومغلق', 'جاري العمل عليه', 'متوقف']
    status_values = [completed_projects, in_progress_projects, stopped_projects]

    return render_template('dashboard.html',
        employees=emp_data, total_employees=len(employees),
        total_projects=total_projects, total_portfolios=total_portfolios,
        total_initiatives=total_initiatives,
        completion_rate=completion_rate, ratings=PERFORMANCE_RATING,
        chart_labels=chart_labels, chart_proj_values=chart_proj_values,
        chart_init_values=chart_init_values, chart_tasks_values=chart_tasks_values,
        port_labels=port_labels, port_values=port_values,
        status_labels=status_labels, status_values=status_values,
        completed_init=completed_init, in_progress_init=in_progress_init)

# ---- EMPLOYEES ----
@app.route('/employees')
@login_required
def employees_page():
    employees = Employee.query.filter_by(is_active=True).all()
    portfolios = Portfolio.query.order_by(Portfolio.name).all()
    return render_template('employees.html', employees=employees, portfolios=portfolios)

@app.route('/employees/add', methods=['POST'])
@login_required
def add_employee():
    name = request.form.get('name', '').strip()
    if not name:
        flash('يرجى إدخال اسم الموظف', 'error')
        return redirect(url_for('employees_page'))
    emp = Employee(name=name)
    port_ids = request.form.getlist('portfolios')
    for pid in port_ids:
        p = Portfolio.query.get(int(pid))
        if p: emp.portfolios.append(p)
    db.session.add(emp)
    db.session.commit()
    flash(f'تم إضافة الموظف {name} بنجاح', 'success')
    return redirect(url_for('employees_page'))

@app.route('/employees/<int:emp_id>/delete', methods=['POST'])
@login_required
def delete_employee(emp_id):
    emp = Employee.query.get_or_404(emp_id)
    emp.is_active = False
    db.session.commit()
    flash(f'تم حذف الموظف {emp.name}', 'success')
    return redirect(url_for('employees_page'))

# ---- PROJECTS ----
@app.route('/projects')
@login_required
def projects_page():
    emp_filter = request.args.get('employee_id', type=int)
    status_filter = request.args.get('status', '')
    q = Project.query
    if emp_filter:
        q = q.filter_by(employee_id=emp_filter)
    if status_filter:
        q = q.filter_by(status=status_filter)
    projects = q.all()
    employees = Employee.query.filter_by(is_active=True).all()
    total = Project.query.count()
    completed = Project.query.filter_by(status='مكتمل ومغلق').count()
    in_progress = Project.query.filter_by(status='جاري العمل عليه').count()
    stopped = Project.query.filter_by(status='متوقف').count()
    return render_template('projects.html', projects=projects, employees=employees,
        total=total, completed=completed, in_progress=in_progress, stopped=stopped,
        emp_filter=emp_filter, status_filter=status_filter)

@app.route('/projects/<int:proj_id>/status', methods=['POST'])
@login_required
def update_project_status(proj_id):
    proj = Project.query.get_or_404(proj_id)
    proj.status = request.form.get('status', proj.status)
    db.session.commit()
    return jsonify({'ok': True})

# ---- WEEKLY TASKS ----
@app.route('/tasks')
@login_required
def tasks_page():
    employees = Employee.query.filter_by(is_active=True).all()
    task_defs = TaskDefinition.query.order_by(TaskDefinition.task_number).all()
    return render_template('weekly_tasks.html', employees=employees, task_defs=task_defs)

@app.route('/api/tasks/<int:project_id>/<week_date>')
@login_required
def get_tasks(project_id, week_date):
    task_defs = TaskDefinition.query.order_by(TaskDefinition.task_number).all()
    responses = {tr.task_def_id: tr for tr in TaskResponse.query.filter_by(project_id=project_id, week_date=week_date).all()}
    result = []
    for td in task_defs:
        tr = responses.get(td.id)
        result.append({
            'task_def_id': td.id, 'task_number': td.task_number,
            'question': td.question_text, 'goal_link': td.goal_link,
            'status': tr.status if tr else '', 'notes': tr.notes if tr else ''
        })
    return jsonify(result)

@app.route('/api/tasks/<int:project_id>/<week_date>', methods=['PUT'])
@login_required
def save_tasks(project_id, week_date):
    data = request.json
    for item in data.get('tasks', []):
        tr = TaskResponse.query.filter_by(project_id=project_id, task_def_id=item['task_def_id'], week_date=week_date).first()
        if not tr:
            tr = TaskResponse(project_id=project_id, task_def_id=item['task_def_id'], week_date=week_date)
            db.session.add(tr)
        tr.status = item.get('status', '')
        tr.notes = item.get('notes', '')
    db.session.commit()

    responses = TaskResponse.query.filter_by(project_id=project_id, week_date=week_date).all()
    total_defs = TaskDefinition.query.count()
    answered = [r for r in responses if r.status in ('نعم', 'لا ينطبق')]
    proj = Project.query.get(project_id)
    if len(answered) == total_defs:
        proj.status = 'مكتمل ومغلق'
    else:
        has_any = any(r.status for r in responses)
        proj.status = 'جاري العمل عليه' if has_any else proj.status
    db.session.commit()
    return jsonify({'ok': True, 'project_status': proj.status})

@app.route('/api/projects_by_employee/<int:emp_id>')
@login_required
def projects_by_employee(emp_id):
    projects = Project.query.filter_by(employee_id=emp_id).all()
    return jsonify([{'id': p.id, 'name': p.name, 'status': p.status} for p in projects])

# ---- INITIATIVES & CUSTOM TASKS ----
@app.route('/initiatives')
@login_required
def initiatives_page():
    employees = Employee.query.filter_by(is_active=True).all()
    emp_filter = request.args.get('employee_id', type=int)
    q_init = Initiative.query
    q_task = CustomTask.query
    if emp_filter:
        q_init = q_init.filter_by(employee_id=emp_filter)
        q_task = q_task.filter_by(employee_id=emp_filter)
    initiatives = q_init.order_by(Initiative.created_at.desc()).all()
    custom_tasks = q_task.order_by(CustomTask.created_at.desc()).all()
    return render_template('initiatives.html', employees=employees, initiatives=initiatives,
        custom_tasks=custom_tasks, emp_filter=emp_filter)

@app.route('/initiatives/add', methods=['POST'])
@login_required
def add_initiative():
    init = Initiative(
        employee_id=int(request.form['employee_id']),
        name=request.form['name'].strip(),
        description=request.form.get('description', '').strip(),
        start_date=request.form['start_date'],
        end_date=request.form['end_date'],
        status=request.form.get('status', 'جاري العمل')
    )
    db.session.add(init)
    db.session.commit()
    flash('تم إضافة المبادرة بنجاح', 'success')
    return redirect(url_for('initiatives_page'))

@app.route('/initiatives/<int:init_id>/delete', methods=['POST'])
@login_required
def delete_initiative(init_id):
    init = Initiative.query.get_or_404(init_id)
    db.session.delete(init)
    db.session.commit()
    flash('تم حذف المبادرة', 'success')
    return redirect(url_for('initiatives_page'))

@app.route('/initiatives/<int:init_id>/status', methods=['POST'])
@login_required
def update_initiative_status(init_id):
    init = Initiative.query.get_or_404(init_id)
    init.status = request.form.get('status', init.status)
    db.session.commit()
    return jsonify({'ok': True})

@app.route('/custom_tasks/add', methods=['POST'])
@login_required
def add_custom_task():
    task = CustomTask(
        employee_id=int(request.form['employee_id']),
        title=request.form['title'].strip(),
        description=request.form.get('description', '').strip(),
        start_date=request.form['start_date'],
        end_date=request_form['end_date'],
        status=request.form.get('status', 'جاري العمل')
    )
    db.session.add(task)
    db.session.commit()
    flash('تم إضافة المهمة بنجاح', 'success')
    return redirect(url_for('initiatives_page'))

@app.route('/custom_tasks/<int:task_id>/delete', methods=['POST'])
@login_required
def delete_custom_task(task_id):
    task = CustomTask.query.get_or_404(task_id)
    db.session.delete(task)
    db.session.commit()
    flash('تم حذف المهمة', 'success')
    return redirect(url_for('initiatives_page'))

@app.route('/custom_tasks/<int:task_id>/status', methods=['POST'])
@login_required
def update_custom_task_status(task_id):
    task = CustomTask.query.get_or_404(task_id)
    task.status = request.form.get('status', task.status)
    db.session.commit()
    return jsonify({'ok': True})

# ---- ATTENDANCE ----
@app.route('/attendance')
@login_required
def attendance_page():
    employees = Employee.query.filter_by(is_active=True).all()
    year = request.args.get('year', datetime.date.today().year, type=int)
    month = request.args.get('month', datetime.date.today().month, type=int)

    cal = calendar.Calendar(firstweekday=6)
    working_days = []
    for d in cal.itermonthdates(year, month):
        if d.month == month and d.weekday() in (6, 0, 1, 2, 3):
            working_days.append(d)

    day_names_map = {6: 'الأحد', 0: 'الإثنين', 1: 'الثلاثاء', 2: 'الأربعاء', 3: 'الخميس'}
    days_info = [{'date': d.isoformat(), 'day_num': d.day, 'day_name': day_names_map[d.weekday()]} for d in working_days]

    records = {}
    for att in Attendance.query.filter(Attendance.date.like(f'{year}-{month:02d}%')).all():
        records[(att.employee_id, att.date)] = att.status

    months_ar = ["يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"]

    att_stats = {}
    for emp in employees:
        total_days = len(days_info)
        present = sum(1 for d in days_info if records.get((emp.id, d['date'])) == 'present')
        absent = sum(1 for d in days_info if records.get((emp.id, d['date'])) in (None, ''))
        att_stats[emp.id] = {'total': total_days, 'present': present, 'rate': round(present/total_days*100) if total_days else 0}

    return render_template('attendance.html', employees=employees, days=days_info,
        records=records, year=year, month=month, month_name=months_ar[month-1],
        months=list(enumerate(months_ar, 1)), att_stats=att_stats)

@app.route('/api/attendance', methods=['POST'])
@login_required
def save_attendance():
    data = request.json
    emp_id = data['employee_id']
    date_str = data['date']

    d = datetime.date.fromisoformat(date_str)
    if d.weekday() in (4, 5):
        return jsonify({'error': 'لا يمكن تسجيل الحضور في عطلة نهاية الأسبوع'}), 400

    att = Attendance.query.filter_by(employee_id=emp_id, date=date_str).first()
    new_status = data['status']
    if att:
        if new_status == '':
            db.session.delete(att)
        else:
            att.status = new_status
    elif new_status:
        att = Attendance(employee_id=emp_id, date=date_str, status=new_status)
        db.session.add(att)
    db.session.commit()

    emp = Employee.query.get(emp_id)
    return jsonify({
        'ok': True,
        'balances': {
            'annual': emp.remaining_annual(),
            'sick': emp.remaining_sick(),
            'remote': emp.remaining_remote()
        }
    })

# ---- ANNUAL GOALS ----
@app.route('/goals')
@login_required
def goals_page():
    if not session.get('goals_authenticated'):
        return render_template('goals_login.html')
    employees = Employee.query.filter_by(is_active=True).all()
    return render_template('goals.html', employees=employees)

@app.route('/goals/login', methods=['POST'])
@login_required
def goals_login():
    password = request.form.get('password', '')
    if password == 'admin123':
        session['goals_authenticated'] = True
        return redirect(url_for('goals_page'))
    flash('كلمة المرور غير صحيحة', 'error')
    return redirect(url_for('goals_page'))

@app.route('/goals/logout')
@login_required
def goals_logout():
    session.pop('goals_authenticated', None)
    return redirect(url_for('goals_page'))

# ---- PRINT REPORT ----
@app.route('/report')
@login_required
def report_page():
    employees = Employee.query.filter_by(is_active=True).all()
    emp_id = request.args.get('employee_id', type=int)
    report_type = request.args.get('type', 'full')
    selected = None
    projects_list = []
    initiatives_list = []
    custom_tasks_list = []
    if emp_id:
        selected = Employee.query.get(emp_id)
        if selected:
            projects_list = list(selected.projects.all())
            initiatives_list = list(selected.initiatives.all())
            custom_tasks_list = list(selected.custom_tasks.all())
    try:
        return render_template('report.html', employees=employees, selected=selected,
            projects_list=projects_list, initiatives_list=initiatives_list,
            custom_tasks_list=custom_tasks_list, report_type=report_type)
    except Exception:
        import traceback
        tb = traceback.format_exc()
        app.logger.error(tb)
        return f"<pre>{tb}</pre>", 500

# ---- PERFORMANCE RATING ----
@app.route('/performance')
@login_required
def performance_page():
    return render_template('performance.html', ratings=PERFORMANCE_RATING)

# Initialize database on startup
try:
    with app.app_context():
        db.create_all()
        seed_database()
except Exception as e:
    print(f"Database initialization error: {e}")

@app.errorhandler(500)
def handle_500(e):
    import traceback, sys
    traceback.print_exc(file=sys.stdout)
    return f"<pre>500 Error:\n{traceback.format_exc()}</pre>", 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, port=port, host='0.0.0.0')
