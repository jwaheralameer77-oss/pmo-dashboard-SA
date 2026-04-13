import os
import json
from flask import Flask
from models import db, Employee, Portfolio, Project, Initiative, CustomTask, Attendance
from app import app

def build_static():
    """Build static HTML files for Surge deployment"""
    
    # Create dist directory
    dist_dir = os.path.join(os.path.dirname(__file__), 'dist')
    os.makedirs(dist_dir, exist_ok=True)
    
    # Copy static files
    static_src = os.path.join(os.path.dirname(__file__), 'static')
    static_dst = os.path.join(dist_dir, 'static')
    
    if os.path.exists(static_src):
        import shutil
        if os.path.exists(static_dst):
            shutil.rmtree(static_dst)
        shutil.copytree(static_src, static_dst)
    
    with app.app_context():
        # Get all data
        employees = Employee.query.filter_by(is_active=True).all()
        projects = Project.query.all()
        initiatives = Initiative.query.all()
        
        # Build dashboard HTML
        dashboard_html = build_dashboard_html(employees, projects, initiatives)
        
        with open(os.path.join(dist_dir, 'index.html'), 'w', encoding='utf-8') as f:
            f.write(dashboard_html)
        
        # Build reports page
        reports_html = build_reports_html(employees)
        
        with open(os.path.join(dist_dir, 'reports.html'), 'w', encoding='utf-8') as f:
            f.write(reports_html)
        
        # Export data as JSON for frontend
        data = {
            'employees': [{'id': e.id, 'name': e.name} for e in employees],
            'projects_count': len(projects),
            'initiatives_count': len(initiatives)
        }
        
        with open(os.path.join(dist_dir, 'data.json'), 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"✅ Static site built in {dist_dir}")
    return dist_dir

def build_dashboard_html(employees, projects, initiatives):
    """Build dashboard HTML"""
    
    emp_rows = ""
    for i, emp in enumerate(employees, 1):
        portfolios = ', '.join([p.name for p in emp.portfolios]) if emp.portfolios else '-'
        proj_count = emp.projects.count()
        emp_rows += f"""
        <tr>
            <td>{i}</td>
            <td>{emp.name}</td>
            <td>{portfolios}</td>
            <td>{proj_count}</td>
        </tr>
        """
    
    return f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PMO Dashboard - مركز التأمين الصحي الوطني</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        * {{ font-family: 'Noto Sans Arabic', sans-serif; }}
        body {{ background: #f5f5f5; }}
        .teal-bg {{ background-color: #0D7377; }}
        .teal-dark {{ background-color: #0A5C5F; }}
        .teal-light {{ background-color: #E0F2F1; }}
    </style>
</head>
<body>
    <!-- Header -->
    <header class="teal-bg text-white p-4 shadow-lg">
        <div class="container mx-auto flex items-center justify-between">
            <div class="flex items-center gap-4">
                <img src="static/img/logo.jpeg" alt="Logo" class="h-12 w-12 rounded-full object-cover" onerror="this.style.display='none'">
                <div>
                    <h1 class="text-xl font-bold">مركز التأمين الصحي الوطني</h1>
                    <p class="text-sm opacity-80">نظام إدارة المشاريع (PMO)</p>
                </div>
            </div>
            <nav class="flex gap-4">
                <a href="index.html" class="px-4 py-2 bg-white text-teal-700 rounded-lg font-semibold">الرئيسية</a>
                <a href="reports.html" class="px-4 py-2 hover:bg-teal-600 rounded-lg">التقارير</a>
            </nav>
        </div>
    </header>

    <!-- Main Content -->
    <main class="container mx-auto p-6">
        <!-- KPI Cards -->
        <div class="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">
            <div class="bg-white rounded-xl shadow-md p-5 border-r-4" style="border-color: #0D7377;">
                <p class="text-gray-500 text-sm">الموظفين</p>
                <p class="text-3xl font-bold" style="color: #0D7377;">{len(employees)}</p>
            </div>
            <div class="bg-white rounded-xl shadow-md p-5 border-r-4" style="border-color: #43A047;">
                <p class="text-gray-500 text-sm">المشاريع</p>
                <p class="text-3xl font-bold text-green-600">{len(projects)}</p>
            </div>
            <div class="bg-white rounded-xl shadow-md p-5 border-r-4" style="border-color: #FB8C00;">
                <p class="text-gray-500 text-sm">المبادرات</p>
                <p class="text-3xl font-bold text-orange-600">{len(initiatives)}</p>
            </div>
            <div class="bg-white rounded-xl shadow-md p-5 border-r-4" style="border-color: #7B1FA2;">
                <p class="text-gray-500 text-sm">المحافظ</p>
                <p class="text-3xl font-bold text-purple-600">{len(set([p.name for emp in employees for p in emp.portfolios]))}</p>
            </div>
        </div>

        <!-- Employees Table -->
        <div class="bg-white rounded-xl shadow-md overflow-hidden">
            <div class="teal-bg text-white px-6 py-4">
                <h2 class="text-lg font-bold">الموظفون</h2>
            </div>
            <div class="overflow-x-auto">
                <table class="w-full text-sm">
                    <thead class="teal-light">
                        <tr>
                            <th class="px-4 py-3 text-right">#</th>
                            <th class="px-4 py-3 text-right">الاسم</th>
                            <th class="px-4 py-3 text-right">المحافظ</th>
                            <th class="px-4 py-3 text-center">المشاريع</th>
                        </tr>
                    </thead>
                    <tbody>
                        {emp_rows}
                    </tbody>
                </table>
            </div>
        </div>
    </main>

    <!-- Footer -->
    <footer class="teal-dark text-white text-center p-4 mt-8">
        <p>© 2026 مركز التأمين الصحي الوطني - إدارة المشاريع</p>
    </footer>
</body>
</html>"""

def build_reports_html(employees):
    """Build reports HTML"""
    
    options = ""
    for emp in employees:
        options += f'<option value="{emp.id}">{emp.name}</option>'
    
    return f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>التقارير - PMO Dashboard</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        * {{ font-family: 'Noto Sans Arabic', sans-serif; }}
        body {{ background: #f5f5f5; }}
        .teal-bg {{ background-color: #0D7377; }}
        .teal-dark {{ background-color: #0A5C5F; }}
        .teal-light {{ background-color: #E0F2F1; }}
    </style>
</head>
<body>
    <!-- Header -->
    <header class="teal-bg text-white p-4 shadow-lg">
        <div class="container mx-auto flex items-center justify-between">
            <div class="flex items-center gap-4">
                <img src="static/img/logo.jpeg" alt="Logo" class="h-12 w-12 rounded-full object-cover" onerror="this.style.display='none'">
                <div>
                    <h1 class="text-xl font-bold">مركز التأمين الصحي الوطني</h1>
                    <p class="text-sm opacity-80">نظام إدارة المشاريع (PMO)</p>
                </div>
            </div>
            <nav class="flex gap-4">
                <a href="index.html" class="px-4 py-2 hover:bg-teal-600 rounded-lg">الرئيسية</a>
                <a href="reports.html" class="px-4 py-2 bg-white text-teal-700 rounded-lg font-semibold">التقارير</a>
            </nav>
        </div>
    </header>

    <!-- Main Content -->
    <main class="container mx-auto p-6">
        <div class="bg-white rounded-xl shadow-md p-6">
            <h2 class="text-2xl font-bold mb-6" style="color: #0D7377;">التقارير</h2>
            
            <div class="mb-6">
                <label class="block text-sm font-semibold text-gray-700 mb-2">اختر الموظف</label>
                <select class="w-full border border-gray-300 rounded-lg px-4 py-2">
                    <option value="">-- اختر موظفاً --</option>
                    {options}
                </select>
            </div>

            <div class="grid grid-cols-2 md:grid-cols-4 gap-4">
                <button class="p-4 bg-teal-50 border-2 border-teal-500 rounded-lg hover:bg-teal-100 transition">
                    <span class="text-2xl">📊</span>
                    <p class="font-semibold mt-2">تقرير كامل</p>
                </button>
                <button class="p-4 bg-green-50 border-2 border-green-500 rounded-lg hover:bg-green-100 transition">
                    <span class="text-2xl">📁</span>
                    <p class="font-semibold mt-2">المشاريع</p>
                </button>
                <button class="p-4 bg-orange-50 border-2 border-orange-500 rounded-lg hover:bg-orange-100 transition">
                    <span class="text-2xl">📅</span>
                    <p class="font-semibold mt-2">الحضور</p>
                </button>
                <button class="p-4 bg-purple-50 border-2 border-purple-500 rounded-lg hover:bg-purple-100 transition">
                    <span class="text-2xl">🏆</span>
                    <p class="font-semibold mt-2">شهادة تميز</p>
                </button>
            </div>
        </div>
    </main>

    <!-- Footer -->
    <footer class="teal-dark text-white text-center p-4 mt-8">
        <p>© 2026 مركز التأمين الصحي الوطني - إدارة المشاريع</p>
    </footer>
</body>
</html>"""

if __name__ == '__main__':
    build_static()
