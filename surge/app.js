// Data Management
let appData = {
    employees: [],
    portfolios: [],
    projects: [],
    taskDefinitions: [],
    annualGoals: [],
    initiatives: JSON.parse(localStorage.getItem('initiatives') || '[]'),
    customTasks: JSON.parse(localStorage.getItem('customTasks') || '[]'),
    taskResponses: JSON.parse(localStorage.getItem('taskResponses') || '{}'),
    attendance: JSON.parse(localStorage.getItem('attendance') || '{}')
};

// Load initial data from JSON file
fetch('data.json')
    .then(r => r.json())
    .then(data => {
        appData = {...appData, ...data};
        appData.taskDefinitions = data.task_definitions;
        appData.annualGoals = data.annual_goals;
        initDashboard();
        initEmployees();
        initProjects();
        initTasksPage();
    });

// Navigation
function showPage(pageId) {
    document.querySelectorAll('.page').forEach(p => p.classList.add('hidden'));
    document.getElementById(pageId).classList.remove('hidden');
}

// Initialize Dashboard
function initDashboard() {
    document.getElementById('totalEmployees').textContent = appData.employees.length;
    document.getElementById('totalProjects').textContent = appData.projects.length;
    document.getElementById('totalPortfolios').textContent = appData.portfolios.length;
    document.getElementById('totalInitiatives').textContent = appData.initiatives.length;
    
    const completed = appData.projects.filter(p => p.status === 'مكتمل ومغلق').length;
    const rate = appData.projects.length ? Math.round((completed / appData.projects.length) * 100) : 0;
    document.getElementById('completionRate').textContent = rate + '%';
    
    document.getElementById('currentDate').textContent = new Date().toLocaleDateString('ar-SA', {
        weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'
    });
    
    // Charts
    initCharts();
}

function initCharts() {
    // Combined Chart
    new Chart(document.getElementById('combinedChart'), {
        type: 'bar',
        data: {
            labels: appData.employees.map(e => e.name),
            datasets: [
                { label: 'المشاريع', data: appData.employees.map(e => appData.projects.filter(p => p.employee_id === e.id).length), backgroundColor: '#0D7377' },
                { label: 'المبادرات', data: appData.employees.map(() => 0), backgroundColor: '#7B1FA2' }
            ]
        },
        options: { responsive: true, plugins: { legend: { position: 'top' } } }
    });
    
    // Status Pie Chart
    const completed = appData.projects.filter(p => p.status === 'مكتمل ومغلق').length;
    const inProgress = appData.projects.filter(p => p.status === 'جاري العمل عليه').length;
    const stopped = appData.projects.filter(p => p.status === 'متوقف').length;
    
    new Chart(document.getElementById('statusChart'), {
        type: 'pie',
        data: {
            labels: ['مكتمل', 'جاري', 'متوقف'],
            datasets: [{ data: [completed, inProgress, stopped], backgroundColor: ['#43A047', '#FB8C00', '#E53935'] }]
        },
        options: { responsive: true }
    });
}

// Employees Page
function initEmployees() {
    const tbody = document.getElementById('employeesTable');
    tbody.innerHTML = appData.employees.map((emp, i) => `
        <tr class="border-b hover:bg-gray-50">
            <td class="px-4 py-3">${i + 1}</td>
            <td class="px-4 py-3 font-medium">${emp.name}</td>
            <td class="px-4 py-3">${emp.portfolios.join('، ')}</td>
            <td class="px-4 py-3">${appData.projects.filter(p => p.employee_id === emp.id).length}</td>
        </tr>
    `).join('');
}

// Projects Page
function initProjects() {
    const tbody = document.getElementById('projectsTable');
    tbody.innerHTML = appData.projects.map((proj, i) => {
        const portfolio = appData.portfolios.find(p => p.id === proj.portfolio_id);
        return `
        <tr class="border-b hover:bg-gray-50">
            <td class="px-4 py-3">${i + 1}</td>
            <td class="px-4 py-3">${proj.name}</td>
            <td class="px-4 py-3">${portfolio ? portfolio.name : ''}</td>
            <td class="px-4 py-3">
                <span class="px-2 py-1 rounded text-xs ${proj.status === 'مكتمل ومغلق' ? 'bg-green-100 text-green-700' : 'bg-orange-100 text-orange-700'}">${proj.status}</span>
            </td>
        </tr>
    `}).join('');
}

// Tasks Page
function initTasksPage() {
    const select = document.getElementById('taskEmployee');
    select.innerHTML = '<option value="">اختر الموظف</option>' + 
        appData.employees.map(e => `<option value="${e.id}">${e.name}</option>`).join('');
}

function loadEmployeeProjects() {
    const empId = parseInt(document.getElementById('taskEmployee').value);
    const select = document.getElementById('taskProject');
    const projects = appData.projects.filter(p => p.employee_id === empId);
    select.innerHTML = '<option value="">اختر المشروع</option>' + 
        projects.map(p => `<option value="${p.id}">${p.name}</option>`).join('');
}

function loadTasks() {
    const projectId = document.getElementById('taskProject').value;
    const weekDate = document.getElementById('taskDate').value;
    
    if (!projectId || !weekDate) return alert('الرجاء اختيار المشروع والتاريخ');
    
    const key = `${projectId}_${weekDate}`;
    const saved = appData.taskResponses[key] || {};
    
    const tbody = document.getElementById('tasksTable');
    tbody.innerHTML = appData.taskDefinitions.map((task, i) => `
        <tr class="border-b hover:bg-gray-50" data-task-id="${task.id}">
            <td class="px-4 py-3">${i + 1}</td>
            <td class="px-4 py-3">${task.question_text}</td>
            <td class="px-4 py-3">
                <select class="task-status border rounded px-2 py-1">
                    <option value="">--</option>
                    <option value="نعم" ${saved[task.id]?.status === 'نعم' ? 'selected' : ''}>نعم</option>
                    <option value="لا" ${saved[task.id]?.status === 'لا' ? 'selected' : ''}>لا</option>
                    <option value="لا ينطبق" ${saved[task.id]?.status === 'لا ينطبق' ? 'selected' : ''}>لا ينطبق</option>
                    <option value="جاري العمل عليه" ${saved[task.id]?.status === 'جاري العمل عليه' ? 'selected' : ''}>جاري العمل عليه</option>
                </select>
            </td>
            <td class="px-4 py-3">
                <input type="text" class="task-notes border rounded px-2 py-1 w-full" value="${saved[task.id]?.notes || ''}">
            </td>
        </tr>
    `).join('');
    
    document.getElementById('tasksContainer').classList.remove('hidden');
}

function saveTasks() {
    const projectId = document.getElementById('taskProject').value;
    const weekDate = document.getElementById('taskDate').value;
    const key = `${projectId}_${weekDate}`;
    
    const responses = {};
    document.querySelectorAll('#tasksTable tr').forEach(row => {
        const taskId = row.dataset.taskId;
        if (taskId) {
            responses[taskId] = {
                status: row.querySelector('.task-status').value,
                notes: row.querySelector('.task-notes').value
            };
        }
    });
    
    appData.taskResponses[key] = responses;
    localStorage.setItem('taskResponses', JSON.stringify(appData.taskResponses));
    alert('تم حفظ المهام بنجاح!');
}

// Show dashboard by default
showPage('dashboard');
