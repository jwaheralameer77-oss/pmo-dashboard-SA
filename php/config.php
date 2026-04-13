<?php
// Database connection
$db = new SQLite3('pmo.db');;

// Create tables if not exist
$db->exec('CREATE TABLE IF NOT EXISTS employees (id INTEGER PRIMARY KEY, name TEXT, is_active INTEGER DEFAULT 1)');
$db->exec('CREATE TABLE IF NOT EXISTS portfolios (id INTEGER PRIMARY KEY, name TEXT)');
$db->exec('CREATE TABLE IF NOT EXISTS projects (id INTEGER PRIMARY KEY, name TEXT, employee_id INTEGER, portfolio_id INTEGER, status TEXT)');
$db->exec('CREATE TABLE IF NOT EXISTS task_definitions (id INTEGER PRIMARY KEY, task_number INTEGER, question_text TEXT, goal_link INTEGER)');

// Seed data if empty
$count = $db->querySingle("SELECT COUNT(*) FROM employees");
if ($count == 0) {
    // Add employees
    $employees = ['شكرية كنكار', 'نايف الجلال', 'عبد العزيز المعجل', 'صالح الغفيلي', 'مشعل محمد'];
    foreach ($employees as $name) {
        $db->exec("INSERT INTO employees (name) VALUES ('$name')");
    }
    
    // Add task definitions
    $tasks = [
        "هل تم تحديث الأنظمة والمنصات المتعلقة بالمشاريع؟",
        "هل تم تحديث الخطة التفصيلية للمشروع؟",
        "هل تم حل التحديات والمخاطر؟",
        "هل تم متابعة وإغلاق أوامر التغيير؟",
        "هل تم إرسال تقرير أسبوعي لمالك المحفظة؟",
        "هل تم التنبيه بثلاثة أسابيع بحلول قرب تسليم المخرجات؟",
        "هل تم متابعة تسليم المخرجات؟",
        "هل تم تحديث الخطة المالية الرئيسية؟",
        "هل تم تصدير شهادات الإنجاز للمالية؟",
        "هل تم متابعة وتحديث المشاريع التشغيلية؟",
        "هل تم رفع جميع الوثائق في نظام ميسر؟ (مخرجات وشهادات إنجاز المعتمدة)",
        "هل يوجد مبادرات وإنجازات تم العمل عليها خلال هذا الأسبوع؟ (إذا كانت الإجابة نعم يتم إضافة اسم المبادرة ووصفها وتاريخ البداية والنهاية وحالة الإنجاز)",
        "هل تم الالتزام بمصفوفة الصلاحيات التشغيلية والحوكمة واتباعها؟"
    ];
    
    foreach ($tasks as $i => $task) {
        $num = $i + 1;
        $db->exec("INSERT INTO task_definitions (task_number, question_text, goal_link) VALUES ($num, '$task', 1)");
    }
}
?>
