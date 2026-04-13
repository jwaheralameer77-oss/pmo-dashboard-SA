from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()

employee_portfolios = db.Table('employee_portfolios',
    db.Column('employee_id', db.Integer, db.ForeignKey('employees.id'), primary_key=True),
    db.Column('portfolio_id', db.Integer, db.ForeignKey('portfolios.id'), primary_key=True)
)

class Employee(db.Model):
    __tablename__ = 'employees'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    annual_leave_total = db.Column(db.Integer, default=25)
    sick_leave_total = db.Column(db.Integer, default=60)
    remote_days_total = db.Column(db.Integer, default=12)
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    portfolios = db.relationship('Portfolio', secondary=employee_portfolios, backref=db.backref('employees', lazy='dynamic'))
    projects = db.relationship('Project', backref='employee', lazy='dynamic')
    attendance_records = db.relationship('Attendance', backref='employee', lazy='dynamic')
    annual_goals = db.relationship('AnnualGoal', backref='employee', lazy='dynamic', order_by='AnnualGoal.id')

    def used_annual(self):
        return Attendance.query.filter_by(employee_id=self.id, status='annual').count()
    def used_sick(self):
        return Attendance.query.filter_by(employee_id=self.id, status='sick').count()
    def used_remote(self):
        return Attendance.query.filter_by(employee_id=self.id, status='remote').count()
    def remaining_annual(self):
        return self.annual_leave_total - self.used_annual()
    def remaining_sick(self):
        return self.sick_leave_total - self.used_sick()
    def remaining_remote(self):
        return self.remote_days_total - self.used_remote()

class Portfolio(db.Model):
    __tablename__ = 'portfolios'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False, unique=True)
    projects = db.relationship('Project', backref='portfolio', lazy='dynamic')

class Project(db.Model):
    __tablename__ = 'projects'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(500), nullable=False)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    portfolio_id = db.Column(db.Integer, db.ForeignKey('portfolios.id'), nullable=False)
    status = db.Column(db.String(50), default='جاري العمل عليه')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    task_responses = db.relationship('TaskResponse', backref='project', lazy='dynamic')

class TaskDefinition(db.Model):
    __tablename__ = 'task_definitions'
    id = db.Column(db.Integer, primary_key=True)
    task_number = db.Column(db.Integer, nullable=False)
    question_text = db.Column(db.String(500), nullable=False)
    goal_link = db.Column(db.Integer, default=1)

    responses = db.relationship('TaskResponse', backref='task_definition', lazy='dynamic')

class TaskResponse(db.Model):
    __tablename__ = 'task_responses'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('projects.id'), nullable=False)
    task_def_id = db.Column(db.Integer, db.ForeignKey('task_definitions.id'), nullable=False)
    week_date = db.Column(db.String(10), nullable=False)
    status = db.Column(db.String(50), default='')
    notes = db.Column(db.Text, default='')

    __table_args__ = (db.UniqueConstraint('project_id', 'task_def_id', 'week_date'),)

class Attendance(db.Model):
    __tablename__ = 'attendance'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    date = db.Column(db.String(10), nullable=False)
    status = db.Column(db.String(20), nullable=False)

    __table_args__ = (db.UniqueConstraint('employee_id', 'date'),)

class AnnualGoal(db.Model):
    __tablename__ = 'annual_goals'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    goal_number = db.Column(db.Integer, nullable=False)
    text = db.Column(db.Text, nullable=False)
    weight = db.Column(db.String(10), default='20%')

class Initiative(db.Model):
    __tablename__ = 'initiatives'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    name = db.Column(db.String(500), nullable=False)
    description = db.Column(db.Text, default='')
    start_date = db.Column(db.String(10), nullable=False)
    end_date = db.Column(db.String(10), nullable=False)
    status = db.Column(db.String(50), default='جاري العمل')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    employee = db.relationship('Employee', backref=db.backref('initiatives', lazy='dynamic'))

class CustomTask(db.Model):
    __tablename__ = 'custom_tasks'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    title = db.Column(db.String(500), nullable=False)
    description = db.Column(db.Text, default='')
    start_date = db.Column(db.String(10), nullable=False)
    end_date = db.Column(db.String(10), nullable=False)
    status = db.Column(db.String(50), default='جاري العمل')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    employee = db.relationship('Employee', backref=db.backref('custom_tasks', lazy='dynamic'))
