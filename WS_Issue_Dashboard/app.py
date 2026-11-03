from flask import Flask, render_template_string, request, redirect, url_for, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
from openpyxl import load_workbook
import io

app = Flask(__name__)
app.config['SECRET_KEY'] = 'dev-secret-key-change-me'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///instance/ws_issues.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True)
    password_hash = db.Column(db.String(256))
    role = db.Column(db.String(20), default='viewer')
    def set_password(self, p): self.password_hash = generate_password_hash(p)
    def check_password(self, p): return check_password_hash(self.password_hash, p)

class Issue(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200))
    description = db.Column(db.Text)
    status = db.Column(db.String(50), default='Open')
    priority = db.Column(db.String(20), default='Medium')
    category = db.Column(db.String(100))
    assigned_to = db.Column(db.Integer, db.ForeignKey('user.id'))
    created_by = db.Column(db.Integer, db.ForeignKey('user.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

@login_manager.user_loader
def load_user(uid): return User.query.get(int(uid))

# ---------- ROUTES ----------
@app.route('/')
@login_required
def index():
    issues = Issue.query.order_by(Issue.created_at.desc()).all()
    return render_template_string(BASE_TEMPLATE, issues=issues)

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        u = User.query.filter_by(username=request.form['username']).first()
        if u and u.check_password(request.form['password']):
            login_user(u); return redirect('/')
        flash('Invalid credentials')
    return LOGIN_TEMPLATE

@app.route('/logout')
@login_required
def logout(): logout_user(); return redirect('/login')

@app.route('/new', methods=['GET','POST'])
@login_required
def new():
    if request.method == 'POST':
        i = Issue(title=request.form['title'], description=request.form['description'],
                  status=request.form['status'], priority=request.form['priority'],
                  category=request.form['category'], created_by=current_user.id)
        db.session.add(i); db.session.commit()
        return redirect('/')
    return NEW_TEMPLATE

@app.route('/issue/<int:id>')
@login_required
def view(id):
    i = Issue.query.get_or_404(id)
    return render_template_string(VIEW_TEMPLATE, issue=i)

@app.route('/import', methods=['GET','POST'])
@login_required
def import_excel():
    if request.method == 'POST':
        f = request.files['file']
        wb = load_workbook(f, data_only=True)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]

        imported = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            data = dict(zip(headers, row))
            i = Issue(
                title=str(data.get('Title', data.get('title', 'Issue'))),
                description=str(data.get('Description', data.get('description', ''))),
                status=str(data.get('Status', data.get('status', 'Open'))),
                priority=str(data.get('Priority', data.get('priority', 'Medium'))),
                category=str(data.get('Category', data.get('category', 'Bug'))),
                created_by=current_user.id
            )
            db.session.add(i)
            imported += 1
        db.session.commit()
        flash(f'Imported {imported} issues')
        return redirect('/')
    return IMPORT_TEMPLATE

# ---------- TEMPLATES ----------
BASE_TEMPLATE = '''
<!DOCTYPE html><html><head><title>WS Issues</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
</head><body class="p-4">
<nav class="navbar navbar-dark bg-primary mb-4"><div class="container-fluid">
<a class="navbar-brand" href="/">WS Issue Dashboard</a>
<div><span class="me-3">{{ current_user.username }} ({{ current_user.role }})</span>
<a class="btn btn-light btn-sm" href="/logout">Logout</a></div></div></nav>
{% with messages = get_flashed_messages() %}{% if messages %}
<div class="alert alert-success">{{ messages[0] }}</div>{% endif %}{% endwith %}
<div class="d-flex justify-content-between mb-3">
<h2>Issues</h2><a class="btn btn-primary" href="/new">New Issue</a> <a class="btn btn-outline-secondary" href="/import">Import Excel</a></div>
<table class="table table-hover">
<thead><tr><th>ID</th><th>Title</th><th>Status</th><th>Priority</th><th>Category</th><th>Assignee</th><th>Actions</th></tr></thead>
<tbody>
{% for i in issues %}
<tr><td>{{ i.id }}</td><td>{{ i.title }}</td>
<td><span class="badge bg-{{ 'warning' if i.status=='Open' else 'info' if i.status=='In Progress' else 'success' }}">{{ i.status }}</span></td>
<td><span class="badge bg-{{ 'danger' if i.priority=='Critical' else 'warning' if i.priority=='High' else 'info' }}">{{ i.priority }}</span></td>
<td>{{ i.category }}</td><td>{{ i.assignee.username if i.assignee else 'Unassigned' }}</td>
<td><a class="btn btn-sm btn-outline-primary" href="/issue/{{ i.id }}">View</a></td></tr>
{% endfor %}</tbody></table>
</body></html>
'''

LOGIN_TEMPLATE = '''
<!DOCTYPE html><html><head><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body class="d-flex align-items-center" style="height:100vh"><div class="container"><div class="row justify-content-center"><div class="col-md-4">
<div class="card"><div class="card-body"><h3 class="text-center mb-4">WS Issue Dashboard</h3>
<form method="POST"><div class="mb-3"><input name="username" class="form-control" placeholder="Username" required></div>
<div class="mb-3"><input name="password" type="password" class="form-control" placeholder="Password" required></div>
<button class="btn btn-primary w-100">Login</button></form></div></div></div></div></body></html>
'''

NEW_TEMPLATE = '''
<!DOCTYPE html><html><head><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body class="p-4"><div class="container" style="max-width:600px">
<h2>New Issue</h2><form method="POST">
<div class="mb-3"><input name="title" class="form-control" placeholder="Title" required></div>
<div class="mb-3"><textarea name="description" class="form-control" placeholder="Description"></textarea></div>
<div class="row"><div class="col"><select name="status" class="form-select"><option>Open</option><option>In Progress</option><option>Resolved</option><option>Closed</option></select></div>
<div class="col"><select name="priority" class="form-select"><option>Low</option><option selected>Medium</option><option>High</option><option>Critical</option></select></div></div>
<div class="mb-3 mt-2"><input name="category" class="form-control" placeholder="Category (Bug, Feature, etc)"></div>
<button class="btn btn-primary">Create</button> <a class="btn btn-secondary" href="/">Cancel</a></form></div></body></html>
'''

VIEW_TEMPLATE = '''
<!DOCTYPE html><html><head><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body class="p-4"><div class="container" style="max-width:800px">
<a class="btn btn-secondary mb-3" href="/">← Back</a>
<h2>#{{ issue.id }} - {{ issue.title }}</h2>
<p><strong>Status:</strong> <span class="badge bg-{{ 'warning' if issue.status=='Open' else 'info' if issue.status=='In Progress' else 'success' }}">{{ issue.status }}</span></p>
<p><strong>Priority:</strong> <span class="badge bg-{{ 'danger' if issue.priority=='Critical' else 'warning' if issue.priority=='High' else 'info' }}">{{ issue.priority }}</span></p>
<p><strong>Category:</strong> {{ issue.category }}</p>
<p><strong>Description:</strong></p><div>{{ issue.description or 'No description' }}</div>
<hr><a class="btn btn-primary" href="/">Back to List</a></div></body></html>
'''

IMPORT_TEMPLATE = '''
<!DOCTYPE html><html><head><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet"></head>
<body class="p-4"><div class="container" style="max-width:600px">
<h2>Import from Excel</h2><form method="POST" enctype="multipart/form-data">
<div class="mb-3"><input name="file" type="file" class="form-control" accept=".xlsx,.xls" required></div>
<div class="alert alert-info">Columns: Title, Description, Status, Priority, Category (case-insensitive)</div>
<button class="btn btn-primary">Import</button> <a class="btn btn-secondary" href="/">Cancel</a></form></div></body></html>
'''

# ---------- INIT ----------
with app.app_context():
    db.create_all()
    if not User.query.filter_by(username='admin').first():
        a = User(username='admin', role='admin'); a.set_password('admin123')
        db.session.add(a); db.session.commit()

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)