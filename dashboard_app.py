"""
WS Issue Dashboard - Flask Application
Reads data from Open_Issues_Wondersoft.xlsx and displays an interactive dashboard
"""

from flask import Flask, render_template, jsonify, request, redirect, session, url_for
from flask_wtf.csrf import CSRFProtect
from openpyxl import load_workbook
from datetime import datetime
import hmac
import json
import os
import threading

app = Flask(__name__)
secret_key = os.environ.get('SECRET_KEY')
if not secret_key:
    raise RuntimeError('SECRET_KEY must be configured before starting the dashboard')

app.config.update(
    SECRET_KEY=secret_key,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    SESSION_COOKIE_SECURE=os.environ.get('SESSION_COOKIE_SECURE', 'false').lower() == 'true',
)
csrf = CSRFProtect(app)

DASHBOARD_USERNAME = os.environ.get('DASHBOARD_USERNAME', '')
DASHBOARD_PASSWORD = os.environ.get('DASHBOARD_PASSWORD', '')
if not DASHBOARD_USERNAME or not DASHBOARD_PASSWORD:
    raise RuntimeError('DASHBOARD_USERNAME and DASHBOARD_PASSWORD must be configured')

EXCEL_PATH = os.path.join(os.path.dirname(__file__), 'Open_Issues_Wondersoft.xlsx')
MEETING_NOTES_PATH = os.path.join(os.path.dirname(__file__), 'meeting_notes.json')

# Thread lock for Excel file writes
excel_lock = threading.Lock()

# Color palette from dataviz skill (validated for CVD safety)
COLORS = {
    'categorical': [
        '#2a78d6',  # blue
        '#eb6834',  # orange
        '#1baf7a',  # aqua
        '#eda100',  # yellow
        '#e87ba4',  # magenta
        '#008300',  # green
        '#6250d6',  # violet
        '#e34948',  # red
    ],
    'status': {
        'good': '#0ca30c',
        'warning': '#fab219',
        'serious': '#ec835a',
        'critical': '#d03b3b',
    },
    'severity': {
        'Critical': '#d03b3b',
        'High': '#ec835a',
        'Medium': '#fab219',
        'Low': '#0ca30c',
    }
}

def load_tracker_data():
    """Load and parse data from Tracker Master sheet"""
    wb = load_workbook(EXCEL_PATH, data_only=True, read_only=True)
    try:
        ws = wb['Tracker Master']

        issues = []
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
            if row[0] is None:  # Skip empty rows
                continue

            issue = {
                's_no': row[0],
                'date_reported': row[1],
                'issue': row[2],
                'description': row[3],
                'integration_flow': row[4],
                'impact': row[5],
                'severity': str(row[6]).strip() if row[6] else 'Unknown',
                'ticket_number': row[7],
                'last_comments': row[8],
                'root_cause': row[9],
                'closure_note': row[10],
                'closed_date': row[11],
                'fixed_by': row[12],
                'status': row[13],
                'responsible': row[14],
                'delay_days': row[15],
                'record_type': row[16],
            }
            issues.append(issue)

        return issues
    finally:
        wb.close()

def get_dashboard_stats(issues):
    """Calculate dashboard statistics"""
    total = len(issues)

    # Status counts
    status_counts = {}
    for issue in issues:
        status = issue['status'] or 'Unknown'
        status_counts[status] = status_counts.get(status, 0) + 1

    # Severity counts
    severity_counts = {}
    for issue in issues:
        sev = issue['severity'] or 'Unknown'
        severity_counts[sev] = severity_counts.get(sev, 0) + 1

    # Integration flow counts
    integration_counts = {}
    for issue in issues:
        flow = issue['integration_flow'] or 'Unknown'
        integration_counts[flow] = integration_counts.get(flow, 0) + 1

    # Responsible counts
    responsible_counts = {}
    for issue in issues:
        resp = issue['responsible'] or 'Unknown'
        responsible_counts[resp] = responsible_counts.get(resp, 0) + 1

    # Record type counts
    record_type_counts = {}
    for issue in issues:
        rt = issue['record_type'] or 'Unknown'
        record_type_counts[rt] = record_type_counts.get(rt, 0) + 1

    # Monthly trend (issues reported per month)
    monthly_counts = {}
    for issue in issues:
        if issue['date_reported']:
            try:
                if isinstance(issue['date_reported'], datetime):
                    dt = issue['date_reported']
                else:
                    dt = datetime.strptime(str(issue['date_reported']), '%Y-%m-%d %H:%M:%S')
                key = dt.strftime('%Y-%m')
                monthly_counts[key] = monthly_counts.get(key, 0) + 1
            except:
                pass

    # Sort monthly
    monthly_sorted = dict(sorted(monthly_counts.items()))

    # Delay analysis
    delays = [i['delay_days'] for i in issues if i['delay_days'] is not None and isinstance(i['delay_days'], (int, float))]
    avg_delay = sum(delays) / len(delays) if delays else 0
    max_delay = max(delays) if delays else 0

    # Open vs Closed
    open_count = sum(1 for i in issues if i['status'] not in ['Closed', 'Resolved'])
    closed_count = sum(1 for i in issues if i['status'] in ['Closed', 'Resolved'])

    # In progress
    in_progress_count = sum(1 for i in issues if i['status'] in ['In Progress', 'In Development', 'In Testing', 'In Production'])

    return {
        'total': total,
        'open': open_count,
        'closed': closed_count,
        'in_progress': in_progress_count,
        'status_counts': status_counts,
        'severity_counts': severity_counts,
        'integration_counts': integration_counts,
        'responsible_counts': responsible_counts,
        'record_type_counts': record_type_counts,
        'monthly_trend': monthly_sorted,
        'avg_delay': round(avg_delay, 1),
        'max_delay': max_delay,
    }

def get_chart_data(stats):
    """Prepare data for Chart.js"""

    # Status Chart (Doughnut)
    status_labels = list(stats['status_counts'].keys())
    status_data = list(stats['status_counts'].values())
    status_colors = []
    for label in status_labels:
        if label == 'Closed':
            status_colors.append(COLORS['status']['good'])
        elif label in ['In Progress', 'In Development', 'In Testing', 'In Production']:
            status_colors.append(COLORS['status']['warning'])
        elif label == 'Resolved':
            status_colors.append(COLORS['status']['good'])
        else:
            status_colors.append(COLORS['categorical'][len(status_colors) % len(COLORS['categorical'])])

    # Severity Chart (Bar)
    severity_order = ['Critical', 'High', 'Medium', 'Low']
    severity_labels = [s for s in severity_order if s in stats['severity_counts']]
    severity_labels += [s for s in stats['severity_counts'] if s not in severity_order]
    severity_data = [stats['severity_counts'][s] for s in severity_labels]
    severity_colors = [COLORS['severity'].get(s, COLORS['categorical'][i % len(COLORS['categorical'])])
                       for i, s in enumerate(severity_labels)]

    # Integration Flow Chart (Bar - top 10)
    integration_sorted = sorted(stats['integration_counts'].items(), key=lambda x: x[1], reverse=True)[:10]
    integration_labels = [x[0] for x in integration_sorted]
    integration_data = [x[1] for x in integration_sorted]
    integration_colors = COLORS['categorical'][:len(integration_labels)]

    # Responsible Chart (Doughnut)
    responsible_labels = list(stats['responsible_counts'].keys())
    responsible_data = list(stats['responsible_counts'].values())
    responsible_colors = COLORS['categorical'][:len(responsible_labels)]

    # Record Type Chart (Doughnut)
    record_type_labels = list(stats['record_type_counts'].keys())
    record_type_data = list(stats['record_type_counts'].values())
    record_type_colors = COLORS['categorical'][:len(record_type_labels)]

    # Monthly Trend (Line)
    monthly_labels = list(stats['monthly_trend'].keys())
    monthly_data = list(stats['monthly_trend'].values())

    return {
        'status': {'labels': status_labels, 'data': status_data, 'colors': status_colors},
        'severity': {'labels': severity_labels, 'data': severity_data, 'colors': severity_colors},
        'integration': {'labels': integration_labels, 'data': integration_data, 'colors': integration_colors},
        'responsible': {'labels': responsible_labels, 'data': responsible_data, 'colors': responsible_colors},
        'record_type': {'labels': record_type_labels, 'data': record_type_data, 'colors': record_type_colors},
        'monthly': {'labels': monthly_labels, 'data': monthly_data},
    }

def load_meeting_notes():
    """Load meeting notes from JSON file"""
    if os.path.exists(MEETING_NOTES_PATH):
        with open(MEETING_NOTES_PATH, 'r') as f:
            return json.load(f)
    return {}

def save_meeting_notes(notes):
    """Save meeting notes to JSON file"""
    with open(MEETING_NOTES_PATH, 'w') as f:
        json.dump(notes, f, indent=2, default=str)

def add_new_issue(issue_data):
    """Add a new issue to the Excel file"""
    with excel_lock:
        with open(EXCEL_PATH, 'rb') as source:
            wb = load_workbook(source)
        try:
            ws = wb['Tracker Master']

            # Find next S.No
            max_s_no = 0
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1, values_only=True):
                if row[0] is not None and isinstance(row[0], (int, float)):
                    max_s_no = max(max_s_no, int(row[0]))

            new_s_no = max_s_no + 1

            # Prepare new row data
            new_row = [
                new_s_no,
                issue_data.get('date_reported', datetime.now()),
                issue_data.get('issue', ''),
                issue_data.get('description', ''),
                issue_data.get('integration_flow', ''),
                issue_data.get('impact', ''),
                issue_data.get('severity', 'Medium'),
                issue_data.get('ticket_number', ''),
                issue_data.get('last_comments', ''),
                issue_data.get('root_cause', ''),
                issue_data.get('closure_note', ''),
                issue_data.get('closed_date', None),
                issue_data.get('fixed_by', ''),
                issue_data.get('status', 'Open'),
                issue_data.get('responsible', ''),
                issue_data.get('delay_days', 0),
                issue_data.get('record_type', 'Open Issue'),
            ]

            ws.append(new_row)
            wb.save(EXCEL_PATH)
            return new_s_no
        finally:
            wb.close()

@app.route('/healthz')
def healthz():
    return jsonify({'status': 'ok'})

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '')
        password = request.form.get('password', '')
        if hmac.compare_digest(username, DASHBOARD_USERNAME) and hmac.compare_digest(password, DASHBOARD_PASSWORD):
            session.clear()
            session['authenticated'] = True
            return redirect(url_for('dashboard'))
        return render_template('login.html', error='Invalid username or password'), 401

    return render_template('login.html')

@app.route('/logout', methods=['POST'])
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.before_request
def require_login():
    if request.endpoint in {'healthz', 'login', 'static'}:
        return None
    if session.get('authenticated'):
        return None
    if request.path.startswith('/api/'):
        return jsonify({'error': 'Authentication required'}), 401
    return redirect(url_for('login'))

@app.route('/')
def dashboard():
    issues = load_tracker_data()
    stats = get_dashboard_stats(issues)
    charts = get_chart_data(stats)
    meeting_notes = load_meeting_notes()

    # Get filter options
    statuses = sorted(set(i['status'] for i in issues if i['status']))
    severities = sorted(set(i['severity'] for i in issues if i['severity']))
    integrations = sorted(set(i['integration_flow'] for i in issues if i['integration_flow']))
    responsibles = sorted(set(i['responsible'] for i in issues if i['responsible']))
    record_types = sorted(set(i['record_type'] for i in issues if i['record_type']))

    return render_template('dashboard.html',
        stats=stats,
        charts=charts,
        meeting_notes=meeting_notes,
        filters={
            'statuses': statuses,
            'severities': severities,
            'integrations': integrations,
            'responsibles': responsibles,
            'record_types': record_types,
        },
        entry_added=request.args.get('entry_added') == '1',
    )

@app.route('/api/issues')
def api_issues():
    """API endpoint for filtered issues"""
    issues = load_tracker_data()

    # Apply filters
    status_filter = request.args.get('status')
    severity_filter = request.args.get('severity')
    integration_filter = request.args.get('integration')
    responsible_filter = request.args.get('responsible')
    record_type_filter = request.args.get('record_type')
    search = request.args.get('search', '').lower()

    filtered = issues
    if status_filter:
        filtered = [i for i in filtered if i['status'] == status_filter]
    if severity_filter:
        filtered = [i for i in filtered if i['severity'] == severity_filter]
    if integration_filter:
        filtered = [i for i in filtered if i['integration_flow'] == integration_filter]
    if responsible_filter:
        filtered = [i for i in filtered if i['responsible'] == responsible_filter]
    if record_type_filter:
        filtered = [i for i in filtered if i['record_type'] == record_type_filter]
    if search:
        filtered = [i for i in filtered if
            search in (i['issue'] or '').lower() or
            search in (i['description'] or '').lower() or
            search in (i['ticket_number'] or '').lower() or
            search in (i['impact'] or '').lower()]

    # Sort by date reported descending
    filtered.sort(key=lambda x: x['date_reported'] or datetime.min, reverse=True)

    # Pagination
    page = int(request.args.get('page', 1))
    per_page = int(request.args.get('per_page', 20))
    total = len(filtered)
    start = (page - 1) * per_page
    end = start + per_page
    paginated = filtered[start:end]

    # Format for JSON
    result = []
    for issue in paginated:
        result.append({
            's_no': issue['s_no'],
            'date_reported': issue['date_reported'].strftime('%Y-%m-%d') if issue['date_reported'] else '',
            'issue': issue['issue'],
            'description': issue['description'],
            'integration_flow': issue['integration_flow'],
            'impact': issue['impact'],
            'severity': issue['severity'],
            'ticket_number': issue['ticket_number'],
            'status': issue['status'],
            'responsible': issue['responsible'],
            'delay_days': issue['delay_days'],
            'record_type': issue['record_type'],
            'fixed_by': issue['fixed_by'],
            'closed_date': issue['closed_date'].strftime('%Y-%m-%d') if issue['closed_date'] else '',
            'last_comments': issue['last_comments'],
            'root_cause': issue['root_cause'],
            'closure_note': issue['closure_note'],
        })

    return jsonify({
        'issues': result,
        'total': total,
        'page': page,
        'per_page': per_page,
        'total_pages': (total + per_page - 1) // per_page
    })

@app.route('/api/stats')
def api_stats():
    """API endpoint for stats"""
    issues = load_tracker_data()
    stats = get_dashboard_stats(issues)
    charts = get_chart_data(stats)
    return jsonify({'stats': stats, 'charts': charts})

@app.route('/api/add_issue', methods=['POST'])
def api_add_issue():
    """API endpoint to add a new issue, requirement, or enhancement"""
    try:
        data = request.get_json(silent=True)
        if not isinstance(data, dict):
            return jsonify({'success': False, 'error': 'Request body must be a JSON object'}), 400

        issue = data.get('issue')
        record_type = data.get('record_type')
        if not isinstance(issue, str) or not issue.strip():
            return jsonify({'success': False, 'error': 'Issue/title is required'}), 400
        if not isinstance(record_type, str) or not record_type.strip():
            return jsonify({'success': False, 'error': 'Record type is required'}), 400

        record_type = record_type.strip()
        if record_type not in {'Issue', 'Requirement', 'Enhancement'}:
            return jsonify({'success': False, 'error': 'Unsupported record type'}), 400

        optional_fields = (
            'description', 'integration_flow', 'impact', 'severity',
            'ticket_number', 'last_comments', 'status', 'responsible',
        )
        for field in optional_fields:
            value = data.get(field)
            if value is not None and not isinstance(value, str):
                return jsonify({'success': False, 'error': f'{field} must be text'}), 400

        # Set defaults
        issue_data = {
            'date_reported': datetime.now(),
            'issue': issue.strip(),
            'description': data.get('description') or '',
            'integration_flow': data.get('integration_flow') or '',
            'impact': data.get('impact') or '',
            'severity': data.get('severity') or 'Medium',
            'ticket_number': data.get('ticket_number') or '',
            'last_comments': data.get('last_comments') or '',
            'root_cause': '',
            'closure_note': '',
            'closed_date': None,
            'fixed_by': '',
            'status': data.get('status') or 'Open',
            'responsible': data.get('responsible') or '',
            'delay_days': 0,
            'record_type': record_type,
        }

        new_s_no = add_new_issue(issue_data)

        return jsonify({
            'success': True,
            's_no': new_s_no,
            'message': 'Record added successfully'
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/meeting_notes', methods=['GET', 'POST'])
def api_meeting_notes():
    """API endpoint for meeting notes"""
    if request.method == 'GET':
        notes = load_meeting_notes()
        return jsonify(notes)
    
    # POST - save meeting notes
    try:
        data = request.get_json()
        notes = load_meeting_notes()
        
        date_key = data.get('date', datetime.now().strftime('%Y-%m-%d'))
        if date_key not in notes:
            notes[date_key] = {
                'date': date_key,
                'issues_discussed': [],
                'enhancements_discussed': [],
                'action_items': [],
                'general_notes': '',
                'attendees': [],
                'created_at': datetime.now().isoformat(),
                'updated_at': datetime.now().isoformat(),
            }
        
        # Update the note
        note = notes[date_key]
        note['issues_discussed'] = data.get('issues_discussed', note['issues_discussed'])
        note['enhancements_discussed'] = data.get('enhancements_discussed', note['enhancements_discussed'])
        note['action_items'] = data.get('action_items', note['action_items'])
        note['general_notes'] = data.get('general_notes', note['general_notes'])
        note['attendees'] = data.get('attendees', note['attendees'])
        note['updated_at'] = datetime.now().isoformat()
        
        save_meeting_notes(notes)
        
        return jsonify({'success': True, 'note': note})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/meeting_notes/<date>', methods=['DELETE'])
def api_delete_meeting_note(date):
    """Delete a meeting note"""
    try:
        notes = load_meeting_notes()
        if date in notes:
            del notes[date]
            save_meeting_notes(notes)
            return jsonify({'success': True})
        return jsonify({'success': False, 'error': 'Note not found'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=False, host='127.0.0.1', port=5001)