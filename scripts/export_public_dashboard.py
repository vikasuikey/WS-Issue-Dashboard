"""Export a privacy-limited aggregate snapshot for the public Pages site."""

import argparse
import json
from collections import Counter
from datetime import date, datetime, timezone
from pathlib import Path

from openpyxl import load_workbook


STATUS_LABELS = (
    'Open',
    'In Progress',
    'In Development',
    'In Testing',
    'In Production',
    'Closed',
    'Resolved',
    'Other',
)
SEVERITY_LABELS = ('Critical', 'High', 'Medium', 'Low', 'Other')
RECORD_TYPE_LABELS = ('Issue', 'Requirement', 'Enhancement', 'Other')
IN_PROGRESS_STATUSES = {'in progress', 'in development', 'in testing', 'in production'}
CLOSED_STATUSES = {'closed', 'resolved'}


def _normalized_label(value, allowed, aliases=None):
    aliases = aliases or {}
    key = str(value).strip().casefold() if value is not None else ''
    canonical = aliases.get(key)
    if canonical:
        return canonical
    for label in allowed:
        if key == label.casefold():
            return label
    return 'Other'


def _month_key(value):
    if isinstance(value, datetime):
        return value.strftime('%Y-%m')
    if isinstance(value, date):
        return value.strftime('%Y-%m')
    if not isinstance(value, str):
        return None

    value = value.strip()
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace('Z', '+00:00')).strftime('%Y-%m')
    except ValueError:
        for date_format in ('%m/%d/%Y', '%d-%m-%Y', '%Y-%m-%d'):
            try:
                return datetime.strptime(value, date_format).strftime('%Y-%m')
            except ValueError:
                continue
    return None


def build_public_summary(workbook_path):
    """Read Tracker Master and return only allowlisted aggregate values."""
    workbook_file = Path(workbook_path).open('rb')
    try:
        workbook = load_workbook(workbook_file, data_only=True, read_only=True)
    except Exception:
        workbook_file.close()
        raise
    rows = None
    try:
        worksheet = workbook['Tracker Master']
        rows = worksheet.iter_rows(values_only=True)
        headers = next(rows, None)
        if not headers:
            raise ValueError('Tracker Master sheet is empty')

        columns = {
            str(value).strip(): index
            for index, value in enumerate(headers)
            if value is not None
        }
        required = ('S.No', 'Date Reported', 'Severity', 'Status', 'Delay in Days', 'Record Type')
        if any(column not in columns for column in required):
            raise ValueError('Tracker Master is missing required aggregate columns')

        total = 0
        status_counts = Counter()
        severity_counts = Counter()
        record_type_counts = Counter()
        monthly_counts = Counter()
        delays = []

        for row in rows:
            serial = row[columns['S.No']] if columns['S.No'] < len(row) else None
            if serial is None:
                continue

            total += 1
            status = _normalized_label(row[columns['Status']], STATUS_LABELS)
            severity = _normalized_label(row[columns['Severity']], SEVERITY_LABELS)
            record_type = _normalized_label(
                row[columns['Record Type']],
                RECORD_TYPE_LABELS,
                aliases={'open issue': 'Issue'},
            )
            status_counts[status] += 1
            severity_counts[severity] += 1
            record_type_counts[record_type] += 1

            month = _month_key(row[columns['Date Reported']])
            if month:
                monthly_counts[month] += 1

            delay = row[columns['Delay in Days']]
            if isinstance(delay, (int, float)) and not isinstance(delay, bool):
                delays.append(delay)

        closed_count = sum(status_counts[label] for label in ('Closed', 'Resolved'))
        in_progress_count = sum(status_counts[label] for label in (
            'In Progress', 'In Development', 'In Testing', 'In Production'
        ))
        avg_delay = round(sum(delays) / len(delays), 1) if delays else 0

        return {
            'generated_at': datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace('+00:00', 'Z'),
            'summary': {
                'total': total,
                'open': total - closed_count,
                'closed': closed_count,
                'in_progress': in_progress_count,
                'avg_delay': avg_delay,
                'max_delay': max(delays) if delays else 0,
            },
            'status_counts': {label: status_counts[label] for label in STATUS_LABELS},
            'severity_counts': {label: severity_counts[label] for label in SEVERITY_LABELS},
            'record_type_counts': {label: record_type_counts[label] for label in RECORD_TYPE_LABELS},
            'monthly_trend': [
                {'month': month, 'count': monthly_counts[month]}
                for month in sorted(monthly_counts)
            ],
        }
    finally:
        if rows is not None:
            rows.close()
        workbook.close()
        workbook_file.close()


def export_public_summary(workbook_path, output_path):
    summary = build_public_summary(workbook_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open('w', encoding='utf-8') as output_file:
        json.dump(summary, output_file, indent=2, ensure_ascii=False)
        output_file.write('\n')
    return summary


def main():
    repository_root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        '--workbook',
        type=Path,
        default=repository_root / 'Open_Issues_Wondersoft.xlsx',
    )
    parser.add_argument('--output', type=Path, default=Path('_site/dashboard-data.json'))
    args = parser.parse_args()

    export_public_summary(args.workbook, args.output)
    print('Wrote aggregate-only dashboard data.')


if __name__ == '__main__':
    main()
