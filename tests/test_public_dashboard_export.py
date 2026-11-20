import json
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook

from scripts.export_public_dashboard import export_public_summary


HEADERS = (
    'S.No', 'Date Reported', 'Issue', 'Issue Description', 'Integration Flow',
    'Impact', 'Severity', 'Ticket Number', 'Last Comments', 'Root Cause',
    'Closure Note', 'Closed Date', 'Fixed By', 'Status', 'Responsible',
    'Delay in Days', 'Record Type',
)


class PublicDashboardExportTests(unittest.TestCase):
    def make_workbook(self, directory):
        path = Path(directory) / 'tracker.xlsx'
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = 'Tracker Master'
        worksheet.append(HEADERS)
        worksheet.append((
            1, datetime(2026, 1, 3), 'PRIVATE-ISSUE-TITLE', 'PRIVATE-DESCRIPTION',
            'PRIVATE-INTEGRATION', 'PRIVATE-IMPACT', 'Critical', 'PRIVATE-TICKET',
            'PRIVATE-COMMENT', 'PRIVATE-ROOT-CAUSE', 'PRIVATE-CLOSURE', None,
            'PRIVATE-FIXED-BY', 'Open', 'PRIVATE-OWNER', 5, 'Issue',
        ))
        worksheet.append((
            2, datetime(2026, 2, 5), 'PRIVATE-REQUIREMENT', 'PRIVATE-DESCRIPTION-2',
            'PRIVATE-INTEGRATION-2', 'PRIVATE-IMPACT-2', 'High', 'PRIVATE-TICKET-2',
            'PRIVATE-COMMENT-2', 'PRIVATE-ROOT-CAUSE-2', '', None, '',
            'In Progress', 'PRIVATE-OWNER-2', 3, 'Requirement',
        ))
        worksheet.append((
            3, datetime(2026, 2, 18), 'PRIVATE-ENHANCEMENT', 'PRIVATE-DESCRIPTION-3',
            'PRIVATE-INTEGRATION-3', 'PRIVATE-IMPACT-3', 'Medium', 'PRIVATE-TICKET-3',
            'PRIVATE-COMMENT-3', 'PRIVATE-ROOT-CAUSE-3', 'PRIVATE-CLOSURE-3',
            datetime(2026, 3, 2), 'PRIVATE-FIXED-BY-3', 'Resolved',
            'PRIVATE-OWNER-3', 2, 'Enhancement',
        ))
        worksheet.append((
            4, 'not-a-date', 'PRIVATE-CUSTOM-STATUS', '', '', '',
            'PRIVATE-CUSTOM-SEVERITY', '', '', '', '', None, '',
            'PRIVATE-CUSTOM-STATUS', 'PRIVATE-OWNER-4', None,
            'PRIVATE-CUSTOM-TYPE',
        ))
        workbook.save(path)
        workbook.close()
        return path

    def test_exports_only_allowlisted_aggregates(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook_path = self.make_workbook(temp_dir)
            output_path = Path(temp_dir) / 'site' / 'dashboard-data.json'
            result = export_public_summary(workbook_path, output_path)
            serialized = output_path.read_text(encoding='utf-8')
            data = json.loads(serialized)

        self.assertEqual(
            set(data),
            {
                'generated_at', 'summary', 'executive_summary', 'status_counts',
                'severity_counts', 'record_type_counts', 'monthly_trend',
                'open_issues_detail', 'open_requirements_enhancements_detail',
            },
        )
        self.assertEqual(data['summary'], {
            'total': 4,
            'open': 3,
            'open_issues': 1,
            'open_requirements_enhancements': 2,
            'closed': 1,
            'in_progress': 1,
            'avg_delay': 3.3,
            'max_delay': 5,
        })
        self.assertEqual(data['executive_summary'], {
            'total_summary': {
                'open_issues': 1,
                'requirements_enhancements': 2,
                'closed_issues': 0,
                'grand_total': 3,
            },
            'open_by_severity': [
                {'severity': 'Critical', 'open_issues': 1, 'open_requirements_enhancements': 0, 'total': 1},
                {'severity': 'High', 'open_issues': 0, 'open_requirements_enhancements': 1, 'total': 1},
                {'severity': 'Medium', 'open_issues': 0, 'open_requirements_enhancements': 1, 'total': 1},
                {'severity': 'Low', 'open_issues': 0, 'open_requirements_enhancements': 0, 'total': 0},
            ],
            'highlights': {
                'active_high_severity': 1,
                'closed_last_7_days': 0,
                'high_severity_closed_last_7_days': 0,
            },
        })
        self.assertEqual(data['status_counts']['Open'], 1)
        self.assertEqual(data['status_counts']['In Progress'], 1)
        self.assertEqual(data['status_counts']['Resolved'], 1)
        self.assertEqual(data['status_counts']['Other'], 1)
        self.assertEqual(data['severity_counts']['Critical'], 1)
        self.assertEqual(data['severity_counts']['Other'], 1)
        self.assertEqual(data['record_type_counts']['Issue'], 1)
        self.assertEqual(data['record_type_counts']['Requirement'], 1)
        self.assertEqual(data['record_type_counts']['Enhancement'], 1)
        self.assertEqual(data['record_type_counts']['Requirement / Enhancement'], 0)
        self.assertEqual(data['record_type_counts']['Other'], 1)
        self.assertEqual(data['monthly_trend'], [
            {'month': '2026-01', 'count': 1},
            {'month': '2026-02', 'count': 2},
        ])
        self.assertTrue(data['generated_at'].endswith('Z'))
        self.assertEqual(result, data)

        # Verify detailed executive summary data is included
        self.assertIn('open_issues_detail', data)
        self.assertIn('open_requirements_enhancements_detail', data)
        # Test data has: 1 open issue (serial 1 Open), 2 open req/enh (serial 2 In Progress, serial 3 Resolved)
        self.assertEqual(len(data['open_issues_detail']), 1)
        self.assertEqual(len(data['open_requirements_enhancements_detail']), 2)

        # Verify structure of open issues detail
        open_issue = data['open_issues_detail'][0]
        self.assertEqual(open_issue['s_no'], 1)
        self.assertEqual(open_issue['severity'], 'Critical')
        self.assertEqual(open_issue['status'], 'Open')
        self.assertEqual(open_issue['issue'], 'PRIVATE-ISSUE-TITLE')

        # Verify structure of requirements/enhancements detail
        req_enh = data['open_requirements_enhancements_detail']
        self.assertEqual(len(req_enh), 2)
        # In Progress Requirement (serial 2)
        self.assertEqual(req_enh[0]['s_no'], 2)
        self.assertEqual(req_enh[0]['severity'], 'High')
        self.assertEqual(req_enh[0]['status'], 'In Progress')
        self.assertEqual(req_enh[0]['issue'], 'PRIVATE-REQUIREMENT')
        # Resolved Enhancement (serial 3)
        self.assertEqual(req_enh[1]['s_no'], 3)
        self.assertEqual(req_enh[1]['severity'], 'Medium')
        self.assertEqual(req_enh[1]['status'], 'Resolved')
        self.assertEqual(req_enh[1]['issue'], 'PRIVATE-ENHANCEMENT')

        # Verify private data that should NOT be in detailed export (description, integration, impact, etc.)
        # Note: 'responsible' is intentionally included for the executive summary table display
        for private_key in (
            'description', 'integration_flow', 'impact', 'root_cause',
            'closure_note', 'fixed_by',
        ):
            # These should not appear as JSON keys in the detailed data
            for item in data['open_issues_detail'] + data['open_requirements_enhancements_detail']:
                self.assertNotIn(private_key, item)

    def test_open_summary_counts_match_excel_categories(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / 'tracker.xlsx'
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = 'Tracker Master'
            worksheet.append(HEADERS)
            for serial, status, record_type in (
                (1, 'Open', 'Open Issue'),
                (2, 'Closed', 'Open Issue'),
                (3, 'In Progress', 'Requirement / Enhancement'),
                (4, 'Resolved', 'Requirement / Enhancement'),
                (5, 'Closed', 'Requirement / Enhancement'),
                (6, 'Resolved', 'Issue'),
                (7, 'In Testing', 'Enhancement'),
            ):
                row = [None] * len(HEADERS)
                row[0] = serial
                row[13] = status
                row[16] = record_type
                worksheet.append(row)
            workbook.save(workbook_path)
            workbook.close()

            data = export_public_summary(workbook_path, Path(temp_dir) / 'summary.json')

        self.assertEqual(data['summary']['open_issues'], 2)
        self.assertEqual(data['summary']['open_requirements_enhancements'], 3)
        self.assertEqual(data['record_type_counts']['Requirement / Enhancement'], 3)

    def test_executive_summary_aggregates_match_excel_formulas(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / 'tracker.xlsx'
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = 'Tracker Master'
            worksheet.append(HEADERS)
            for serial, status, record_type, severity, closed_date in (
                (1, 'Open', 'Open Issue', 'Critical', None),
                (2, 'Closed', 'Open Issue', 'High', date(2026, 11, 6)),
                (3, 'In Progress', 'Requirement / Enhancement', 'High', None),
                (4, 'Resolved', 'Requirement / Enhancement', 'Low', date(2026, 11, 4)),
                (5, 'Closed', 'Requirement / Enhancement', 'Medium', date(2026, 10, 29)),
                (6, 'Resolved', 'Issue', 'Medium', None),
                (7, 'In Testing', 'Enhancement', 'High', None),
            ):
                row = [None] * len(HEADERS)
                row[0] = serial
                row[6] = severity
                row[11] = closed_date
                row[13] = status
                row[16] = record_type
                worksheet.append(row)
            workbook.save(workbook_path)
            workbook.close()

            data = export_public_summary(
                workbook_path,
                Path(temp_dir) / 'summary.json',
                as_of=date(2026, 11, 6),
            )

        executive = data['executive_summary']
        self.assertEqual(executive['total_summary'], {
            'open_issues': 2,
            'requirements_enhancements': 3,
            'closed_issues': 2,
            'grand_total': 7,
        })
        self.assertEqual(executive['open_by_severity'], [
            {'severity': 'Critical', 'open_issues': 1, 'open_requirements_enhancements': 0, 'total': 1},
            {'severity': 'High', 'open_issues': 0, 'open_requirements_enhancements': 2, 'total': 2},
            {'severity': 'Medium', 'open_issues': 1, 'open_requirements_enhancements': 0, 'total': 1},
            {'severity': 'Low', 'open_issues': 0, 'open_requirements_enhancements': 1, 'total': 1},
        ])
        self.assertEqual(executive['highlights'], {
            'active_high_severity': 2,
            'closed_last_7_days': 1,
            'high_severity_closed_last_7_days': 1,
        })

    def test_fails_when_required_aggregate_columns_are_missing(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / 'tracker.xlsx'
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = 'Tracker Master'
            worksheet.append(('S.No', 'Issue', 'Description'))
            workbook.save(workbook_path)
            workbook.close()

            with self.assertRaisesRegex(ValueError, 'required aggregate columns'):
                export_public_summary(workbook_path, Path(temp_dir) / 'data.json')


if __name__ == '__main__':
    unittest.main()
