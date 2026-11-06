import json
import tempfile
import unittest
from datetime import datetime
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
            {'generated_at', 'summary', 'status_counts', 'severity_counts', 'record_type_counts', 'monthly_trend'},
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

        for private_value in (
            'PRIVATE-ISSUE-TITLE', 'PRIVATE-DESCRIPTION', 'PRIVATE-INTEGRATION',
            'PRIVATE-IMPACT', 'PRIVATE-TICKET', 'PRIVATE-COMMENT',
            'PRIVATE-ROOT-CAUSE', 'PRIVATE-CLOSURE', 'PRIVATE-FIXED-BY',
            'PRIVATE-OWNER', 'PRIVATE-CUSTOM-STATUS', 'PRIVATE-CUSTOM-SEVERITY',
            'PRIVATE-CUSTOM-TYPE',
        ):
            self.assertNotIn(private_value, serialized)

        for private_key in (
            'issue', 'description', 'integration_flow', 'impact', 'ticket_number',
            'last_comments', 'root_cause', 'closure_note', 'fixed_by', 'responsible',
        ):
            self.assertNotIn(f'"{private_key}"', serialized)

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
