import os
import re
import tempfile
import unittest

os.environ.setdefault('SECRET_KEY', 'test-secret-key')
os.environ.setdefault('DASHBOARD_USERNAME', 'test-user')
os.environ.setdefault('DASHBOARD_PASSWORD', 'test-password')
os.environ.setdefault('SESSION_COOKIE_SECURE', 'false')

import dashboard_app


class DashboardAuthTests(unittest.TestCase):
    def setUp(self):
        dashboard_app.app.config.update(TESTING=True, WTF_CSRF_ENABLED=True)
        self.client = dashboard_app.app.test_client()

    def csrf_token(self, response, pattern):
        match = re.search(pattern, response.get_data(as_text=True))
        self.assertIsNotNone(match)
        return match.group(1)

    def test_dashboard_and_api_require_login(self):
        self.assertEqual(self.client.get('/').status_code, 302)
        self.assertEqual(self.client.get('/api/issues').status_code, 401)
        self.assertEqual(self.client.get('/healthz').status_code, 200)

    def test_login_and_csrf_protect_meeting_note_writes(self):
        login_page = self.client.get('/login')
        login_token = self.csrf_token(
            login_page, r'name="csrf_token" value="([^"]+)"'
        )
        response = self.client.post('/login', data={
            'csrf_token': login_token,
            'username': 'test-user',
            'password': 'test-password',
        })
        self.assertEqual(response.status_code, 302)

        dashboard_page = self.client.get('/')
        dashboard_token = self.csrf_token(
            dashboard_page, r'name="csrf-token" content="([^"]+)"'
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            original_path = dashboard_app.MEETING_NOTES_PATH
            dashboard_app.MEETING_NOTES_PATH = os.path.join(temp_dir, 'meeting_notes.json')
            try:
                rejected = self.client.post('/api/meeting_notes', json={
                    'date': '2026-11-03',
                    'general_notes': 'test',
                })
                self.assertEqual(rejected.status_code, 400)

                accepted = self.client.post(
                    '/api/meeting_notes',
                    json={'date': '2026-11-03', 'general_notes': 'test'},
                    headers={'X-CSRFToken': dashboard_token},
                )
                self.assertEqual(accepted.status_code, 200)
            finally:
                dashboard_app.MEETING_NOTES_PATH = original_path


if __name__ == '__main__':
    unittest.main()
