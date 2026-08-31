# -*- coding: utf-8 -*-
import unittest
from unittest.mock import patch

from app import create_app


class ShutdownTests(unittest.TestCase):
    def setUp(self):
        self.app = create_app()
        self.app.config.update(TESTING=True)
        self.client = self.app.test_client()

    @patch("app._agendar_encerramento")
    def test_local_user_can_shutdown(self, schedule_shutdown):
        response = self.client.post(
            "/api/shutdown",
            environ_base={"REMOTE_ADDR": "127.0.0.1"},
        )

        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.get_json()["ok"])
        schedule_shutdown.assert_called_once_with()

    @patch("app._agendar_encerramento")
    def test_remote_user_cannot_shutdown(self, schedule_shutdown):
        response = self.client.post(
            "/api/shutdown",
            environ_base={"REMOTE_ADDR": "192.168.1.50"},
        )

        self.assertEqual(response.status_code, 403)
        self.assertFalse(response.get_json()["ok"])
        schedule_shutdown.assert_not_called()


if __name__ == "__main__":
    unittest.main()
