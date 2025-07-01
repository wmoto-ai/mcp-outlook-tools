import unittest
from datetime import datetime, timedelta
from src.outlook_tools.calendar_service import OutlookCalendarService

class TestOutlookCalendarService(unittest.TestCase):
    def setUp(self):
        self.service = OutlookCalendarService()

    def test_get_calendar_items(self):
        start_date = datetime.now() - timedelta(days=1)
        end_date = datetime.now() + timedelta(days=1)
        items = self.service.get_calendar_items(start_date, end_date)
        self.assertIsInstance(items, list)

    def test_add_appointment(self):
        subject = "Test Appointment"
        start = datetime.now() + timedelta(hours=1)
        end = start + timedelta(hours=1)
        result = self.service.add_appointment(subject, start, end, "Test Location", "Test Body")
        self.assertTrue(result)

if __name__ == "__main__":
    unittest.main()
