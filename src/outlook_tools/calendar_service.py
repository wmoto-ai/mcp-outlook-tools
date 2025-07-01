import win32com.client
from datetime import datetime
from typing import List, Dict, Any

class OutlookCalendarService:
    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.calendar = self.namespace.GetDefaultFolder(9)

    def get_calendar_items(self, start_date: datetime, end_date: datetime) -> List[Dict[str, Any]]:
        items = self.calendar.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        restriction = "[Start] >= '{}' AND [End] <= '{}'".format(
            start_date.strftime("%m/%d/%Y %H:%M %p"),
            end_date.strftime("%m/%d/%Y %H:%M %p")
        )
        restricted_items = items.Restrict(restriction)

        result = []
        for appointment in restricted_items:
            result.append({
                "subject": appointment.Subject,
                "start": appointment.Start.strftime("%Y-%m-%d %H:%M"),
                "end": appointment.End.strftime("%Y-%m-%d %H:%M"),
                "location": appointment.Location,
                "body": appointment.Body,
                "categories": appointment.Categories,
                "busy_status": appointment.BusyStatus
            })
        return result

    def add_appointment(self, subject: str, start: datetime, end: datetime, 
                       location: str = "", body: str = "", categories: str = "", busy_status: int = 1) -> bool:
        try:
            appointment = self.calendar.Items.Add()
            appointment.Subject = subject
            appointment.Start = start
            appointment.End = end
            appointment.Location = location
            appointment.Body = body
            appointment.Categories = categories
            appointment.BusyStatus = busy_status
            appointment.Save()
            if categories or busy_status != 1:
                appointment.Send()
            return True
        except Exception as e:
            print(f"Error adding appointment: {e}")
            return False
