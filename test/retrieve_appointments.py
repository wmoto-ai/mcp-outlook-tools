from datetime import datetime
from src.outlook_tools.calendar_service import OutlookCalendarService

def main():
    service = OutlookCalendarService()
    start_date = datetime(2025, 1, 17, 0, 0)
    end_date = datetime(2025, 1, 17, 23, 59)
    appointments = service.get_calendar_items(start_date, end_date)
    
    if appointments:
        for appointment in appointments:
            print(f"Subject: {appointment['subject']}")
            print(f"Start: {appointment['start']}")
            print(f"End: {appointment['end']}")
            print(f"Location: {appointment['location']}")
            print(f"Body: {appointment['body']}\n")
    else:
        print("No appointments found for the specified date.")

if __name__ == "__main__":
    main()
