from datetime import datetime
from dateutil.parser import parse
from dateutil.relativedelta import relativedelta
import json
from typing import Optional

from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.utilities.logging import get_logger
# 相対インポートから絶対インポートに変更
from outlook_tools.calendar_service import OutlookCalendarService
from outlook_tools.search_service import OutlookSearchService

mcp = FastMCP("Outlook Calendar")
calendar_service = OutlookCalendarService()

@mcp.tool()
async def add_appointment(
    subject: str,
    start_time: Optional[str] = None,
    end_time: Optional[str] = None,
    location: str = "",
    description: str = "",
    categories: str = "",
    busy_status: int = 1
) -> str:
    """Add a new appointment to Outlook calendar"""
    try:
        if not start_time or not end_time:
            return "I need both start time and end time. Please provide them."

        start_dt = parse(start_time) + relativedelta(hours=9)
        end_dt = parse(end_time) + relativedelta(hours=9)

        if calendar_service.add_appointment(subject, start_dt, end_dt, location, description, categories, busy_status):
            return f"Successfully added appointment: {subject}"
        else:
            return "Failed to add appointment"
    except ValueError:
        return "Invalid date/time format. Please provide dates in YYYY-MM-DD HH:MM format"

@mcp.tool()
async def get_calendar(start_date: str, end_date: str) -> str:
    """Get calendar items for the specified date range"""
    try:
        start_dt = parse(start_date)
        end_dt = parse(end_date) + relativedelta(days=1)
        items = calendar_service.get_calendar_items(start_dt, end_dt)
        
        if not items:
            return "No appointments found for the specified period."
        
        result = ["Calendar appointments:"]
        for item in items:
            result.append("\n---")
            result.append(f"Subject: {item['subject']}")
            result.append(f"Start: {item['start']}")
            result.append(f"End: {item['end']}")
            result.append(f"Location: {item['location']}")
            result.append(f"Details: {item['body'][:100]}...")
            result.append(f"Categories: {item.get('categories', 'N/A')}")
            result.append(f"Busy Status: {item.get('busy_status', 'N/A')}")
        
        return "\n".join(result)
    except ValueError:
        return "Invalid date format. Please provide dates in YYYY-MM-DD format"

search_service = OutlookSearchService()

@mcp.tool()
async def send_email(
    to: str,
    cc: str,
    subject: str,
    body: str
) -> str:
    """Send an email with the specified details and display it before sending"""
    import win32com.client

    try:
        # Outlookアプリケーションをインスタンス化
        outlook = win32com.client.Dispatch("Outlook.Application")

        # メールオブジェクトの作成
        mail = outlook.CreateItem(0)  # 0: メールアイテム

        mail.to = to
        mail.cc = cc
        mail.subject = subject
        mail.bodyFormat = 1
        mail.body = body

        # 送信前に確認（Outlookが起動）
        mail.display(True)

        # メール送信
        mail.Send()

        return "Email sent successfully."
    except Exception as e:
        return f"Failed to send email: {str(e)}"

@mcp.tool()
async def search_contact(name: str) -> str:
    """Search for a contact in Outlook by name"""
    return search_service.search_user(name)

@mcp.tool()
async def search_email(date: str, keyword: str) -> str:
    """
    指定した日付 (YYYY-MM-DD形式) に受信し、
    件名または本文にキーワードが含まれる Outlook のメールを検索するツールです。
    """
    try:
        # 入力された文字列を日付オブジェクトに変換
        target_date = datetime.strptime(date, "%Y-%m-%d").date()
    except ValueError:
        return "Invalid date format. Please use YYYY-MM-DD."

    try:
        # Outlook の COM オブジェクトを取得
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        # 受信トレイ (olFolderInbox は 6)
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items

        filtered_emails = []
        for message in messages:
            try:
                # MailItem かどうかの確認（43 は MailItem の定数）
                if message.Class == 43:
                    received_time = message.ReceivedTime  # datetime.datetime オブジェクト
                    if received_time.date() == target_date:
                        subject = message.Subject or ""
                        body = message.Body or ""
                        if keyword.lower() in subject.lower() or keyword.lower() in body.lower():
                            filtered_emails.append(message)
            except Exception:
                # メール以外のアイテム等の例外は無視する
                continue

        if not filtered_emails:
            return f"No emails found on {target_date} with keyword '{keyword}'."

        result_lines = [f"Found {len(filtered_emails)} email(s) on {target_date} with keyword '{keyword}':"]
        for idx, email in enumerate(filtered_emails, start=1):
            result_lines.append("-----")
            result_lines.append(f"Email {idx}:")
            result_lines.append(f"Sender: {email.Sender}")
            result_lines.append(f"Subject: {email.Subject}")
            result_lines.append(f"Received: {email.ReceivedTime}")
            body_text = email.Body.strip() if email.Body else ""
            # 先頭200文字まで表示
            preview = body_text.replace("\r\n", " ")[:200] + ("..." if len(body_text) > 200 else "")
            result_lines.append(f"Body Preview: {preview}")
        return "\n".join(result_lines)
    except Exception as e:
        return f"Error occurred during email search: {str(e)}"

if __name__ == "__main__":
    mcp.run()
