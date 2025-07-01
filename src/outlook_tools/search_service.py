import logging
from typing import Optional
from dataclasses import dataclass

import pythoncom
import win32com.client
from mcp.server.fastmcp.utilities.logging import get_logger

logger = get_logger(__name__)

@dataclass
class UserInfo:
    name: str
    email: Optional[str] = None
    department: Optional[str] = None
    job_title: Optional[str] = None
    company: Optional[str] = None
    phone: Optional[str] = None
    location: Optional[str] = None
    manager: Optional[str] = None

def fix_encoding(text: Optional[str]) -> Optional[str]:
    """Fix encoding issues with Japanese text from Outlook"""
    if not text:
        return None
    try:
        for encoding in ['shift_jis', 'cp932', 'iso-2022-jp']:
            try:
                encoded = text.encode(encoding, errors='ignore')
                decoded = encoded.decode(encoding, errors='ignore')
                if decoded and not all('?' in c for c in decoded):
                    return decoded
            except Exception:
                continue
        return text
    except Exception as e:
        logger.warning(f"Error fixing encoding: {e}")
        return text

import datetime

class OutlookSearchService:
    def search_emails(self, target_date: datetime.date, keyword: str):
        """指定した日付とキーワードにマッチするメールを検索"""
        try:
            # Outlook アプリケーションの COM オブジェクトを取得
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            
            # 受信トレイはフォルダ番号6（olFolderInbox）に相当する
            inbox = outlook.GetDefaultFolder(6)
            messages = inbox.Items

            # Outlook の Items は COM オブジェクトなので、全件を走査する
            filtered_emails = []
            for message in messages:
                try:
                    # MailItem かどうかのチェック（43 は MailItem の定数）
                    if message.Class == 43:
                        received_time = message.ReceivedTime  # datetime.datetime オブジェクト
                        # 日付部分だけを比較
                        if received_time.date() == target_date:
                            # キーワード検索（大文字小文字を無視）
                            subject = message.Subject or ""
                            body = message.Body or ""
                            if keyword.lower() in subject.lower() or keyword.lower() in body.lower():
                                sender = message.Sender
                                recipients = [recipient.Address for recipient in message.Recipients]
                                filtered_emails.append({
                                    "subject": subject,
                                    "received_time": received_time,
                                    "sender": sender,
                                    "recipients": recipients,
                                    "body_preview": body.strip().replace("\r\n", " ")[:200]
                                })
                except Exception as e:
                    logger.warning(f"メールの処理中にエラーが発生しました: {e}")
            return filtered_emails
        except Exception as e:
            logger.error(f"メール検索中にエラーが発生しました: {e}")
            return []
    def __init__(self):
        # Initialize COM
        pythoncom.CoInitialize()
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.mail = self.outlook.CreateItem(0)
    
    def cleanup(self):
        """Clean up resources"""
        try:
            if self.mail:
                self.mail.Close(1)  # 1: olDiscard
                self.mail = None
            if self.outlook:
                self.outlook = None
            pythoncom.CoUninitialize()
        except Exception as e:
            logger.error(f"Error during cleanup: {e}")
    
    def search_user(self, name: str) -> str:
        try:
            recipient = self.mail.Recipients.Add(name)
            
            if not recipient.Resolve():
                logger.warning(f"Could not resolve user: {name}")
                return f"User not found: {name}"
            
            address_entry = recipient.AddressEntry
            exchange_user = address_entry.GetExchangeUser()
            
            if not exchange_user:
                logger.warning(f"Could not get Exchange user details for: {name}")
                return f"Could not retrieve user information: {name}"
            
            response = [
                "User Information:",
                f"Name: {exchange_user.Name}",
                f"Email: {exchange_user.PrimarySmtpAddress}" if exchange_user.PrimarySmtpAddress else "",
                f"Department: {fix_encoding(exchange_user.Department)}" if exchange_user.Department else "",
                f"Job Title: {fix_encoding(exchange_user.JobTitle)}" if exchange_user.JobTitle else "",
                f"Company: {fix_encoding(exchange_user.CompanyName)}" if exchange_user.CompanyName else "",
                f"Phone: {exchange_user.BusinessTelephoneNumber}" if exchange_user.BusinessTelephoneNumber else "",
                f"Location: {fix_encoding(exchange_user.OfficeLocation)}" if exchange_user.OfficeLocation else "",
                f"Manager: {fix_encoding(exchange_user.Manager)}" if exchange_user.Manager else ""
            ]
            
            return "\n".join(line for line in response if line)
            
        except Exception as e:
            logger.error(f"Error searching user: {str(e)}", exc_info=True)
            return f"Error searching Outlook: {str(e)}"
