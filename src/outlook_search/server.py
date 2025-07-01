import sys
import logging
from typing import Optional
from dataclasses import dataclass

from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.utilities.logging import get_logger

logger = get_logger(__name__)

try:
    import win32com.client
    import pythoncom
except ImportError:
    logger.error("Failed to import win32com. Please ensure pywin32 is installed correctly.")
    sys.exit(1)

def fix_encoding(text: Optional[str]) -> Optional[str]:
    """Fix encoding issues with Japanese text from Outlook"""
    if not text:
        return None
    try:
        # Try different encodings
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

# Create MCP server
mcp = FastMCP("Outlook Search", dependencies=["pywin32"])

class OutlookWrapper:
    """Class to manage connection with Outlook"""
    
    def __init__(self):
        self.outlook = None
        self.mail = None
        
    def __enter__(self):
        try:
            # Initialize COM for each thread
            pythoncom.CoInitialize()
            # Connect to Outlook application
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.mail = self.outlook.CreateItem(0)  # 0: olMailItem
            return self
        except Exception as e:
            logger.error(f"Error initializing Outlook: {e}")
            self.cleanup()
            raise
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.cleanup()
    
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

@mcp.tool()  
async def search_outlook(name: str) -> str:
    """
    Search for a user in Outlook and return their information.
    
    Args:
        name: Name or email to search for
    """
    try:
        logger.info(f"Searching for user: {name}")
        
        with OutlookWrapper() as outlook_wrapper:
            if not outlook_wrapper.outlook:
                return "Failed to connect to Outlook"
            
            recipient = outlook_wrapper.mail.Recipients.Add(name)
            
            if not recipient.Resolve():
                logger.warning(f"Could not resolve user: {name}")
                return f"User not found: {name}"
            
            address_entry = recipient.AddressEntry
            exchange_user = address_entry.GetExchangeUser()
            
            if not exchange_user:
                logger.warning(f"Could not get Exchange user details for: {name}")
                return f"Could not retrieve user information: {name}"
            
            # Create debug log
            logger.debug(f"Raw department value: {exchange_user.Department}")
            logger.debug(f"Raw company value: {exchange_user.CompanyName}")
            
            # Format response with encoding fix for specific fields
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
            
            result = "\n".join(line for line in response if line)
            logger.debug(f"Final formatted result: {result}")
            return result
            
    except Exception as e:
        logger.error(f"Error during Outlook search: {str(e)}", exc_info=True)
        return f"Error searching Outlook: {str(e)}"

def main():
    logging.basicConfig(level=logging.DEBUG)
    logger.info("Starting Outlook Search MCP Server")
    mcp.run()

if __name__ == "__main__":
    main()