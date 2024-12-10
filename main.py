import os
import imaplib
import email
from datetime import datetime
import yaml
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle
from dotenv import load_dotenv

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/drive'
]

class ExpenseTracker:
    def __init__(self):
        self.load_config()
        self.authenticate_google()
        
    def load_config(self):
        """Load configuration from config.yaml."""
        with open('config.yaml', 'r') as file:
            self.config = yaml.safe_load(file)
    
    def authenticate_google(self):
        """Authenticate with Google APIs."""
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        
        self.sheets_service = build('sheets', 'v4', credentials=creds)
        self.drive_service = build('drive', 'v3', credentials=creds)
    
    def process_emails(self):
        """Process new emails with receipts/invoices."""
        try:
            # Connect to email
            mail = imaplib.IMAP4_SSL(
                self.config['email']['imap_server'],
                self.config['email']['imap_port']
            )
            mail.login(
                os.getenv('EMAIL_USER'),
                os.getenv('EMAIL_PASSWORD')
            )
            
            print(f"Connected to email: {os.getenv('EMAIL_USER')}")
            
            # Select inbox
            mail.select('INBOX')
            
            # Search for unread emails with receipt/invoice in subject
            _, messages = mail.search(None, '(UNSEEN SUBJECT "receipt" OR SUBJECT "invoice")')
            
            if not messages[0]:
                print("No new receipts or invoices found.")
                return
            
            print(f"Found {len(messages[0].split())} new messages to process.")
            
            # Process each email
            for msg_id in messages[0].split():
                try:
                    _, msg_data = mail.fetch(msg_id, '(RFC822)')
                    email_message = email.message_from_bytes(msg_data[0][1])
                    
                    print(f"\nProcessing email: {email_message['subject']}")
                    
                    # Process the email and extract expense data
                    expense_data = self.extract_expense_data(email_message)
                    
                    if expense_data:
                        # Add to Google Sheet
                        self.add_to_sheet(expense_data)
                        
                        # Mark email as read
                        mail.store(msg_id, '+FLAGS', '\\Seen')
                        print(f"Successfully processed: {expense_data}")
                    
                except Exception as e:
                    print(f"Error processing message: {str(e)}")
                    continue
            
        except Exception as e:
            print(f"Error connecting to email: {str(e)}")
        
        finally:
            try:
                mail.logout()
            except:
                pass
    
    def extract_expense_data(self, email_message):
        """Extract expense information from email."""
        # Basic extraction for testing
        subject = email_message['subject']
        date = datetime.now().strftime('%Y-%m-%d')
        
        return {
            'date': date,
            'description': subject,
            'amount': 0.0,  # You would implement amount extraction here
            'category': 'Other'
        }
    
    def add_to_sheet(self, expense_data):
        """Add expense data to Google Sheet."""
        spreadsheet_id = self.get_or_create_spreadsheet()
        range_name = f"{self.config['google_sheets']['worksheet_name']}!A:D"
        
        values = [[
            expense_data['date'],
            expense_data['description'],
            expense_data['amount'],
            expense_data['category']
        ]]
        
        body = {
            'values': values
        }
        
        self.sheets_service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
    
    def get_or_create_spreadsheet(self):
        """Get existing or create new expense tracking spreadsheet."""
        # Search for existing spreadsheet
        results = self.drive_service.files().list(
            q=f"name='{self.config['google_sheets']['spreadsheet_name']}' and mimeType='application/vnd.google-apps.spreadsheet'",
            spaces='drive',
            fields='files(id)'
        ).execute()
        
        if results.get('files'):
            return results['files'][0]['id']
        
        # Create new spreadsheet
        spreadsheet = {
            'properties': {
                'title': self.config['google_sheets']['spreadsheet_name']
            },
            'sheets': [{
                'properties': {
                    'title': self.config['google_sheets']['worksheet_name'],
                    'gridProperties': {
                        'frozenRowCount': 1
                    }
                }
            }]
        }
        
        spreadsheet = self.sheets_service.spreadsheets().create(body=spreadsheet).execute()
        
        # Format header row
        header = [['Date', 'Description', 'Amount', 'Category']]
        self.sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet['spreadsheetId'],
            range='A1:D1',
            valueInputOption='RAW',
            body={'values': header}
        ).execute()
        
        return spreadsheet['spreadsheetId']
    
    def run(self):
        """Main execution loop."""
        print(f"Starting expense tracker at {datetime.now()}")
        self.process_emails()

if __name__ == '__main__':
    load_dotenv()
    tracker = ExpenseTracker()
    tracker.run()