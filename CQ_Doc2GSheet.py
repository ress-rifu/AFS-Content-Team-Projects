import os
import re
from docx import Document
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox

# ============================
# Configuration
# ============================

# If modifying these scopes, delete the file token.json
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# ============================
# Functions for Direct .docx Parsing and Google Sheets Interaction
# ============================

def parse_docx(docx_path):
    """Parses the Word document to extract Creative Questions with Answers."""
    print(f"Parsing the Word document: {docx_path}")
    document = Document(docx_path)
    cq_data = []
    serial_number = ''
    passage_text = ''
    board_institute = ''
    cq_questions = {'Question 1': '', 'Question 2': '', 'Question 3': '', 'Question 4': ''}
    cq_answers = {'Answer 1': '', 'Answer 2': '', 'Answer 3': '', 'Answer 4': ''}
    current_question = None  # Tracks the current question being processed
    current_answer = None    # Tracks the current answer being processed
    is_creative = False
    skip_next = False  # Flag to skip the next paragraph if it's already processed

    # Regular expressions
    # Pattern to detect Creative Questions: starts with "প্রশ্ন" followed by serial number
    cq_start_pattern = re.compile(r'^প্রশ্ন\s+([০-৯]+)[.)]\s*(.*)', re.UNICODE)
    # Pattern to detect questions: starts with ক.), খ.), গ.), ঘ.)
    question_pattern = re.compile(r'^([ক-ঘ])\.\s*(.*)', re.UNICODE)
    # Pattern to detect answers: starts with উত্তর (ক)., উত্তর (খ)., উত্তর (গ)., উত্তর (ঘ).
    answer_pattern = re.compile(r'^উত্তর\s*\(([ক-ঘ])\)\.\s*(.*)', re.UNICODE)
    # Pattern to detect Board/Institute: enclosed in square brackets
    board_institute_pattern = re.compile(r'^\[(.*?)\]$')

    paragraphs = document.paragraphs
    total_paragraphs = len(paragraphs)
    i = 0

    while i < total_paragraphs:
        para = paragraphs[i]
        text = para.text.strip()
        if not text:
            i += 1
            continue

        print(f"Processing paragraph {i+1}/{total_paragraphs}: {text}")

        # Check for Creative Question start
        cq_start_match = cq_start_pattern.match(text)
        if cq_start_match:
            # If there's an ongoing CQ, save it first
            if is_creative and serial_number:
                cq_data.append([
                    serial_number,
                    passage_text.strip(),
                    board_institute.strip(),
                    cq_questions['Question 1'].strip(),
                    cq_answers['Answer 1'].strip(),
                    cq_questions['Question 2'].strip(),
                    cq_answers['Answer 2'].strip(),
                    cq_questions['Question 3'].strip(),
                    cq_answers['Answer 3'].strip(),
                    cq_questions['Question 4'].strip(),
                    cq_answers['Answer 4'].strip()
                ])
                print(f"Appended Creative Question {serial_number}")

            # Start new Creative Question
            serial_number = f"প্রশ্ন {cq_start_match.group(1)}"
            remaining_text = cq_start_match.group(2).strip()

            # Extract Passage and possibly Board/Institute if present
            board_institute_match = board_institute_pattern.match(remaining_text)
            if board_institute_match:
                board_institute = board_institute_match.group(1)
                passage_text = replace_latex_delimiters('')
            else:
                # Passage might be followed by Board/Institute in the next paragraph
                passage_text = replace_latex_delimiters(remaining_text)
                board_institute = ''
                # Check if next paragraph contains [Board or Institute]
                if (i + 1) < total_paragraphs:
                    next_para = paragraphs[i + 1].text.strip()
                    board_match = board_institute_pattern.match(next_para)
                    if board_match:
                        board_institute = board_match.group(1)
                        skip_next = True  # Skip processing this in the next iteration
                        print(f"Detected Board/Institute: {board_institute}")
            is_creative = True
            cq_questions = {'Question 1': '', 'Question 2': '', 'Question 3': '', 'Question 4': ''}
            cq_answers = {'Answer 1': '', 'Answer 2': '', 'Answer 3': '', 'Answer 4': ''}
            current_question = None
            current_answer = None
            i += 1
            if skip_next:
                i += 1
                skip_next = False
            continue

        # Check for question
        question_match = question_pattern.match(text)
        if question_match and is_creative:
            key = question_match.group(1)
            question_text = replace_latex_delimiters(question_match.group(2).strip())
            question_number = {'ক':1, 'খ':2, 'গ':3, 'ঘ':4}[key]
            cq_questions[f"Question {question_number}"] = question_text
            current_question = f"Question {question_number}"
            current_answer = f"Answer {question_number}"
            print(f"Detected Creative Question {serial_number} - Question {key}: {question_text}")
            i += 1
            continue

        # Check for answer
        answer_match = answer_pattern.match(text)
        if answer_match and is_creative:
            key = answer_match.group(1)
            answer_text = replace_latex_delimiters(answer_match.group(2).strip())
            answer_number = {'ক':1, 'খ':2, 'গ':3, 'ঘ':4}[key]
            cq_answers[f"Answer {answer_number}"] = answer_text
            print(f"Detected Creative Question {serial_number} - Answer ({key}): {answer_text}")
            current_answer = f"Answer {answer_number}"
            i += 1
            continue

        # If none of the above, it's either part of the passage, questions, or answers
        if is_creative and serial_number:
            # Determine if we're appending to a question or an answer
            if current_answer:
                # Append to the current answer if it's not empty
                if cq_answers[current_answer]:
                    # Append to existing answer
                    additional_text = replace_latex_delimiters(text)
                    cq_answers[current_answer] += ' ' + additional_text
                    print(f"Appending to Creative Question {serial_number} - {current_answer}: {text}")
                else:
                    # Set as the current answer
                    cq_answers[current_answer] = replace_latex_delimiters(text)
                    print(f"Setting Creative Question {serial_number} - {current_answer}: {text}")
            elif current_question:
                # Append to the current question
                additional_text = replace_latex_delimiters(text)
                cq_questions[current_question] += ' ' + additional_text
                print(f"Appending to Creative Question {serial_number} - {current_question}: {text}")
            else:
                # Append to the passage text
                passage_text += ' ' + replace_latex_delimiters(text)
                print(f"Appending to Creative Question {serial_number} - Passage: {text}")
        i += 1

    # After processing all paragraphs, append the last CQ
    if is_creative and serial_number:
        cq_data.append([
            serial_number,
            passage_text.strip(),
            board_institute.strip(),
            cq_questions['Question 1'].strip(),
            cq_answers['Answer 1'].strip(),
            cq_questions['Question 2'].strip(),
            cq_answers['Answer 2'].strip(),
            cq_questions['Question 3'].strip(),
            cq_answers['Answer 3'].strip(),
            cq_questions['Question 4'].strip(),
            cq_answers['Answer 4'].strip()
        ])
        print(f"Appended Creative Question {serial_number}")

    print(f"Total Creative Questions parsed: {len(cq_data)}")
    return cq_data

def replace_latex_delimiters(text):
    """Cleans the text by removing or replacing LaTeX-specific commands."""
    # Replace inline math delimiters
    text = re.sub(r'\\\((.*?)\\\)', r'$\1$', text)

    # Remove \textquotesingle and similar LaTeX commands
    text = re.sub(r'\\textquotesingle', '', text)
    text = re.sub(r'\\textbf\{(.*?)\}', r'\1', text)  # Remove \textbf{}
    text = re.sub(r'\\textit\{(.*?)\}', r'\1', text)  # Remove \textit{}
    text = re.sub(r'\\[^ ]+\{([^}]*)\}', r'\1', text)  # Remove other LaTeX commands with braces
    text = re.sub(r'\\[^ ]+', '', text)  # Remove other LaTeX commands without braces
    # Add more substitutions as needed

    # Remove extra spaces
    text = re.sub(r'\s+', ' ', text)

    return text.strip()

def authenticate_google_sheets(credentials_path, token_path):
    """Authenticates with Google Sheets API using OAuth2."""
    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        print("Loaded credentials from token file.")
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            print("Refreshed expired credentials.")
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
            print("Obtained new credentials via OAuth flow.")
        with open(token_path, 'w', encoding='utf-8') as token:
            token.write(creds.to_json())
            print("Saved credentials to token file.")
    return creds

def write_to_google_sheets(cq_data, creds):
    """Creates a new Google Sheet and writes the Creative Questions data."""
    service = build('sheets', 'v4', credentials=creds)

    # Create a new spreadsheet
    spreadsheet = {
        'properties': {
            'title': 'Creative Questions Repository'
        }
    }
    try:
        spreadsheet = service.spreadsheets().create(body=spreadsheet, fields='spreadsheetId').execute()
        spreadsheet_id = spreadsheet.get('spreadsheetId')
        print(f"Created new spreadsheet with ID: {spreadsheet_id}")
    except Exception as e:
        print(f"Error creating spreadsheet: {e}")
        return

    # Prepare data for Creative Questions
    if cq_data:
        cq_values = [
            ['Serial', 'Passage', 'Board/Institute', 
             'Question 1', 'Answer 1', 
             'Question 2', 'Answer 2', 
             'Question 3', 'Answer 3', 
             'Question 4', 'Answer 4']
        ] + cq_data
        print("Sample Creative Questions data to be written to Google Sheets:")
        for row in cq_data[:5]:
            print(row)
    else:
        cq_values = []
        print("No Creative Questions to write.")

    # Add a new sheet for Creative Questions
    if cq_data:
        requests = [{
            'addSheet': {
                'properties': {
                    'title': 'Creative Questions'
                }
            }
        }]
        try:
            batch_update_request = {'requests': requests}
            response = service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=batch_update_request
            ).execute()
            print("Added new sheet for Creative Questions.")
        except Exception as e:
            print(f"Error adding sheet: {e}")
            return

        # Write Creative Questions data
        try:
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range='Creative Questions!A1',
                valueInputOption='RAW',
                body={'values': cq_values}
            ).execute()
            print("Creative Questions data written successfully.")
        except Exception as e:
            print(f"Error writing Creative Questions data: {e}")

    # Open the spreadsheet in the default browser
    sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
    webbrowser.open(sheet_url)
    print("Opened the spreadsheet in your default browser.")

# ============================
# GUI Application
# ============================

class DocxToGsheetApp:
    def __init__(self, master):
        self.master = master
        master.title("Docx to Google Sheets Converter (CQs Only)")

        self.docx_path = tk.StringVar()
        self.credentials_path = tk.StringVar()
        self.token_path = 'token.json'  # Fixed token path

        # Docx File Selection
        self.docx_label = tk.Label(master, text="Word Document:")
        self.docx_label.grid(row=0, column=0, padx=5, pady=10, sticky="w")

        self.docx_entry = tk.Entry(master, textvariable=self.docx_path, width=50)
        self.docx_entry.grid(row=0, column=1, padx=5, pady=10)

        self.docx_button = tk.Button(master, text="Browse", command=self.browse_docx)
        self.docx_button.grid(row=0, column=2, padx=5, pady=10)

        # Credentials File Selection
        self.credentials_label = tk.Label(master, text="Credentials File (JSON):")
        self.credentials_label.grid(row=1, column=0, padx=5, pady=10, sticky="w")

        self.credentials_entry = tk.Entry(master, textvariable=self.credentials_path, width=50)
        self.credentials_entry.grid(row=1, column=1, padx=5, pady=10)

        self.credentials_button = tk.Button(master, text="Browse", command=self.browse_credentials)
        self.credentials_button.grid(row=1, column=2, padx=5, pady=10)

        # Convert Button
        self.convert_button = tk.Button(master, text="Convert to Google Sheets", command=self.convert, width=25)
        self.convert_button.grid(row=2, column=1, padx=5, pady=20)

    def browse_docx(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        self.docx_path.set(filename)

    def browse_credentials(self):
        filename = filedialog.askopenfilename(
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        self.credentials_path.set(filename)

    def convert(self):
        docx_file = self.docx_path.get()
        credentials_file = self.credentials_path.get()

        if not docx_file or not credentials_file:
            messagebox.showerror("Error", "Please select both a Word document and a credentials file.")
            return

        if not os.path.exists(docx_file):
            messagebox.showerror("Error", "The selected Word document does not exist.")
            return

        if not os.path.exists(credentials_file):
            messagebox.showerror("Error", "The selected credentials file does not exist.")
            return

        try:
            cq_data = parse_docx(docx_file)
        except Exception as e:
            print(f"Error during parsing: {e}")
            messagebox.showerror("Error", f"Failed to parse the Word document.\n\n{e}")
            return

        if not cq_data:
            messagebox.showerror("Error", "No Creative Questions were parsed from the document.")
            return

        try:
            creds = authenticate_google_sheets(credentials_file, self.token_path)
        except Exception as e:
            print(f"Error during authentication: {e}")
            messagebox.showerror("Error", f"Failed to authenticate with Google Sheets.\n\n{e}")
            return

        try:
            write_to_google_sheets(cq_data, creds)
        except Exception as e:
            print(f"Error during writing to Google Sheets: {e}")
            messagebox.showerror("Error", f"Failed to write data to Google Sheets.\n\n{e}")
            return

        messagebox.showinfo("Success", "Successfully converted and uploaded Creative Questions to Google Sheets!")

# ============================
# Main Execution
# ============================

def main():
    root = tk.Tk()
    app = DocxToGsheetApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
