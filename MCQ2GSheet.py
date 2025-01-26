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
    """Parses the Word document to extract MCQs."""
    print(f"Parsing the Word document: {docx_path}")
    document = Document(docx_path)
    mcq_data = []
    serial_numbers_seen = set()
    serial_number = ''
    question_text = ''
    options = {'ক': '', 'খ': '', 'গ': '', 'ঘ': ''}
    answer = ''
    board_institute = ''
    current_option = None  # Tracks the current option being processed
    expect_board_institute = False  # Flag to capture Board/Institute for MCQs

    # Regular expressions
    # Pattern to detect MCQ question number: starts with Bengali numerals followed by ., ), বা ।
    mcq_question_pattern = re.compile(r'^([০-৯]+)[.,)।]\s*(.*)', re.UNICODE)
    # Pattern to detect main options: starts with ক), খ), গ), ঘ) or ক., খ., গ., ঘ.
    option_pattern_main = re.compile(r'^([ক-ঘ])[).]\s*(.*)', re.UNICODE)
    # Pattern to detect sub-options: starts with i., ii., iii., etc.
    option_pattern_sub = re.compile(r'^([iI]{1,3})[.)]\s*(.*)', re.UNICODE)
    # Pattern to detect answers: starts with উত্তর: or উত্তরঃ
    answer_pattern = re.compile(r'^উত্তর[:ঃ]\s*(.*)', re.UNICODE)
    # Pattern to detect Board/Institute: enclosed in square brackets [ ]
    board_institute_pattern = re.compile(r'^\[(.*?)\]$', re.UNICODE)

    for idx, para in enumerate(document.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        print(f"Processing paragraph {idx+1}: {text}")

        # Check if the previous paragraph expects Board/Institute for MCQ
        if expect_board_institute:
            board_institute_match = board_institute_pattern.match(text)
            if board_institute_match:
                board_institute = replace_latex_delimiters(board_institute_match.group(1).strip())
                print(f"Captured Board/Institute for MCQ {serial_number}: {board_institute}")
            else:
                board_institute = ''
                print(f"No Board/Institute found for MCQ {serial_number}, leaving it empty.")
                # Since this paragraph is not Board/Institute, process it normally
            expect_board_institute = False
            if not board_institute_match:
                # Continue processing this paragraph as it might be part of the question or options
                pass
            else:
                # If Board/Institute was captured, skip further processing of this paragraph
                continue

        # Check for MCQ question number
        mcq_question_match = mcq_question_pattern.match(text)
        if mcq_question_match:
            # If there's an ongoing question, save it first
            if serial_number:
                mcq_data.append([
                    serial_number,
                    question_text.strip(),
                    board_institute.strip(),
                    options['ক'].strip(),
                    options['খ'].strip(),
                    options['গ'].strip(),
                    options['ঘ'].strip(),
                    answer.strip()
                ])
                print(f"Appended MCQ {serial_number}")

            # Start new MCQ
            serial_number = mcq_question_match.group(1)
            question_text = replace_latex_delimiters(mcq_question_match.group(2).strip())
            board_institute = ''
            options = {'ক': '', 'খ': '', 'গ': '', 'ঘ': ''}
            answer = ''
            current_option = None
            serial_numbers_seen.add(serial_number)
            print(f"Detected MCQ {serial_number}: {question_text}")

            # Expect the next paragraph to be Board/Institute
            expect_board_institute = True
            continue

        # Check for answer in MCQ
        answer_match = answer_pattern.match(text)
        if answer_match:
            answer_text = replace_latex_delimiters(answer_match.group(1).strip())
            print(f"Found answer line: {text}")
            # Assign the answer exactly as it is after "উত্তরঃ "
            # If the answer is a single option label (e.g., 'ক'), map it to the corresponding option text
            single_option_match = re.match(r'^([ক-ঘ])$', answer_text)
            if single_option_match and options.get(single_option_match.group(1)):
                answer = options[single_option_match.group(1)].strip()
                print(f"Mapped Answer from option label '{single_option_match.group(1)}' to '{answer}'")
            else:
                # If the answer is not a single option label, assign it as-is
                answer = answer_text
                print(f"Assigned Answer: {answer}")
            continue

        # Check for main options (ক), খ), গ), ঘ)) or ক., খ., গ., ঘ.
        option_match_main = option_pattern_main.match(text)
        if option_match_main:
            key = option_match_main.group(1)
            opt_text = replace_latex_delimiters(option_match_main.group(2).strip())
            options[key] = opt_text
            print(f"Detected MCQ {serial_number} - Option {key}: {opt_text}")
            current_option = key  # Update current_option
            continue

        # Check for multiple main options in the same paragraph
        option_line_matches = re.findall(r'([ক-ঘ])[).]\s*([^ক-ঘ]*)', text, re.UNICODE)
        if option_line_matches:
            for opt in option_line_matches:
                key = opt[0]
                opt_text = replace_latex_delimiters(opt[1].strip())
                options[key] = opt_text
                print(f"Detected MCQ {serial_number} - Option {key}: {opt_text}")
                current_option = key  # Update current_option
            continue

        # Check for sub-options (i., ii., iii., etc.)
        option_match_sub = option_pattern_sub.match(text)
        if option_match_sub:
            # Treat sub-options as part of the question text
            sub_key = option_match_sub.group(1)
            sub_text = replace_latex_delimiters(option_match_sub.group(2).strip())
            question_text += f" {sub_key}. {sub_text}"
            print(f"Appending sub-option to MCQ {serial_number} - Question: {sub_key}. {sub_text}")
            continue

        # If none of the above, it's either part of the passage or multi-line option
        if serial_number:
            if current_option:
                # Append to the current MCQ option
                additional_text = replace_latex_delimiters(text)
                options[current_option] += ' ' + additional_text
                print(f"Appending to MCQ {serial_number} - Option {current_option}: {text}")
            else:
                # Append to the question text
                question_text += ' ' + replace_latex_delimiters(text)
                print(f"Appending to MCQ {serial_number} - Question: {text}")

    # After processing all paragraphs, append the last question
    if serial_number:
        mcq_data.append([
            serial_number,
            question_text.strip(),
            board_institute.strip(),
            options['ক'].strip(),
            options['খ'].strip(),
            options['গ'].strip(),
            options['ঘ'].strip(),
            answer.strip()
        ])
        print(f"Appended MCQ {serial_number}")

    print(f"Total MCQs parsed: {len(mcq_data)}")
    return mcq_data

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

def write_to_google_sheets(mcq_data, creds, sheet_name):
    """Creates a new Google Sheet with the specified name and writes the MCQ data to Sheet1."""
    service = build('sheets', 'v4', credentials=creds)

    # Create a new spreadsheet with the specified sheet name
    spreadsheet = {
        'properties': {
            'title': sheet_name
        }
    }
    try:
        spreadsheet = service.spreadsheets().create(body=spreadsheet, fields='spreadsheetId').execute()
        spreadsheet_id = spreadsheet.get('spreadsheetId')
        print(f"Created new spreadsheet '{sheet_name}' with ID: {spreadsheet_id}")
    except Exception as e:
        print(f"Error creating spreadsheet: {e}")
        return

    # Prepare data for MCQs
    if mcq_data:
        mcq_values = [
            ['Serial Number', 'Question', 'Board/Institute', 'Option ক', 'Option খ', 'Option গ', 'Option ঘ', 'Answer']
        ] + mcq_data
        print("Sample MCQ data to be written to Google Sheets:")
        for row in mcq_data[:5]:
            print(row)
    else:
        mcq_values = []
        print("No MCQs to write.")

    # Write MCQs data to Sheet1
    if mcq_data:
        try:
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range='Sheet1!A1',  # Write to Sheet1
                valueInputOption='RAW',
                body={'values': mcq_values}
            ).execute()
            print("MCQs data written successfully to Sheet1.")
        except Exception as e:
            print(f"Error writing MCQs data: {e}")

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
        master.title("Docx to Google Sheets Converter (MCQs Only)")

        self.docx_path = tk.StringVar()
        self.credentials_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
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

        # Sheet Name Input
        self.sheet_name_label = tk.Label(master, text="Google Sheet Name:")
        self.sheet_name_label.grid(row=2, column=0, padx=5, pady=10, sticky="w")

        self.sheet_name_entry = tk.Entry(master, textvariable=self.sheet_name, width=50)
        self.sheet_name_entry.grid(row=2, column=1, padx=5, pady=10)

        # Convert Button
        self.convert_button = tk.Button(master, text="Convert to Google Sheets", command=self.convert, width=25)
        self.convert_button.grid(row=3, column=1, padx=5, pady=20)

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
        sheet_name_input = self.sheet_name.get().strip()

        if not docx_file or not credentials_file:
            messagebox.showerror("Error", "Please select both a Word document and a credentials file.")
            return

        if not sheet_name_input:
            messagebox.showerror("Error", "Please enter a name for the Google Sheet.")
            return

        if not os.path.exists(docx_file):
            messagebox.showerror("Error", "The selected Word document does not exist.")
            return

        if not os.path.exists(credentials_file):
            messagebox.showerror("Error", "The selected credentials file does not exist.")
            return

        try:
            mcq_data = parse_docx(docx_file)
        except Exception as e:
            print(f"Error during parsing: {e}")
            messagebox.showerror("Error", f"Failed to parse the Word document.\n\n{e}")
            return

        if not mcq_data:
            messagebox.showerror("Error", "No MCQs were parsed from the document.")
            return

        try:
            creds = authenticate_google_sheets(credentials_file, self.token_path)
        except Exception as e:
            print(f"Error during authentication: {e}")
            messagebox.showerror("Error", f"Failed to authenticate with Google Sheets.\n\n{e}")
            return

        try:
            write_to_google_sheets(mcq_data, creds, sheet_name_input)
        except Exception as e:
            print(f"Error during writing to Google Sheets: {e}")
            messagebox.showerror("Error", f"Failed to write data to Google Sheets.\n\n{e}")
            return

        messagebox.showinfo("Success", f"Successfully converted and uploaded MCQs to Google Sheets named '{sheet_name_input}'!")

# ============================
# Main Execution
# ============================

def main():
    root = tk.Tk()
    app = DocxToGsheetApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
