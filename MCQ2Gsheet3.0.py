import os
import re
import subprocess
import tempfile
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox
# Google API imports
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

TOKEN_FILE = 'token.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

class DocxToGsheetPandocGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Docx to Google Sheets (with Pandoc)")

        # StringVars for user inputs
        self.docx_path = tk.StringVar()
        self.creds_path = tk.StringVar()
        self.sheet_name = tk.StringVar()

        # Layout
        self.create_widgets()

    def create_widgets(self):
        # Row 0: Word doc
        tk.Label(self.master, text="Word Document:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        tk.Entry(self.master, textvariable=self.docx_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self.master, text="Browse", command=self.browse_docx).grid(row=0, column=2, padx=5, pady=5)

        # Row 1: Credentials
        tk.Label(self.master, text="Credentials JSON:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        tk.Entry(self.master, textvariable=self.creds_path, width=50).grid(row=1, column=1, padx=5, pady=5)
        tk.Button(self.master, text="Browse", command=self.browse_credentials).grid(row=1, column=2, padx=5, pady=5)

        # Row 2: Sheet name
        tk.Label(self.master, text="Google Sheet Name:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        tk.Entry(self.master, textvariable=self.sheet_name, width=50).grid(row=2, column=1, padx=5, pady=5)

        # Row 3: Convert button
        tk.Button(self.master, text="Convert & Upload", command=self.on_convert_click, width=20)\
            .grid(row=3, column=1, pady=15)

    def browse_docx(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if file_path:
            self.docx_path.set(file_path)

    def browse_credentials(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        if file_path:
            self.creds_path.set(file_path)

    def on_convert_click(self):
        docx_file = self.docx_path.get().strip()
        creds_file = self.creds_path.get().strip()
        sheet_name_input = self.sheet_name.get().strip()

        # Basic checks
        if not docx_file or not os.path.exists(docx_file):
            messagebox.showerror("Error", "Please select a valid .docx file.")
            return
        if not creds_file or not os.path.exists(creds_file):
            messagebox.showerror("Error", "Please select a valid credentials JSON file.")
            return
        if not sheet_name_input:
            messagebox.showerror("Error", "Please enter a Google Sheet name.")
            return

        # Convert docx -> .tex using pandoc
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tex_path = os.path.join(tmpdir, "converted.tex")
                cmd = ["pandoc", docx_file, "-o", tex_path]
                subprocess.run(cmd, check=True)

                # Parse the generated .tex for MCQs in the desired structure
                mcq_data = self.parse_latex_for_mcqs(tex_path)
        except FileNotFoundError:
            messagebox.showerror("Pandoc Error", "pandoc not found. Please install pandoc.")
            return
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Pandoc Error", f"Pandoc failed to convert:\n{e}")
            return
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error during LaTeX parsing:\n{e}")
            return

        if not mcq_data:
            messagebox.showinfo("No MCQs", "No MCQs found in the document.")
            return

        # Authenticate Google Sheets
        try:
            creds = self.authenticate_google_sheets(creds_file, TOKEN_FILE)
        except Exception as e:
            messagebox.showerror("Authentication Error", f"Failed to authenticate:\n{e}")
            return

        # Create sheet & write data
        try:
            self.write_to_google_sheets(mcq_data, creds, sheet_name_input)
        except Exception as e:
            messagebox.showerror("Sheets Error", f"Failed to upload data:\n{e}")
            return

        messagebox.showinfo("Success", f"Uploaded MCQs to Google Sheet: {sheet_name_input}")

    # ---------------------------------------------------------------------
    # Helper Methods
    # ---------------------------------------------------------------------
    def parse_latex_for_mcqs(self, latex_file):
        """
        Reads the LaTeX line by line, searching for the structure:

            ১২১. <Question text...>
            ক. <Option text>
            খ. <Option text>
            গ. <Option text>
            ঘ. <Option text>
            উত্তর: <Answer text>

        We also convert $...$ (LaTeX math) to naive Unicode on each line.
        Returns a list of rows, each row = [
            serial_number,
            question_text,
            "",                 # (placeholder for 'board_institute' - not used)
            option_ক,
            option_খ,
            option_গ,
            option_ঘ,
            answer
        ].
        """

        with open(latex_file, "r", encoding="utf-8") as f:
            lines = f.readlines()

        # Regex patterns for your EXACT MCQ format
        # 1) question line: "১২১. some text" => group(1)=serial, group(2)=question
        re_question = re.compile(r'^([\u09E6-\u09EF0-9]+)[.,]\s+(.*)$', re.UNICODE)
        # 2) option lines: "ক. something" => group(1)=ক, group(2)=text
        re_option = re.compile(r'^([ক-ঘ])[.)]\s+(.*)$', re.UNICODE)
        # 3) answer line: "উত্তর: something"
        re_answer = re.compile(r'^উত্তর[:ঃ]\s+(.*)$', re.UNICODE)

        mcq_data = []

        # Temporary state
        serial_number = ""
        question_text = ""
        options = {"ক": "", "খ": "", "গ": "", "ঘ": ""}
        answer = ""

        def commit_mcq():
            """Append the current MCQ to mcq_data, then reset."""
            if serial_number:
                mcq_data.append([
                    serial_number.strip(),
                    question_text.strip(),
                    "",  # board_institute placeholder
                    options["ক"].strip(),
                    options["খ"].strip(),
                    options["গ"].strip(),
                    options["ঘ"].strip(),
                    answer.strip()
                ])

        # We'll read line-by-line, detect question or option or answer
        for line in lines:
            text = line.strip()
            if not text:
                continue

            # Convert any LaTeX inline math to naive Unicode
            text = self.convert_inline_equations_to_unicode(text)

            # Check if question line
            mq = re_question.match(text)
            if mq:
                # If we have an ongoing MCQ, commit it
                if serial_number:
                    commit_mcq()

                serial_number = mq.group(1)
                question_text = mq.group(2)
                options = {"ক": "", "খ": "", "গ": "", "ঘ": ""}
                answer = ""
                continue

            # Check if option line
            mo = re_option.match(text)
            if mo:
                letter = mo.group(1)
                opt_text = mo.group(2)
                options[letter] = opt_text
                continue

            # Check if answer line
            ma = re_answer.match(text)
            if ma:
                answer_text = ma.group(1)
                answer = answer_text
                continue

            # Otherwise, assume it's extra text continuing the question
            if serial_number:
                question_text += " " + text

        # End of file => commit last MCQ
        if serial_number:
            commit_mcq()

        return mcq_data

    def convert_inline_equations_to_unicode(self, text):
        """
        Finds $...$ or $$...$$ blocks in 'text' and converts the inside to naive Unicode.
        """
        pattern = re.compile(r'(\${1,2})(.*?)(\1)', re.DOTALL)

        def replacer(m):
            eq_content = m.group(2).strip()
            return self.convert_latex_to_unicode(eq_content)

        return pattern.sub(replacer, text)

    def convert_latex_to_unicode(self, eq_text):
        """
        Naive approach: \alpha->α, x^2->x², x_2->x₂, etc.
        """
        greek_map = {
            r'\\alpha': 'α',
            r'\\beta': 'β',
            r'\\gamma': 'γ',
            r'\\delta': 'δ',
            r'\\theta': 'θ',
            r'\\mu': 'μ',
            r'\\pi': 'π',
            r'\\sigma': 'σ',
            r'\\phi': 'φ',
            r'\\omega': 'ω'
        }
        for latex_g, uni_g in greek_map.items():
            eq_text = eq_text.replace(latex_g, uni_g)

        # x^2 -> x²
        eq_text = re.sub(
            r'([A-Za-z0-9])\^([A-Za-z0-9])',
            lambda m: m.group(1) + self.to_superscript(m.group(2)),
            eq_text
        )
        # x_2 -> x₂
        eq_text = re.sub(
            r'([A-Za-z0-9])_([A-Za-z0-9])',
            lambda m: m.group(1) + self.to_subscript(m.group(2)),
            eq_text
        )

        # Common math symbols
        eq_text = eq_text.replace(r'\times', '×')
        eq_text = eq_text.replace(r'\cdot', '·')
        eq_text = eq_text.replace(r'\pm', '±')
        eq_text = eq_text.replace(r'\approx', '≈')
        eq_text = eq_text.replace(r'\neq', '≠')
        # optional spacing for '='
        eq_text = eq_text.replace(r'=', ' = ')

        return eq_text.strip()

    def to_superscript(self, char):
        supers = {
            '0': '⁰', '1': '¹', '2': '²', '3': '³',
            '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷',
            '8': '⁸', '9': '⁹',
            'n': 'ⁿ', 'i': 'ⁱ', '+': '⁺', '-': '⁻'
        }
        return supers.get(char, '^' + char)

    def to_subscript(self, char):
        subs = {
            '0': '₀', '1': '₁', '2': '₂', '3': '₃',
            '4': '₄', '5': '₅', '6': '₆', '7': '₇',
            '8': '₈', '9': '₉',
            '+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎'
        }
        return subs.get(char, '_' + char)

    def authenticate_google_sheets(self, creds_file, token_file):
        """Authenticate with Google Sheets, storing/using a local token."""
        creds = None
        if os.path.exists(token_file):
            creds = Credentials.from_authorized_user_file(token_file, SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(creds_file, SCOPES)
                creds = flow.run_local_server(port=0)
            with open(token_file, 'w', encoding='utf-8') as tfile:
                tfile.write(creds.to_json())
        return creds

    def write_to_google_sheets(self, mcq_data, creds, sheet_name):
        """Create a new Google Sheet, write MCQs, open in browser."""
        service = build('sheets', 'v4', credentials=creds)

        # Create
        spreadsheet_body = {"properties": {"title": sheet_name}}
        ss = service.spreadsheets().create(body=spreadsheet_body, fields="spreadsheetId").execute()
        ssid = ss.get("spreadsheetId")

        # Prepare data
        header = ["Serial", "Question", "Board/Inst", "Option ক", "Option খ", "Option গ", "Option ঘ", "Answer"]
        values = [header] + mcq_data

        # Write
        service.spreadsheets().values().update(
            spreadsheetId=ssid,
            range="Sheet1!A1",
            valueInputOption="RAW",
            body={"values": values}
        ).execute()

        # Open
        url = f"https://docs.google.com/spreadsheets/d/{ssid}"
        webbrowser.open(url)


def main():
    root = tk.Tk()
    app = DocxToGsheetPandocGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
