import os
import re
import subprocess
import tempfile
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox
import base64
import requests
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
        tk.Button(self.master, text="Convert & Upload", command=self.on_convert_click, width=20).grid(row=3, column=1, pady=15)

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

                # Check for images in the temp directory
                image_files = []
                for entry in os.listdir(tmpdir):
                    if entry.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                        image_files.append(os.path.join(tmpdir, entry))

                image_map = {}
                if image_files:
                    image_map = self.upload_images_to_base64(image_files)
                    if not image_map:
                        messagebox.showerror("Error", "Failed to upload one or more images.")
                        return

                # Parse the generated .tex for MCQs
                mcq_data = self.parse_latex_for_mcqs(tex_path, image_map)
        except FileNotFoundError:
            messagebox.showerror("Pandoc Error", "pandoc not found. Please install pandoc.")
            return
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Pandoc Error", f"Pandoc failed to convert:\n{e}")
            return
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error during processing:\n{e}")
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

    def upload_images_to_base64(self, image_paths):
        """Convert images to base64 and return a dict of {filename: base64_str}."""
        image_map = {}
        for img_path in image_paths:
            try:
                with open(img_path, "rb") as f:
                    img_data = f.read()
            except Exception as e:
                print(f"Error reading {img_path}: {e}")
                continue  # Skip unreadable files

            b64_data = base64.b64encode(img_data).decode('utf-8')
            filename = os.path.basename(img_path)
            image_map[filename] = b64_data
        print("Image Map:", image_map)  # Debugging log
        return image_map

    def replace_image_commands(self, text, image_map):
        """Replace LaTeX includegraphics commands with base64-encoded image strings."""
        pattern = re.compile(r'\\includegraphics(\[.*?\])?{([^}]+)}')
        def replacer(match):
            filename = match.group(2).strip()
            base64_str = image_map.get(filename, "Upload failed")
            print(f"Replacing image: {filename} -> {base64_str}")  # Debugging log
            return f' [Image: {base64_str} ] '
        return pattern.sub(replacer, text)

    def parse_latex_for_mcqs(self, latex_file, image_map):
        """Parse LaTeX content for MCQs, replacing images using image_map."""
        with open(latex_file, "r", encoding="utf-8") as f:
            lines = f.readlines()

        # Match question serial number and question
        re_question = re.compile(r'^([\u09E6-\u09EF0-9]+)[.,):।]\s*(.*)', re.UNICODE)
        # Match options
        re_option = re.compile(r'^([ক-ঘ])[.)]\s+(.*)$', re.UNICODE)
        # Match answers
        re_answer = re.compile(r'^উত্তর[:ঃ]\s+(.*)$', re.UNICODE)
        # Match explanations
        re_explanation = re.compile(r'^ব্যাখ্যাঃ\s+(.*)$', re.UNICODE)

        mcq_data = []
        serial_number = ""
        question_text = ""
        board_institute = ""
        topic = ""
        options = {"ক": "", "খ": "", "গ": "", "ঘ": ""}
        answer = ""
        explanation = ""
        question_image = ""
        option_images = {"ক": "", "খ": "", "গ": "", "ঘ": ""}
        explanation_image = ""

        def commit_mcq():
            if serial_number:
                mcq_data.append([
                    serial_number.strip(),
                    "", "", "", "",
                    question_text.strip(),
                    topic.strip(),
                    board_institute.strip(),
                    option_images["ক"],
                    option_images["খ"],
                    option_images["গ"],
                    option_images["ঘ"],
                    answer.strip(),
                    explanation.strip(),
                    question_image,
                    explanation_image
                ])

        for line in lines:
            text = line.strip()
            if not text:
                continue

            # Process images and equations
            text = self.replace_image_commands(text, image_map)
            text = self.convert_inline_equations_to_unicode(text)

            # Question match
            mq = re_question.match(text)
            if mq:
                if serial_number:
                    commit_mcq()
                serial_number = mq.group(1)
                q_text = mq.group(2)
                base_text, new_board, new_topic = self.parse_bracket_tokens(q_text)
                question_text = base_text
                board_institute = new_board
                topic = new_topic
                options = {"ক": "", "খ": "", "গ": "", "ঘ": ""}
                answer = ""
                explanation = ""
                question_image = ""
                explanation_image = ""
                continue

            # Option match
            if serial_number:
                mo = re_option.match(text)
                if mo:
                    option_letter = mo.group(1)
                    option_text = mo.group(2)
                    option_images[option_letter] = self.replace_image_commands(option_text, image_map)
                    options[option_letter] = option_text
                    continue

                # Answer match
                ma = re_answer.match(text)
                if ma:
                    answer = ma.group(1)
                    continue

                # Explanation match
                me = re_explanation.match(text)
                if me:
                    explanation = me.group(1)
                    continue

                question_text += " " + text

        if serial_number:
            commit_mcq()

        return mcq_data

    def parse_bracket_tokens(self, text):
        """Extracts topic and board/institute info from text inside brackets."""
        bracket_pattern = re.compile(r'\{\[}(.*?)\{]}')
        matches = bracket_pattern.findall(text)
        base_text = bracket_pattern.sub('', text).strip()

        board_institute = ""
        topic = ""
        for m in matches:
            stripped_m = m.strip()
            if stripped_m.startswith("টপিক:"):
                topic = stripped_m.replace("টপিক:", "").strip()
            else:
                board_institute = stripped_m
        return base_text, board_institute, topic

    def convert_inline_equations_to_unicode(self, text):
        """Convert inline LaTeX equations to Unicode (e.g., math symbols)."""
        pattern = re.compile(r'(\${1,2})(.*?)(\1)', re.DOTALL)
        def replacer(m):
            return self.convert_latex_to_unicode(m.group(2).strip())
        return pattern.sub(replacer, text)

    def convert_latex_to_unicode(self, eq_text):
        """Convert LaTeX math symbols to Unicode."""
        greek_map = {
            r'\\alpha': 'α', r'\\beta': 'β', r'\\gamma': 'γ',
            r'\\delta': 'δ', r'\\theta': 'θ', r'\\mu': 'μ',
            r'\\pi': 'π', r'\\sigma': 'σ', r'\\phi': 'φ',
            r'\\omega': 'ω'
        }
        for latex_g, uni_g in greek_map.items():
            eq_text = eq_text.replace(latex_g, uni_g)

        eq_text = re.sub(r'([A-Za-z0-9])\^([A-Za-z0-9])',
            lambda m: m.group(1) + self.to_superscript(m.group(2)), eq_text)
        eq_text = re.sub(r'([A-Za-z0-9])_([A-Za-z0-9])',
            lambda m: m.group(1) + self.to_subscript(m.group(2)), eq_text)

        eq_text = eq_text.replace(r'\times', '×').replace(r'\cdot', '·')
        eq_text = eq_text.replace(r'\pm', '±').replace(r'\approx', '≈')
        eq_text = eq_text.replace(r'\neq', '≠').replace(r'=', ' = ')
        return eq_text.strip()

    def to_superscript(self, char):
        """Convert character to superscript."""
        supers = {'0':'⁰','1':'¹','2':'²','3':'³','4':'⁴','5':'⁵','6':'⁶','7':'⁷','8':'⁸','9':'⁹','n':'ⁿ','i':'ⁱ','+':'⁺','-':'⁻'}
        return supers.get(char, f'^{char}')

    def to_subscript(self, char):
        """Convert character to subscript."""
        subs = {'0':'₀','1':'₁','2':'₂','3':'₃','4':'₄','5':'₅','6':'₆','7':'₇','8':'₈','9':'₉','+':'₊','-':'₋','=':'₌','(':'₍',')':'₎'}
        return subs.get(char, f'_{char}')

    def authenticate_google_sheets(self, creds_file, token_file):
        creds = None
        if os.path.exists(token_file):
            creds = Credentials.from_authorized_user_file(token_file, SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(creds_file, SCOPES)
                creds = flow.run_local_server(port=0)
            with open(token_file, 'w') as tfile:
                tfile.write(creds.to_json())
        return creds

    def write_to_google_sheets(self, mcq_data, creds, sheet_name):
        service = build('sheets', 'v4', credentials=creds)
        spreadsheet_body = {"properties": {"title": sheet_name}}
        ss = service.spreadsheets().create(body=spreadsheet_body, fields="spreadsheetId").execute()
        ssid = ss.get("spreadsheetId")

        header = [
            "Serial", "For Class Slide", "For Lecture sheet", "For Quiz (Daily)",
            "For Quiz (Weekly)", "Question", "Topic", "Board/Inst", "Option ক",
            "Option খ", "Option গ", "Option ঘ", "Answer", "Explanation", 
            "QuestionIMG", "ExplanationIMG"
        ]
        values = [header] + mcq_data

        service.spreadsheets().values().update(
            spreadsheetId=ssid,
            range="Sheet1!A1",
            valueInputOption="RAW",
            body={"values": values}
        ).execute()

        webbrowser.open(f"https://docs.google.com/spreadsheets/d/{ssid}")

def main():
    root = tk.Tk()
    app = DocxToGsheetPandocGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
