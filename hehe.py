import os
import re
import subprocess
import tempfile
import webbrowser
import tkinter as tk
import requests
from tkinter import filedialog, messagebox
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import io
from PIL import Image
import base64

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
        self.imgbb_api_key = tk.StringVar()
        
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
        
        # Row 3: ImgBB API Key
        tk.Label(self.master, text="ImgBB API Key:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        tk.Entry(self.master, textvariable=self.imgbb_api_key, width=50).grid(row=3, column=1, padx=5, pady=5)

        # Row 4: Convert button
        tk.Button(self.master, text="Convert & Upload", command=self.on_convert_click, width=20).grid(row=4, column=1, pady=15)

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
        imgbb_key = self.imgbb_api_key.get().strip()

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
        if not imgbb_key:
            messagebox.showerror("Error", "Please enter an ImgBB API key.")
            return

        # Convert docx -> .tex using pandoc
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tex_path = os.path.join(tmpdir, "converted.tex")
                cmd = ["pandoc", docx_file, "-o", tex_path, "--extract-media", tmpdir]
                subprocess.run(cmd, check=True)

                # Debug - list all files in the temp directory
                print("Files in temp directory:")
                for root, dirs, files in os.walk(tmpdir):
                    for file in files:
                        print(os.path.join(root, file))

                # Collect all images in the temp directory and subdirectories
                image_files = []
                for root, _, files in os.walk(tmpdir):
                    for file in files:
                        if file.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                            image_files.append(os.path.join(root, file))

                print(f"Found {len(image_files)} images")
                
                # Authenticate Google Sheets
                try:
                    creds = self.authenticate_google_sheets(creds_file, TOKEN_FILE)
                except Exception as e:
                    messagebox.showerror("Authentication Error", f"Failed to authenticate:\n{e}")
                    return
                
                # Upload images to ImgBB and get URLs
                image_map = self.upload_images_to_imgbb(image_files, tmpdir, imgbb_key)
                if image_files and not image_map:
                    messagebox.showerror("Error", "Failed to process images.")
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

        # Create sheet & write data
        try:
            self.write_to_google_sheets(mcq_data, creds, sheet_name_input)
        except Exception as e:
            messagebox.showerror("Sheets Error", f"Failed to upload data:\n{e}")
            return

        messagebox.showinfo("Success", f"Uploaded MCQs to Google Sheet: {sheet_name_input}")

    def upload_images_to_imgbb(self, image_paths, tmpdir, api_key):
        """Upload images to ImgBB and return a map of image URLs."""
        image_map = {}
        
        for img_path in image_paths:
            try:
                # Processed image path
                processed_img_path = os.path.join(tmpdir, f"processed_{os.path.basename(img_path)}")
                
                with Image.open(img_path) as img:
                    # Convert to RGB if needed
                    if img.mode == 'RGBA' and not img_path.lower().endswith('.png'):
                        background = Image.new('RGB', img.size, (255, 255, 255))
                        background.paste(img, (0, 0), img)
                        img = background
                    
                    # Resize if necessary
                    max_dimension = 800
                    if img.width > max_dimension or img.height > max_dimension:
                        if img.width > img.height:
                            new_width = max_dimension
                            new_height = int(img.height * (max_dimension / img.width))
                        else:
                            new_height = max_dimension
                            new_width = int(img.width * (max_dimension / img.height))
                        img = img.resize((new_width, new_height), Image.LANCZOS)
                    
                    # Save with compression
                    if img_path.lower().endswith('.png'):
                        img.save(processed_img_path, format='PNG', optimize=True)
                    else:
                        img.save(processed_img_path, format='JPEG', quality=70, optimize=True)
                
                # Check file size - if still too large, compress more
                file_size = os.path.getsize(processed_img_path)
                if file_size > 10240:  # 10KB
                    with Image.open(processed_img_path) as img:
                        # Reduce dimensions more
                        scale_factor = min(1.0, (10240 / file_size) ** 0.5) * 0.9  # Added safety margin
                        new_width = int(img.width * scale_factor)
                        new_height = int(img.height * scale_factor)
                        img = img.resize((new_width, new_height), Image.LANCZOS)
                        
                        # Save with more compression
                        if processed_img_path.lower().endswith('.png'):
                            img.save(processed_img_path, format='PNG', optimize=True)
                        else:
                            img.save(processed_img_path, format='JPEG', quality=50, optimize=True)
                
                # Upload to ImgBB
                with open(processed_img_path, "rb") as image_file:
                    # Convert image to base64
                    encoded_string = base64.b64encode(image_file.read())
                    
                    # Upload to ImgBB
                    url = "https://api.imgbb.com/1/upload"
                    payload = {
                        "key": api_key,
                        "image": encoded_string,
                        "name": os.path.basename(processed_img_path)
                    }
                    
                    response = requests.post(url, payload)
                    
                    if response.status_code == 200:
                        json_data = response.json()
                        if json_data["success"]:
                            # Get direct image URL - using display_url which is for embedding
                            image_url = json_data["data"]["display_url"]
                            
                            # Store mappings with multiple keys for better matching
                            rel_path = os.path.relpath(img_path, tmpdir).replace(os.path.sep, '/')
                            filename = os.path.basename(img_path)
                            media_path = f"media/{filename}"
                            
                            image_map[rel_path] = image_url
                            image_map[filename] = image_url
                            image_map[media_path] = image_url
                            
                            print(f"Added image mapping for: {rel_path}, {filename}, {media_path}")
                        else:
                            print(f"ImgBB API error: {json_data.get('error', {}).get('message', 'Unknown error')}")
                    else:
                        print(f"ImgBB API request failed with status code: {response.status_code}")
                
            except Exception as e:
                print(f"Error processing image {img_path}: {e}")
                continue
                
        return image_map

    def parse_latex_for_mcqs(self, latex_file, image_map):
        """Parse LaTeX content for MCQs with image handling."""
        with open(latex_file, "r", encoding="utf-8") as f:
            content = f.read()

        # Split into lines and process
        lines = content.split('\n')
        mcq_data = []
        current_q = {}
        
        re_question = re.compile(r'^(\d+)[.)]\s*(.*)')
        re_option = re.compile(r'^([ক-ঘ])[.)]\s+(.*)$')
        re_answer = re.compile(r'^উত্তর[:]\s+(.*)$')
        re_explanation = re.compile(r'^ব্যাখ্যা[:]\s+(.*)$')
        re_image = re.compile(r'\\includegraphics(?:\[.*?\])?\{(.*?)\}')
        
        in_explanation = False
        current_images = {'question': '', 'explanation': ''}

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # Check for images in line
            img_matches = re_image.findall(line)
            for img_path in img_matches:
                img_path = img_path.strip()
                
                # Try multiple keys for image matching
                for key in [img_path, os.path.basename(img_path), f"media/{os.path.basename(img_path)}"]:
                    if key in image_map:
                        img_url = image_map[key]
                        if in_explanation:
                            current_images['explanation'] = img_url
                        else:
                            current_images['question'] = img_url
                        print(f"Found image match for {key} in {'explanation' if in_explanation else 'question'}")
                        break
                else:
                    print(f"Could not find image match for {img_path}")

            # Question detection
            q_match = re_question.match(line)
            if q_match:
                if current_q:
                    # Add images to the current question before moving to the next
                    current_q['images'] = current_images
                    mcq_data.append(self.format_question(current_q))
                    current_images = {'question': '', 'explanation': ''}
                
                in_explanation = False
                current_q = {
                    'serial': q_match.group(1),
                    'question': q_match.group(2),
                    'options': {},
                    'answer': '',
                    'explanation': '',
                    'images': {'question': '', 'explanation': ''}
                }
                continue

            # Option detection
            opt_match = re_option.match(line)
            if opt_match and current_q:
                current_q['options'][opt_match.group(1)] = opt_match.group(2)
                continue

            # Answer detection
            ans_match = re_answer.match(line)
            if ans_match and current_q:
                current_q['answer'] = ans_match.group(1)
                continue

            # Explanation detection
            exp_match = re_explanation.match(line)
            if exp_match and current_q:
                current_q['explanation'] = exp_match.group(1)
                in_explanation = True
                continue

            # Accumulate question or explanation text
            if current_q:
                if in_explanation:
                    current_q['explanation'] += ' ' + line
                else:
                    current_q['question'] += ' ' + line

        # Process the last question
        if current_q:
            current_q['images'] = current_images
            mcq_data.append(self.format_question(current_q))

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

    def format_question(self, q):
        """Format question dictionary into sheet row."""
        # Extract topic and board info from the question
        base_question, board, topic = self.parse_bracket_tokens(q['question'])
        
        # Process equations
        base_question = self.convert_inline_equations_to_unicode(base_question)
        explanation = self.convert_inline_equations_to_unicode(q['explanation'])
        
        # Clean LaTeX commands like \includegraphics from text
        base_question = re.sub(r'\\includegraphics(?:\[.*?\])?\{.*?\}', '', base_question)
        explanation = re.sub(r'\\includegraphics(?:\[.*?\])?\{.*?\}', '', explanation)
        
        # Use direct image URLs instead of =IMAGE() formulas
        question_img_url = q['images']['question']
        explanation_img_url = q['images']['explanation']
        
        return [
            q['serial'],
            "", "", "", "",  # Empty columns for class flags
            base_question.strip(),
            topic, board,  # Topic and board 
            q['options'].get('ক', ''),
            q['options'].get('খ', ''),
            q['options'].get('গ', ''),
            q['options'].get('ঘ', ''),
            q['answer'],
            explanation.strip(),
            question_img_url,  # Direct URL instead of formula
            explanation_img_url  # Direct URL instead of formula
        ]

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
            valueInputOption="USER_ENTERED",
            body={"values": values}
        ).execute()

        webbrowser.open(f"https://docs.google.com/spreadsheets/d/{ssid}")

def main():
    root = tk.Tk()
    app = DocxToGsheetPandocGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()