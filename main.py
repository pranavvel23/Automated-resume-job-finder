import tkinter as tk
from tkinter import filedialog, messagebox
import re
import fitz  # PyMuPDF
import pandas as pd
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os

# -------------------------------
# Resume Parsing Functions
# -------------------------------

def parse_resume(resume_path):
    doc = fitz.open(resume_path)
    return "".join(page.get_text() for page in doc)

def extract_email(resume_text):
    match = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", resume_text)
    return match.group(0) if match else None

def extract_name(resume_text):
    lines = resume_text.split("\n")
    for line in lines:
        if len(line.strip().split()) in [2, 3] and not line.strip().isdigit():
            return line.strip()
    return "Candidate"

def extract_skills(resume_text):
    skills_list = ["Python", "Java", "SQL", "Machine Learning", "Communication", "Teamwork"]
    return [skill for skill in skills_list if skill.lower() in resume_text.lower()]

def load_jobs(path):
    return pd.read_excel(path)

def match_jobs(skills, jobs_df):
    return jobs_df[jobs_df['key_skills'].apply(
        lambda job_skills: any(skill.lower() in str(job_skills).lower() for skill in skills)
    )][['company', 'job_role', 'location', 'key_skills']]

# -------------------------------
# PDF Report Generator
# -------------------------------

class PDFReport(FPDF):
    def header(self):
        self.set_font("Arial", "B", 16)
        self.set_text_color(30, 30, 60)
        self.cell(0, 15, "Job Recommendations", ln=True, align="C")
        self.set_draw_color(100, 100, 150)
        self.set_line_width(0.8)
        self.line(10, 27, 200, 27)
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"Page {self.page_no()}", 0, 0, "C")

def clean_text(text):
    return text.replace('\u2013', '-').replace('\u2014', '-').encode('latin-1', 'replace').decode('latin-1')

def generate_pdf(data, name, output):
    pdf = PDFReport()
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, f"Job Recommendations for {name}", ln=True, align="C")
    pdf.ln(5)
    pdf.set_font("Arial", "", 12)

    for _, row in data.iterrows():
        pdf.set_font("Arial", "B", 12)
        pdf.set_text_color(0, 102, 204)
        pdf.cell(0, 10, clean_text(str(row['company'])), ln=True)

        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", "", 11)
        pdf.cell(0, 8, f"Job Title: {clean_text(str(row['job_role']))}", ln=True)
        pdf.cell(0, 8, f"Location: {clean_text(str(row['location']))}", ln=True)
        pdf.multi_cell(0, 8, f"Required Skills: {clean_text(str(row['key_skills']))}")
        pdf.ln(3)

    pdf.output(output)

# -------------------------------
# Email Sender
# -------------------------------

def send_email(receiver_email, pdf_path):
    SENDER_EMAIL = "samplebot.job@gmail.com"
    SENDER_PASSWORD = "hfiz sgci hqyp rknu"
  # ← Replace with your app password

    subject = "Your Job Recommendations Report"
    body = "Hello,\n\nPlease find attached your personalized job recommendations.\n\nBest Regards,\nJob Bot"

    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with open(pdf_path, "rb") as f:
        mime = MIMEBase('application', 'pdf')
        mime.set_payload(f.read())
        encoders.encode_base64(mime)
        mime.add_header('Content-Disposition', f'attachment; filename={os.path.basename(pdf_path)}')
        msg.attach(mime)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(SENDER_EMAIL, SENDER_PASSWORD)
    server.send_message(msg)
    server.quit()

# -------------------------------
# GUI Logic
# -------------------------------

def run_resume_process(file_path):
    try:
        resume_text = parse_resume(file_path)
        email = extract_email(resume_text)
        name = extract_name(resume_text)

        if not email:
            messagebox.showerror("No Email", "Could not find an email address in the resume.")
            return

        skills = extract_skills(resume_text)
        jobs_df = load_jobs("data/jobs_data.xlsx")
        matched = match_jobs(skills, jobs_df)

        if matched.empty:
            messagebox.showinfo("No Matches", "No matching jobs found for the candidate's skills.")
            return

        pdf_path = f"Job_Report_{name.replace(' ', '_')}.pdf"
        generate_pdf(matched, name, pdf_path)
        send_email(email, pdf_path)
        messagebox.showinfo("Success", f"Email sent to {email}")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# -------------------------------
# GUI Setup
# -------------------------------

def browse_resume():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        run_resume_process(file_path)

root = tk.Tk()
root.title("Resume Job Recommender")
root.geometry("400x200")
root.resizable(False, False)

label = tk.Label(root, text="Select a resume PDF to recommend jobs", font=("Arial", 12), pady=20)
label.pack()

button = tk.Button(root, text="Choose Resume", command=browse_resume, bg="#007acc", fg="white", padx=15, pady=8)
button.pack()

root.mainloop()
