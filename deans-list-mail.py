'''
Author: Caleb Harris
Title: IT Professional
Date: 7/19/23
Description: Send deans list letter documents to each user passed in from a csv file
customized for each user with a customized email for each person.
'''

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus import Image, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import fonts
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os
from getpass import getpass
import tkinter as tk
from tkinter import ttk, PhotoImage
from tkinter import filedialog, messagebox
from ttkthemes import ThemedTk
from tkinter import StringVar
import webbrowser
import csv
import smtplib, ssl
from threading import Thread
import queue

def create_pdf(name, gpa, address, award_type_value):

    fonts.addMapping('Times-Roman', 0, 0, 'Times-Roman')
    filename = name + ".pdf"
    pdf = canvas.Canvas(filename, pagesize=letter)
    pdf.drawImage("UCCS Logo Signature.png", x=inch, y=9 * inch, width=4 * 72, height=0.5 * inch)
    pdf.setFont("Times-Roman", 12)
    message_x = 1 * inch
    message_y = letter[1] - 2.4 * inch
    if (award_type_value == 'deans_list'):
        message_text = f"""July 13, 2023,<br/><br/><br/>
        {name}<br/>
        {address}<br/>
        Colorado Springs, CO 80907<br/><br/><br/>
        Dear {name.strip().split(" ")[0]},<br/><br/>
        Congratulations! I am most pleased to announce that you have been named to the Dean's List for the Spring 2023 semester.
        In order to receive this honor, a student must have earned between a 3.75-3.99 GPA, while completing at least 12 credit hours.
        Your GPA from Spring 2023 was {gpa}.<br/><br/>
        Your outstanding academic performance for this semester is a source of considerable pride for the College of Letters, Arts, and Sciences at the University of Colorado Colorado Springs.<br/><br/>
        Please accept my sincere congratulations for this well-deserved honor and my hope for your continued academic success. Keep up the great work.<br/><br/><br/>
        Yours Truly,"""
    elif (award_type_value == 'presidents_list'):
        message_text = f"""July 13, 2023,<br/><br/><br/>
        {name}<br/>
        {address}<br/>
        Colorado Springs, CO 80907<br/><br/><br/>
        Dear {name.strip().split(" ")[0]},<br/><br/>
        Congratulations! I am most pleased to announce that you have been named to the President's List for the Spring 2023 semester.
        In order to receive this honor, a student must have earned a 4.0 GPA, while completing at least 12 credit hours.
        Your GPA from Spring 2023 was {gpa}.<br/><br/>
        Your outstanding academic performance for this semester is a source of considerable pride for the College of Letters, Arts, and Sciences at the University of Colorado Colorado Springs.<br/><br/>
        Please accept my sincere congratulations for this well-deserved honor and my hope for your continued academic success. Keep up the great work.<br/><br/><br/>
        Yours Truly,"""
    else:
        raise Exception("Unknown radio option selection!")

    signature_text = """K. Alex Ilyasova<br/>
    Associate Dean<br/>
    College of Letters, Arts, and Sciences"""

    styles = getSampleStyleSheet()
    paragraph_style = styles["Normal"]
    paragraph_style.fontName = "Times-Roman"  # Set font to Times New Roman
    paragraph_style.fontSize = 12  # Set font size to 12
    paragraph_style.spaceBefore = 12
    paragraph_style.leading = 14.4  # Set line spacing to 1.15
    paragraph_style.spaceAfter = 13.86  # Set paragraph spacing after

    footer_text = """College of Letters, Arts & Sciences • Office of the Dean<br/>
    Columbine Hall 2025 • 1420 Austin Bluffs Pkwy • Colorado Springs, CO 80918<br/> 
    t 719-255-4550 • f 719-255-4200"""

    text_width = letter[0] - 2 * inch
    text_height = 4 * inch

    paragraph = Paragraph(message_text, style=paragraph_style)

    wrapped_text_height = paragraph.wrap(text_width, text_height)[1]

    start_y = message_y - wrapped_text_height

    paragraph.drawOn(pdf, message_x, start_y)

    signature_image_y = 3.33 * inch  
    signature_text_y = signature_image_y - 0.01 * inch  

    pdf.drawImage("deanslistlettersignature.jpg", x=inch, y=signature_image_y, width=2.07 * 72, height=0.82 * inch)

    signature_paragraph = Paragraph(signature_text, style=paragraph_style)

    signature_text_height = signature_paragraph.wrap(text_width, text_height)[1]

    signature_start_y = signature_text_y - signature_text_height

    signature_paragraph.drawOn(pdf, message_x, signature_start_y)

    paragraph_style.fontSize = 9.96
    paragraph_style.alignment = 1  # Center-justified

    footer_paragraph = Paragraph(footer_text, paragraph_style)

    footer_text_height = footer_paragraph.wrap(text_width, text_height)[1]

    footer_text_y = 0.75 * inch

    footer_paragraph.drawOn(pdf, 0.5 * (letter[0] - text_width), footer_text_y)

    pdf.save()
    return filename

def send_email(sender_email, password, receiver_email, name, gpa, address, award_type_value):

    if (award_type_value == 'deans_list'):
        type_value = "Dean's"
    else:
        type_value = "President's"

    # create email object
    message = MIMEMultipart("alternative")
    message["Subject"] = f"LAS {type_value} List Recipient | Spring2023 - DEVELOPMENT"
    message["From"] = sender_email
    message["To"] = receiver_email

    # create pdf file
    filename = create_pdf(name, gpa, address, award_type_value)

    # Attach the pdf
    with open(filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)

    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )

    message.attach(part)

    # Create HTML message
    html = f"""\
    <html>
    <body>
        <p>
            Hello {name},<br><br>
            Congratulations! The College of Letters, Arts & Sciences has named you a {type_value} List honoree for the Spring 2023 semester.<br><br>
            Associate Dean Ilyasova would like to commend you in the attached letter.<br><br>
            You've been featured in UCCS' Communique should you like to share the news with family and friends. https://communique.uccs.edu/?p=148841#LAS<br><br>

            -The Letters, Arts & Sciences Dean's Office Team
        </p>
    </body>
    </html>
    """

    # Turn these into plain/html MIMEText objects
    email_message = MIMEText(html, "html")

    # Add HTML to MIMEMultipart message
    message.attach(email_message)

    # Create a secure connection with the server
    context = ssl.create_default_context()
    with smtplib.SMTP("smtp.office365.com", 587) as server:
        server.starttls(context=context)
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())

    os.remove(filename)

class Application(ttk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.email_label = ttk.Label(self, text="Email:")
        self.email_label.pack()
        self.email_entry = ttk.Entry(self)
        self.email_entry.pack()

        self.pass_label = ttk.Label(self, text="Password:")
        self.pass_label.pack()
        self.pass_entry = ttk.Entry(self, show="*")
        self.pass_entry.pack()

        self.choose_file = ttk.Button(self)
        self.choose_file["text"] = "Choose CSV file"
        self.choose_file["command"] = self.load_file
        self.choose_file.pack()

        self.file_label = ttk.Label(self, text="")
        self.file_label.pack()

        self.send_emails = ttk.Button(self)
        self.send_emails["text"] = "Send Emails"
        self.send_emails["command"] = self.send_emails_func
        self.send_emails.pack()

        self.progress = ttk.Progressbar(self, length=100, mode='indeterminate')
        self.progress.pack()

        self.award_type = StringVar()
        self.award_type.set("deans_list") 
        self.radio_button1 = ttk.Radiobutton(
            self,
            text="Dean's List Award Message",
            variable=self.award_type,
            value="deans_list"
        )
        self.radio_button1.pack()

        self.radio_button2 = ttk.Radiobutton(
            self,
            text="President's List Award Message",
            variable=self.award_type,
            value="presidents_list"
        )
        self.radio_button2.pack()

    def load_file(self):
        self.filename = filedialog.askopenfilename(initialdir = "/", title = "Select file",
                                                   filetypes = (("CSV files","*.csv"),("all files","*.*")))
        if self.filename:
            # Update the file_label text to the filename instead of opening a new window.
            self.file_label["text"] = "Chosen file: " + os.path.basename(self.filename)

    def send_emails_func(self):
        self.progress.start()
        self.queue = queue.Queue()
        Thread(target=self.send_emails_thread).start()
        self.master.after(100, self.process_queue)

    def send_emails_thread(self):
        data = []
        with open(self.filename, 'r') as file:
            reader = csv.reader(file)

            for row in reader:
                data.append(row)

        award_type_value = self.award_type.get()

        for index, row in enumerate(data[1:]):
            send_email(self.email_entry.get(), self.pass_entry.get(), row[0], row[1], str(row[2]), row[3], award_type_value)
            self.queue.put((index+1)*100/len(data))

        self.queue.put(None)

    def process_queue(self):
        try:
            progress = self.queue.get(0)
            if progress is None:
                self.progress.stop()
                messagebox.showinfo("Done", "All emails sent successfully")
            else:
                self.progress['value'] = progress
            self.master.after(100, self.process_queue)
        except queue.Empty:
            self.master.after(100, self.process_queue)


def open_help():
  webbrowser.open('https://oit.uccs.edu/get-help')


def main():
  root = ThemedTk(theme="arc")
  root.title("Dean's and President's List Email App")
  root.geometry("400x300")

  menubar = tk.Menu(root)
  filemenu = tk.Menu(menubar, tearoff=0)
  filemenu.add_command(label="Exit", command=root.quit)
  menubar.add_cascade(label="File", menu=filemenu)

  helpmenu = tk.Menu(menubar, tearoff=0)
  helpmenu.add_command(label="Get Help", command=open_help)
  menubar.add_cascade(label="Help", menu=helpmenu)

  root.config(menu=menubar)

  app = Application(master=root)
  app.mainloop()

if __name__ == '__main__':
  main()
