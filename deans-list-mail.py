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
import webbrowser
import csv
import smtplib, ssl
from threading import Thread
import queue

def create_pdf(name, gpa, address):
    # Register the Times New Roman font
    fonts.addMapping('Times-Roman', 0, 0, 'Times-Roman')
    fonts.addMapping('Arial', 0, 0, 'Arial')

    # Create a PDF canvas with letter page size
    filename = name + ".pdf"
    pdf = canvas.Canvas(filename, pagesize=letter)

    # Add image with a 1-inch margin from the top and left
    pdf.drawImage("UCCS Logo Signature.png", x=inch, y=9 * inch, width=4 * 72, height=0.5 * inch)

    # Set font properties
    pdf.setFont("Times-Roman", 12)

    # Add congratulatory message
    message_x = 1 * inch
    message_y = letter[1] - 2.4 * inch  # Adjusted position for the paragraph
    message_text = f"""July 13, 2023,<br/><br/><br/>
    {name}<br/>
    {address}<br/>
    Colorado Springs, CO 80907<br/><br/><br/>
    Dear {name.split(" ")[0]},<br/><br/>
    Congratulations! I am most pleased to announce that you have been named to the Dean's List for the Spring 2023 semester.
    In order to receive this honor, a student must have earned between a 3.75-3.99 GPA, while completing at least 12 credit hours.
    Your GPA from Spring 2023 was {gpa}.<br/><br/>
    Your outstanding academic performance for this semester is a source of considerable pride for the College of Letters, Arts, and Sciences at the University of Colorado Colorado Springs.<br/><br/>
    Please accept my sincere congratulations for this well-deserved honor and my hope for your continued academic success. Keep up the great work.<br/><br/><br/>
    Yours Truly,"""

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

    # Footer text
    footer_text = """College of Letters, Arts & Sciences • Office of the Dean<br/>
    Columbine Hall 2025 • 1420 Austin Bluffs Pkwy • Colorado Springs, CO 80918<br/> 
    t 719-255-4550 • f 719-255-4200"""

    # Calculate the text area width and height
    text_width = letter[0] - 2 * inch
    text_height = 4 * inch

    # Create a Paragraph object for the message text
    paragraph = Paragraph(message_text, style=paragraph_style)

    # Calculate the height required for the wrapped text
    wrapped_text_height = paragraph.wrap(text_width, text_height)[1]

    # Calculate the starting y-coordinate for the wrapped text
    start_y = message_y - wrapped_text_height

    # Draw the wrapped text within the text area
    paragraph.drawOn(pdf, message_x, start_y)

    # insert signature
    signature_image_y = 3.33 * inch  # Adjusted position for the signature image
    signature_text_y = signature_image_y - 0.01 * inch  # Adjusted position for the signature text

    # Draw the signature image
    pdf.drawImage("deanslistlettersignature.jpg", x=inch, y=signature_image_y, width=2.07 * 72, height=0.82 * inch)

    # Create a Paragraph object for the signature text
    signature_paragraph = Paragraph(signature_text, style=paragraph_style)

    # Calculate the height required for the signature text
    signature_text_height = signature_paragraph.wrap(text_width, text_height)[1]

    # Calculate the starting y-coordinate for the signature text
    signature_start_y = signature_text_y - signature_text_height

    # Draw the signature text
    signature_paragraph.drawOn(pdf, message_x, signature_start_y)

    # Create a Paragraph style for the footer text
    paragraph_style.fontSize = 9.96
    paragraph_style.alignment = 1  # Center-justified

    # Create a Paragraph object for the footer text
    footer_paragraph = Paragraph(footer_text, paragraph_style)

    # Calculate the height required for the footer text
    footer_text_height = footer_paragraph.wrap(text_width, text_height)[1]

    # Calculate the starting y-coordinate for the footer text
    footer_text_y = 0.75 * inch

    # Draw the centered, center-justified footer text
    footer_paragraph.drawOn(pdf, 0.5 * (letter[0] - text_width), footer_text_y)

    # Save the PDF
    pdf.save()
    return filename

def send_email(sender_email, password, receiver_email, name, gpa, address):

  # create email object
  message = MIMEMultipart("alternative")
  message["Subject"] = "Deans List Emails - DEVELOPMENT"
  message["From"] = sender_email
  message["To"] = receiver_email

  # create pdf file
  filename = create_pdf(name, gpa, address)

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

  # Create HTML version of message

  html = f"""\
  <html>
    <body>
      <p>
          Hello {name},<br><br>
          Congratulations! You have been placed on the Dean's list for academic success with a GPA of {gpa}!
          Enter rest of message as seen fit.<br><br>

          Sincerely, Office of Letters Arts and Sciences
      </p>
    </body>
  </html>
  """

  # Turn these into plain/html MIMEText objects
  part2 = MIMEText(html, "html")

  # Add HTML/plain-text parts to MIMEMultipart message
  # The email client will try to render the last part first
  message.attach(part2)

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

        for index, row in enumerate(data[1:]):
            send_email(self.email_entry.get(), self.pass_entry.get(), row[0], row[1], str(row[2]), row[3])
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
  root.title("Dean's List Email App")
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
