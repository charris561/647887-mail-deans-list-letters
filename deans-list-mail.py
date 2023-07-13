'''
Author: Caleb Harris
Title: IT Professional
Date: 7/19/23
Description: Send deans list letter documents to each user passed in from a csv file
customized for each user with a customized email for each person.
'''

from fpdf import FPDF
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
  pdf = FPDF()

  pdf.add_page()
  pdf.image('UCCS Logo Signature.png', x=10, y=8, w=33)

  pdf.set_font("Arial", size=12)
  pdf.cell(200, 10, txt="Dear "+name+",", ln=True)
  pdf.cell(200, 10, txt="Your GPA is "+gpa+".", ln=True)
  pdf.cell(200, 10, txt="We have your address as: "+address, ln=True)

  filename = name+".pdf"
  pdf.output(filename)
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

  # background image
  #bg = PhotoImage(file="uccs.gif")
  #label1 = ttk.Label(root, image=bg)
  #label1.place(x=0, y=0)

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
