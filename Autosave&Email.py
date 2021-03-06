import openpyxl
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Font
from copy import copy, deepcopy
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import *
from tkinter.ttk import *
from datetime import datetime, date
import time
from smtplib import SMTP_SSL as SMTP
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

root= tk.Tk()

root.title('Autosave&Mail')

root.geometry('550x550')
root.minsize(550, 550)
root.configure(background='#fff')

def getFile ():

    import_file_path = filedialog.askopenfilename()
    
    wb = openpyxl.load_workbook(import_file_path)
    
    labelinfo = wb.create_sheet()
    labelinfo.title = "Info"

    labelinfo['B3'] = 'Prepared by Gianmarco Ercolani'

    labelinfo['B4'] = 'Analytics Team'

    labelinfo['B6'] = 'If you have any further questions, please contact me directly:'
    labelinfo['B6'].font = Font(italic=True)

    labelinfo['B7'] = 'gianmarco.ercolani@xxx.com'
    labelinfo['B7'].font = Font(underline='single')

    for i in range (3, 7):
        labelinfo.cell(row=3, column=i).border = Border(top=Side(style='thin'))
    for i in range (3, 7):
        labelinfo.cell(row=7, column=i).border = Border(bottom=Side(style='thin'))
    for i in range (4, 7):
        labelinfo.cell(row=i, column=2).border = Border(left=Side(style='thin'))
    for i in range (4, 7):
        labelinfo.cell(row=i, column=7).border = Border(right=Side(style='thin'))

    labelinfo.cell(row=3, column=2).border = Border(top=Side(style='thin'), left=Side(style='thin'))
    labelinfo.cell(row=3, column=7).border = Border(top=Side(style='thin'), right=Side(style='thin'))
    labelinfo.cell(row=7, column=2).border = Border(bottom=Side(style='thin'), left=Side(style='thin'))
    labelinfo.cell(row=7, column=7).border = Border(bottom=Side(style='thin'), right=Side(style='thin'))
    labelinfo.sheet_view.showGridLines = False

    today = date.today()
    d1 = today.strftime("%m.%Y")
    d2 = today.strftime("%Y.%m.%d")
    d3 = today.strftime("%Y")
        
    labelinfo['B10'] = (f'Structure: {structure.get()}')
    labelinfo['B10'].font = Font(bold=True)
    labelinfo['B11'] = (f'Date as of: {d2}')
    labelinfo['B11'].font = Font(bold=True)
    labelinfo['B12'] = (f'Datasource: {sourcein.get()}')
    labelinfo['B12'].font = Font(bold=True)

    labelinfo.protection.sheet = True
    labelinfo.protection.enable()
    labelinfo.protection.password = 'ProtezioneAttiva' #password
    
    if location.get() == 'PL' and var.get() == 1:
        os.chdir(f"K:\\2.6 NOC HRO\\11. Processes\\Reporting\\2. Sent reports\\PL\\{d3}\\{d1}\\DLP")
        wb.save(os.path.basename(import_file_path))
    elif location.get() == 'PL' and var.get() == 0:
        os.chdir(f"K:\\2.6 NOC HRO\\11. Processes\\Reporting\\2. Sent reports\\PL\\{d3}\\{d1}")
        wb.save(os.path.basename(import_file_path))
    elif location.get() == 'SE' and var.get() == 1:
        os.chdir(f"K:\\2.6 NOC HRO\\11. Processes\\Reporting\\2. Sent reports\\SE\\{d3}\\{d1}\\DLP")
        wb.save(os.path.basename(import_file_path))
    elif location.get() == 'SE' and var.get() == 0:
        os.chdir(f"K:\\2.6 NOC HRO\\11. Processes\\Reporting\\2. Sent reports\\SE\\{d3}\\{d1}")
        wb.save(os.path.basename(import_file_path))
    elif location.get() == 'DK' and var.get() == 1:
        os.chdir(f"K:\\2.6 NOC HRO\\11. Processes\\Reporting\\2. Sent reports\\DK\\{d3}\\{d1}\\DLP")
        wb.save(os.path.basename(import_file_path))
    elif location.get() == 'DK' and var.get() == 0:
        os.chdir(f"K:\\2.6 NOC HRO\\11. Processes\\Reporting\\2. Sent reports\\DK\\{d3}\\{d1}")
        wb.save(os.path.basename(import_file_path))
    elif location.get() == 'NO' and var.get() == 1:
        os.chdir(f"K:\\2.6 NOC HRO\\11. Processes\\Reporting\\2. Sent reports\\NO\\{d3}\\{d1}\\DLP")
        wb.save(os.path.basename(import_file_path))
    elif location.get() == 'NO' and var.get() == 0:
        os.chdir(f"K:\\2.6 NOC HRO\\11. Processes\\Reporting\\2. Sent reports\\NO\\{d3}\\{d1}")
        wb.save(os.path.basename(import_file_path))
    elif location.get() == 'FI' and var.get() == 1:
        os.chdir(f"K:\\2.6 NOC HRO\\11. Processes\\Reporting\\2. Sent reports\\FI\\{d3}\\{d1}\\DLP")
        wb.save(os.path.basename(import_file_path))
    elif location.get() == 'FI' and var.get() == 0:
        os.chdir(f"K:\\2.6 NOC HRO\\11. Processes\\Reporting\\2. Sent reports\\FI\\{d3}\\{d1}")
        wb.save(os.path.basename(import_file_path))
    elif location.get() == 'EE' and var.get() == 1:
        os.chdir(f"K:\\2.6 NOC HRO\\11. Processes\\Reporting\\2. Sent reports\\EE\\{d3}\\{d1}\\DLP")
        wb.save(os.path.basename(import_file_path))
    elif location.get() == 'EE' and var.get() == 0:
        os.chdir(f"K:\\2.6 NOC HRO\\11. Processes\\Reporting\\2. Sent reports\\EE\\{d3}\\{d1}")
        wb.save(os.path.basename(import_file_path))
    elif location.get() == 'Learning Reports' and var.get() == 1:
        os.chdir(f"K:\\2.6 NOC HRO\\11. Processes\\Reporting\\2. Sent reports\\Learning Reports\\{d3}\\{d1}\\DLP")
        wb.save(os.path.basename(import_file_path))
    elif location.get() == 'Learning Reports' and var.get() == 0:
        os.chdir(f"K:\\2.6 NOC HRO\\11. Processes\\Reporting\\2. Sent reports\\Learning Reports\\{d3}\\{d1}")
        wb.save(os.path.basename(import_file_path))
    elif location.get() == 'Other' and var.get() == 1:
        os.chdir(f"K:\\2.6 NOC HRO\\11. Processes\\Reporting\\2. Sent reports\\Other\\{d3}\\{d1}\\DLP")
        wb.save(os.path.basename(import_file_path))
    elif location.get() == 'Other' and var.get() == 0:
        os.chdir(f"K:\\2.6 NOC HRO\\11. Processes\\Reporting\\2. Sent reports\\Other\\{d3}\\{d1}")
        wb.save(os.path.basename(import_file_path))

    def smtp_endpoint():
        smtp = os.getenv('EMAIL_NOTIFICATIONS_SMTP_ENDPOINT', '*IP O SMTP ADDRESS*:25')
        return tuple(smtp.split(':'))

    def sender():
        return str("*gianmarco.ercolani@xxx.com*")
        #return getpass.getuser() + '@' + socket.gethostname()

    def send_email(send_to, subject, message, file_path=None, file_name=None):
        send_from = sender()
        msg = MIMEMultipart('alternative')
        msg['From'] = send_from
        msg['To'] = send_to
        msg["Cc"] = cc.get()
        msg["Bcc"] = bcc.get()
        msg['Subject'] = subject
        msg.attach(MIMEText(message, 'html'))
        if file_path:
            f = open(file_path, 'rb')
            attachment = MIMEApplication(f.read())
            #attachment = MIMEText(f.read())
            email_filename = file_name if file_name else file_path
            attachment.add_header('Content-Disposition', 'attachment', filename=email_filename)
            msg.attach(attachment)
        smtp = smtp_endpoint()
        server = smtplib.SMTP(smtp[0], smtp[1])
        server.sendmail(send_from, msg['To'].split(";") + msg["Cc"].split(",") + msg["Bcc"].split(","), msg.as_string())
        server.close()

    time.sleep(3)
    message = """\
    <html>
    Hi,
    <br><br>
    Please find attached the requested report.
    <br><br>
    <div style="font-size:90%; color:#0606bf;">
        <b>NEED A REPORT?</b>
    </div>
    <div style="font-size:90%;">
        <i>I really enjoy helping our people  to be sure that you get your report as soon as possible, please use THIS tool instead of sending request via e-mail.</i>
    </div>
    <br>
    Kind Regards,
    <br><br>
    <b>Gianmarco Ercolani</b>
    <br>
    <div style="font-size:90%;">
        Reporting Specialist
    <div>
    <br>
    <div style="font-size:90%;">
        <b style="color:#0606bf;">xxx</b> | Analytics & Reporting
    </div>
    </div>
    <div style="font-size:90%;">
        Visit me: Address
    </div>
    <div style="font-size:90%;">
        E-mail: gianmarco.ercolani@xxx.com
    </div>
    <div style="font-size:90%;">
        Web:
        <a href="www.example.com" target="_blank">
            example.com
        </a>
    </div>
    <br>
    <div style="font-size:85%; color:#888">
        Company info
        <br><br>
        This e-mail may contain confidential information. If you receive this e-mail by mistake, please inform the sender, delete the e-mail and do not share or copy it.
    </div>
    </html>
    """
    send_email(email.get(), f"{os.path.basename(import_file_path)} {d1}", message, f"{os.path.abspath(os.path.basename(import_file_path))}", f"{os.path.basename(import_file_path)}")


structure = Combobox(root, width=8)
structure ['values']= ("MCC", "Org")
structure.configure(font=10)
structure.current(0)

location = Combobox(root, width=8)
location ['values']= ("PL", "SE", "DK", "NO", "FI", "EE", "Other", "Learning Reports")
location["state"]= "readonly"
location.configure(font=10)
location.current(0)

infolbl = Label(root, text="Signature", font=('helvetica 8 bold'), foreground='#008000', background='#fff')
infolbl.grid(row=0, sticky='W', padx=(10, 0))
Separator(root, orient="horizontal")
Separator().grid(row=0, column= 0, pady= (20, 0), sticky=EW, columnspan=2)

structurelbl = Label(root, text = "Structure", font=('helvetica', 10), background='#fff')
structurelbl.grid(row=2, sticky='E', padx= 15, pady=(20, 0))

source = Label(root, text = "Datasource", font=('helvetica', 10), background='#fff')
source.grid(row= 3, sticky='E', padx= 15, pady=20)

locationlbl = Label(root, text = "Folder", font=('helvetica', 10), background='#fff')
locationlbl.grid(row=5, column= 0, sticky='E', padx= 15, pady=20)

infolbl = Label(root, text="Save to", font=('helvetica 8 bold'), foreground='#008000', background='#fff')
infolbl.grid(row=4, sticky='W', padx=(10, 0))
Separator(root, orient="horizontal")
Separator().grid(row=4, column= 0, pady=(20, 0), sticky=EW, columnspan=2)

structure.grid(row= 2, column= 1, padx=15, pady=(20, 0))

sourcein = Entry(root, width=25, font=('helvetica', 10))
sourcein.grid(row= 3, column= 1, padx=(15, 50), pady=20, sticky='E')

location.grid(row= 5, column= 1, padx=15, pady=20)

var = IntVar()
dlp = tk.Checkbutton (root, text='DLP', variable=var, bg='#fff', font=('helvetica', 10))
dlp.grid(row=5, column=1, padx=(0, 50), pady=20, sticky='E')

infolbl = Label(root, text="Send Email", font=('helvetica 8 bold'), foreground='#008000', background='#fff')
infolbl.grid(row=7, sticky='W', padx=(10, 0))
Separator(root, orient="horizontal")
Separator().grid(row=7, column= 0, pady= (20, 0), sticky=EW, columnspan=2)

emailto = Label(root, text = "To", font=('helvetica', 10), background='#fff')
emailto.grid(row=8, column=0, sticky='E', padx= 15, pady=(20, 0))
email = Entry(root, width=25, font=('helvetica', 10))
email.grid(row=8, column= 1, padx=(15, 50), pady=(20, 0), sticky='E')

ccto = Label(root, text = "Cc", font=('helvetica', 10), background='#fff')
ccto.grid(row=9, column=0, sticky='E', padx= 15, pady=(20, 0))
cc = Entry(root, width=25, font=('helvetica', 10))
cc.grid(row=9, column= 1, padx=(15, 50), pady=(20, 0), sticky='E')

bccto = Label(root, text = "Bcc", font=('helvetica', 10), background='#fff')
bccto.grid(row=10, column=0, sticky='E', padx= 15, pady=(20, 0))
bcc = Entry(root, width=25, font=('helvetica', 10))
bcc.grid(row=10, column= 1, padx=(15, 50), pady=(20, 0), sticky='E')

browseButton_Excel = tk.Button(text='Choose File', command=getFile, bg='#008000', fg='#fff', font=('helvetica', 12))
browseButton_Excel.grid(columnspan=2, pady= 30)

root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=2)
root.rowconfigure(11, weight=1)

root.mainloop()
