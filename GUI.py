import gc
import time
import tkinter as tk
import imaplib, email
# intsall:oauth2client, gspread, PyOpenSSL, gspread-formatting
import gspread
from gspread_formatting import *
from httplib2 import ServerNotFoundError

from oauth2client.service_account import ServiceAccountCredentials

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from datetime import date
import datetime


class Window:
    def __init__(self):
        self.scope = ['https://spreadsheets.google.com/feeds',
                      'https://www.googleapis.com/auth/drive']
        self.credsdict = {'type': 'service_account',
                          'project_id': 'lyons-email-updatesheet-script',
                          'private_key_id': '2414046f0774321dd439b121749c6eff1db8713b',
                          'private_key': '-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQDrHJwzqsP88DVg\nHplpDCpKYwUYGUEbRg0sqFx7rJy9mYOXD7/TpducJlInQs5PbTiF/yQ0o/IjAaSB\nWWqaQbCv+COSVg5udAcbrP6xSTv+3L2crqFMDXnHbKHHfhvNTQSyfh2JSIZY78le\nn6dcq0g3MvHEt7Qy89bdgC966XsXXVe+va0t8lI26IgPf9ZTSFm4a4TYStsT3mP3\nKl4RegTn03bMHYdPmOm2B/I5E814RG8Wjsk0FIYhW6reYNhLuBLrS14744bWmd6f\nYvZLpKmEf5EN91nD/veWoKB2qLhESfmpP1iipieoVpm+VvfDlWmcusl8s+w0/ETf\nkk1BMo5dAgMBAAECggEAE3cw744p3907bhPae7oIHlSIbXBZ1Zo9KP9feNXXvFLj\ndDRXm3xV7F2324xKbIUMcvum0bzpJUDTj+oJS3A44rjWqRz64OY2WHJAPAlmMDmy\ncTB8JkHPXVV/J3cnch34T5bldyJMDTz9HRp2ztNXjUpoffL/tmA93+TnCXQfPtXQ\n9vo1nj2n3Pix4zSZSwMk18Ll6vERbjKHffaSc/xdvMuHgEFa48ze3cOmrpFyJHM1\nrdZ6qhCUQPwgaHHFTgKcIDU5ILVlmuIm1DnUCo3K0ocDdov7zwyU/J/5gUoFzOpW\nCA7pwy2QbLZ3XWGL01gV6d1mGqiF9ajNxLBezu7AwQKBgQD4/44ywpno0nbEXlNj\n8zLpWGJyFkVDYwdPtcPlQZzPdq/UJd0bfApSnRaNDTdNwTbYtUSyT3QlXUOgzjai\nn2zKf6cXRtaGScGnmF4acJFvBSw9xTaaDaLs7IVoGD/CaCPwwJJ71OBLRTPxvIev\ntMfWWsANpYy/6pipuJI3Ujb/ZQKBgQDxuRdLN8Ox1OVQxAXUbykqyeuNt5h+h/vk\n9SdrYu72KmdHIUHgfu1L9N9yFp/KpGZ4CnhlkGX0LXoDnqQKd1yjiLxGrWgLk76P\n/dXCpLs4N7av94je5Z2e2Q1v+1qQ8cAs13UCcwemHe1WxRfJqsj8kP3ZX2FFGUke\nJgzZGpkPmQKBgDISWgsVHRQ3tpB4k3ZnApbwIiPlHJqXgHHkEHe6wQjrSiJ0Vslf\nIUhJtK46uSNWtmvPz/e3iJi275GXxl7fhmYWU4iXwy4QCPRl7I6OkoBr3uCxFvDV\nyyyvx4gOUEwM2yVf5FUoks4wJWj4S6TmysTtTO+xmeNCDt8acbTUQKENAoGAPH0k\n5x29SvMLr3peOxrWIm8FEyGud3tv/YubobPQOKnDznj0E0mv+CH/CH3A3uTk/4Uf\nO8s2uDPpJJ6+TiAwfnvpIYajUsJWHZJXu62dbCQFA2PeTGkJWIbYZf1wXHUishX4\nofRHJbq3ec84dK7YPNvLqmnD3ZbGRVUgQfP1+YECgYBiwykzHoHeY4e1yAbDzPzU\nwvlxE6bGgJXJLKObhRJjoFn5zyAEteMFOdTh7OUfcWIxi3/HdxIR8k8FvupNAOAi\nBQYLiMjIRLCJ3ljR3JHN3fDrwVpNOBEdtT1S7i8jWTA0tb2XSn9mN3IDjrdE5aqU\n2hm0MtKP0tWBhzfllvvOHg==\n-----END PRIVATE KEY-----\n',
                          'client_email': 'id-d-print-request@lyons-email-updatesheet-script.iam.gserviceaccount.com',
                          'client_id': '117542574436878496508',
                          'auth_uri': 'https://accounts.google.com/o/oauth2/auth',
                          'token_uri': 'https://oauth2.googleapis.com/token',
                          'auth_provider_x509_cert_url': 'https://www.googleapis.com/oauth2/v1/certs',
                          'client_x509_cert_url': 'https://www.googleapis.com/robot/v1/metadata/x509/id-d-print-request%40lyons-email-updatesheet-script.iam.gserviceaccount.com'}

        self.fmtreadypickup = CellFormat(
            backgroundColor=Color(0.078, 0.616, 1),
        )
        self.fmtdenied = CellFormat(
            backgroundColor=Color(0.878, 0.4, 0.4),
        )
        self.fmtfailed = CellFormat(
            backgroundColor=Color(0.957, 0.8, 0.8),
        )
        self.fmtclarification = CellFormat(
            backgroundColor=Color(1, 0.851, 0.4),
        )
        self.fmtcancelled = CellFormat(
            backgroundColor=Color(0.71, 0.37, 0.02)
        )
        self.fmtpickedup = CellFormat(
            backgroundColor=Color(0.85, 0.85, 0.85),
        )
        self.fmtneverpickedup = CellFormat(
            backgroundColor=Color(0.4, 0.31, 0.65),
        )
        self.credentials = ""
        self.wks = ""
        self.workSheet = ""
        self.Ticketnum = ""
        self.row_number = ""
        self.z = ""
        self.log = ""
        self.rowstr = ""
        # Define Name Parameter
        self.name = ""
        self.a = ""
        # Define Patron Email parameter
        self.patron_email = ""
        self.reasonEntry = ""
        # Assign Patron Email to variable
        self.b = ""
        # define message for email
        self.msg = MIMEMultipart()
        self.c = ""
        self.x1 = ""
        self.initialDate = ""
        self.lastDate = ""
        # Setup window
        self.window = tk.Tk()
        self.window.resizable(1, 1)
        self.titleFrame = ""
        # Initiate Login
        self.LoginMenu()
        self.window.mainloop()

    def Authorize(self):
        self.credentials = ServiceAccountCredentials.from_json_keyfile_dict(self.credsdict, self.scope)
        try:
            self.gc = gspread.authorize(self.credentials)
        except ServerNotFoundError:
            connection = tk.Label(self.titleFrame, text="Login Failed, No Internet Connection")
            connection.pack()
            connection.update()
            self.titleFrame.after(2000, connection.destroy())
            self.titleFrame.destroy()
            # Re-Initiate Login
            self.LoginMenu()
            self.wifi = 0

        else:
            self.sh = self.gc.open('3D Printing Requests')
            self.worksheet_list = self.sh.worksheets()
            self.worksheet_str = [str(i) for i in self.worksheet_list]
            self.worksheet = [x.replace("<Worksheet '", '').split("' id:", 1)[0] for x in self.worksheet_str]
            self.mail = imaplib.IMAP4_SSL('imap.gmail.com')
            self.wifi = 1

    def LoginMenu(self):
        password = tk.StringVar()
        self.window.title("3D Print Request - Login")
        self.window.geometry("500x500")
        self.titleFrame = tk.Frame(self.window)
        self.titleFrame.pack()
        Account = 'lyons.newmedia@gmail.com'
        Sender = tk.StringVar(self.titleFrame, value=Account)
        tk.Label(self.titleFrame, text="Account:").pack()
        tk.Entry(self.titleFrame, textvariable=Sender, width=30).pack()
        tk.Label(self.titleFrame, text='Enter Password:').pack()
        tk.Entry(self.titleFrame, textvariable=password, width=30).pack()
        tk.Button(self.titleFrame, text='Enter', command=lambda: self.PasswordEntry(Sender, password)).pack()

    def PasswordEntry(self, Sender, password):
        self.User = str(Sender.get())
        self.Password = str(password.get())
        self.Authorize()
        if self.wifi == 1:
            # Exception Handling for when there's no match
            try:
                self.mail.login(self.User, self.Password)
            except Exception as e:
                self.log = '0'
            else:
                self.log = '1'
            if self.log == '0':
                statusLabel = tk.Label(self.titleFrame, text="Login Failed, Invalid Email/Password")
                statusLabel.pack()
                statusLabel.update()
                self.titleFrame.after(2000, statusLabel.destroy())
                self.titleFrame.destroy()
                # Re-Initiate Login
                self.LoginMenu()
            elif self.log == '1':
                statusLabel = tk.Label(self.titleFrame, text="Login Success")
                statusLabel.pack()
                statusLabel.update()
                self.titleFrame.after(1000, statusLabel.destroy())
                self.titleFrame.destroy()
                # Initiate Menu
                self.StartMenu()

    def StartMenu(self):
        self.window.title("3D Print Request - Menu")
        self.window.geometry("500x800")
        self.titleFrame = tk.Frame(self.window)
        self.titleFrame.pack()
        tk.Label(self.titleFrame, text="Select the time period of the Print Request").pack()
        self.workSheet = tk.StringVar(self.titleFrame)
        self.workSheet.set(list(self.worksheet)[0])  # default value
        tk.OptionMenu(self.titleFrame, self.workSheet, *self.worksheet).pack()
        tk.Label(self.titleFrame, text="").pack()
        tk.Button(self.titleFrame, text="New Submission Processing", width="40", pady="5",
                  command=self.getInfoNewEntry).pack()
        tk.Button(self.titleFrame, text="Ready For Pickup", width="40", pady="5",
                  command=lambda: self.getInfo(self.readyForPickup, "Send Email",
                                               "3D Print Request - Ready For Pickup")).pack()
        tk.Button(self.titleFrame, text="Delayed Printing", width="40", pady="5",
                  command=lambda: self.getInfo4(self.DelayedPrinting, "Send Email",
                                                "3D Print Request - Delayed Printing")).pack()
        tk.Button(self.titleFrame, text="Denied", width="40", pady="5",
                  command=lambda: self.getInfo2(self.Denied, "Send Email", "3D Print Request - Denied")).pack()
        tk.Button(self.titleFrame, text="Dimensions Clarification - Skewed Print", width="40", pady="5",
                  command=lambda: self.getInfo(self.Clarification_Skewed, "Send Email",
                                               "3D Print Request - Skewed Print")).pack()
        tk.Button(self.titleFrame, text="Dimensions Clarification - Large Print", width="40", pady="5",
                  command=lambda: self.getInfo(self.Clarification_Large, "Send Email",
                                               "3D Print Request - Large Print")).pack()
        tk.Button(self.titleFrame, text="Reminder", width="40", pady="5",
                  command=lambda: self.getInfo3(self.Reminder, "Send Email", "3D Print Request - Reminder")).pack()
        tk.Button(self.titleFrame, text="Failed", width="40", pady="5",
                  command=lambda: self.getInfo2(self.Failed, "Send Email", "3D Print Request - Failed")).pack()
        tk.Button(self.titleFrame, text="Picked Up", width="40", pady="5",
                  command=lambda: self.getInfo(self.pickedUp, "Update Spreadsheet",
                                               "3D Print Request - Picked Up")).pack()
        tk.Button(self.titleFrame, text="Never Picked Up", width="40", pady="5",
                  command=lambda: self.getInfo(self.nevPickedUp, "Update Spreadsheet",
                                               "3D Print Request - Never Picked Up")).pack()
        tk.Button(self.titleFrame, text="Cancelled", width="40", pady="5",
                  command=lambda: self.getInfo2(self.cancelled, "Update Spreadsheet",
                                                "3D Print Request - Cancelled")).pack()

    def backToMenu(self):
        self.infoFrame.destroy()
        self.wks = ""
        self.Ticketnum = ""
        self.reasonEntry = ""
        self.row_number = ""
        self.rowstr = ""
        # Define Name Parameter
        self.name = ""
        self.a = ""
        # Define Patron Email parameter
        self.patron_email = ""
        self.b = ""
        self.z = ""
        # define message for email
        self.msg = MIMEMultipart()
        self.c = ""
        self.x1 = ""
        self.StartMenu()

    def getInfo(self, function, text, title):
        self.window.title(title)
        self.wks = self.sh.get_worksheet(list(self.worksheet).index(self.workSheet.get()))
        self.titleFrame.destroy()
        self.infoFrame = tk.Frame(self.window)
        self.infoFrame.pack()
        self.nameEntry = tk.StringVar(self.infoFrame, value=self.name)
        ticketNumEntry = tk.StringVar(self.infoFrame, value=self.Ticketnum)
        self.emailEntry = tk.StringVar(self.infoFrame, value=self.patron_email)
        tk.Label(self.infoFrame, text="Enter Ticket #:").pack()
        tk.Entry(self.infoFrame, textvariable=ticketNumEntry).pack()
        self.ticket = str(ticketNumEntry)
        tk.Button(self.infoFrame, text="Search",
                  command=lambda: self.findTicket(ticketNumEntry, function, text, title)).pack()
        tk.Label(self.infoFrame, text="Enter Patron Name:").pack()
        tk.Entry(self.infoFrame, textvariable=self.nameEntry).pack()
        tk.Label(self.infoFrame, text="Enter Patron Email:").pack()
        tk.Entry(self.infoFrame, textvariable=self.emailEntry).pack()
        tk.Button(self.infoFrame, text=text, command=lambda: function(ticketNumEntry)).pack()
        tk.Button(self.infoFrame, text="Back to Menu", command=self.backToMenu).pack()

    def getInfo2(self, function, text, title):
        self.window.title(title)
        self.wks = self.sh.get_worksheet(list(self.worksheet).index(self.workSheet.get()))
        self.titleFrame.destroy()
        self.infoFrame = tk.Frame(self.window)
        self.infoFrame.pack()
        self.nameEntry = tk.StringVar(self.infoFrame, value=self.name)
        ticketNumEntry = tk.StringVar(self.infoFrame, value=self.Ticketnum)
        self.emailEntry = tk.StringVar(self.infoFrame, value=self.patron_email)
        reasonEntry = tk.StringVar()
        tk.Label(self.infoFrame, text="Enter Ticket #:").pack()
        tk.Entry(self.infoFrame, textvariable=ticketNumEntry).pack()
        tk.Button(self.infoFrame, text="Search",
                  command=lambda: self.findTicket2(ticketNumEntry, function, text, title)).pack()
        tk.Label(self.infoFrame, text="Enter Patron Name:").pack()
        tk.Entry(self.infoFrame, textvariable=self.nameEntry).pack()
        tk.Label(self.infoFrame, text="Enter Patron Email:").pack()
        tk.Entry(self.infoFrame, textvariable=self.emailEntry).pack()
        tk.Label(self.infoFrame, text="Enter Reason:").pack()
        tk.Entry(self.infoFrame, textvariable=reasonEntry, width=40).pack()
        tk.Button(self.infoFrame, text=text, command=lambda: function(ticketNumEntry, reasonEntry)).pack()
        tk.Button(self.infoFrame, text="Back to Menu", command=self.backToMenu).pack()

    def getInfo3(self, function, text, title):
        self.window.title(title)
        self.wks = self.sh.get_worksheet(list(self.worksheet).index(self.workSheet.get()))
        self.titleFrame.destroy()
        self.infoFrame = tk.Frame(self.window)
        self.infoFrame.pack()
        self.nameEntry = tk.StringVar(self.infoFrame, value=self.name)
        ticketNumEntry = tk.StringVar(self.infoFrame, value=self.Ticketnum)
        self.emailEntry = tk.StringVar(self.infoFrame, value=self.patron_email)
        userDateString = 'mm/dd/yyyy'
        dateEntry1 = tk.StringVar(self.infoFrame, value=userDateString)
        dateEntry2 = tk.StringVar(self.infoFrame, value=userDateString)
        tk.Label(self.infoFrame, text="Enter Ticket #:").pack()
        tk.Entry(self.infoFrame, textvariable=ticketNumEntry).pack()
        tk.Button(self.infoFrame, text="Search",
                  command=lambda: self.findTicket3(ticketNumEntry, function, text, title)).pack()
        tk.Label(self.infoFrame, text="Enter Patron Name:").pack()
        tk.Entry(self.infoFrame, textvariable=self.nameEntry).pack()
        tk.Label(self.infoFrame, text="Enter Patron Email:").pack()
        tk.Entry(self.infoFrame, textvariable=self.emailEntry).pack()
        tk.Label(self.infoFrame, text="Date of original message:").pack()
        tk.Entry(self.infoFrame, textvariable=dateEntry1, width=15).pack()
        tk.Label(self.infoFrame, text="Last date to pickup print:").pack()
        tk.Entry(self.infoFrame, textvariable=dateEntry2, width=15).pack()
        tk.Button(self.infoFrame, text=text, command=lambda: function(ticketNumEntry, dateEntry1, dateEntry2)).pack()
        tk.Button(self.infoFrame, text="Back to Menu", command=self.backToMenu).pack()

    def getInfo4(self, function, text, title):
        self.window.title(title)
        self.wks = self.sh.get_worksheet(list(self.worksheet).index(self.workSheet.get()))
        self.titleFrame.destroy()
        self.infoFrame = tk.Frame(self.window)
        self.infoFrame.pack()
        self.nameEntry = tk.StringVar(self.infoFrame, value=self.name)
        ticketNumEntry = tk.StringVar(self.infoFrame, value=self.Ticketnum)
        self.emailEntry = tk.StringVar(self.infoFrame, value=self.patron_email)
        userDateString = 'Month, DD, YYYY'
        responseDate = tk.StringVar(self.infoFrame, value=userDateString)
        tk.Label(self.infoFrame, text="Enter Ticket #:").pack()
        tk.Entry(self.infoFrame, textvariable=ticketNumEntry).pack()
        tk.Button(self.infoFrame, text="Search",
                  command=lambda: self.findTicket4(ticketNumEntry, function, text, title)).pack()
        tk.Label(self.infoFrame, text="Enter Patron Name:").pack()
        tk.Entry(self.infoFrame, textvariable=self.nameEntry).pack()
        tk.Label(self.infoFrame, text="Enter Patron Email:").pack()
        tk.Entry(self.infoFrame, textvariable=self.emailEntry).pack()
        tk.Label(self.infoFrame, text="Cancel Request if Patron doesn't respond by:").pack()
        tk.Entry(self.infoFrame, textvariable=responseDate, width=20).pack()
        tk.Button(self.infoFrame, text=text, command=lambda: function(ticketNumEntry, responseDate)).pack()
        tk.Button(self.infoFrame, text="Back to Menu", command=self.backToMenu).pack()

    def getInfoNewEntry(self):
        self.window.title("3D Print Request - New Submission")
        self.wks = self.sh.get_worksheet(list(self.worksheet).index(self.workSheet.get()))
        self.titleFrame.destroy()
        self.infoFrame = tk.Frame(self.window)
        self.infoFrame.pack()
        nameEntry = tk.StringVar(self.infoFrame, value=self.name)
        ticketNumEntry = tk.StringVar(self.infoFrame, value=self.Ticketnum)
        emailEntry = tk.StringVar(self.infoFrame, value=self.patron_email)
        StaffInitials = tk.StringVar(self.infoFrame)
        dateToday = tk.StringVar(self.infoFrame, value=date.today().strftime("%m/%d/%Y"))
        CourseYN = tk.IntVar(self.infoFrame)
        CourseCode = tk.StringVar(self.infoFrame)
        affiliation = tk.StringVar(self.infoFrame)
        department = tk.StringVar(self.infoFrame)
        research = tk.IntVar(self.infoFrame)
        ownC = tk.IntVar(self.infoFrame)
        consent = tk.IntVar(self.infoFrame)
        handle = tk.StringVar(self.infoFrame)
        SD = tk.StringVar(self.infoFrame)
        Fname = tk.StringVar(self.infoFrame)
        Ptime = tk.StringVar(self.infoFrame)
        tk.Label(self.infoFrame, text="Ticket #:").pack()
        tk.Entry(self.infoFrame, textvariable=ticketNumEntry).pack()
        tk.Label(self.infoFrame, text="Patron Name:").pack()
        tk.Entry(self.infoFrame, textvariable=nameEntry).pack()
        tk.Label(self.infoFrame, text="Patron Email:").pack()
        tk.Entry(self.infoFrame, textvariable=emailEntry).pack()
        tk.Label(self.infoFrame, text="Date:").pack()
        tk.Entry(self.infoFrame, textvariable=dateToday).pack()
        tk.Label(self.infoFrame, text="Staff Initials:").pack()
        tk.Entry(self.infoFrame, textvariable=StaffInitials).pack()
        tk.Label(self.infoFrame, text="Is it for a course?").pack()
        CourseYNF = tk.LabelFrame(self.infoFrame)
        CourseYNF.pack()
        tk.Radiobutton(CourseYNF, text="Yes", padx=20, variable=CourseYN, value=1).pack(side="left")
        tk.Radiobutton(CourseYNF, text="No", padx=20, variable=CourseYN, value=0).pack(side="left")
        tk.Label(self.infoFrame, text="Course Code:").pack()
        tk.Entry(self.infoFrame, textvariable=CourseCode).pack()
        tk.Label(self.infoFrame, text="Affiliation:").pack()
        tk.Entry(self.infoFrame, textvariable=affiliation).pack()
        tk.Label(self.infoFrame, text="Department:").pack()
        tk.Entry(self.infoFrame, textvariable=department).pack()
        tk.Label(self.infoFrame, text="Is it for research?").pack()
        researchF = tk.LabelFrame(self.infoFrame)
        researchF.pack()
        tk.Radiobutton(researchF, text="Yes", padx=20, variable=research, value=1).pack(side="left")
        tk.Radiobutton(researchF, text="No", padx=20, variable=research, value=0).pack(side="left")
        tk.Label(self.infoFrame, text="Did you create this model?").pack()
        createF = tk.LabelFrame(self.infoFrame)
        createF.pack()
        tk.Radiobutton(createF, text="Yes", padx=20, variable=ownC, value=1).pack(side="left")
        tk.Radiobutton(createF, text="No", padx=20, variable=ownC, value=0).pack(side="left")
        tk.Label(self.infoFrame, text="Do you consent to Instagram Post?").pack()
        consentF = tk.LabelFrame(self.infoFrame)
        consentF.pack()
        tk.Radiobutton(consentF, text="Yes", padx=20, variable=consent, value=1).pack(side="left")
        tk.Radiobutton(consentF, text="No", padx=20, variable=consent, value=0).pack(side="left")
        tk.Label(self.infoFrame, text="Instagram Handle:").pack()
        tk.Entry(self.infoFrame, textvariable=handle).pack()
        tk.Label(self.infoFrame, text="File Name:").pack()
        tk.Entry(self.infoFrame, textvariable=Fname).pack()
        tk.Label(self.infoFrame, text="SD card:").pack()
        tk.Entry(self.infoFrame, textvariable=SD).pack()
        tk.Label(self.infoFrame, text="Print Time:").pack()
        tk.Entry(self.infoFrame, textvariable=Ptime).pack()
        tk.Button(self.infoFrame, text="Submit",
                  command=lambda: self.defineNewPatronInfo(nameEntry.get(), ticketNumEntry.get(), emailEntry.get(),
                                                           dateToday.get(), StaffInitials.get(), CourseYN.get(),
                                                           CourseCode.get(), affiliation.get(), department.get(),
                                                           research.get(), ownC.get(), consent.get(), handle.get(),
                                                           SD.get(), Fname.get(), Ptime.get())).pack()
        tk.Button(self.infoFrame, text="Back to Menu", command=self.backToMenu).pack()

    def defineNewPatronInfo(self, nameEntry, ticketEntry, emailEntry, dateToday, StaffInitials, CourseYN, CourseCode,
                            affiliation, department, research, ownC, consent, handle, SD, Fname, Ptime):
        all = self.wks.get_all_values()
        end_row = len(all) + 1
        if CourseYN == 0:
            CourseYN = "N"
        else:
            CourseYN = "Y"
        if research == 0:
            research = "N"
        else:
            research = "Y"
        if ownC == 0:
            OwnC = "N"
        else:
            OwnC = "Y"
        if consent == 0:
            consent = "N"
        else:
            consent = "Y"
        self.wks.update_cell(end_row, 1, ticketEntry)
        self.wks.update_cell(end_row, 2, nameEntry)
        self.wks.update_cell(end_row, 3, emailEntry)
        self.wks.update_cell(end_row, 4, dateToday)
        self.wks.update_cell(end_row, 5, StaffInitials)
        self.wks.update_cell(end_row, 6, CourseYN)
        self.wks.update_cell(end_row, 7, CourseCode)
        self.wks.update_cell(end_row, 8, affiliation)
        self.wks.update_cell(end_row, 9, department)
        self.wks.update_cell(end_row, 10, research)
        self.wks.update_cell(end_row, 11, OwnC)
        self.wks.update_cell(end_row, 12, consent)
        self.wks.update_cell(end_row, 13, handle)
        self.wks.update_cell(end_row, 14, SD)
        self.wks.update_cell(end_row, 15, Fname)
        self.wks.update_cell(end_row, 16, Ptime)
        infoLab1 = tk.Label(self.infoFrame, text="Submitted!")
        infoLab1.pack()
        infoLab1.update()
        time.sleep(1)
        infoLab1.destroy()
        self.infoFrame.destroy()
        self.StartMenu()

    def readyForPickup(self, ticketNumEntry):
        if (self.z == "1"):
            self.row_number = (self.wks.find(self.name).row)
            self.rowstr = str(self.row_number)
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.Ticketnum = str(ticketNumEntry.get())
        else:
            self.name = self.nameEntry.get()
            self.patron_email = self.emailEntry.get()
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        subject = "3D Print Request - Ready for Pickup"
        self.msg += "Subject: " + subject + '\n\n'
        body1 = "Hi " + self.name + ",\n\nGood news! The following requested 3D print job has been printed successfully:\n\n"
        body1 += "Ticket #: " + self.Ticketnum + "\n\nPlease bring this email and your McMaster ID card with you to the Help Desk " \
                                                 "in Lyons New Media Centre (Mills Library, 4th floor) to retrieve your item.\n\n"
        body1 += "You will be required to sign for it, so a proxy cannot come to pick this up for you.\n\nWe will hold this " \
                 "item for no more than 30 days from today's date before it is reclaimed and/or recycled.  " \
                 "If you cannot make it into the Centre due to work/being home etc., please let us know and we can arrange to " \
                 "hold onto it until you can make it in.\n\nSincerely,\n\nLyons New Media Centre Staff\n\n"
        body1 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
        self.msg += body1
        LNMC = """library.mcmaster.ca/spaces/lyons"""
        self.msg += LNMC
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(self.User, self.Password)
        server.sendmail(self.User, self.patron_email, self.msg)
        server.quit()
        if (self.z == "1"):
            self.wks.update_cell(self.row_number, 17, "Y")
            format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtreadypickup)
            infoLab2 = tk.Label(self.infoFrame, text="Message Sent & Spreadsheet Updated!")
            infoLab2.pack()
            infoLab2.update()
            self.infoFrame.after(2000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.StartMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Message Sent!")
            infoLab1.pack()
            infoLab1.update()
            self.infoFrame.after(2000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.StartMenu()

    def DelayedPrinting(self, ticketNumEntry, responseDate):
        if (self.z == "1"):
            self.row_number = self.wks.find(self.name).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.Ticketnum = str(ticketNumEntry.get())
            self.rowstr = str(self.row_number)
        else:
            self.name = self.nameEntry.get()
            self.patron_email = self.emailEntry.get()
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        subject = "3D Print Request - Delayed Printing"
        self.msg += "Subject: " + subject + '\n\n'
        body2 = "Hi " + self.name + ",\n\n"
        body2 += "This is in regards to 3D Print Ticket#: " + self.Ticketnum + ".\n\n"
        body2 += "We have had an unusual amount of course-related print requests submitted this term, and are prioritizing " \
                 "those requests before regular requests. Because of this, we may not be able to complete your request " \
                 "by the end of April (last day of exams), so it may be completed as we're going into May " \
                 "(during the summer months).\n\n"
        body2 += "We need to know if you would still like this ticket to be printed, knowing that there is a delay that it " \
                 "may not be printed before the term is over. If you still want it to be printed, but you are not able to " \
                 "pick it up immediately since it may be completed during the summer, you can let us know to hold it for " \
                 "you till you can.\n\n"
        body2 += "Please respond to this email by " + str(
            responseDate.get()) + ". If we do not hear from you by that date, we will " \
                                  "assume it is unwanted and will cancel the request.\n\nThank you\n\nLyons New Media Centre\n\n"
        body2 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
        self.msg += body2
        LNMC = """library.mcmaster.ca/spaces/lyons"""
        self.msg += LNMC
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(self.User, self.Password)
        server.sendmail(self.User, self.patron_email, self.msg)
        server.quit()
        if (self.z == "1"):
            self.wks.update_cell(self.row_number, 17, "Y")
            infoLab2 = tk.Label(self.infoFrame, text="Message Sent & Spreadsheet Updated!")
            infoLab2.pack()
            infoLab2.update()
            self.infoFrame.after(2000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.StartMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Message Sent!")
            infoLab1.pack()
            infoLab1.update()
            self.infoFrame.after(2000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.StartMenu()

    def Denied(self, ticketNumEntry, reasonEntry):
        if (self.z == "1"):
            self.row_number = self.wks.find(self.name).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.Ticketnum = str(ticketNumEntry.get())
            self.rowstr = str(self.row_number)
        else:
            self.name = self.nameEntry.get()
            self.patron_email = self.emailEntry.get()
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        subject = "3D Print Request - Denied"
        self.msg += "Subject: " + subject + '\n\n'
        body3 = "Hi " + self.name + ",\n\n"
        body3 += "This is in regards to 3D Print Ticket#: " + self.Ticketnum + ".\n\n"
        body3 += "We are sorry but the following print request has been denied - Ticket#: " + self.Ticketnum + "\n\n"
        body3 += "The reasoning: " + str(reasonEntry.get()) + \
                 "\n\nSincerely,\n\nThe Lyons New Media Centre Staff\n\n"
        body3 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
        self.msg += body3
        LNMC = """library.mcmaster.ca/spaces/lyons"""
        self.msg += LNMC
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(self.User, self.Password)
        server.sendmail(self.User, self.patron_email, self.msg)
        server.quit()
        if (self.z == "1"):
            self.wks.update_cell(self.row_number, 17, "Y")
            format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtdenied)
            infoLab2 = tk.Label(self.infoFrame, text="Message Sent & Spreadsheet Updated!")
            infoLab2.pack()
            infoLab2.update()
            self.infoFrame.after(2000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.StartMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Message Sent!")
            infoLab1.pack()
            infoLab1.update()
            self.infoFrame.after(2000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.StartMenu()

    def Clarification_Skewed(self, ticketNumEntry):
        if (self.z == "1"):
            self.row_number = self.wks.find(self.name).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.Ticketnum = str(ticketNumEntry.get())
            self.rowstr = str(self.row_number)
        else:
            self.name = self.nameEntry.get()
            self.patron_email = self.emailEntry.get()
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        subject = "3D Print Request - Dimensions Clarifications Needed"
        self.msg += "Subject: " + subject + '\n\n'
        body4 = "Hi " + self.name + ",\n\n"
        body4 += "I'm looking at your print request - Ticket#: " + self.Ticketnum + ".\n\n"
        body4 += "Unfortunately, the dimensions you have submitted appear to skew the 3D model. " \
                 "You can use Cura, a free software to double check your dimensions." \
                 "\n\nOnce you have double checked, feel free to simply reply to this email with the " \
                 "new dimensions in this format:\n\n"
        body4 += "Width (x):\n\nDepth (y):\n\nHeight (z):\n\n"
        body4 += "Once we receive your clarification we will reprocess the print for you!\n\nSincerely,\n\nThe Lyons New Media Centre Staff\n\n"
        body4 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
        self.msg += body4
        LNMC = """library.mcmaster.ca/spaces/lyons"""
        self.msg += LNMC
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(self.User, self.Password)
        server.sendmail(self.User, self.patron_email, self.msg)
        server.quit()
        if (self.z == "1"):
            self.wks.update_cell(self.row_number, 17, "Y")
            format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtclarification)
            infoLab2 = tk.Label(self.infoFrame, text="Message Sent & Spreadsheet Updated!")
            infoLab2.pack()
            infoLab2.update()
            self.infoFrame.after(2000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.StartMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Message Sent!")
            infoLab1.pack()
            infoLab1.update()
            self.infoFrame.after(2000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.StartMenu()

    def Clarification_Large(self, ticketNumEntry):
        if (self.z == "1"):
            self.row_number = self.wks.find(self.name).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.Ticketnum = str(ticketNumEntry.get())
            self.rowstr = str(self.row_number)

        else:
            self.name = self.nameEntry.get()
            self.patron_email = self.emailEntry.get()
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        subject = "3D Print Request - Dimensions Clarifications Needed"
        self.msg += "Subject: " + subject + '\n\n'
        body5 = "Hi " + self.name + ",\n\n"
        body5 += "I'm looking at your print request - Ticket#: " + self.Ticketnum + ".\n\n"
        body5 += "Unfortunately, the dimensions you have submitted are too large for our 3D printers to handle. " \
                 "You can use Cura, a free software to resize your model to a size that fits within 5-6 hours.\n\n" \
                 " Once you have double checked the dimensions feel free to simply reply to this email with the new " \
                 "dimensions in this format:\n\n"
        body5 += "Width (x):\n\nDepth (y):\n\nHeight (z):\n\n"
        body5 += "Once we receive your clarification we will reprocess the print for you!\n\nSincerely,\n\nThe Lyons New Media Centre Staff\n\n"
        body5 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
        self.msg += body5
        LNMC = """library.mcmaster.ca/spaces/lyons"""
        self.msg += LNMC
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(self.User, self.Password)
        server.sendmail(self.User, self.patron_email, self.msg)
        server.quit()
        if (self.z == "1"):
            self.wks.update_cell(self.row_number, 17, "Y")
            format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtclarification)
            infoLab2 = tk.Label(self.infoFrame, text="Message Sent & Spreadsheet Updated!")
            infoLab2.pack()
            infoLab2.update()
            self.infoFrame.after(2000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.StartMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Message Sent!")
            infoLab1.pack()
            infoLab1.update()
            self.infoFrame.after(2000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.StartMenu()

    def Reminder(self, ticketNumEntry, dateEntry1, dateEntry2):
        if (self.z == "1"):
            self.row_number = self.wks.find(self.name).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.Ticketnum = str(ticketNumEntry.get())
            self.rowstr = str(self.row_number)
        else:
            self.name = self.nameEntry.get()
            self.patron_email = self.emailEntry.get()
        self.initialDate = str(dateEntry1.get())
        self.lastDate = str(dateEntry2.get())
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        subject = "3D Print Request - Reminder"
        self.msg += "Subject: " + subject + '\n\n'
        body6 = "Hi " + self.name + ",\n\n"
        body6 += "This is a reminder that your 3D print job - Ticket#:" + self.Ticketnum + " is ready for pickup\n\n"
        body6 += "Please see the original message sent on " + self.initialDate + \
                 " with instructions on picking up the item. If the item is not picked up by " + \
                 self.lastDate + ", we will discard it.\n\n"
        body6 += "If you cannot make it into the Centre due to work/being home etc., please let us know so we can " \
                 "arrange to hold onto it until you can make it in.\n\nSincerely,\n\nThe Lyons New Media Centre Staff\n\n"
        body6 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
        self.msg += body6
        LNMC = """library.mcmaster.ca/spaces/lyons"""
        self.msg += LNMC
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(self.User, self.Password)
        server.sendmail(self.User, self.patron_email, self.msg)
        server.quit()
        if (self.z == "1"):
            self.wks.update_cell(self.row_number, 17, "Y")
            infoLab2 = tk.Label(self.infoFrame, text="Message Sent & Spreadsheet Updated!")
            infoLab2.pack()
            infoLab2.update()
            self.infoFrame.after(2000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.StartMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Message Sent!")
            infoLab1.pack()
            infoLab1.update()
            self.infoFrame.after(2000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.StartMenu()

    def Failed(self, ticketNumEntry, reasonEntry):
        if (self.z == "1"):
            self.row_number = self.wks.find(self.name).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.Ticketnum = str(ticketNumEntry.get())
            self.rowstr = str(self.row_number)
        else:
            self.name = self.nameEntry.get()
            self.patron_email = self.emailEntry.get()
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        subject = "3D Print Request - Failed"
        self.msg += "Subject: " + subject + '\n\n'
        body7 = "Hi " + self.name + ",\n\n"
        body7 += "We are sorry but the following print request has not printed properly - Ticket#: " + self.Ticketnum + "\n\n"
        body7 += "What happened / suggestions for printing: " + str(reasonEntry.get()) \
                 + "\n\nSincerely,\n\nThe Lyons New Media Centre Staff\n\n"
        body7 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
        self.msg += body7
        LNMC = """library.mcmaster.ca/spaces/lyons"""
        self.msg += LNMC
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(self.User, self.Password)
        server.sendmail(self.User, self.patron_email, self.msg)
        server.quit()
        if (self.z == "1"):
            self.wks.update_cell(self.row_number, 17, "Y")
            format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtfailed)
            infoLab2 = tk.Label(self.infoFrame, text="Message Sent & Spreadsheet Updated!")
            infoLab2.pack()
            infoLab2.update()
            self.infoFrame.after(2000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.StartMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Message Sent!")
            infoLab1.pack()
            infoLab1.update()
            self.infoFrame.after(2000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.StartMenu()

    def pickedUp(self, ticketNumEntry):
        if (self.z == "1"):
            self.row_number = self.wks.find(self.name).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.Ticketnum = str(ticketNumEntry.get())
            self.rowstr = str(self.row_number)
            dateToday = date.today().strftime("%m/%d/%Y")
            format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtpickedup)
            self.wks.update_cell(self.row_number, 18, dateToday)
            infoLab2 = tk.Label(self.infoFrame, text="Spreadsheet updated")
            infoLab2.pack()
            infoLab2.update()
            self.infoFrame.after(2000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.StartMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Ticket not found, unable to update spreadsheet")
            infoLab1.pack()
            infoLab1.update()
            self.infoFrame.after(2000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.StartMenu()

    def nevPickedUp(self, ticketNumEntry):
        if (self.z == "1"):
            self.row_number = self.wks.find(self.name).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.Ticketnum = str(ticketNumEntry.get())
            self.rowstr = str(self.row_number)
            self.wks.update_cell(self.row_number, 19, "Reminder email sent but never picked up")
            format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtneverpickedup)
            infoLab2 = tk.Label(self.infoFrame, text="Spreadsheet updated")
            infoLab2.pack()
            infoLab2.update()
            self.infoFrame.after(2000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.StartMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Ticket not found, unable to update spreadsheet")
            infoLab1.pack()
            infoLab1.update()
            self.infoFrame.after(2000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.StartMenu()

    def cancelled(self, ticketNumEntry, reasonEntry):
        if (self.z == "1"):
            self.row_number = self.wks.find(self.name).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.Ticketnum = str(ticketNumEntry.get())
            self.rowstr = str(self.row_number)
            self.wks.update_cell(self.row_number, 19, str(reasonEntry.get()))
            format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtcancelled)
            print("3D Print has been cancelled")
            print("Spreadsheet Updated")
            infoLab2 = tk.Label(self.infoFrame, text="Spreadsheet updated")
            infoLab2.pack()
            infoLab2.update()
            self.infoFrame.after(2000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.StartMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Ticket not found, unable to update spreadsheet")
            infoLab1.pack()
            infoLab1.update()
            self.infoFrame.after(2000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.StartMenu()

    def findTicket(self, ticketNumEntry, function, text, title):
        self.Ticketnum = str(ticketNumEntry.get())
        # Exception Handling for when there's no match
        try:
            self.row_number = self.wks.find(self.Ticketnum).row
        except Exception as e:
            self.z = '0'
            infoLab1 = tk.Label(self.infoFrame, text="No matching Ticket Number found, Enter Patron info manually")
            infoLab1.pack()
            infoLab1.update()
            infoLab1.after(3000, infoLab1.destroy())
        else:
            if (self.Ticketnum == ''):
                self.z = '0'
                infoLab1 = tk.Label(self.infoFrame, text="No matching Ticket Number found, Enter Patron info manually")
                infoLab1.pack()
                infoLab1.update()
                infoLab1.after(3000, infoLab1.destroy())
            else:
                self.z = '1'
                self.name = self.wks.cell(self.row_number, 2).value
                self.patron_email = self.wks.cell(self.row_number, 3).value
                self.infoFrame.destroy()
                self.getInfo(function, text, title)
                infoLab1 = tk.Label(self.infoFrame, text="Found Matching Ticket Number")
                infoLab1.pack()
                infoLab1.update()
                infoLab1.after(2000, infoLab1.destroy())

    def findTicket2(self, ticketNumEntry, function, text, title):
        self.Ticketnum = str(ticketNumEntry.get())
        # Exception Handling for when there's no match
        try:
            self.row_number = self.wks.find(self.Ticketnum).row
        except Exception as e:
            self.z = '0'
            infoLab1 = tk.Label(self.infoFrame, text="No matching Ticket Number found, Enter Patron info manually")
            infoLab1.pack()
            infoLab1.update()
            infoLab1.after(3000, infoLab1.destroy())
        else:
            if (self.Ticketnum == ''):
                self.z = '0'
                infoLab1 = tk.Label(self.infoFrame, text="No matching Ticket Number found, Enter Patron info manually")
                infoLab1.pack()
                infoLab1.update()
                infoLab1.after(3000, infoLab1.destroy())
            else:
                self.z = '1'
                self.name = self.wks.cell(self.row_number, 2).value
                self.patron_email = self.wks.cell(self.row_number, 3).value
                self.infoFrame.destroy()
                self.getInfo2(function, text, title)
                infoLab1 = tk.Label(self.infoFrame, text="Found Matching Ticket Number")
                infoLab1.pack()
                infoLab1.update()
                infoLab1.after(2000, infoLab1.destroy())

    def findTicket3(self, ticketNumEntry, function, text, title):
        self.Ticketnum = str(ticketNumEntry.get())
        # Exception Handling for when there's no match
        try:
            self.row_number = self.wks.find(self.Ticketnum).row
        except Exception as e:
            self.z = '0'
            infoLab1 = tk.Label(self.infoFrame, text="No matching Ticket Number found, Enter Patron info manually")
            infoLab1.pack()
            infoLab1.update()
            infoLab1.after(3000, infoLab1.destroy())
        else:
            if (self.Ticketnum == ''):
                self.z = '0'
                infoLab1 = tk.Label(self.infoFrame, text="No matching Ticket Number found, Enter Patron info manually")
                infoLab1.pack()
                infoLab1.update()
                infoLab1.after(3000, infoLab1.destroy())
            else:
                self.z = '1'
                self.name = self.wks.cell(self.row_number, 2).value
                self.patron_email = self.wks.cell(self.row_number, 3).value
                self.infoFrame.destroy()
                self.getInfo3(function, text, title)
                infoLab1 = tk.Label(self.infoFrame, text="Found Matching Ticket Number")
                infoLab1.pack()
                infoLab1.update()
                infoLab1.after(2000, infoLab1.destroy())

    def findTicket4(self, ticketNumEntry, function, text, title):
        self.Ticketnum = str(ticketNumEntry.get())
        # Exception Handling for when there's no match
        try:
            self.row_number = self.wks.find(self.Ticketnum).row
        except Exception as e:
            self.z = '0'
            infoLab1 = tk.Label(self.infoFrame, text="No matching Ticket Number found, Enter Patron info manually")
            infoLab1.pack()
            infoLab1.update()
            infoLab1.after(3000, infoLab1.destroy())
        else:
            if (self.Ticketnum == ''):
                self.z = '0'
                infoLab1 = tk.Label(self.infoFrame, text="No matching Ticket Number found, Enter Patron info manually")
                infoLab1.pack()
                infoLab1.update()
                infoLab1.after(3000, infoLab1.destroy())
            else:
                self.z = '1'
                self.name = self.wks.cell(self.row_number, 2).value
                self.patron_email = self.wks.cell(self.row_number, 3).value
                self.infoFrame.destroy()
                self.getInfo4(function, text, title)
                infoLab1 = tk.Label(self.infoFrame, text="Found Matching Ticket Number")
                infoLab1.pack()
                infoLab1.update()
                infoLab1.after(2000, infoLab1.destroy())


Window()
