import gc
import time
import tkinter as tk
# intsall:oauth2client, gspread, PyOpenSSL, gspread-formatting
import gspread
from gspread_formatting import *

from oauth2client.service_account import ServiceAccountCredentials

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from datetime import date
#boopdffd

class Window():
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

        self.credentials = ServiceAccountCredentials.from_json_keyfile_dict(self.credsdict, self.scope)

        self.gc = gspread.authorize(self.credentials)

        self.sh = self.gc.open('3D Printing Requests')

        self.worksheet_list = self.sh.worksheets()

        self.wks = ""

        self.Ticketnum = ""

        self.row_number = ""

        self.z = ""

        self.rowstr = ""

        # Define Name Parameter
        self.name = ""

        # Assign Name to variable
        # a = getPatronName(name)
        self.a = ""

        # Define Patron Email parameter
        self.patron_email = ""

        # Assign Patron Email to variable
        # b = (getPatronEmail(patron_email))
        self.b = ""

        # define message for email
        self.msg = MIMEMultipart()

        self.c = ""

        self.x1 = ""

        self.workSDict = {}

        for i in self.worksheet_list:
            self.workSDict.update({str(i) : i})

        # Functions

        # Setup window
        self.window = tk.Tk()
        self.window.title("3D Print Request")
        self.window.geometry("500x800")
        self.window.resizable(1, 1)

        # Initiate Menu
        self.StartMenu()

        (self.window).mainloop()

    def StartMenu(self):
        self.titleFrame = tk.Frame(self.window)
        self.titleFrame.pack()

        tk.Label(self.titleFrame, text="Choose the spreadsheet you would like to edit").pack()

        self.workSheet = tk.StringVar(self.titleFrame)
        self.workSheet.set(self.worksheet_list[0])  # default value
        tk.OptionMenu(self.titleFrame, self.workSheet, *self.worksheet_list).pack()

        tk.Label(self.titleFrame, text="").pack()

        tk.Button(self.titleFrame, text="New Submission Processing", width="40", pady="5", command = self.getInfoNewEntry).pack()
        tk.Button(self.titleFrame, text="Ready For Pickup", width="40", pady="5", command=lambda:self.getInfo(self.readyForPickup,"Send Email")).pack()
        tk.Button(self.titleFrame, text="Delay Printing", width="40", pady="5", command=lambda:self.getInfo(self.DelayedPrinting,"Send Email")).pack()
        tk.Button(self.titleFrame, text="Denied", width="40", pady="5", command=self.getInfo).pack()
        tk.Button(self.titleFrame, text="Dimensions Clarification - Skewed Print", width="40", pady="5",
                  command=self.getInfo).pack()
        tk.Button(self.titleFrame, text="Dimensions Clarification - Large Print", width="40", pady="5",
                  command=self.getInfo).pack()
        tk.Button(self.titleFrame, text="Reminder", width="40", pady="5", command=self.getInfo).pack()
        tk.Button(self.titleFrame, text="Failed", width="40", pady="5", command=self.getInfo).pack()
        tk.Button(self.titleFrame, text="Picked Up", width="40", pady="5", command=self.getInfo).pack()
        tk.Button(self.titleFrame, text="Never Picked Up", width="40", pady="5", command=self.getInfo).pack()
        tk.Button(self.titleFrame, text="Cancelled", width="40", pady="5", command=self.getInfo).pack()

    def backToMenu(self):
        self.infoFrame.destroy()
        self.wks = ""

        self.Ticketnum = ""

        self.row_number = ""

        self.z = ""

        self.rowstr = ""

        # Define Name Parameter
        self.name = ""

        # Assign Name to variable
        # a = getPatronName(name)
        self.a = ""

        # Define Patron Email parameter
        self.patron_email = ""

        # Assign Patron Email to variable
        # b = (getPatronEmail(patron_email))
        self.b = ""

        # define message for email
        self.msg = MIMEMultipart()

        self.c = ""

        self.x1 = ""
        self.StartMenu()

    def getInfo(self,function,text):
        self.wks = self.workSDict[self.workSheet.get()]
        self.titleFrame.destroy()
        self.infoFrame = tk.Frame(self.window)
        self.infoFrame.pack()
        nameEntry = tk.StringVar(self.infoFrame, value=self.name)
        ticketNumEntry = tk.StringVar(self.infoFrame, value = self.Ticketnum)
        emailEntry = tk.StringVar(self.infoFrame, value = self.patron_email)
        tk.Label(self.infoFrame, text="Enter Ticket #:").pack()
        tk.Entry(self.infoFrame, textvariable=ticketNumEntry).pack()
        tk.Button(self.infoFrame, text="Search", command=lambda:self.findTicket(ticketNumEntry,function,text)).pack()
        tk.Label(self.infoFrame, text="Enter Patron Name:").pack()
        tk.Entry(self.infoFrame, textvariable=nameEntry).pack()
        tk.Label(self.infoFrame, text="Enter Patron Email:").pack()
        tk.Entry(self.infoFrame, textvariable=emailEntry).pack()
        tk.Button(self.infoFrame, text = text, command = lambda:function(ticketNumEntry)).pack()
        tk.Button(self.infoFrame, text="Back to Menu", command=self.backToMenu).pack()
    def getInfoNewEntry(self):
        self.wks = self.workSDict[self.workSheet.get()]
        self.titleFrame.destroy()
        self.infoFrame = tk.Frame(self.window)
        self.infoFrame.pack()
        nameEntry = tk.StringVar(self.infoFrame, value=self.name)
        ticketNumEntry = tk.StringVar(self.infoFrame, value = self.Ticketnum)
        emailEntry = tk.StringVar(self.infoFrame, value = self.patron_email)
        StaffInitials = tk.StringVar(self.infoFrame)
        dateToday = tk.StringVar(self.infoFrame, value = date.today().strftime("%m/%d/%Y"))
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
        tk.Radiobutton(CourseYNF,text="Yes", padx=20,variable=CourseYN, value=1).pack(side = "left")
        tk.Radiobutton(CourseYNF, text="No",padx=20,variable=CourseYN,value=0).pack(side = "left")
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
        tk.Button(self.infoFrame, text = "Submit", command = lambda:self.defineNewPatronInfo(nameEntry.get(),ticketNumEntry.get(),emailEntry.get(),dateToday.get(),StaffInitials.get(),CourseYN.get(),CourseCode.get(),affiliation.get(),department.get(),research.get(),ownC.get(),consent.get(),handle.get(),SD.get(),Fname.get(),Ptime.get())).pack()
        tk.Button(self.infoFrame, text="Back to Menu", command=self.backToMenu).pack()

    def defineNewPatronInfo(self,nameEntry,ticketEntry,emailEntry,dateToday,StaffInitials,CourseYN,CourseCode,affiliation,department,research,ownC,consent,handle,SD,Fname,Ptime):
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
    def Denied(self, ticketNumEntry):
        self.row_number = self.wks.find(self.name).row
        self.name = self.wks.cell(self.row_number, 2).value
        self.patron_email = self.wks.cell(self.row_number, 3).value
        self.Ticketnum = str(ticketNumEntry.get())

        self.rowstr = str(self.row_number)
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        x1 = 1
        subject = "3D Print Request - Delayed Printing"
        self.msg += "Subject: " + subject + '\n\n'
        body2 = "Hi " + self.name + ",\n\n"
        body2 += "This is in regards to 3D Print Ticket#: " + self.Ticketnum + ".\n\n"
        body2 += "We have had an unusual amount of course-related print requests submitted this term, and are prioritizing " \
                 "those requests before regular requests. Because of this, we may not be able to complete your request " \
                 "by April 26th (last day of exams), so it may be completed as we're going into May " \
                 "(during the summer months).\n\n"
        body2 += "We need to know if you would still like this ticket to be printed, knowing that there is a delay that it " \
                 "may not be printed before the term is over. If you still want it to be printed, but you are not able to " \
                 "pick it up immediately since it may be completed during the summer, you can let us know to hold it for " \
                 "you till you can.\n\n"
        body2 += "Please respond to this email by Friday, May 4th, 2018. If we do not hear from you by that date, we will " \
                 "assume it is unwanted and will cancel the request.\n\nThank you\n\nLyons New Media Centre\n\n"
        body2 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
        self.msg += body2
        LNMC = """library.mcmaster.ca/spaces/lyons"""
        self.msg += LNMC
        print("\n" + self.msg)
        print(self.rowstr)
        sender = "lyons.newmedia@gmail.com"
        password = "DigitalM3dia"
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender, password)
        server.sendmail(sender, self.patron_email, self.msg)
        server.quit()

        infoLab1 = tk.Label(self.infoFrame, text="Message Sent!")
        infoLab1.pack()
        infoLab1.update()
        time.sleep(1)
        infoLab1.destroy()
        self.infoFrame.destroy()
        self.StartMenu()

        print("Spreadsheet Updated")
        self.wks.update_cell(self.row_number, 17, "Y")
        print("Message Sent")
    def DelayedPrinting(self, ticketNumEntry):
        self.row_number = self.wks.find(self.name).row
        self.name = self.wks.cell(self.row_number, 2).value
        self.patron_email = self.wks.cell(self.row_number, 3).value
        self.Ticketnum = str(ticketNumEntry.get())

        self.rowstr = str(self.row_number)
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        x1 = 1
        subject = "3D Print Request - Delayed Printing"
        self.msg += "Subject: " + subject + '\n\n'
        body2 = "Hi " + self.name + ",\n\n"
        body2 += "This is in regards to 3D Print Ticket#: " + self.Ticketnum + ".\n\n"
        body2 += "We have had an unusual amount of course-related print requests submitted this term, and are prioritizing " \
                 "those requests before regular requests. Because of this, we may not be able to complete your request " \
                 "by April 26th (last day of exams), so it may be completed as we're going into May " \
                 "(during the summer months).\n\n"
        body2 += "We need to know if you would still like this ticket to be printed, knowing that there is a delay that it " \
                 "may not be printed before the term is over. If you still want it to be printed, but you are not able to " \
                 "pick it up immediately since it may be completed during the summer, you can let us know to hold it for " \
                 "you till you can.\n\n"
        body2 += "Please respond to this email by Friday, May 4th, 2018. If we do not hear from you by that date, we will " \
                 "assume it is unwanted and will cancel the request.\n\nThank you\n\nLyons New Media Centre\n\n"
        body2 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
        self.msg += body2
        LNMC = """library.mcmaster.ca/spaces/lyons"""
        self.msg += LNMC
        print("\n" + self.msg)
        print(self.rowstr)
        sender = "lyons.newmedia@gmail.com"
        password = "DigitalM3dia"
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender, password)
        server.sendmail(sender, self.patron_email, self.msg)
        server.quit()

        infoLab1 = tk.Label(self.infoFrame, text="Message Sent!")
        infoLab1.pack()
        infoLab1.update()
        time.sleep(1)
        infoLab1.destroy()
        self.infoFrame.destroy()
        self.StartMenu()

        print("Spreadsheet Updated")
        self.wks.update_cell(self.row_number, 17, "Y")
        print("Message Sent")

    def readyForPickup(self,ticketNumEntry):
        self.row_number = self.wks.find(self.name).row
        self.name = self.wks.cell(self.row_number, 2).value
        self.patron_email = self.wks.cell(self.row_number, 3).value
        self.Ticketnum = str(ticketNumEntry.get())

        self.rowstr = str(self.row_number)
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        x1 = 1
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
        print("\n" + self.msg)
        print(self.rowstr)
        format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtreadypickup)
        sender = "lyons.newmedia@gmail.com"
        password = "DigitalM3dia"
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender, password)
        server.sendmail(sender, self.patron_email, self.msg)
        server.quit()

        infoLab1 = tk.Label(self.infoFrame, text="Message Sent!")
        infoLab1.pack()
        infoLab1.update()
        time.sleep(1)
        infoLab1.destroy()
        self.infoFrame.destroy()
        self.StartMenu()




        print("Spreadsheet Updated")
        self.wks.update_cell(self.row_number, 17, "Y")
        print("Message Sent")

    def findTicket(self,ticketNumEntry,function,text):
        self.Ticketnum = str(ticketNumEntry.get())
        # Exception Handling for when there's no match
        try:
            self.row_number = self.wks.find(self.Ticketnum).row
        except Exception as e:
            self.z += '0'
            infoLab1 = tk.Label(self.infoFrame,text="No matching Ticket Number, Enter Patron info manually")
            infoLab1.pack()
            infoLab1.update()
            time.sleep(1)
            infoLab1.destroy()

        else:
            if(self.Ticketnum == ''):
                self.z += '0'
                infoLab1 = tk.Label(self.infoFrame, text="No matching Ticket Number, Enter Patron info manually")
                infoLab1.pack()
                infoLab1.update()
                time.sleep(1)
                infoLab1.destroy()
            else:
                self.z += '1'
                self.name = self.wks.cell(self.row_number, 2).value
                self.patron_email = self.wks.cell(self.row_number, 3).value
                self.infoFrame.destroy()
                self.getInfo(function,text)
                infoLab1 = tk.Label(self.infoFrame, text="Found Matching Ticket Number")
                infoLab1.pack()
                infoLab1.update()
                time.sleep(1)
                infoLab1.destroy()

Window()
