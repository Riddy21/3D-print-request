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
        self.window.geometry("500x500")
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

        tk.Button(self.titleFrame, text="New Submission Processing", width="40", pady="5").pack()
        tk.Button(self.titleFrame, text="New Manual Submission Processing", width="40", pady="5").pack()
        tk.Button(self.titleFrame, text="Ready For Pickup", width="40", pady="5", command=self.getInfo).pack()
        tk.Button(self.titleFrame, text="Delay Printing", width="40", pady="5", command=self.getInfo).pack()
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

    def getInfo(self):
        self.wks = self.workSDict[self.workSheet.get()]
        self.titleFrame.destroy()
        self.infoFrame = tk.Frame(self.window)
        self.infoFrame.pack()
        nameEntry = tk.StringVar(self.infoFrame, value=self.name)
        ticketNumEntry = tk.StringVar(self.infoFrame, value = self.Ticketnum)
        emailEntry = tk.StringVar(self.infoFrame, value = self.patron_email)
        tk.Label(self.infoFrame, text="Enter Ticket #:").pack()
        tk.Entry(self.infoFrame, textvariable=ticketNumEntry).pack()
        tk.Button(self.infoFrame, text="Search", command=lambda:self.findTicket(ticketNumEntry)).pack()
        tk.Label(self.infoFrame, text="Enter Patron Name:").pack()
        tk.Entry(self.infoFrame, textvariable=nameEntry).pack()
        tk.Label(self.infoFrame, text="Enter Patron Email:").pack()
        tk.Entry(self.infoFrame, textvariable=emailEntry).pack()
        tk.Button(self.infoFrame, text = "Go", command = lambda:self.definePatronInfo(ticketNumEntry)).pack()
        tk.Button(self.infoFrame, text="Back to Menu", command=self.backToMenu).pack()
    def definePatronInfo(self,ticketNumEntry):
        self.name = self.wks.cell(self.row_number, 2).value
        self.patron_email = self.wks.cell(self.row_number, 3).value
        self.Ticketnum = str(ticketNumEntry.get())
        print(self.name,self.patron_email,self.Ticketnum)

    def findTicket(self,ticketNumEntry):
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
                self.getInfo()
                infoLab1 = tk.Label(self.infoFrame, text="Found Matching Ticket Number")
                infoLab1.pack()
                infoLab1.update()
                time.sleep(1)
                infoLab1.destroy()


Window()
