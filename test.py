# intsall:oauth2client, gspread, PyOpenSSL, gspread-formatting
import gspread
from gspread_formatting import *

from oauth2client.service_account import ServiceAccountCredentials

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credsdict = {'type': 'service_account',
             'project_id': 'lyons-email-updatesheet-script',
             'private_key_id': '2414046f0774321dd439b121749c6eff1db8713b',
             'private_key': '-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQDrHJwzqsP88DVg\nHplpDCpKYwUYGUEbRg0sqFx7rJy9mYOXD7/TpducJlInQs5PbTiF/yQ0o/IjAaSB\nWWqaQbCv+COSVg5udAcbrP6xSTv+3L2crqFMDXnHbKHHfhvNTQSyfh2JSIZY78le\nn6dcq0g3MvHEt7Qy89bdgC966XsXXVe+va0t8lI26IgPf9ZTSFm4a4TYStsT3mP3\nKl4RegTn03bMHYdPmOm2B/I5E814RG8Wjsk0FIYhW6reYNhLuBLrS14744bWmd6f\nYvZLpKmEf5EN91nD/veWoKB2qLhESfmpP1iipieoVpm+VvfDlWmcusl8s+w0/ETf\nkk1BMo5dAgMBAAECggEAE3cw744p3907bhPae7oIHlSIbXBZ1Zo9KP9feNXXvFLj\ndDRXm3xV7F2324xKbIUMcvum0bzpJUDTj+oJS3A44rjWqRz64OY2WHJAPAlmMDmy\ncTB8JkHPXVV/J3cnch34T5bldyJMDTz9HRp2ztNXjUpoffL/tmA93+TnCXQfPtXQ\n9vo1nj2n3Pix4zSZSwMk18Ll6vERbjKHffaSc/xdvMuHgEFa48ze3cOmrpFyJHM1\nrdZ6qhCUQPwgaHHFTgKcIDU5ILVlmuIm1DnUCo3K0ocDdov7zwyU/J/5gUoFzOpW\nCA7pwy2QbLZ3XWGL01gV6d1mGqiF9ajNxLBezu7AwQKBgQD4/44ywpno0nbEXlNj\n8zLpWGJyFkVDYwdPtcPlQZzPdq/UJd0bfApSnRaNDTdNwTbYtUSyT3QlXUOgzjai\nn2zKf6cXRtaGScGnmF4acJFvBSw9xTaaDaLs7IVoGD/CaCPwwJJ71OBLRTPxvIev\ntMfWWsANpYy/6pipuJI3Ujb/ZQKBgQDxuRdLN8Ox1OVQxAXUbykqyeuNt5h+h/vk\n9SdrYu72KmdHIUHgfu1L9N9yFp/KpGZ4CnhlkGX0LXoDnqQKd1yjiLxGrWgLk76P\n/dXCpLs4N7av94je5Z2e2Q1v+1qQ8cAs13UCcwemHe1WxRfJqsj8kP3ZX2FFGUke\nJgzZGpkPmQKBgDISWgsVHRQ3tpB4k3ZnApbwIiPlHJqXgHHkEHe6wQjrSiJ0Vslf\nIUhJtK46uSNWtmvPz/e3iJi275GXxl7fhmYWU4iXwy4QCPRl7I6OkoBr3uCxFvDV\nyyyvx4gOUEwM2yVf5FUoks4wJWj4S6TmysTtTO+xmeNCDt8acbTUQKENAoGAPH0k\n5x29SvMLr3peOxrWIm8FEyGud3tv/YubobPQOKnDznj0E0mv+CH/CH3A3uTk/4Uf\nO8s2uDPpJJ6+TiAwfnvpIYajUsJWHZJXu62dbCQFA2PeTGkJWIbYZf1wXHUishX4\nofRHJbq3ec84dK7YPNvLqmnD3ZbGRVUgQfP1+YECgYBiwykzHoHeY4e1yAbDzPzU\nwvlxE6bGgJXJLKObhRJjoFn5zyAEteMFOdTh7OUfcWIxi3/HdxIR8k8FvupNAOAi\nBQYLiMjIRLCJ3ljR3JHN3fDrwVpNOBEdtT1S7i8jWTA0tb2XSn9mN3IDjrdE5aqU\n2hm0MtKP0tWBhzfllvvOHg==\n-----END PRIVATE KEY-----\n',
             'client_email': 'id-d-print-request@lyons-email-updatesheet-script.iam.gserviceaccount.com',
             'client_id': '117542574436878496508',
             'auth_uri': 'https://accounts.google.com/o/oauth2/auth',
             'token_uri': 'https://oauth2.googleapis.com/token',
             'auth_provider_x509_cert_url': 'https://www.googleapis.com/oauth2/v1/certs',
             'client_x509_cert_url': 'https://www.googleapis.com/robot/v1/metadata/x509/id-d-print-request%40lyons-email-updatesheet-script.iam.gserviceaccount.com'}

fmtreadypickup = CellFormat(
    backgroundColor=Color(0.078, 0.616, 1),
)
fmtdenied = CellFormat(
    backgroundColor=Color(0.878, 0.4, 0.4),
)
fmtfailed = CellFormat(
    backgroundColor=Color(0.957, 0.8, 0.8),
)
fmtclarification = CellFormat(
    backgroundColor=Color(1, 0.851, 0.4),
)
fmtcancelled = CellFormat(
    backgroundColor=Color(0.71, 0.37, 0.02)
)
fmtpickedup = CellFormat(
    backgroundColor=Color(0.85, 0.85, 0.85),
)
fmtneverpickedup = CellFormat(
    backgroundColor=Color(0.4, 0.31, 0.65),
)

credentials = ServiceAccountCredentials.from_json_keyfile_dict(credsdict, scope)

gc = gspread.authorize(credentials)

sh = gc.open('3D Printing Requests')

worksheet_list = sh.worksheets()

wks = ""

Ticketnum = ""

row_number = ""

z = ""

rowstr = ""

# Define Name Parameter
name = ""

# Assign Name to variable
# a = getPatronName(name)
a = ""

# Define Patron Email parameter
patron_email = ""

# Assign Patron Email to variable
# b = (getPatronEmail(patron_email))
b = ""

# define message for email
msg = MIMEMultipart()

c = ""

x1 = ""


# Functions

# Select Worksheet to find Ticket Number
def get_worksheet():
    global wks
    # using enumerate to get index value of list
    max_index = (len(worksheet_list))
    index_count = str(list(range(max_index)))
    print("Worksheet index-values are:" + index_count + "\n")
    user = ""

    for index, value in enumerate(worksheet_list):
        print(index, value)
    while True:
        try:
            user = (int(input("\nEnter index value according to the time period of the print request:")))
        except ValueError:
            print("\nInvalid Entry, please enter the index value corresponding to the worksheet you want to select")
            continue

        if user < 0:
            print(
                "\nInvalid Entry, the index value cannot be negative, please enter the index value corresponding to the worksheet you want to select.")
            continue

        elif user > max_index - 1:
            print(
                "\nInvalid Entry, the value you entered is too large, please enter the index value corresponding to the worksheet you want to select.")
            continue
        else:
            break
    wks = sh.get_worksheet(user)


# Search spreadsheet for Ticket#
def get_ticketNumber():
    global z, row_number, Ticketnum
    Ticketnum = input('\nEnter Ticket #:')
    print(type(wks))
    # Exception Handling for when there's no match
    try:
        row_number = wks.find(Ticketnum).row
    except Exception as e:
        z += '0'
        print("\nNo matching Ticket Number, Enter Patron info manually")
    else:
        z += '1'
        print("\nFound Matching Ticket Number")


# Search spreadsheet for Patron Name
def get_patronName():
    global a, name, z
    if z == '0':
        name += input('\nEnter Patron Name:')
    else:
        while True:
            try:
                dialog = int(input("\nEnter '1' to search spreadsheet for Patron Name\n\nEnter '2' to enter manually"))
            except ValueError:
                print("\nInvalid Entry, please enter either 1 or 2")
                continue

            if dialog < 1:
                print("\nInvalid Entry, please enter either 1 or 2")
                continue

            elif dialog > 2:
                print("\nInvalid Entry, please enter either 1 or 2")
                continue

            if dialog == 1:
                name += wks.cell(row_number, 2).value
                break

            elif dialog == 2:
                name += input('\nEnter Patron Name:')
                break
    print(name)
    a = name  # Assign Name to variable


# Search spreadsheet for Patron Email
def get_patronEmail():
    global b, patron_email, z
    if z == '0':
        patron_email += input('\nEnter Patron Email:')
    else:
        while True:
            try:
                dialog = int(input("\nEnter '1' to search spreadsheet for Patron Email\n\nEnter '2' to enter manually"))
            except ValueError:
                print("\nInvalid Entry, please enter either 1 or 2")
                continue

            if dialog < 1:
                print("\nInvalid Entry, please enter either 1 or 2")
                continue

            elif dialog > 2:
                print("\nInvalid Entry, please enter either 1 or 2")
                continue

            elif dialog == 1:
                patron_email += wks.cell(row_number, 3).value
                break

            elif dialog == 2:
                patron_email += input('\nEnter Patron Email:')
                break
    print(patron_email)
    b = patron_email


# build message for Email
def message_builder():
    global msg, c, x1, Ticketnum
    msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
    # select message type
    while True:
        try:
            msgtype = int(
                input('\nEnter the number corresponding to the required message/action\n\n1 = Ready for Pickup'
                      '\n\n2 = Delayed Printing\n\n3 = Denied\n\n4 = Dimension Clarification - Skewed Print\n\n5= '
                      'Dimension Clarification - Large Print\n\n6 = Reminder'
                      '\n\n7 = Failed\n\n8 = Picked Up\n\n9 = Never Picked Up\n\n10 = Cancelled'))
        except ValueError:
            print("\nInvalid Entry, please enter a number from 1 to 10")
            continue

        if msgtype < 1:
            print("\nThe number you entered is less than 1, please enter a number from 1 to 10")
            continue

        elif msgtype > 10:
            print("\nThe number you entered is more than 10, please enter a number from 1 to 10")
            continue

        elif msgtype == 1:  # Message type: Ready for Pickup
            x1 = 1
            subject = "3D Print Request - Ready for Pickup"
            msg += "Subject: " + subject + '\n\n'
            body1 = "Hi " + a + ",\n\nGood news! The following requested 3D print job has been printed successfully:\n\n"
            body1 += "Ticket #: " + Ticketnum + "\n\nPlease bring this email and your McMaster ID card with you to the Help Desk " \
                                                "in Lyons New Media Centre (Mills Library, 4th floor) to retrieve your item.\n\n"
            body1 += "You will be required to sign for it, so a proxy cannot come to pick this up for you.\n\nWe will hold this " \
                     "item for no more than 30 days from today's date before it is reclaimed and/or recycled.  " \
                     "If you cannot make it into the Centre due to work/being home etc., please let us know and we can arrange to " \
                     "hold onto it until you can make it in.\n\nSincerely,\n\nLyons New Media Centre Staff\n\n"
            body1 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
            msg += body1
            LNMC = """library.mcmaster.ca/spaces/lyons"""
            msg += LNMC
            print("\n" + msg)
            break

        elif msgtype == 2:  # Message type: Delayed Printing
            x1 = 2
            subject = "3D Print Request - Delayed Printing"
            msg += "Subject: " + subject + '\n\n'
            body2 = "Hi " + a + ",\n\n"
            body2 += "This is in regards to 3D Print Ticket#: " + Ticketnum + ".\n\n"
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
            msg += body2
            LNMC = """library.mcmaster.ca/spaces/lyons"""
            msg += LNMC
            print("\n" + msg)
            break

        elif msgtype == 3:  # Message type: Denied
            x1 = 3
            subject = "3D Print Request - Denied"
            msg += "Subject: " + subject + '\n\n'
            body3 = "Hi " + a + ",\n\n"
            body3 += "We are sorry but the following print request has been denied - Ticket#: " + Ticketnum + "\n\n"
            body3 += "The reasoning: " + input('enter reasoning for denied 3d print request') + \
                     "\n\nSincerely,\n\nThe Lyons New Media Centre Staff\n\n"
            body3 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
            msg += body3
            LNMC = """library.mcmaster.ca/spaces/lyons"""
            msg += LNMC
            print("\n" + msg)
            break

        elif msgtype == 4:  # Message type: Dimension Clarification - Skewed Print
            x1 = 4
            subject = "3D Print Request - Dimensions Clarifications Needed"
            msg += "Subject: " + subject + '\n\n'
            body4 = "Hi " + a + ",\n\n"
            body4 += "I'm looking at your print request - Ticket#: " + Ticketnum + ".\n\n"
            body4 += "Unfortunately, the dimensions you have submitted appear to skew the 3D model. " \
                     "You can use Cura, a free software to double check your dimensions." \
                     "\n\nOnce you have double checked, feel free to simply reply to this email with the " \
                     "new dimensions in this format:\n\n"
            body4 += "Width (x):\n\nDepth (y):\n\nHeight (z):\n\n"
            body4 += "Once we receive your clarification we will reprocess the print for you!\n\nSincerely,\n\nThe Lyons New Media Centre Staff\n\n"
            body4 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
            msg += body4
            LNMC = """library.mcmaster.ca/spaces/lyons"""
            msg += LNMC
            print("\n" + msg)
            break

        elif msgtype == 5:  # Message type: Dimension Clarification - Large Print
            x1 = 5
            subject = "3D Print Request - Dimensions Clarifications Needed"
            msg += "Subject: " + subject + '\n\n'
            body4 = "Hi " + a + ",\n\n"
            body4 += "I'm looking at your print request - Ticket#: " + Ticketnum + ".\n\n"
            body4 += "Unfortunately, the dimensions you have submitted are too large for our 3D printers to handle. " \
                     "You can use Cura, a free software to resize your model to a size that fits within 5-6 hours.\n\n" \
                     " Once you have double checked the dimensions feel free to simply reply to this email with the new " \
                     "dimensions in this format:\n\n"
            body4 += "Width (x):\n\nDepth (y):\n\nHeight (z):\n\n"
            body4 += "Once we receive your clarification we will reprocess the print for you!\n\nSincerely," \
                     "\n\nThe Lyons New Media Centre Staff\n\n"
            body4 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
            msg += body4
            LNMC = """library.mcmaster.ca/spaces/lyons"""
            msg += LNMC
            print("\n" + msg)
            break

        elif msgtype == 6:  # Message type: Reminder
            x1 = 6
            subject = "3D Print Request - Reminder"
            msg += "Subject: " + subject + '\n\n'
            body5 = "Hi " + a + ",\n\n"
            body5 += "This is a reminder that your 3D print job - Ticket#:" + Ticketnum + " is ready for pickup\n\n"
            body5 += "Please see the original message sent on " + input('Enter date of original message (dd/mm/yyyy)') + \
                     " with instructions on picking up the item. If the item is not picked up by " + \
                     input('Enter date of last date to pickup item (dd/mm/yyyy)') + ", we will discard it.\n\n"
            body5 += "If you cannot make it into the Centre due to work/being home etc., please let us know so we can " \
                     "arrange to hold onto it until you can make it in.\n\nSincerely,\n\nThe Lyons New Media Centre Staff\n\n"
            body5 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
            msg += body5
            LNMC = """library.mcmaster.ca/spaces/lyons"""
            msg += LNMC
            print("\n" + msg)
            break

        elif msgtype == 7:  # Message type: Failed
            x1 = 7
            subject = "3D Print Request - Failed"
            msg += "Subject: " + subject + '\n\n'
            body6 = "Hi " + a + ",\n\n"
            body6 += "We are sorry but the following print request has not printed properly - Ticket#: " + Ticketnum + "\n\n"
            body6 += "What happened / suggestions for printing: " + input(
                'Enter reason for 3D Print failure/suggestions for printing') \
                     + "\n\nSincerely,\n\nThe Lyons New Media Centre Staff\n\n"
            body6 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
            msg += body6
            LNMC = """library.mcmaster.ca/spaces/lyons"""
            msg += LNMC
            print("\n" + msg)
            break

        elif msgtype == 8:  # Message type: Picked Up
            x1 = 8
            msg += ''
            print("\n3D Print has been picked up")
            break

        elif msgtype == 9:  # Message type: Never Picked Up
            x1 = 9
            msg += ''
            print("\n3D Print has never been picked up")
            break

        elif msgtype == 10:  # Message type: Cancelled
            x1 = 10
            msg += ''
            print("\n3D Print has been cancelled")
            break

    # Assign message to variable
    c += msg


# Send Email
def send_mail():
    global z, rowstr, row_number, x1
    # Email Account Login
    sender = "lyons.newmedia@gmail.com"
    password = "DigitalM3dia"
    # confirm submission or cancel message
    rowstr = str(row_number)
    if 0 < x1 < 8:
        while True:
            try:
                send = int(input("\nEnter '1' to send email to patron\n\nEnter '2' to cancel"))

            except ValueError:
                print("\nInvalid Entry, please enter either 1 or 2")
                continue

            if send < 1:
                print("\nThe number you entered is less than 1, please enter either 1 or 2")
                continue

            elif send > 2:
                print("\nThe number you entered is more than 2, please enter either 1 or 2")
                continue

            elif send == 1:  # Confirmation of sent email
                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.starttls()
                server.login(sender, password)
                server.sendmail(sender, b, c)
                server.quit()
                if z == '1':
                    if x1 == 1:
                        format_cell_range(wks, 'A' + rowstr + ':AC' + rowstr, fmtreadypickup)
                    elif x1 == 3:
                        format_cell_range(wks, 'A' + rowstr + ':AC' + rowstr, fmtdenied)
                    elif x1 == 4:
                        format_cell_range(wks, 'A' + rowstr + ':AC' + rowstr, fmtclarification)
                    elif x1 == 5:
                        format_cell_range(wks, 'A' + rowstr + ':AC' + rowstr, fmtclarification)
                    elif x1 == 7:
                        format_cell_range(wks, 'A' + rowstr + ':AC' + rowstr, fmtfailed)
                    print("\nSpreadsheet Updated")
                    wks.update_cell(row_number, 17, "Y")
                elif z == '0':
                    print("\nSpreadsheet formatting not required")
                print("\nMessage Sent")
                break

            elif send == 2:  # Confirmation of cancelled email
                print("\nMessage cancelled")
                break
    elif x1 == 8:
        if z == '1':
            date = input('\nEnter date of pickup (month/date/year)')
            format_cell_range(wks, 'A' + rowstr + ':AC' + rowstr, fmtpickedup)
            wks.update_cell(row_number, 18, date)
            print("\nSpreadsheet Updated")
        elif z == '0':
            print("\nUnable to update spreadsheet as ticket number wasn't found")
    elif x1 == 9:
        if z == '1':
            wks.update_cell(row_number, 19, "Reminder email sent but never picked up")
            format_cell_range(wks, 'A' + rowstr + ':AC' + rowstr, fmtneverpickedup)
            print("\nSpreadsheet Updated")
        elif z == '0':
            print("\nUnable to update spreadsheet as ticket number wasn't found")
    elif x1 == 10:
        if z == '1':
            r = input("\nEnter reason for print cancellation")
            wks.update_cell(row_number, 19, r)
            format_cell_range(wks, 'A' + rowstr + ':AC' + rowstr, fmtcancelled)
            print("\nSpreadsheet Updated")
        elif z == '0':
            print("\nUnable to update spreadsheet as ticket number wasn't found")
    input("\nProcess completed! Press return to exit")


# Main Function
def main():
    get_worksheet()
    get_ticketNumber()
    get_patronName()
    get_patronEmail()
    message_builder()
    send_mail()


main()
