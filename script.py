# coding=utf-8
import smtplib
import xlrd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

server = smtplib.SMTP('smtp-mail.outlook.com', 587)


def start():
    """Establishes a connection with the Outlook SMTP server and logs into your account."""

    try:
        server.starttls()
        print("Successful connection to Outlook server")
        print("--------------------------")
        sender = input("Enter your Outlook email address: ")
        pwd = input("Enter your Outlook password: ")
        print("--------------------------")
        server.login(sender, pwd)
        print("Successfully logged into Outlook")
        print("--------------------------")
        return sender
    except Exception as e:
        print("Unable to login. Check that the login information is correct")
        print(e)
        print("--------------------------")
        quit()


def get_info():
    """Reads a spreadsheet from the given path and locates the column with email addresses."""

    path = "data.xlsx"  # change path depending on the name and location of the file
    xl_book = xlrd.open_workbook(path)
    xl_sheet = xl_book.sheet_by_index(0)  # selects the first sheet in the spreadsheet
    emails = xl_sheet.col_values(1, 1)  # emails are in second column
    names = xl_sheet.col_values(0, 1)  # client names are in first column
    return emails, names


def format_message(to_email, to_name):
    """Formats the content of the email."""

    message = MIMEMultipart("alternative")
    message["From"] = my_email
    message["To"] = to_email

    # This message is formatted in HTML since it's the only way to embed links
    with open('message.html', 'r') as f:
        message["Subject"] = f.readline().split('Subject: ', 1)[-1]
        raw = f.read() % to_name
    content = MIMEText(raw, "html")
    message.attach(content)
    return message


if __name__ == '__main__':
    my_email = start()
    recipient_emails, recipient_names = get_info()
    for email in recipient_emails:
        name = recipient_names.pop(0)
        msg = format_message(email, name)
        try:
            server.sendmail(my_email, email, msg.as_string())
            print("Email sent to {} successfully".format(name))
        except Exception as exception:
            print("Email did not successfully send to {}".format(name))
            print("Error message: " + str(exception))
        finally:
            print("--------------------------")
    server.quit()
    input("Press ENTER to exit...")
