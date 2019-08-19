"""
usage: generate_email.py [-h] -t TO -s SUBJECT -b BODY
                         [-a ATTACHMENTS [ATTACHMENTS ...]]

Generates an email and opens in Outlook

optional arguments:
  -h, --help            show this help message and exit
  -t TO, --to TO        Recipient email address
  -s SUBJECT, --subject SUBJECT
                        Subject line for email
  -b BODY, --body BODY  Email body
  -a ATTACHMENTS [ATTACHMENTS ...], --attachments ATTACHMENTS [ATTACHMENTS ...]
                        File paths for attachments
"""
import argparse
import win32com.client as win32

def process_arguments():
    """
    Uses argparse module to define and handle command line input arguments and help menu
    """
    # Create ArgumentParser object. Description message will be displayed as part of help message if script is run with -h flag
    parser = argparse.ArgumentParser(description='Generates an email and opens in Outlook')
    # Define the arguments that will be taken.
    parser.add_argument('-t', '--to', required=True, type=str, help='Recipient email address')
    parser.add_argument('-s', '--subject', required=True, type=str, help='Subject line for email')
    parser.add_argument('-b', '--body', required=True, type=str, help='Email body')
    parser.add_argument('-a', '--attachments', required=False, type=str, nargs='+', help='File paths for attachments')
    # Return the arguments
    return parser.parse_args()

def generate_email(to_address, subject, body, attachments):
    '''
    Populates an Outlook email and opens in separate window
    '''
    # Create Outlook message object
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    # Set email attributes
    mail.To = to_address
    mail.Subject = subject
    mail.HtmlBody = body
    # Attach files
    if attachments:
        [mail.Attachments.Add(Source=attachment) for attachment in attachments]
    # Open the email in outlook. False argument prevents the Outlook window from blocking the script 
    mail.Display(False)

def main():
    args = process_arguments()
    # Populate an outlook email addressed to clinican with results attached 
    generate_email(args.to, args.subject, args.body, args.attachments)

if __name__ == '__main__':
    main()
