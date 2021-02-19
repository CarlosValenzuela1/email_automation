# to get our current directory where this script is located
import os

# To store/retrieve password from OS Credential Manager
import keyring

# to access all our necessary constants
from constants import MY_EMAIL

# to process emails individually, this is where the logic to process emails is
from email_processor import EmailProcessor

# to send email notifications
from email_sender import EmailSender

# to connect and handle the emails from the ews
from exchangelib import Credentials, Account, Message

# to decrypt the password needed to login to inbox
from cryptography.fernet import Fernet

# to identify our current directory where we are located
current_dir = os.path.abspath(os.path.dirname(__file__))

def main():
    """
    STARTING PLACE OF OUR APPLICATION
    """
    # to initiate the credentials and create the account to parse through the EWS Emails
    account = setup_ews_client()

    # initiate configs, this will contain the recipients for the emails to be sent
    configs = {"recipient_low_salary": MY_EMAIL, "recipient_high_salary": MY_EMAIL}

    # instanciating EmailProcessor to handle the logic applied to each email
    email_processor = EmailProcessor(configs)

    # pylint: disable=maybe-no-member
    if(account.inbox.all().count() == 0):
        print(f"There are currently no emails to process")

    # iterate through each item on the inbox
    # pylint: disable=maybe-no-member
    for item in account.inbox.all().order_by('-datetime_received'):
        #try:
        item.is_read = True
        item.save()

        print_email_details(item)

        # check if email is related to Salary, if yes then process email else continue with the next processing logic
        if(is_salary_processing(email_processor, item)):
            return

        #TODO
        # check if email is related to job descriptions
        if(is_job_profiler()):
            return

        # except Exception as e:
        #     mail_processing_failed()
        #     continue

def setup_ews_client():
    # using Keyring to obtain password from Microsoft Credential Manager for security reasons
    credentials = Credentials(MY_EMAIL, keyring.get_password("demo", MY_EMAIL))

    try:
        # establish connection to EWS and returning account to be able to parse through the mailbox, contacts, todo, etc
        account = Account('carlos.valenzuela.cs@outlook.com', credentials=credentials, autodiscover=True)
    except Exception as e:
        print(f"Error when attempting to connect to Exchange Web Server: {e}")
    return account

def print_email_details(item: Message):
    """Standard method to print relevant values of each email"""
    print(f"Sender: {item.sender.email_address}")
    print(f"Subject: {item.subject}")
    print(f"Date: {item.datetime_received}")

def is_salary_processing(email_processor: EmailProcessor, item: Message):
    """
    Method responsible to call email processor with the appropriate  parameters
    It will also return False, if the email processed was not related to Salary, so other functions can continue processing the same email
    """
    delete_email, total_salary = email_processor._get_total_salary(item.subject, item.sender.email_address, item.attachments)

    # shorcut for if delete_email == True: OR if delete_email != None
    if delete_email:
        email_sender = EmailSender(total_salary)
        email_sender._notify_service_operations(email_processor.configs, item.account, item.attachments)

        item.move_to_trash()
        return True
    else:
        return False

# TODO
def is_job_profiler():
    print(f"Hello World")

# TODO
def mail_processing_failed():
    print(f"Error")

if __name__ == "__main__":
    main()
