# to get our current directory where this script is located
import os

# to store/retrieve password from OS Credential Manager
import keyring

# to access all our necessary constants
from src.scripts.constants import MY_EMAIL

# to process emails individually, this is where the logic to process emails is
from src.scripts.email_processor import EmailProcessor

# to send email notifications
from src.scripts.email_sender import EmailSender

# to connect and handle the emails from the ews
from exchangelib import Credentials, Account, Message, DELEGATE

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

    # checking if inbox has any emails
    if(account.inbox.all().count() != 0):

        # iterate through each item on the inbox
        for item in account.inbox.all().order_by('-datetime_received'):
            print_email_details(item)

            if item.categories != ['Red Category']:
                try:
                    item.is_read = True
                    item.save()

                    # check if email is related to Salary, if yes then process email else continue with the next processing logic
                    if(is_salary_processing(email_processor, item)):
                        return

                except Exception as e:
                    mail_processing_failed(item, e)
                    continue
            
            else:
                print(f"Mail was not processed because its category is {item.categories}")
    else:
        print(f"There are currently no emails to process")

def setup_ews_client():
    # using Keyring to obtain password from Microsoft Credential Manager for security reasons
    credentials = Credentials(MY_EMAIL, keyring.get_password("demo", MY_EMAIL))

    try:
        # establish connection to EWS and returning account to be able to parse through the mailbox, contacts, todo, etc
        account = Account('carlos.valenzuela.cs@outlook.com', credentials=credentials, autodiscover=True, access_type=DELEGATE)
    except Exception as e:
        print(f"Error when attempting to connect to Exchange Web Server: {e}")

    return account

def print_email_details(item: Message):
    """Standard method to print relevant values of each email"""
    print(f"")
    print(f"Processing the following email: ")
    print(f"Sender: {item.sender.email_address}")
    print(f"Subject: {item.subject}")
    print(f"Date: {item.datetime_received}")
    print(f"")

def is_salary_processing(email_processor: EmailProcessor, item: Message):
    """
    Method responsible to call email processor with the appropriate  parameters
    It will also return False, if the email processed was not related to Salary, so other functions can continue processing the same email
    """
    delete_email, total_salary = email_processor._get_total_salary(item.subject, item.sender.email_address, item.attachments)

    if delete_email:
        email_sender = EmailSender(total_salary)
        email_sender._notify_service_operations(email_processor.configs, item.account, item.attachments, total_salary)
        
        return True
    else:
        return False

def mail_processing_failed(item: Message, e: TypeError):
    item.is_read = True
    item.categories = ['Red Category']
    item.save()

    print(f"Type of e: {type(e)}")
    print(f"AN ERROR HAS OCURRED WHILE PROCESSING THE FOLLOWING EMAIL:")
    print_email_details(item)
    print(f"Please see below exception: \n{e}")