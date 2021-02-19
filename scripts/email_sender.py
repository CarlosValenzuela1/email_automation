from exchangelib import Account, Message
from constants import LOW_SALARY_RESPONSE_SUBJECT, HIGH_SALARY_RESPONSE_SUBJECT

class EmailSender(object):

    def __init__(self, total_salary=0):
        print(f"Creating Email Sender")
        self.total_salary = total_salary
    
    def _notify_service_operations(self, configs: dict, account: Account, attachments: list):
        """
        Will send an email notification based on the salary, if its less than 15M or greater than 15M
        """

        # initiate Message object
        message = Message(account=account)

        # add all attachments from the original email
        message.attach(attachments)

        if self.total_salary <= 15000000:
            # we customize the criteria to be used when salary is less than 15M. i.e, send email to diff recipients, have different subject or body on the notif, etc.
            print(f"Sending notification of Low Salary")
            recipients = configs["recipient_low_salary"].split(";")
            message.subject=LOW_SALARY_RESPONSE_SUBJECT
            message.to_recipients = recipients
            body = "This is a Low Salary Notification"
            print(f"Attempting to send email to: {recipients}")

        elif self.total_salary > 15000000:
            # we customize the criteria to be used when salary is greater than 15M. i.e, send email to diff recipients, have different subject or body on the notif, etc.
            print(f"Sending notification of High Salary")
            recipients = configs["recipient_high_salary"].split(";")
            message.subject=HIGH_SALARY_RESPONSE_SUBJECT
            body = "This is a High Salary Notification"
            message.to_recipients = recipients
            print(f"Attempting to send email to: {recipients}")

        message.body = body
        message.send()


