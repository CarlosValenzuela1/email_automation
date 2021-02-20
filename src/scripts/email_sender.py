from exchangelib import Account, Message, HTMLBody
import src.scripts.constants as c


class EmailSender(object):

    def __init__(self, total_salary=0):
        print(f"Creating Email Sender")
        self.total_salary = total_salary
    
    def _notify_service_operations(self, configs: dict, account: Account, attachments: list, salary: int):
        """
        Will send an email notification based on the salary, if its less than 15M or greater than 15M
        """

        # initiate Message object
        message = Message(account=account)

        # add all attachments from the original email
        message.attach(attachments)

        # LOW SALARY
        if self.total_salary <= int(c.THRESHOLD):
            # we customize the criteria to be used when salary is less than 15M. i.e, send email to diff recipients, have different subject or body on the notif, etc.
            print(f"Sending notification of Low Salary")
            recipients = configs["recipient_low_salary"].split(";")
            message.subject=c.LOW_SALARY_RESPONSE_SUBJECT
            message.to_recipients = recipients

            # for more details on how to format from 1000 -> $1,000.00 see this: https://docs.python.org/3/library/string.html#format-specification-mini-language
            body = c.LOW_BODY + f"${salary:,.2f}" + c.EMAIL_SIGNATURE
            print(f"Attempting to send email to: {recipients}")

        # HIGH SALARY
        elif self.total_salary > int(c.THRESHOLD):
            # we customize the criteria to be used when salary is greater than 15M. i.e, send email to diff recipients, have different subject or body on the notif, etc.
            print(f"Sending notification of High Salary")
            recipients = configs["recipient_high_salary"].split(";")
            message.subject=c.HIGH_SALARY_RESPONSE_SUBJECT
            # for more details on how to format from 1000 -> $1,000.00 see this: https://docs.python.org/3/library/string.html#format-specification-mini-language
            body = c.HIGH_BODY + f"${salary:,.2f}" + c.EMAIL_SIGNATURE
            message.to_recipients = recipients
            print(f"Attempting to send email to: {recipients}")

        message.body = HTMLBody(body)
        message.send()