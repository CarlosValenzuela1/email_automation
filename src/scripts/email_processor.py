from src.scripts.constants import SALARY_SUBJECT, SALARY_SENDER, SALARY_ATTACHMENT

# to convert from exchangelib.FileAttachment -> Bytesio -> excel -> panda
from io import BytesIO
from exchangelib import FileAttachment
from pandas import DataFrame, read_excel


class EmailProcessor(object):
    """
    Responsible for all business logic of the application
    i.e., filter a specific range in an email attachment, analyze email body, isolate url links, etc
    """

    def __init__(self, configs=None):
        print(f"Creating Email Processor")

        # contains the list of recipients for low and high salary notifications
        self.configs = configs

    def _get_total_salary(self, subject:str, sender:str, attachments:list):
        """
        Main business logic to process the salary email. 
        Process will verify email attributes, extract appropriate attachment, convert attachment to usable obj, extract total salary
        """
        salary = None
        delete = False

        # we verify the sender and subject of the email are appropriate, otherwise we skip the email
        if (sender == SALARY_SENDER) and (subject == SALARY_SUBJECT):
            # since sender and subject are valid, then we will delete the email from inbox after its processing
            delete = True
            print(f"The Sender and the Subject values have now been verified.")

            # extracting only salary related attachment, all other attachments will be ignored
            attachment = self._get_valid_attachment(attachments)

            # converting attachment stream to panda.dataframe
            df = self._convert_attachment_to_df(attachment)

            # extracting salary from dataframe
            salary = self._find_total_salary(df)

        # scenario happens when excel structure has changed. i,e., doesn't contain correct column, data has a non-supported format, etc
        if (salary == None) and (delete == True):
            print(f"The attachments are corrupt, please verify the integrity of the attachments")
        
        # scenario happens when email does not comply with ANY of our rules which define a Salary email
        # Rule 1: Sender must be carlos.valenzuela.cs@outlook.com
        # Rule 2: Subject must contain 'Salary Report from Oct 2077' 
        # Rule 3: Email must contain attachments and at least 1 excel attachment called 'Salaries from Oct 2077'
        elif (salary == None) and (delete == False):
            print(f"This email is not a Salary related email, continuing the process")

        return delete, salary
        
    def _get_valid_attachment(self, attachments: list):
        """ iterate through the attachments and return the attachment called 'Salaries from Oct 2077' """
        valid_attachment = None

        # exchangelib handles attachments as a simple list of FileAttachment files
        for attachment in attachments:
            print(f"Current attachment being analysed: {attachment.name}")

            if SALARY_ATTACHMENT in attachment.name:
                print(f"The correct attachment has been identified")
                valid_attachment = attachment

        return valid_attachment

    def _convert_attachment_to_df(self, attachment: FileAttachment):
        dataframe = None

        if attachment != None:
            # converting attachment stream to excel compatible type
            io_excel = BytesIO(attachment.content)

            # converting io_excel to pandas.Dataframe with pandas.read_excel
            dataframe = read_excel(io_excel)
        
        return dataframe

    def _find_total_salary(self, df: DataFrame):
        """ extract the total salary field from the dataframe and return -1 if not found"""
        total_salary = -1

        if not df.empty:
            # total salary is a sum of all the other salaries in the excel, so it will always be the biggest number
            total_salary = df["Salary"].max()
            print(f"Total Salary ({total_salary}) has been extracted from the attachment successfully")
        
        return total_salary