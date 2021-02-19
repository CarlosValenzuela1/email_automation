from constants import SALARY_SUBJECT, SALARY_SENDER, SALARY_ATTACHMENT

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
        delete = False
        salary = None

        if (sender == SALARY_SENDER) and (subject == SALARY_SUBJECT):
            delete = True
            print(f"The Sender and the Subject values have now been verified.")
            salary = self._identify_attachments(attachments)
            
        if (salary == None) and (delete == True):
            print(f"The attachments are corrupt, please verify the integrity of the attachments")
        elif (salary == None) and (delete == False):
            print(f"This email is not a Salary related email, continuing the process")
        return delete, salary
        
    def _identify_attachments(self, attachments: list):
        salary = None
        for attachment in attachments:
            print(f"type of attach: {type(attachment)}")
            print(f"Current attachment being analysed: {attachment.name}")
            if SALARY_ATTACHMENT in attachment.name:
                print(f"The correct attachment has been identified")
                salary = self._process_salary_attachments(attachment)
        
        return salary

    def _process_salary_attachments(self, attachment: FileAttachment):
        salary = None
        # converting attachment stream to excel compatible type
        io_excel = BytesIO(attachment.content)

        # converting io_excel to pandas.Dataframe with pandas.read_excel
        df = read_excel(io_excel)
        if not df.empty:
            salary = self._find_total_salary(df)
        
        return salary


    def _find_total_salary(self, df: DataFrame):
        total_salary = df["Salary"].max()
        print(f"Total Salary ({total_salary}) has been extracted from the attachment successfully")
        return total_salary
