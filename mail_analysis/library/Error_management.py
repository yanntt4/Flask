# -*- coding: utf-8 -*-
"""
Created on Tue Jun  4 10:55:10 2024

@author: ythiriet
"""

# Global librairies
import sys
from datetime import date
import matplotlib.pyplot as plot
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


# Error Management object
class Error_management():
    """Class used to deal with various errors
    
    Attributes
    ----------
    ERROR_TEXT : str
        Text to indicate the error met
    MAIL_CONNEXION_ERROR : bool
        Indicator for a connexion error with mailbox
    MAIL_ANALYSIS_ERROR : bool
        Indicator for a mail analysis error
    MAIL_SENDING_ERROR : bool
        Indicator for a API, Webhook or a mail sending error

    """

    def __init__(self):
        self.ERROR_TEXT = ""
        self.MAIL_CONNEXION_ERROR = False
        self.MAIL_ANALYSIS_ERROR = False
        self.MAIL_SENDING_ERROR = False
        

    def error_text_creation(self, MOMENT_ANALYSIS_FAIL, COMPANY_SENDER, ERROR_TYPE):
        
        self.ERROR_TEXT += f"\n {str(date.today())} // Error for {COMPANY_SENDER} {ERROR_TYPE}\n"
        
        if MOMENT_ANALYSIS_FAIL == 1:
            self.ERROR_TEXT += "GENERIC ANALYSIS FAILURE\n"
        elif MOMENT_ANALYSIS_FAIL == 2:
            self.ERROR_TEXT += "SPECIFIC ANALYSIS FAILURE\n"
        elif MOMENT_ANALYSIS_FAIL == 3:
            self.ERROR_TEXT += "SPECIFIC ANALYSIS FINDING ERROR\n"
        elif MOMENT_ANALYSIS_FAIL == 4:
            self.ERROR_TEXT += "SENDING_FAILURE\n"
    
    
    def sending_back_mail_1(self, IMAP_SSL, SMTP_SSL, SENDER_EMAIL, RECEIVER_EMAIL, EMAIL_UID):
        
        # Sending back problematic mail
        resp_code, email_data = IMAP_SSL.uid('fetch', EMAIL_UID, '(RFC822)')
        SMTP_SSL.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, email_data[0][1])
        
        # Pause before sending another mail
        plot.pause(3)


    def sending_back_mail_2(self, IMAP_SSL, SMTP_SSL, SENDER_EMAIL, RECEIVER_EMAIL, EMAIL_UID):
        
        # Sending error mail
        resp_code, email_data = IMAP_SSL.fetch(EMAIL_UID, '(RFC822)')
        SMTP_SSL.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, email_data[0][1])
        
        # Pause before sending another mail
        plot.pause(3)
        
    
    def sending_error_mail(self, IMAP_SSL, SMTP_SSL, SENDER_EMAIL, RECEIVER_EMAIL, SHIPMENT, LINE_NO, EXC_OBJ):
        
        # Sending error mail
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = RECEIVER_EMAIL
        msg['Subject'] = f"Erreur lors de l'analyse des mails Alliance et Speedmove // Shipment {SHIPMENT}"

        msg.attach(MIMEText(str(f"Une erreur s'est produite à la ligne {LINE_NO} : {EXC_OBJ}"), 'plain'))
        TEXT = msg.as_string()

        SMTP_SSL.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, TEXT)
        
        # Pause before sending another mail
        plot.pause(3)


    def error_dealing(self, IMAP_SSL, SMTP_SSL, SENDER_EMAIL, RECEIVER_EMAIL,
                      EMAIL_UID, COMPANY_SENDER, COMPANY_RECEIVER, MOMENT_ANALYSIS_FAIL, 
                      EXCEPTION = "", ERROR_TYPE = "", SHIPMENT = "Not Found", DEBUG_MODE = True):

        """
        Returning a mail (report mail) if an error occurs during connexion, analysis or sending

        Parameters
        ----------
        IMAP_SSL
            In-connexion, to the mailbox, to indicate the problematic mail
        SMTP_SSL
            Out-connexion, to a mailbox, to report the issue
        SENDER_EMAIL
            Mail address, indicating the sender of the report mail
        RECEIVER_EMAIL
            Mail address which will receive the report mail
        EMAIL_UID
            UID for the problematic mail
        COMPANY_SENDER
            Name of the company which sent the problematic mail
        COMPANY_RECEIVER
            Name of the company which receive the problematic mail
        EXCEPTION
            Exception which has resulted in the error
        ERROR_TYPE
            Indicator for the error type (CONNEXION, ANALYSIS, ...)
        SHIPMENT
            Shipment number if available

        """

        # Init
        LINE_NO = 0
        EXC_OBJ = ""

        # Flag to keep the error type
        if ERROR_TYPE == "CONNEXION":
            self.MAIL_ANALYSIS_ERROR = True
        elif ERROR_TYPE == "ANALYSIS":
            self.MAIL_ANALYSIS_ERROR = True
        elif ERROR_TYPE == "SENDING":
            self.MAIL_SENDING_ERROR = True

        # Exception details
        if isinstance(EXCEPTION, str) == False:
            EXC_OBJ, exc_obj, tb = sys.exc_info()
            LINE_NO = tb.tb_lineno
        
        # Creating error text
        self.error_text_creation(MOMENT_ANALYSIS_FAIL, COMPANY_SENDER, ERROR_TYPE)

        # Message to the user
        print(self.ERROR_TEXT)

        # Error report creation
        Error_report = open("Error_report.txt", "a")
        Error_report.write(self.ERROR_TEXT)
        
        # # Condition to send back mail
        # if DEBUG_MODE == False:
            
        # Sending error mail
        self.sending_error_mail(IMAP_SSL, SMTP_SSL, SENDER_EMAIL, RECEIVER_EMAIL, SHIPMENT, LINE_NO, EXC_OBJ)

        # Sending the mail to me for further analysis
        try:
            self.sending_back_mail_1(IMAP_SSL, SMTP_SSL, SENDER_EMAIL, RECEIVER_EMAIL, EMAIL_UID)
        except:
            try:
                self.sending_back_mail_2(IMAP_SSL, SMTP_SSL, SENDER_EMAIL, RECEIVER_EMAIL, EMAIL_UID)
    
            except Exception as e:
    
                # Message to the user
                print(e)
                print("Additionnal error while trying to send back the mail for further analysis")
    
                # Error report added
                Error_report.write("Additionnal error while trying to send back the mail for further analysis")