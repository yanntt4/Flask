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
    MAIL_CONNEXION_ERROR : bool
        Indicator for a connexion error with mailbox
    MAIL_ANALYSIS_ERROR : bool
        Indicator for a mail analysis error
    MAIL_SENDING_ERROR : bool
        Indicator for a API, Webhook or a mail sending error

    """

    def __init__(self):
        self.MAIL_CONNEXION_ERROR = False
        self.MAIL_ANALYSIS_ERROR = False
        self.MAIL_SENDING_ERROR = False


    def error_mail_return(self, IMAP_SSL, SMTP_SSL, SENDER_EMAIL, RECEIVER_EMAIL,
                          EMAIL_UID, COMPANY_SENDER, COMPANY_RECEIVER, EXCEPTION = "",
                          ERROR_TYPE = "", SHIPMENT = "Not Found"):

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
        lineno = 0
        exc_obj = ""

        # Flag to keep the error type
        if ERROR_TYPE == "CONNEXION":
            self.MAIL_ANALYSIS_ERROR = True
        elif ERROR_TYPE == "ANALYSIS":
            self.MAIL_ANALYSIS_ERROR = True
        elif ERROR_TYPE == "SENDING":
            self.MAIL_SENDING_ERROR = True

        # Exception details
        if isinstance(EXCEPTION, str) == False:
            exc_type, exc_obj, tb = sys.exc_info()
            lineno = tb.tb_lineno

        # Message to the user
        print(f"Error with {COMPANY_RECEIVER} {COMPANY_SENDER} {ERROR_TYPE}")

        # Error report creation
        Error_report = open("Error_report.txt", "a")
        Error_report.write(f"\n {str(date.today())} // Error for {COMPANY_SENDER} {ERROR_TYPE}")

        # Sending the mail to me for further analysis
        try:
            resp_code, email_data = IMAP_SSL.uid('fetch', EMAIL_UID, '(RFC822)')
            SMTP_SSL.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, email_data[0][1])

            plot.pause(5)

            msg = MIMEMultipart()
            msg['From'] = SENDER_EMAIL
            msg['To'] = RECEIVER_EMAIL
            msg['Subject'] = f"Erreur lors de l'analyse des mails Alliance et Speedmove // {ERROR_TYPE} // Shipment {SHIPMENT}"

            msg.attach(MIMEText(str(f"Une erreur s'est produite à la ligne {lineno} : {exc_obj}"), 'plain'))
            TEXT = msg.as_string()

            SMTP_SSL.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, TEXT)

        except:
            try:
                resp_code, email_data = IMAP_SSL.fetch(EMAIL_UID, '(RFC822)')
                SMTP_SSL.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, email_data[0][1])

                plot.pause(5)

                msg = MIMEMultipart()
                msg['From'] = SENDER_EMAIL
                msg['To'] = RECEIVER_EMAIL
                msg['Subject'] = f"Erreur lors de l'analyse des mails Alliance et Speedmove : Shipment {SHIPMENT}"

                msg.attach(MIMEText(str(f"Une erreur s'est produite à la ligne {lineno} : {exc_obj}"), 'plain'))
                text = msg.as_string()

                # SMTP_SSL.set_debuglevel(1)
                SMTP_SSL.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, text)

            except Exception as e:

                # Message to the user
                print(e)
                print("Additionnal error while trying to send back the mail for further analysis")

                # Error report added
                Error_report.write("Additionnal error while trying to send back the mail for further analysis")