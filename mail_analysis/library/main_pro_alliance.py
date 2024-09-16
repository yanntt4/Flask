# -*- coding: utf-8 -*-
"""
Created on Mon Feb  5 09:26:56 2024

@author: ythiriet
"""


def main_pro_alliance(global_parameters, general_mail_connexion):


    # Author : THIRIET Yann

    # date : 2023 - 06 - 28

    # Details : This function is made to search and analyse
    # specific automatic mails


    #
    # Global init
    
    # Global librairies
    import os
    import sys
    import numpy as np
    import pandas as pd
    import io

    
    # Local librairies
    sys.path.append(f"{os.getcwd()}/Library")

    # Object for mail Analysis
    from Mail_analysis import Mail_analysis, generic_mail_analysis, specific_mail_analysis
    # Object for Sending result
    from Sending_object import Sending_object
    # Object for Error Report
    from Error_management import Error_management


    # Keywords used for mail Analysis
    class Specific_analysis():
        def __init__(self):
            self.EXTRACT_MAIL = []
            self.DF_EXTRACT_MAIL = []
            self.MAIL_PROCESSED = False
    
        def alliance_pro_init(self):
            self.COMPANY_SENDER = "ALLIANCE"
            self.COMPANY_RECEIVER = "PROLINAIR"
            self.KEYWORD_TITLE = 'SUBJECT "PROLINAIR - Demande"'
            self.KEYWORD_DATA = [[0,"livraison",0],[0,"enlèvement",0],
                                 [1,"Donneur d’Ordre : ",10], [1,"Réf Alliance Logistics : ",8],
                                 [1,"Objet : ",55], [2,["Tarif : ", "€ HT"],0]]

    #
    # Keyword definition


    # Keyword definition for Alliance Logistics Analysis
    alliance_analysis = Specific_analysis()
    alliance_analysis.alliance_pro_init()

    # Error Management Init
    Error = Error_management()

    # Separating mail following company (PRO/CNL and Alliance/Speed Move)
    general_mail_connexion.list_uid = general_mail_connexion.mail_Search(alliance_analysis.KEYWORD_TITLE)
    list_mail = np.zeros([len(general_mail_connexion.list_uid[0].split())], dtype = Mail_analysis)
    alliance_analysis.EXTRACT_MAIL = []

    # mail Analysis for PROLINAIR Alliance
    for i, mail in enumerate(list_mail):
    
        # Init
        alliance_analysis.EXTRACT_MAIL = []
        
        # Getting mail
        mail = Mail_analysis()
        MAIL_UID = general_mail_connexion.list_uid[0].split()[i]
    
        # Generic analysis
        try:
            mail = generic_mail_analysis(
                mail, MAIL_UID, general_mail_connexion.IMAP_SSL,
                global_parameters.SUBJECTS, global_parameters.SUBJECT_END)
        
        # Exception
        except Exception as e:
            Error.error_mail_return(
                general_mail_connexion.IMAP_SSL, general_mail_connexion.SMTP_SSL,
                global_parameters.SENDER_EMAIL, global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                mail.uid, alliance_analysis.COMPANY_SENDER, alliance_analysis.COMPANY_RECEIVER,
                EXCEPTION = e, ERROR_TYPE = "ANALYSIS", shipment = alliance_analysis.EXTRACT_MAIL[2])
    

        # Specific analysis
        try:
            alliance_analysis.EXTRACT_MAIL = specific_mail_analysis(
                mail, alliance_analysis.EXTRACT_MAIL, alliance_analysis.KEYWORD_DATA)
    
        # Exception
        except Exception as e:
            Error.error_mail_return(
                general_mail_connexion.IMAP_SSL, general_mail_connexion.SMTP_SSL,
                global_parameters.SENDER_EMAIL, global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                mail.uid, alliance_analysis.COMPANY_SENDER, alliance_analysis.COMPANY_RECEIVER,
                EXCEPTION = e, ERROR_TYPE = "ANALYSIS", shipment = alliance_analysis.EXTRACT_MAIL[2])
        
        
        # Searching if previous shipment have been analysed already
        PREVIOUS_SHIPMENT = global_parameters.list_shipment_compare(alliance_analysis.EXTRACT_MAIL[2])
    
        # Checking that information have been found
        if (len(alliance_analysis.EXTRACT_MAIL[2]) > 1 and
            len(alliance_analysis.EXTRACT_MAIL[5]) > 0 and
            PREVIOUS_SHIPMENT == False):
    
            # Changing , into . to avoid wrong amounts
            alliance_analysis.EXTRACT_MAIL[5] = alliance_analysis.EXTRACT_MAIL[5].replace(",",".")
    
            # Creating dataframe
            alliance_analysis.DF_EXTRACT_MAIL = pd.DataFrame(
                np.transpose(np.expand_dims(np.array(alliance_analysis.EXTRACT_MAIL),axis = 1)),
                columns = ["EXPORT", "IMPORT", "SHIPMENT", "CONTAINER", "HEADER", "PRICING"])

    
            # Dealing with Exception with mail Sending
            try:
    
                # Object Creation to send information to LOGITUDE/CARGOWISE
                alliance_sending = Sending_object()
                alliance_sending.CORE_MAIL = mail.CORE_MAIL
    
                # Creating PDF
                alliance_sending.pdf_creation(
                    global_parameters.GENERAL_FOLDER_PATH,
                    f"{alliance_analysis.COMPANY_RECEIVER}_{alliance_analysis.COMPANY_SENDER}_{alliance_analysis.DF_EXTRACT_MAIL['SHIPMENT'][0]}")
                
                # Sending Payable lines into LOGITUDE
                if (global_parameters.LG_PRO_MODE and alliance_analysis.DF_EXTRACT_MAIL["SHIPMENT"][0][3:4].isnumeric()):
    
                    # JSON creation
                    alliance_sending.json_creation_alliance(alliance_analysis.DF_EXTRACT_MAIL)
    
                    # Sending into FTP to send into LOGITUDE
                    if global_parameters.DEBUG_MODE == False:
                        
                        # FTP connexion
                        alliance_sending.ftp_connexion(
                            global_parameters.FTP_HOST, global_parameters.FTP_USER, global_parameters.FTP_PASS)
                        
                        # Saving file into FTP folder
                        bio = io.BytesIO(f"{alliance_sending.JSON}".encode())
                        alliance_sending.ftp.storbinary(f'STOR {alliance_analysis.DF_EXTRACT_MAIL["SHIPMENT"][0]}_PROLINAIR_ALLIANCE.json',
                                                        bio)
                        
                        # Saving PDF into FTP folder
                        with open(f"{alliance_analysis.COMPANY_RECEIVER}_{alliance_analysis.COMPANY_SENDER}_{alliance_analysis.DF_EXTRACT_MAIL['SHIPMENT'][0]}.pdf", "rb") as f:
                            alliance_sending.ftp.storbinary(f'STOR {alliance_analysis.DF_EXTRACT_MAIL["SHIPMENT"][0]}_PROLINAIR_ALLIANCE.pdf', f)
                        
                        # Removing PDF
                        alliance_sending.pdf_removal(
                            f"{alliance_analysis.COMPANY_RECEIVER}_{alliance_analysis.COMPANY_SENDER}_{alliance_analysis.DF_EXTRACT_MAIL['SHIPMENT'][0]}")
                        
                        # Completing shipment List
                        global_parameters.list_shipment_completion(alliance_analysis.DF_EXTRACT_MAIL["SHIPMENT"][0])
    
                        # Sending CONFIRMATION mail
                        alliance_sending.sent_confirmation_mail(
                            general_mail_connexion.SMTP_SSL,
                            global_parameters.SENDER_EMAIL,
                            global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                            alliance_analysis.DF_EXTRACT_MAIL["SHIPMENT"][0])
                    
                    # Saving JSON for checking
                    else:
                        
                        # Saving JSON file inside local folder
                        with open(f"{alliance_analysis.DF_EXTRACT_MAIL['SHIPMENT'][0]}_PROLINAIR_ALLIANCE.json", 'w') as f:
                            f.write(alliance_sending.JSON)

    
                # Sending Payable lines into CARGOWISE
                else:
    
                    # XML creation
                    alliance_sending.xml_alliance_cargowise(global_parameters.DEBUG_MODE, alliance_analysis.DF_EXTRACT_MAIL,
                                                            alliance_analysis.COMPANY_RECEIVER)
                    if global_parameters.DEBUG_MODE == False:
                        # Sending line into CW
                        alliance_sending.cargowise_send_line(
                            global_parameters.LOGIN_CARGOWISE, global_parameters.PASSWORD_CARGOWISE)
    
                        # If sending line failed, send back the mail
                        if "Error" in alliance_sending.response_cw.text:
                            alliance_sending.Sent_Error_mail(
                                    general_mail_connexion.SMTP_SSL,
                                    global_parameters.SENDER_EMAIL,
                                    global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                                    alliance_sending.response_cw.text,
                                    general_mail_connexion.IMAP_SSL,
                                    mail.uid)
    
                        # Sending the PDF
                        else:
                            alliance_sending.cargowise_send_pdf(
                                    global_parameters.SENDER_EMAIL,
                                    global_parameters.RECEIVER_EMAIL,
                                    general_mail_connexion.SMTP_SSL,
                                    f"{alliance_analysis.COMPANY_RECEIVER}_{alliance_analysis.COMPANY_SENDER}_{alliance_analysis.DF_EXTRACT_MAIL['SHIPMENT'][0]}",
                                    alliance_analysis.COMPANY_SENDER,
                                    alliance_analysis.DF_EXTRACT_MAIL["SHIPMENT"][0])
    
                            # Completing shipment list
                            global_parameters.list_shipment_completion(alliance_analysis.DF_EXTRACT_MAIL["SHIPMENT"][0])
    
                        # Removing PDF
                        alliance_sending.pdf_removal(
                            f"{alliance_analysis.COMPANY_RECEIVER}_{alliance_analysis.COMPANY_SENDER}_{alliance_analysis.DF_EXTRACT_MAIL['SHIPMENT'][0]}")
            
            # Exception
            except Exception as e:
                Error.error_mail_return(
                    general_mail_connexion.IMAP_SSL, general_mail_connexion.SMTP_SSL,
                    global_parameters.SENDER_EMAIL, global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                    mail.uid, alliance_analysis.COMPANY_SENDER, alliance_analysis.COMPANY_RECEIVER,
                    EXCEPTION = e, ERROR_TYPE = "SENDING", SHIPMENT = alliance_analysis.DF_EXTRACT_MAIL["SHIPMENT"][0])

        # Saving mail analysed
        list_mail[i] = mail

    # End of Running program
    global_parameters.list_shipment_saving()

    # Exit
    return list_mail


