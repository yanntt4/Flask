# -*- coding: utf-8 -*-
"""
Created on Mon Feb  5 09:26:56 2024

@author: ythiriet
"""


def main_analysis_sending(global_parameters, general_mail_connexion, sender_analysis):


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
    from Mail_analysis import Mail_analysis, generic_mail_analysis, specific_mail_analysis, specific_mail_analysis_speedmove
    # Object for Sending result
    from Sending_object import Sending_object
    # Object for Error Report
    from Error_management import Error_management


    # Init
    if sender_analysis.COMPANY_RECEIVER == "PROLINAIR":
        LG_MODE = global_parameters.LG_PRO_MODE
    elif sender_analysis.COMPANY_RECEIVER == "CNL":
        LG_MODE = global_parameters.LG_CNL_MODE

    # Error Management Init
    Error = Error_management()

    # Separating mail following company (PRO/CNL and Alliance/Speed Move)
    general_mail_connexion.list_uid = general_mail_connexion.mail_Search(sender_analysis.KEYWORD_TITLE)
    list_mail = np.zeros([len(general_mail_connexion.list_uid[0].split())], dtype = Mail_analysis)
    sender_analysis.EXTRACT_MAIL = []

    # Analysing mail found following title
    for i, mail in enumerate(list_mail):
    
        # Init
        sender_analysis.EXTRACT_MAIL = []
        
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
                mail.uid, sender_analysis.COMPANY_SENDER, sender_analysis.COMPANY_RECEIVER,
                EXCEPTION = e, ERROR_TYPE = "ANALYSIS", shipment = sender_analysis.EXTRACT_MAIL[2])
    

        # Specific analysis
        try:
            if sender_analysis.COMPANY_SENDER == "ALLIANCE":
                sender_analysis.EXTRACT_MAIL = specific_mail_analysis(
                    mail, sender_analysis.EXTRACT_MAIL, sender_analysis.KEYWORD_DATA)
            elif sender_analysis.COMPANY_SENDER == "SPEEDMOVE":
                sender_analysis.EXTRACT_MAIL = specific_mail_analysis_speedmove(
                    mail, sender_analysis.EXTRACT_MAIL, sender_analysis.KEYWORD_DATA)
    
        # Exception
        except Exception as e:
            Error.error_mail_return(
                general_mail_connexion.IMAP_SSL, general_mail_connexion.SMTP_SSL,
                global_parameters.SENDER_EMAIL, global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                mail.uid, sender_analysis.COMPANY_SENDER, sender_analysis.COMPANY_RECEIVER,
                EXCEPTION = e, ERROR_TYPE = "ANALYSIS", shipment = sender_analysis.EXTRACT_MAIL[2])
        
        
        # Searching if previous shipment have been analysed already
        PREVIOUS_SHIPMENT = global_parameters.list_shipment_compare(sender_analysis.EXTRACT_MAIL[2])
    
        # Checking that information have been found
        if (len(sender_analysis.EXTRACT_MAIL[2]) > 1 and
            len(sender_analysis.EXTRACT_MAIL[5]) > 0 and
            PREVIOUS_SHIPMENT == False):
    
            # Changing , into . to avoid wrong amounts
            sender_analysis.EXTRACT_MAIL[5] = sender_analysis.EXTRACT_MAIL[5].replace(",",".")
    
            # Creating dataframe
            if sender_analysis.COMPANY_SENDER == "ALLIANCE":
                sender_analysis.DF_EXTRACT_MAIL = pd.DataFrame(
                    np.transpose(np.expand_dims(np.array(sender_analysis.EXTRACT_MAIL),axis = 1)),
                    columns = ["EXPORT", "IMPORT", "SHIPMENT", "CONTAINER", "HEADER", "PRICING"])
            
            elif sender_analysis.COMPANY_SENDER == "SPEEDMOVE":
                sender_analysis.DF_EXTRACT_MAIL = pd.DataFrame(
                    np.transpose(np.expand_dims(np.array(sender_analysis.EXTRACT_MAIL),axis = 1)),
                    columns = ["EXPORT", "IMPORT", "SHIPMENT", "CONTAINER", "HEADER", "PRICING", "DATE", "MONTH_DATE",
                               "GASOLE_TAXE_1", "GASOLE_TAXE_2", "MONTH_GASOLE_TAXE_1", "MONTH_GASOLE_TAXE_2"])

    
            # Dealing with Exception with mail Sending
            try:
    
                # Object Creation to send information to LOGITUDE/CARGOWISE
                global_sending = Sending_object()
                global_sending.CORE_MAIL = mail.CORE_MAIL
    
                # Creating PDF
                global_sending.pdf_creation(
                    global_parameters.GENERAL_FOLDER_PATH,
                    f"{sender_analysis.DF_EXTRACT_MAIL['SHIPMENT'][0]}_{sender_analysis.COMPANY_RECEIVER}_{sender_analysis.COMPANY_SENDER}")
                
                # Sending Payable lines into LOGITUDE
                if (LG_MODE and sender_analysis.DF_EXTRACT_MAIL["SHIPMENT"][0][3:4].isnumeric()):
    
                    # JSON creation
                    if sender_analysis.COMPANY_SENDER == "ALLIANCE":
                        global_sending.json_creation_alliance(sender_analysis.DF_EXTRACT_MAIL)
                    elif sender_analysis.COMPANY_SENDER == "SPEEDMOVE":
                        global_sending.json_creation_speedmove(sender_analysis.DF_EXTRACT_MAIL)
    
                    # Sending into FTP to send into LOGITUDE
                    if global_parameters.DEBUG_MODE == False:
                        
                        # FTP connexion
                        global_sending.ftp_connexion(
                            global_parameters.FTP_HOST, global_parameters.FTP_USER, global_parameters.FTP_PASS)
                        
                        # Saving file into FTP folder
                        bio = io.BytesIO(f"{global_sending.JSON}".encode())
                        global_sending.ftp.storbinary(f'STOR {sender_analysis.DF_EXTRACT_MAIL["SHIPMENT"][0]}_{sender_analysis.COMPANY_RECEIVER}_{sender_analysis.COMPANY_SENDER}.json',
                                                        bio)
                        
                        # Saving PDF into FTP folder
                        with open(f"{sender_analysis.DF_EXTRACT_MAIL['SHIPMENT'][0]}_{sender_analysis.COMPANY_RECEIVER}_{sender_analysis.COMPANY_SENDER}.pdf", "rb") as f:
                            global_sending.ftp.storbinary(f'STOR {sender_analysis.DF_EXTRACT_MAIL["SHIPMENT"][0]}_{sender_analysis.COMPANY_RECEIVER}_{sender_analysis.COMPANY_SENDER}.pdf', f)
                        
                        # Removing PDF
                        global_sending.pdf_removal(
                            f"{sender_analysis.DF_EXTRACT_MAIL['SHIPMENT'][0]}_{sender_analysis.COMPANY_RECEIVER}_{sender_analysis.COMPANY_SENDER}")
                        
                        # Completing shipment List
                        global_parameters.list_shipment_completion(sender_analysis.DF_EXTRACT_MAIL["SHIPMENT"][0])
    
                        # Sending CONFIRMATION mail
                        global_sending.sent_confirmation_mail(
                            general_mail_connexion.SMTP_SSL,
                            global_parameters.SENDER_EMAIL,
                            global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                            sender_analysis.DF_EXTRACT_MAIL["SHIPMENT"][0])
                    
                    # Saving JSON for checking
                    else:
                        
                        # Saving JSON file inside local folder
                        with open(f"{sender_analysis.DF_EXTRACT_MAIL['SHIPMENT'][0]}_{sender_analysis.COMPANY_RECEIVER}_{sender_analysis.COMPANY_SENDER}.json", 'w') as f:
                            f.write(global_sending.JSON)

    
                # Sending Payable lines into CARGOWISE
                else:
    
                    # XML creation
                    if sender_analysis.COMPANY_SENDER == "ALLIANCE":
                        global_sending.xml_alliance_cargowise(global_parameters.DEBUG_MODE, sender_analysis.DF_EXTRACT_MAIL,
                                                                sender_analysis.COMPANY_RECEIVER)
                    elif sender_analysis.COMPANY_SENDER == "SPEEDMOVE":
                        global_sending.xml_speedmove_cargowise(global_parameters.DEBUG_MODE, sender_analysis.DF_EXTRACT_MAIL,
                                                                sender_analysis.COMPANY_RECEIVER)
                        
                    if global_parameters.DEBUG_MODE == False:
                        # Sending line into CW
                        global_sending.cargowise_send_line(
                            global_parameters.LOGIN_CARGOWISE, global_parameters.PASSWORD_CARGOWISE)
    
                        # If sending line failed, send back the mail
                        if "Error" in global_sending.response_cw.text:
                            global_sending.Sent_Error_mail(
                                    general_mail_connexion.SMTP_SSL,
                                    global_parameters.SENDER_EMAIL,
                                    global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                                    global_sending.response_cw.text,
                                    general_mail_connexion.IMAP_SSL,
                                    mail.uid)
    
                        # Sending the PDF
                        else:
                            global_sending.cargowise_send_pdf(
                                    global_parameters.SENDER_EMAIL,
                                    global_parameters.RECEIVER_EMAIL,
                                    general_mail_connexion.SMTP_SSL,
                                    f"{sender_analysis.COMPANY_RECEIVER}_{sender_analysis.COMPANY_SENDER}_{sender_analysis.DF_EXTRACT_MAIL['SHIPMENT'][0]}",
                                    sender_analysis.COMPANY_SENDER,
                                    sender_analysis.DF_EXTRACT_MAIL["SHIPMENT"][0])
    
                            # Completing shipment list
                            global_parameters.list_shipment_completion(sender_analysis.DF_EXTRACT_MAIL["SHIPMENT"][0])
    
                        # Removing PDF
                        global_sending.pdf_removal(
                            f"{sender_analysis.COMPANY_RECEIVER}_{sender_analysis.COMPANY_SENDER}_{sender_analysis.DF_EXTRACT_MAIL['SHIPMENT'][0]}")
            
            # Exception
            except Exception as e:
                Error.error_mail_return(
                    general_mail_connexion.IMAP_SSL, general_mail_connexion.SMTP_SSL,
                    global_parameters.SENDER_EMAIL, global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                    mail.uid, sender_analysis.COMPANY_SENDER, sender_analysis.COMPANY_RECEIVER,
                    EXCEPTION = e, ERROR_TYPE = "SENDING", SHIPMENT = sender_analysis.DF_EXTRACT_MAIL["SHIPMENT"][0])

        # Saving mail analysed
        list_mail[i] = mail

    # End of Running program
    global_parameters.list_shipment_saving()

    # Exit
    return list_mail


