# -*- coding: utf-8 -*-
"""
Created on Mon Jun  3 14:59:43 2024

@author: ythiriet
"""


# Global librairies
import re
import base64
import numpy as np
import email


# Class creation
class Mail_analysis:
    def __init__(self):
        self.uid = 0
        self.mail = ""
        self.mail_decode_bytes = ""
        self.mail_decode = ""
        self.subject = ""
        self.CORE_MAIL = ""
        self.CORE_MAIL_SEARCH_ANALYSIS = True
        self.JSON = ""
        self.Previous_shipment = False
        self.DATE = ""
        self.DATE_MONTH = ""


    def mail_decoding(self, IMAP_SSL):
        
        try:
            resp_code, Email_All_Data = IMAP_SSL.uid(None,"All")
            self.mail = Email_All_Data[0][1]
            self.mail_decode_bytes = email.message_from_bytes(self.mail)
            self.mail_decode = email.message_from_bytes(self.mail).as_string()
        except:
            try:
                resp_code, Email_All_Data = IMAP_SSL.uid('fetch', self.uid, '(RFC822)')
                self.mail = Email_All_Data[0][1]
                self.mail_decode_bytes = email.message_from_bytes(self.mail)
                self.mail_decode = email.message_from_bytes(self.mail).as_string()
            except:
                try:
                    resp_code, Email_All_Data = IMAP_SSL.fetch(self.uid, '(RFC822)')
                    self.mail = Email_All_Data[0][1]
                    self.mail_decode_bytes = email.message_from_bytes(self.mail)
                    self.mail_decode = email.message_from_bytes(self.mail).as_string()
                except:

                    # Print an error report in the console
                    print("Error checking mails inside specific mailbox")

                    # Error report creation
                    Error_Report = open("Error_Report.txt", "a")
                    Error_Report.write("Error while analysing mail content")


    def subject_search(self, keywords, keyword_end):
        
        # self.subject = str(email.header.make_header(email.header.decode_header(self.mail['Subject'])))
        # print(self.subject)
    
        for keyword in keywords:
            potential_SUBJECTS = re.split(f'{keyword}', self.mail_decode)
            if len(potential_SUBJECTS) < 2:
                continue

        if len(potential_SUBJECTS) > 1:
            self.subject = re.split(f'{keyword_end}', potential_SUBJECTS[1])[0]

            if self.subject[:8] =="=?utf-8?":
                self.subject = re.split("=", self.subject)[1]
                self.subject = base64.b64decode(
                    self.subject[9:-1]).decode("utf-8", errors = "ignore")


    def date_search(self, keyword_begin, keyword_end):
        potential_dates = re.split(keyword_begin, self.mail_decode)
        if len(potential_dates) > 2:
            for potential_date in potential_dates:
                list_date = re.split(keyword_end, potential_date)
                if len(list_date) > 1:
                    if len(list_date[1]) > 10:
                        self.DATE = list_date[1]
                        break


    def month_identification(self):
        self.DATE_MONTH = self.DATE.split(" ")[3]
        MONTH_CONVERSION = np.array([["Jan", "JANVIER"],["Feb", "FEVRIER"],["Mar", "MARS"],
                            ["Apr", "AVRIL"],["May", "MAI"],["Jun", "JUIN"],
                            ["Jul", "JUILLET"],["Aug", "AOUT"],["Sep", "SEPTEMBRE"],
                            ["Oct", "OCTOBRE"],["Nov", "NOVEMBRE"],["Dec", "DECEMBRE"]],
                                    dtype = object)
        for j in range(MONTH_CONVERSION.shape[0]):
            if self.DATE_MONTH == MONTH_CONVERSION[j,0]:
                self.DATE_MONTH = MONTH_CONVERSION[j,1]
                break


    def core_mail_search(self):

        if self.mail_decode_bytes.is_multipart():
            for part in self.mail_decode_bytes.walk():
                ctype = part.get_content_type()
                cdispo = str(part.get('Content-Disposition'))
        
                # skip any text/plain (txt) attachments
                if ctype == 'text/plain' and 'attachment' not in cdispo:
                    self.CORE_MAIL = part.get_payload(decode=True)  # decode
                    break
        # not multipart - i.e. plain text, no attachments, keeping fingers crossed
        else:
            self.CORE_MAIL = self.mail_decode_bytes.get_payload(decode=True)


    def subject_keyword_search(self, keyword):

        # Init
        result = True

        match = re.search(keyword, self.subject)
        if match is None:
            result = False

        return result


    def inside_mail_search(self, keyword, length_data):

        split_research = re.split(keyword, self.CORE_MAIL)
        if len(split_research) > 1:
            data_found = split_research[1][:length_data]
        else:
            data_found = ""

        # Exit if nothing found
        return data_found


    def inside_mail_search_Between(self, keyword_begin, keyword_end):

        split_research = re.split(f'{keyword_begin}|{keyword_end}', self.CORE_MAIL)
        if len(split_research) > 1:
            data_found = split_research[1]
        else:
            data_found = ""

        # Exit
        return data_found


    def taxe_gazole_search(self):

        # Init
        gasole_taxe_1 = -1
        gasole_taxe_2 = -1
        month_gasole_taxe_1 = ""
        month_gasole_taxe_2 = ""

        # Searching for Both Taxes
        both_taxes = re.split("€", self.CORE_MAIL)
        all_taxes = re.split("%", both_taxes[1])

        # Getting First tax and month
        taxe_1 = re.split("%", both_taxes[1])[0]

        # Getting Second tax and month if found
        if len(all_taxes) > 2:
            taxe_2 = re.split("%", both_taxes[1])[1]


        months = ["JANVIER", "FEVRIER", "MARS", "AVRIL", "MAI", "JUIN",
                 "JUILLET", "AOUT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DECEMBRE"]

        # Splitting First month and tax
        for month in months:
            if re.search(month, taxe_1) is not None:
                month_gasole_taxe_1 = month
                gasole_taxe_1 = float(re.split(month, taxe_1)[-1][1:])
                break

        # Splitting Second month and tax
        if len(all_taxes) > 2:
            for month in months:
                if re.search(month, taxe_2) is not None:
                    month_gasole_taxe_2 = month
                    gasole_taxe_2 = float(re.split(month, taxe_2)[-1][1:])
                    break

        # Exit
        return gasole_taxe_1, gasole_taxe_2, month_gasole_taxe_1, month_gasole_taxe_2


# Getting generic information from mail (CORE, SUBJECT)
def generic_mail_analysis(mail, MAIL_UID, IMAP_SSL, SUBJECTS_KEYWORD, SUBJECT_END_KEYWORD):
    
    # UID from mailbox
    # mail.uid = general_mail_connexion.list_uid[0].split()[i]
    mail.uid = MAIL_UID

    # Decoding mail with utf-8 (bytes to string)
    mail.mail_decoding(IMAP_SSL)

    # Subject
    mail.subject_search(SUBJECTS_KEYWORD, SUBJECT_END_KEYWORD)

    # Core mail
    mail.core_mail_search()

    # Core mail decoding if necessary
    if isinstance(mail.CORE_MAIL,str) == False:
        mail.CORE_MAIL = mail.CORE_MAIL.decode("utf-8", errors = "ignore")
    
    # Exit
    return mail


# Getting specific information from mail
def specific_mail_analysis(mail, EXTRACT_MAIL, KEYWORD_DATA):
    
    for j in range(len(KEYWORD_DATA)):
        if KEYWORD_DATA[j][0] == 0:
            EXTRACT_MAIL.append(mail.subject_keyword_search(
                KEYWORD_DATA[j][1]))
        elif KEYWORD_DATA[j][0] == 1:
            EXTRACT_MAIL.append(mail.inside_mail_search(
                KEYWORD_DATA[j][1], KEYWORD_DATA[j][2]))
        else:
            EXTRACT_MAIL.append(mail.inside_mail_search_Between(
                KEYWORD_DATA[j][1][0], KEYWORD_DATA[j][1][1]))
    
    # Exit
    return EXTRACT_MAIL

# Getting specific information from mail
def specific_mail_analysis_speedmove(mail, EXTRACT_MAIL, KEYWORD_DATA):
    
    for j in range(len(KEYWORD_DATA)):
        if KEYWORD_DATA[j][0] == 0:
            EXTRACT_MAIL.append(mail.subject_keyword_search(
                KEYWORD_DATA[j][1]))
        elif KEYWORD_DATA[j][0] == 1:
            EXTRACT_MAIL.append(mail.inside_mail_search(
                KEYWORD_DATA[j][1], KEYWORD_DATA[j][2]))
        else:
            EXTRACT_MAIL.append(mail.inside_mail_search_Between(
                KEYWORD_DATA[j][1][0], KEYWORD_DATA[j][1][1]))
    
    #
    # Searching for Gazole Tax

    # Adding month found in mail
    EXTRACT_MAIL.append(mail.DATE_MONTH)

    # Searching for Gazole Tax
    gasole_taxe_1, gasole_taxe_2, month_gasole_taxe_1, month_gasole_taxe_2 = mail.taxe_gazole_search()
    EXTRACT_MAIL.append(gasole_taxe_1)
    EXTRACT_MAIL.append(gasole_taxe_2)
    EXTRACT_MAIL.append(month_gasole_taxe_1)
    EXTRACT_MAIL.append(month_gasole_taxe_2)
    
    # Exit
    return EXTRACT_MAIL