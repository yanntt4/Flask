# -*- coding: utf-8 -*-
"""
Created on Mon Feb  5 09:26:56 2024

@author: ythiriet
"""


def main_pro_speedmove(global_parameters, general_mail_connexion):


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
    import requests
    import matplotlib.pyplot as plot
    import joblib
    import base64
    from fpdf import FPDF
    import json
    import xml.dom.minidom
    from random import randint
    from email.message import EmailMessage
    import pandas as pd
    import time
    from ftplib import FTP
    import io

    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart

    # Local librairies
    sys.path.append(f"{os.getcwd()}/Library")

    # Object for Mail Connexion
    from Mail_connexion import Mail_connexion
    # Object for Mail Analysis
    from Mail_analysis import Mail_analysis
    # Object for Error Report
    from Error_management import Error_management


    # Existing class modification to add encoding
    class XML(xml.dom.minidom.Document):

        "Polymorphism of class XML to modify version if necessary"

        def writexml(self, writer, indent="", addindent="", newl="", encoding=None,
                      standalone=None):
            declarations = []

            if encoding:
                declarations.append(f'encoding="{encoding}"')
            if standalone is not None:
                declarations.append(f'standalone="{"yes" if standalone else "no"}"')

            writer.write(f'<?xml version="1.0" {" ".join(declarations)}?>{newl}')

            for node in self.childNodes:
                node.writexml(writer, indent, addindent, newl)


    # Existing class modification to use PDF from Python
    from fpdf.output import OutputProducer
    from pathlib import Path
    class FPDF_Modif(FPDF):

        def output(
            self, name="", dest="", output_producer_class=OutputProducer
        ):
            """
            Output PDF to some destination.
            The method first calls [close](close.md) if necessary to terminate the document.
            After calling this method, content cannot be added to the document anymore.
    
            By default the bytearray buffer is returned.
            If a `name` is given, the PDF is written to a new file.
    
            Args:
                name (str): optional File object or file path where to save the PDF under
                dest (str): [**DEPRECATED since 2.3.0**] unused, will be removed in a later version
                output_producer_class (class): use a custom class for PDF file generation
            """

            # Finish document if necessary:
            if not self.buffer:
                if self.page == 0:
                    self.add_page()
                # Generating final page footer:
                self.in_footer = True
                self.footer()
                self.in_footer = False
                # Generating .buffer based on .pages:
                if self._toc_placeholder:
                    self._insert_table_of_contents()
                if self.str_alias_nb_pages:
                    self._substitute_page_number()
                output_producer = output_producer_class(self)
                self.buffer = output_producer.bufferize()
            if name:
                if isinstance(name, os.PathLike):
                    name.write_bytes(self.buffer)
                elif isinstance(name, str):
                    Path(name).write_bytes(self.buffer)
                else:
                    name.write(self.buffer)
                # return None
            return self.buffer


    #
    # Class creation


    # Object for Sending Information
    class Sending_object():
        def __init__(self, URL_LINE, URL_2_TOKEN, URL_2_DOC,
                     URL_2_DOC_FILLING, ACCESS_KEY, PRIMARY_KEY):

            self.CORE_MAIL = ""
            self.JSON = ""
            self.PDF = ""

            self.XML_BLOB = ""
            self.XML_BLOB_FILLING = ""
            self.XML_CARGOWISE = ""

            self.BLOB_ID = 0
            self.TOKEN = ""

            self.URL_LINE = URL_LINE
            self.URL_2_TOKEN = URL_2_TOKEN
            self.URL_2_DOC = URL_2_DOC
            self.URL_2_DOC_FILLING = URL_2_DOC_FILLING
            self.ACCESS_KEY = ACCESS_KEY
            self.PRIMARY_KEY = PRIMARY_KEY
            # Pu besoin de demander le token car il n'expire plus

            self.URL_CARGOWISE = "https://nxcprdservices.wisegrid.net/eAdaptor"

            self.RESPONSE_CW = ""
            self.RESPONSE_LG_WEBHOOK = ""
            self.RESPONSE_LG_DOC = ""
            self.RESPONSE_LG_DOC_FILLING = ""


        def pdf_creation(self, GENERAL_FOLDER_PATH, TITLE):

            # Removing problematic characters
            for i_chara in range(len(self.CORE_MAIL)):
                if ord(self.CORE_MAIL[i_chara]) > 60000:
                    self.CORE_MAIL = self.CORE_MAIL.replace(self.CORE_MAIL[i_chara]," ",1)

            #
            # Creating PDF file to attach to the mail
            Pdf_Object = FPDF_Modif()

            # Add a page
            Pdf_Object.add_page()

            # Set style and size
            Pdf_Object.add_font(
                "dejavu-sans-condensed", style="",
                fname=f"{GENERAL_FOLDER_PATH}/fonts/DejaVuSansCondensed.ttf")
            Pdf_Object.set_font('dejavu-sans-condensed', size = 8)

            # Splitting PDF according to line break
            CORE_MAIL_SPLIT_LINES = self.CORE_MAIL.splitlines()

            # Writing PDF following line break
            for i_split_lines in range(len(CORE_MAIL_SPLIT_LINES)):
                Pdf_Object.cell(100, 4, text = CORE_MAIL_SPLIT_LINES[i_split_lines],
                                ln = 2, border = 0, align = "", fill = False)

            # Saving PDF
            self.PDF = Pdf_Object.output(f"{TITLE}.pdf")


        def json_creation_speedmove(self, IMPORT, DATE_MONTH, SHIPMENT, PRICING,
                                    MONTH_GASOLE_TAX_2, GASOLE_TAXE_1, GASOLE_TAXE_2):

            # Difference case following Import/Export
            CHARGECODE = "OTPT"
            CHARGECODE_DIESEL = "OFUE"
            if IMPORT:
                CHARGECODE = "DTPT"
                CHARGECODE_DIESEL = "DFUE"

            #
            # Creating dictionnary

            # First case : Month date of and in the mail correspond
            if DATE_MONTH == MONTH_GASOLE_TAX_2:
                Dict = {"ShipmentNumber": SHIPMENT,
                     "Payables": [
                         {"ChargeType": CHARGECODE,
                          "Currency": "EUR",
                          "Vendor": "1004",
                          "UnitPrice": float(PRICING),
                          "Quantity": 1
                          },
                         {"ChargeType": CHARGECODE_DIESEL,
                          "Currency": "EUR",
                          "Vendor": "1004",
                          "UnitPrice": GASOLE_TAXE_2,
                          "Quantity": 1
                          }
                         ]}

            # Second case : Month date of and in the mail don't correspond
            else:
                Dict = {"ShipmentNumber": SHIPMENT,
                     "Payables": [
                         {"ChargeType": CHARGECODE,
                          "UOM": "FIXD",
                          "Currency": "EUR",
                          "Vendor": "1004",
                          "UnitPrice": float(PRICING),
                          "Quantity": 1
                          },
                         {"ChargeType": CHARGECODE_DIESEL,
                          "UOM": "FIXD",
                          "Currency": "EUR",
                          "Vendor": "1004",
                          "UnitPrice": GASOLE_TAXE_1,
                          "Quantity": 1
                          }
                         ]}

            # Converting dictionnary to JSON file
            self.JSON = json.dumps(Dict)
            
            
        # Funcion to create connexion with FTP
        def ftp_connexion(self, FTP_HOST, FTP_USER, FTP_PASS):
            self.ftp = FTP(host = FTP_HOST, user = FTP_USER, passwd = FTP_PASS)


        def cargowise_send_line(self, LOGIN_CARGOWISE, PASSWORD_CARGOWISE):

            # Sending XML using HTTP request
            self.RESPONSE_CW = requests.post(
                self.URL_CARGOWISE,
                auth =(LOGIN_CARGOWISE,PASSWORD_CARGOWISE),
                data = self.XML_CARGOWISE)

            # Text response
            print("\n Request result for CARGOWISE Sending")
            print(f"\n Text : \n {self.RESPONSE_CW.text}")
            print(f"\n Encoding : {self.RESPONSE_CW.encoding}")
            print(f"\n Status code : {self.RESPONSE_CW.status_code}")
            print(f"\n headers : {self.RESPONSE_CW.headers}")


        # Sending mail for attachment into CARGOWISE
        def cargowise_send_pdf(self, SENDER_EMAIL, RECEIVER_EMAIL, SMTP_SSL, TITLE, COMPANY_SENDER ,SHIPMENT):

            # Creating mail object
            msg = EmailMessage()

            # Add PDF attachment
            with open(f"{TITLE}.pdf", "rb") as content_file:
                content = content_file.read()
                msg.add_attachment(
                    content,
                    maintype='application',
                    subtype='pdf',
                    filename=f"{COMPANY_SENDER}_{SHIPMENT}.pdf")

            # Defining Subject, sender and receiver
            msg['Subject'] = f'[ediDocManager SHP ACH {SHIPMENT}]'
            msg['From'] = SENDER_EMAIL
            msg['To'] = RECEIVER_EMAIL

            # Sending mail to attach message into CARGOWISE
            SMTP_SSL.send_message(msg)

            # Changing receiver to send the mail to me for further analysis
            del msg['To']
            msg['To'] = "ythiriet@prolinair.com"

            # Sending mail for CONFIRMATION
            SMTP_SSL.send_message(msg)


        def xml_speedmove_cargowise(self, DEBUG_MODE, COMPANY_RECEIVER, IMPORT, SHIPMENT, PRICING,
                                    DATE_MONTH, GASOLE_TAXE_1, GASOLE_TAXE_2, MONTH_GASOLE_TAX_2):


            # Author : THIRIET Yann

            # date : 2023 - 06 -28

            # Details : This function is made to create the XML following
            # mail analysis for SpeedMove


            # Getting company code for XML
            if COMPANY_RECEIVER == "PROLINAIR":
                COMPANY_CODE = "PRO"
            elif COMPANY_RECEIVER == "CNL":
                COMPANY_CODE = "CNL"

            #
            # Creating XML
    
            # Creating  XML base
            doc = xml.dom.minidom.parseString("<Universalshipment/>")
            root = doc.documentElement
            root.setAttribute("xmlns", "http://www.cargowise.com/Schemas/Universal/2011/11")
            root.setAttribute("version", "1.1")
    
            # Balise ouvrante
            shipment = doc.createElement("shipment")
            root.appendChild(shipment)
    
            # Balise ouvrante
            DataContext = doc.createElement("DataContext")
            shipment.appendChild(DataContext)
    
            # Balise ouvrante
            DataTargetCollection = doc.createElement("DataTargetCollection")
            DataContext.appendChild(DataTargetCollection)
    
            # Balise ouvrante
            DataTarget = doc.createElement("DataTarget")
            DataTargetCollection.appendChild(DataTarget)
    
            # Element : Forwarding shipment
            Type_Key = doc.createElement("Type")
            Type_Key.appendChild(doc.createTextNode("Forwardingshipment"))
            DataTarget.appendChild(Type_Key)
    
            # Element : Numéro de shipment
            Key = doc.createElement("Key")
            Key.appendChild(doc.createTextNode(SHIPMENT))
            DataTarget.appendChild(Key)
    
            # Balise fermante
            # Balise fermante
    
            # Balise ouvrante
            Company = doc.createElement("Company")
            DataContext.appendChild(Company)
    
            # Element : Company Code
            Code_Company = doc.createElement("Code")
            Code_Company.appendChild(doc.createTextNode(COMPANY_CODE))
            Company.appendChild(Code_Company)
    
            # Balise fermante
    
            # Element : EntrepriseID
            EntrepriseID = doc.createElement("EnterpriseID")
            EntrepriseID.appendChild(doc.createTextNode("NXC"))
            DataContext.appendChild(EntrepriseID)
    
            # Element : ServerID
            ServerID = doc.createElement("ServerID")
            ServerID.appendChild(doc.createTextNode("PRD"))
            DataContext.appendChild(ServerID)
    
            # Balise fermante
    
            # Balise ouvrante
            JobCosting = doc.createElement("JobCosting")
            shipment.appendChild(JobCosting)
    
            # Balise ouvrante
            ChargeLineCollection = doc.createElement("ChargeLineCollection")
            JobCosting.appendChild(ChargeLineCollection)
    
            # Balise ouvrante
            ChargeLine = doc.createElement("ChargeLine")
            ChargeLineCollection.appendChild(ChargeLine)
    
            # Balise ouvrante
            ImportMetaData = doc.createElement("ImportMetaData")
            ChargeLine.appendChild(ImportMetaData)
    
            # Element : Instruction
            Instruction = doc.createElement("Instruction")
            Instruction.appendChild(doc.createTextNode("Insert"))
            ImportMetaData.appendChild(Instruction)
    
            # Balise fermante
    
            # Balise ouvrante
            CHARGECODE = doc.createElement("CHARGECODE")
            ChargeLine.appendChild(CHARGECODE)
    
            # Element : Charge Code
            Code = doc.createElement("Code")
            if IMPORT:
                Code.appendChild(doc.createTextNode("DTPT"))
            else:
                Code.appendChild(doc.createTextNode("OTPT"))
            CHARGECODE.appendChild(Code)
    
            # Balise fermante
    
            # Element : Cost OS Amount
            CostOSAmount = doc.createElement("CostOSAmount")
            CostOSAmount.appendChild(doc.createTextNode(PRICING))
            ChargeLine.appendChild(CostOSAmount)
    
            # Balise ouvrante
            CostOSCurrency = doc.createElement("CostOSCurrency")
            ChargeLine.appendChild(CostOSCurrency)
    
            # Element : Currency Code
            Currency_Code = doc.createElement("Code")
            Currency_Code.appendChild(doc.createTextNode("EUR"))
            CostOSCurrency.appendChild(Currency_Code)
    
            # Balise fermante
    
            # Balise ouvrante
            Creditor = doc.createElement("Creditor")
            ChargeLine.appendChild(Creditor)
    
            # Element : Type
            Type = doc.createElement("Type")
            Type.appendChild(doc.createTextNode("Organization"))
            Creditor.appendChild(Type)
    
            # Element : Organization Key
            Key_Orga = doc.createElement("Key")
            Key_Orga.appendChild(doc.createTextNode("SPEEDMOVE"))
            Creditor.appendChild(Key_Orga)
    
            # Balise fermante
            # Balise ouvrante
            ChargeLine = doc.createElement("ChargeLine")
            ChargeLineCollection.appendChild(ChargeLine)
    
            # Balise ouvrante
            ImportMetaData = doc.createElement("ImportMetaData")
            ChargeLine.appendChild(ImportMetaData)
    
            # Element : Instruction
            Instruction = doc.createElement("Instruction")
            Instruction.appendChild(doc.createTextNode("Insert"))
            ImportMetaData.appendChild(Instruction)
    
            # Balise fermante
    
            # Balise ouvrante
            CHARGECODE = doc.createElement("CHARGECODE")
            ChargeLine.appendChild(CHARGECODE)
    
            # Element : Charge Code 2
            Code = doc.createElement("Code")
            if IMPORT:
                Code.appendChild(doc.createTextNode("DFUE"))
            else:
                Code.appendChild(doc.createTextNode("OFUE"))
            CHARGECODE.appendChild(Code)
    
            # Balise fermante
    
            # Element : CostOSAmount
            CostOSAmount = doc.createElement("CostOSAmount")
            if DATE_MONTH == MONTH_GASOLE_TAX_2:
                CostOSAmount.appendChild(doc.createTextNode(GASOLE_TAXE_2))
            else:
                CostOSAmount.appendChild(doc.createTextNode(GASOLE_TAXE_1))
            ChargeLine.appendChild(CostOSAmount)
    
            # Balise ouvrante
            CostOSCurrency = doc.createElement("CostOSCurrency")
            ChargeLine.appendChild(CostOSCurrency)
    
            # Element : Currency Code
            Currency_Code = doc.createElement("Code")
            Currency_Code.appendChild(doc.createTextNode("EUR"))
            CostOSCurrency.appendChild(Currency_Code)
    
            # Balise fermante
    
            # Balise ouvrante
            Creditor = doc.createElement("Creditor")
            ChargeLine.appendChild(Creditor)
    
            # Element : Type
            Type = doc.createElement("Type")
            Type.appendChild(doc.createTextNode("Organization"))
            Creditor.appendChild(Type)
    
            # Element : Organization Key
            Key_Orga = doc.createElement("Key")
            Key_Orga.appendChild(doc.createTextNode("SPEEDMOVE"))
            Creditor.appendChild(Key_Orga)
    
            # Balise fermante
            # Balise fermante
            # Balise fermante
            # Balise fermante
            # Balise fermante
            # Balise fermante
    
            # Affichage du résultat du XML
            self.XML_CARGOWISE = root.toprettyxml()
    
            # Ecriture dans un fichier XML si en debug mode
            if DEBUG_MODE:
                doc.writexml(open(
                    f"{SHIPMENT}.xml", 'w'),
                    indent ="",
                    addindent ="    ",
                    newl='\n')
    
    
        def pdf_removal(self, TITLE):
    
            # Removing PDF file used
            os.remove(f"{TITLE}.pdf")
    
    
        def sent_confirmation_mail(
                self, SMTP_SSL, SENDER_EMAIL, RECEIVER_EMAIL, SHIPMENT, CC_EMAILS = None):
    
            msg = MIMEMultipart()
            msg['From'] = SENDER_EMAIL
            msg['To'] = RECEIVER_EMAIL
            msg['Subject'] = f"CONFIRMATION d'intégration du shipment {SHIPMENT}"
    
            # Si cc_emails est fourni, ajoutez-le à l'entête 'Cc'
            if CC_EMAILS:
                msg['Cc'] = ", ".join(CC_EMAILS)
    
            body = f"The shipment {SHIPMENT} has been correctly integrated into LOGITUDE"
            msg.attach(MIMEText(body, 'plain'))
    
            # Attacher chaque fichier à l'e-mail
            text = msg.as_string()
    
            # Ajoutez les e-mails CC à la liste des destinataires
            if CC_EMAILS:
                recipients = [RECEIVER_EMAIL]
                recipients.extend(CC_EMAILS)
    
                # Sending mail
                SMTP_SSL.sendmail(SENDER_EMAIL, recipients, text)
    
            else:
                SMTP_SSL.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, text)
    
    
        def Sent_Error_mail(
                self, SMTP_SSL, SENDER_EMAIL, RECEIVER_EMAIL, BODY,
                IMAP_SSL, Email_Uid):
    
            msg = MIMEMultipart()
            msg['From'] = SENDER_EMAIL
            msg['To'] = RECEIVER_EMAIL
            msg['Subject'] = f"Erreur lors de l'intégration du shipment {self.shipment}"
    
            msg.attach(MIMEText(BODY, 'plain'))
            text = msg.as_string()
    
            # Sending error mail
            SMTP_SSL.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, text)
    
            # Sending the mail to me for further analysis
            try:
                resp_code, email_data = IMAP_SSL.uid(
                    'fetch', Email_Uid, '(RFC822)')
    
                SMTP_SSL.sendmail(
                    SENDER_EMAIL, RECEIVER_EMAIL, email_data[0][1])
    
            except:
                try:
                    resp_code, email_data = IMAP_SSL.fetch(
                        Email_Uid, '(RFC822)')
    
                    SMTP_SSL.sendmail(
                        SENDER_EMAIL, RECEIVER_EMAIL, email_data[0][1])
    
                except Exception as e:
                    # Message to the user
                    print(e)
                    print("Additionnal error while trying to send back the mail for further analysis")


    # Keywords used for mail Analysis
    class Specific_analysis():
        def __init__(self):
            self.EXTRACT_MAIL = []
            self.DF_EXTRACT_MAIL = []
            self.INFO_HEADER = ""
            self.INFO_PRICING = 0
            self.INFO_CONTAINER = ""
            self.INFO_SHIPMENT = ""
            self.INFO_IMPORT = False
            self.INFO_EXPORT = False
            self.MAIL_PROCESSED = False
            self.INFO_DATE = ""
            self.INFO_MONTH_DATE = ""
    
        def alliance_pro_init(self):
            self.COMPANY_SENDER = "ALLIANCE"
            self.COMPANY_RECEIVER = "PROLINAIR"
            self.KEYWORD_TITLE = 'SUBJECT "PROLINAIR - Demande"'
            self.KEYWORD_MAIN_SEARCH = "confirmons"
            self.KEYWORD_MAIN_REMOVE = "<br>"
            self.KEYWORD_DATA = [[0,"livraison",0],[0,"enlèvement",0],
                                 [1,"Donneur d’Ordre : ",10], [1,"Réf Alliance Logistics : ",8],
                                 [1,"Objet : ",55], [2,["Tarif : ", "€ HT"],0]]
    
        def alliance_cnl_init(self):
            self.COMPANY_SENDER = "ALLIANCE"
            self.COMPANY_RECEIVER = "CNL"
            self.KEYWORD_TITLE = 'SUBJECT "CN LOGISTICS FRANCE - Demande"'
            self.KEYWORD_MAIN_SEARCH = "confirmons"
            self.KEYWORD_MAIN_REMOVE = "<br>"
            self.KEYWORD_DATA = [[0,"livraison",0],[0,"enlèvement",0],
                                 [1,"Donneur d’Ordre : ",10], [1,"Réf Alliance Logistics : ",8],
                                 [1,"Objet : ",55], [2,["Tarif : ", "€ HT"],0]]
    
        def speedmove_init(self):
            self.COMPANY_SENDER = "SPEEDMOVE"
            self.COMPANY_RECEIVER = "PROLINAIR"
            self.KEYWORD_TITLE = 'SUBJECT "PROLINAIR - Demande"'
            self.KEYWORD_MAIN_SEARCH = "accusons"
            self.KEYWORD_MAIN_REMOVE = "<br>"
            self.KEYWORD_DATA = [[0,"livraison",0],[0,"enlèvement",0],
                                 [1,"Votre référence dossier est : ",10], [1,"Objet : ",55],
                                 [2,["montant de ce transport : ", " €"],0], [2,["Envoyé : ", "\nÀ :"],0]]
    
        def event_init(self):
            self.KEYWORD_TITLE = 'SUBJECT "[Ci5 MRS] Tracing"'
            self.KEYWORD_MAIN_SEARCH = "automatic"
            self.KEYWORD_MAIN_REMOVE = "<br>"
            self.KEYWORD_SHIPMENT = "reference : "
            self.KEYWORD_CONTAINER = "cargo "
            self.KEYWORD_DATE = "at "
    
        def alliance_info_extract(self):
            self.INFO_SHIPMENT = self.DF_EXTRACT_MAIL["Shipment"][0]
            self.INFO_IMPORT = self.DF_EXTRACT_MAIL["Import"][0]
            self.INFO_EXPORT = self.DF_EXTRACT_MAIL["Export"][0]
            self.INFO_PRICING = self.DF_EXTRACT_MAIL["Pricing"][0]
    
        def speedmove_info_extract(self):
            self.INFO_SHIPMENT = self.DF_EXTRACT_MAIL["Shipment"][0]
            self.INFO_IMPORT = self.DF_EXTRACT_MAIL["Import"][0]
            self.INFO_EXPORT = self.DF_EXTRACT_MAIL["Export"][0]
            self.INFO_PRICING = self.DF_EXTRACT_MAIL["Pricing"][0]
            self.INFO_DATE = self.DF_EXTRACT_MAIL["Date"][0]
            self.INFO_MONTH_DATE = self.DF_EXTRACT_MAIL["Month date"][0]
            self.INFO_GASOLE_TAXE_1 = self.DF_EXTRACT_MAIL["Gasole taxe 1"][0]
            self.INFO_GASOLE_TAXE_2 = self.DF_EXTRACT_MAIL["Gasole taxe 2"][0]
            self.INFO_MONTH_GASOLE_TAX_1 = self.DF_EXTRACT_MAIL["Month gasole taxe 1"][0]
            self.INFO_MONTH_GASOLE_TAX_2 = self.DF_EXTRACT_MAIL["Month gasole taxe 2"][0]


    #
    # Keyword definition

    # Keyword definition for Speedmove analysis
    speedmove_analysis = Specific_analysis()
    speedmove_analysis.speedmove_init()

    # Error Management Init
    Error = Error_management()

    # Separating mail following company (PRO/CNL and All/Spe)
    general_mail_connexion.list_uid = general_mail_connexion.mail_Search(speedmove_analysis.KEYWORD_TITLE)
    list_mail = np.zeros([len(general_mail_connexion.list_uid[0].split())], dtype = Mail_analysis)
    speedmove_analysis.EXTRACT_MAIL = []


    # mail Analysis for PROLINAIR Speedmove
    for i, mail in enumerate(list_mail):

        # Init
        speedmove_analysis.EXTRACT_MAIL = []
    
        #
        # Generic analysis
        try:
    
            mail = Mail_analysis()
            mail.uid = general_mail_connexion.list_uid[0].split()[i]
    
            # Decoding mail with utf-8 (bytes to string)
            mail.mail_decoding(general_mail_connexion.IMAP_SSL)
    
            # Subject
            mail.subject_search(global_parameters.SUBJECTS, global_parameters.SUBJECT_END)
    
            # Date
            mail.date_search(global_parameters.DATE, global_parameters.DATE_END)
            if len(mail.DATE) > 1:
                mail.month_identification()
            else:
                Error.error_mail_return(
                    general_mail_connexion.IMAP_SSL, general_mail_connexion.SMTP_SSL,
                    global_parameters.SENDER_EMAIL, global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                    mail.uid, speedmove_analysis.COMPANY_SENDER, speedmove_analysis.COMPANY_RECEIVER, 
                    ERROR_TYPE = "DATE")
                continue
    
            # Core mail
            for DECODE in global_parameters.DECODE:
                mail.core_mail_search(DECODE, global_parameters.DECODE_END)
                if mail.CORE_MAIL_SEARCH_ANALYSIS:
                    break
    
            if mail.CORE_MAIL_SEARCH_ANALYSIS == False:
                continue
    
            # Core mail decoding if necessary
            if isinstance(mail.CORE_MAIL,str) == False:
                mail.CORE_MAIL = mail.CORE_MAIL.decode("utf-8", errors = "ignore")
    
    
        # Exception
        except Exception as e:
            Error.error_mail_return(
                general_mail_connexion.IMAP_SSL, general_mail_connexion.SMTP_SSL,
                global_parameters.SENDER_EMAIL, global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                mail.uid, speedmove_analysis.COMPANY_SENDER, speedmove_analysis.COMPANY_RECEIVER,
                EXCEPTION = e, ERROR_TYPE = "ANALYSIS")
    
        #
        # Specific analysis
        try:
    
            for j in range(len(speedmove_analysis.KEYWORD_DATA)):
                if speedmove_analysis.KEYWORD_DATA[j][0] == 0:
                    speedmove_analysis.EXTRACT_MAIL.append(mail.subject_keyword_search(
                        speedmove_analysis.KEYWORD_DATA[j][1]))
                elif speedmove_analysis.KEYWORD_DATA[j][0] == 1:
                    speedmove_analysis.EXTRACT_MAIL.append(mail.inside_mail_search(
                        speedmove_analysis.KEYWORD_DATA[j][1], speedmove_analysis.KEYWORD_DATA[j][2]))
                else:
                    speedmove_analysis.EXTRACT_MAIL.append(mail.inside_mail_search_Between(
                        speedmove_analysis.KEYWORD_DATA[j][1][0], speedmove_analysis.KEYWORD_DATA[j][1][1]))
    
            # Searching if previous shipment have been analysed already
            PREVIOUS_SHIPMENT = global_parameters.list_shipment_compare(speedmove_analysis.EXTRACT_MAIL[2])
    
        # Exception
        except Exception as e:
            Error.error_mail_return(
                general_mail_connexion.IMAP_SSL, general_mail_connexion.SMTP_SSL,
                global_parameters.SENDER_EMAIL, global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                mail.uid, speedmove_analysis.COMPANY_SENDER, speedmove_analysis.COMPANY_RECEIVER,
                EXCEPTION = e, ERROR_TYPE = "ANALYSIS", shipment = speedmove_analysis.EXTRACT_MAIL[2])
    
    
        # Checking that information have been found
        if (len(speedmove_analysis.EXTRACT_MAIL[2]) > 1 and len(speedmove_analysis.EXTRACT_MAIL[4]) > 0 and PREVIOUS_SHIPMENT == False):
    
            # Adding date found in mail
            speedmove_analysis.EXTRACT_MAIL.append(mail.DATE_MONTH)
    
            # Searching for Gazole Tax
            gasole_taxe_1, gasole_taxe_2, month_gasole_taxe_1, month_gasole_taxe_2 = mail.taxe_gazole_search()
            speedmove_analysis.EXTRACT_MAIL.append(gasole_taxe_1)
            speedmove_analysis.EXTRACT_MAIL.append(gasole_taxe_2)
            speedmove_analysis.EXTRACT_MAIL.append(month_gasole_taxe_1)
            speedmove_analysis.EXTRACT_MAIL.append(month_gasole_taxe_2)
    
            # Changing , into . to avoid wrong amounts
            speedmove_analysis.EXTRACT_MAIL[4] = speedmove_analysis.EXTRACT_MAIL[4].replace(",",".")
    
            # Creating dataframe
            speedmove_analysis.DF_EXTRACT_MAIL = pd.DataFrame(
                np.transpose(np.expand_dims(np.array(speedmove_analysis.EXTRACT_MAIL),axis = 1)),
                columns = ["Export", "Import", "Shipment", "Header", "Pricing", "Date", "Month date",
                           "Gasole taxe 1", "Gasole taxe 2", "Month gasole taxe 1", "Month gasole taxe 2"])

            # Extracting information from dataframe
            speedmove_analysis.speedmove_info_extract()

            # Dealing with Exception with mail Sending
            try:

                # Object Creation to send information to LOGITUDE/CARGOWISE
                speedmove_sending = Sending_object(
                    global_parameters.URL_LINE, global_parameters.URL_2_TOKEN,
                    global_parameters.URL_2_DOC, global_parameters.URL_2_DOC_FILLING,
                    global_parameters.ACCESS_KEY, global_parameters.PRIMARY_KEY)
                speedmove_sending.CORE_MAIL = mail.CORE_MAIL
    
                # Creating PDF
                speedmove_sending.pdf_creation(
                    global_parameters.GENERAL_FOLDER_PATH,
                    f"{speedmove_analysis.COMPANY_RECEIVER}_{speedmove_analysis.COMPANY_SENDER}_{speedmove_analysis.INFO_SHIPMENT}")
    
                # Sending Payable lines into LOGITUDE
                if (global_parameters.LG_PRO_MODE and speedmove_analysis.INFO_SHIPMENT[3:4].isnumeric()):
    
                    # JSON creation
                    speedmove_sending.json_creation_speedmove(
                        speedmove_analysis.INFO_IMPORT, speedmove_analysis.INFO_MONTH_DATE,
                        speedmove_analysis.INFO_SHIPMENT, speedmove_analysis.INFO_PRICING,
                        speedmove_analysis.INFO_MONTH_GASOLE_TAX_2, speedmove_analysis.INFO_GASOLE_TAXE_1,
                        speedmove_analysis.INFO_GASOLE_TAXE_2)
    
                    # Sending into FTP to send into LOGITUDE
                    if global_parameters.DEBUG_MODE == False:
                        
                        # FTP connexion
                        speedmove_sending.ftp_connexion(global_parameters.FTP_HOST,
                                                       global_parameters.FTP_USER,
                                                       global_parameters.FTP_PASS)
                        
                        # Saving file into FTP folder
                        bio = io.BytesIO(f"{speedmove_sending.JSON}".encode())
                        speedmove_sending.ftp.storbinary(f'STOR {speedmove_analysis.INFO_SHIPMENT}_PROLINAIR_SPEEDMOVE.json',
                                                        bio)
                        
                        # Saving PDF into FTP folder
                        with open(f"{speedmove_analysis.COMPANY_RECEIVER}_{speedmove_analysis.COMPANY_SENDER}_{speedmove_analysis.INFO_SHIPMENT}.pdf", "rb") as f:
                            speedmove_sending.ftp.storbinary(f'STOR {speedmove_analysis.INFO_SHIPMENT}_PROLINAIR_SPEEDMOVE.pdf', f)
                        
                        # Removing PDF
                        speedmove_sending.pdf_removal(
                            f"{speedmove_analysis.COMPANY_RECEIVER}_{speedmove_analysis.COMPANY_SENDER}_{speedmove_analysis.INFO_SHIPMENT}")
    
                        # Completing shipment List
                        global_parameters.list_shipment_completion(speedmove_analysis.INFO_SHIPMENT)
    
                        # Sending CONFIRMATION mail
                        speedmove_sending.sent_confirmation_mail(
                            general_mail_connexion.SMTP_SSL,
                            global_parameters.SENDER_EMAIL,
                            global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                            speedmove_analysis.INFO_SHIPMENT)
                    
                    # Saving JSON for checking
                    else:
                        
                        # Saving JSON file inside local folder
                        with open(f"{speedmove_analysis.INFO_SHIPMENT}_PROLINAIR_SPEEDMOVE.json", "w") as f:
                            f.write(speedmove_sending.JSON)
    
    
                # Sending Payable lines into CARGOWISE
                else:
    
                    # XML creation
                    speedmove_sending.xml_speedmove_cargowise(
                        global_parameters.DEBUG_MODE, speedmove_analysis.COMPANY_RECEIVER,
                        speedmove_analysis.INFO_IMPORT, speedmove_analysis.INFO_SHIPMENT,
                        speedmove_analysis.INFO_PRICING, speedmove_analysis.INFO_MONTH_DATE,
                        speedmove_analysis.INFO_GASOLE_TAXE_1, speedmove_analysis.INFO_GASOLE_TAXE_2,
                        speedmove_analysis.INFO_MONTH_GASOLE_TAX_2)
    
                    if global_parameters.DEBUG_MODE == False:
                        # Sending line into CW
                        speedmove_sending.cargowise_send_line(
                            global_parameters.LOGIN_CARGOWISE, global_parameters.PASSWORD_CARGOWISE)
    
                        # If sending line failed, send back the mail
                        if "Error" in speedmove_sending.response_cw.text:
                            speedmove_sending.Sent_Error_mail(
                                    general_mail_connexion.SMTP_SSL,
                                    global_parameters.SENDER_EMAIL,
                                    global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                                    speedmove_sending.response_cw.text,
                                    general_mail_connexion.IMAP_SSL,
                                    mail.uid)
    
                        # Sending the PDF
                        else:
                            speedmove_sending.cargowise_send_pdf(
                                    global_parameters.SENDER_EMAIL,
                                    global_parameters.RECEIVER_EMAIL,
                                    general_mail_connexion.SMTP_SSL,
                                    f"{speedmove_analysis.COMPANY_RECEIVER}_{speedmove_analysis.COMPANY_SENDER}_{speedmove_analysis.INFO_SHIPMENT}",
                                    speedmove_analysis.COMPANY_SENDER,
                                    speedmove_analysis.INFO_SHIPMENT)
    
                            # Completing shipment list
                            global_parameters.list_shipment_completion(
                                speedmove_analysis.INFO_SHIPMENT)
    
                        # Removing PDF
                        speedmove_sending.pdf_removal(
                            f"{speedmove_analysis.COMPANY_RECEIVER}_{speedmove_analysis.COMPANY_SENDER}_{speedmove_analysis.INFO_SHIPMENT}")

            # Exception
            except Exception as e:
                Error.error_mail_return(
                    general_mail_connexion.IMAP_SSL, general_mail_connexion.SMTP_SSL,
                    global_parameters.SENDER_EMAIL, global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                    mail.uid, speedmove_analysis.COMPANY_SENDER, speedmove_analysis.COMPANY_RECEIVER,
                    EXCEPTION = e, ERROR_TYPE = "SENDING", SHIPMENT = speedmove_analysis.INFO_SHIPMENT)

        list_mail[i] = mail


    # End of Running program
    global_parameters.list_shipment_saving()

    # Exit
    return list_mail