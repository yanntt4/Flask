# -*- coding: utf-8 -*-
"""
Created on Mon Feb  5 09:26:56 2024

@author: ythiriet
"""


def main_cnl_alliance(global_parameters, general_mail_connexion):


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
    from datetime import date
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
    
    # Object for mail Connexion
    from Mail_connexion import Mail_connexion
    # Object for mail Analysis
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
        def __init__(self, URL_LINE, URL_2_TOKEN, URL_2_Doc,
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
            self.URL_2_DOC = URL_2_Doc
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
    
    
        def json_creation_alliance(self, IMPORT, SHIPMENT, PRICING):
    
            # Difference case following Import/Export
            CHARGECODE = "OTPT"
            if IMPORT:
                CHARGECODE = "DTPT"
    
            # Creating dictionnary
            Dict = {"ShipmentNumber": SHIPMENT,
                  "Payables": [
                      {"ChargeType": CHARGECODE,
                      "Currency": "EUR",
                      "Vendor": "1003",
                      "UnitPrice": float(PRICING),
                      "Quantity": 1
                      }]}
    
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
    
    
        def xml_alliance_cargowise(self, DEBUG_MODE, COMPANY_RECEIVER, IMPORT, SHIPMENT, PRICING):
    
            # Author : THIRIET Yann
    
            # date : 2024 - 03 -07
    
            # Details : This function is made to create the XML following
            # mail analysis for Alliance
    
            # Getting company code for XML
            if COMPANY_RECEIVER == "PROLINAIR":
                COMPANY_CODE = "PRO"
            elif COMPANY_RECEIVER == "CNL":
                COMPANY_CODE = "CNL"
    
    
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
            Key_Orga.appendChild(doc.createTextNode("ALLIANCE"))
            Creditor.appendChild(Key_Orga)
    
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
    
            # # Attacher chaque fichier à l'e-mail
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
            self.INFO_date = ""
            self.INFO_DATE_MONTH = ""
    
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
    
        def speedMove_Init(self):
            self.KEYWORD_TITLE = 'SUBJECT "PROLINAIR - Demande"'
            self.KEYWORD_MAIN_SEARCH = "accusons"
            self.KEYWORD_MAIN_REMOVE = "<br>"
            self.KEYWORD_IMPORT = "livraison"
            self.KEYWORD_EXPORT = "enlèvement"
            self.KEYWORD_SHIPMENT = "Votre référence dossier est : "
            self.KEYWORD_PRICING_TOP = "montant de ce transport : "
            self.KEYWORD_PRICING_DOWN = " €"
            self.KEYWORD_HEADER = "Objet : "
            self.KEYWORD_date_TOP = "Envoyé : "
            self.KEYWORD_date_DOWN = "\nÀ :"
            self.KEYWORD_GAZOLE_TAX_TOP = np.array([
                " JANVIER "," FEVRIER "," MARS ", " AVRIL "," MAI "," JUIN ",
                " JUILLET "," AOUT "," SEPTEMBRE ", " OCTOBRE "," NOVEMBRE "," DECEMBRE "])
            self.KEYWORD_GAZOLE_TAX_DOWN = " %"
    
            self.INFO_MONTH_GASOLE_TAX_1 = ""
            self.GASOLE_TAXE_1 = -1
            self.MONTH_GASOLE_TAX_2 = ""
            self.GASOLE_TAXE_2 = -1
    
        def event_Init(self):
            self.KEYWORD_TITLE = 'SUBJECT "[Ci5 MRS] Tracing"'
            self.KEYWORD_MAIN_SEARCH = "automatic"
            self.KEYWORD_MAIN_REMOVE = "<br>"
            self.KEYWORD_SHIPMENT = "reference : "
            self.KEYWORD_CONTAINER = "cargo "
            self.KEYWORD_date = "at "
    
        def alliance_info_extract(self):
            self.INFO_SHIPMENT = self.DF_EXTRACT_MAIL["Shipment"][0]
            self.INFO_IMPORT = self.DF_EXTRACT_MAIL["Import"][0]
            self.INFO_EXPORT = self.DF_EXTRACT_MAIL["Export"][0]
            self.INFO_PRICING = self.DF_EXTRACT_MAIL["Pricing"][0]
    
    
    #
    # Keyword definition
    
    # Keyword definition for Alliance Logistics Analysis
    alliance_analysis = Specific_analysis()
    alliance_analysis.alliance_cnl_init()
    
    
    #
    # Error Management Init
    Error = Error_management()
    
    
    # Separating mail following company (PRO/CNL and All/Spe)
    general_mail_connexion.list_uid = general_mail_connexion.mail_Search(alliance_analysis.KEYWORD_TITLE)
    list_mail = np.zeros([len(general_mail_connexion.list_uid[0].split())], dtype = Mail_analysis)
    
    
    # mail Analysis for PROLINAIR Alliance
    for i, mail in enumerate(list_mail):
    
        # Init
        alliance_analysis.EXTRACT_MAIL = []
    
        #
        # Generic analysis
        try:

            mail = Mail_analysis()
            mail.uid = general_mail_connexion.list_uid[0].split()[i]

            # Decoding mail with utf-8 (bytes to string)
            mail.mail_decoding(general_mail_connexion.IMAP_SSL)

            # Subject
            mail.subject_search(global_parameters.SUBJECTS, global_parameters.SUBJECT_END)

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
                mail.uid, alliance_analysis.COMPANY_SENDER, alliance_analysis.COMPANY_RECEIVER,
                EXCEPTION = e, ERROR_TYPE = "ANALYSIS")
    
    
        #
        # Specific analysis
        try:
    
            for j in range(len(alliance_analysis.KEYWORD_DATA)):
                if alliance_analysis.KEYWORD_DATA[j][0] == 0:
                    alliance_analysis.EXTRACT_MAIL.append(mail.subject_keyword_search(
                        alliance_analysis.KEYWORD_DATA[j][1]))
                elif alliance_analysis.KEYWORD_DATA[j][0] == 1:
                    alliance_analysis.EXTRACT_MAIL.append(mail.inside_mail_search(
                        alliance_analysis.KEYWORD_DATA[j][1], alliance_analysis.KEYWORD_DATA[j][2]))
                else:
                    alliance_analysis.EXTRACT_MAIL.append(mail.inside_mail_search_Between(
                        alliance_analysis.KEYWORD_DATA[j][1][0], alliance_analysis.KEYWORD_DATA[j][1][1]))
    
            # Searching if previous shipment have been analysed already
            PREVIOUS_SHIPMENT = global_parameters.list_shipment_compare(alliance_analysis.EXTRACT_MAIL[2])
    
        # Exception
        except Exception as e:
            Error.error_mail_return(
                general_mail_connexion.IMAP_SSL, general_mail_connexion.SMTP_SSL,
                global_parameters.SENDER_EMAIL, global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                mail.uid, alliance_analysis.COMPANY_SENDER, alliance_analysis.COMPANY_RECEIVER,
                EXCEPTION = e, ERROR_TYPE = "ANALYSIS", shipment = alliance_analysis.EXTRACT_MAIL[2])
    
    
        # Checking that information have been found
        if (len(alliance_analysis.EXTRACT_MAIL[2]) > 1 and len(alliance_analysis.EXTRACT_MAIL[5]) > 0 and PREVIOUS_SHIPMENT == False):
    
            # Changing , into . to avoid wrong amounts
            alliance_analysis.EXTRACT_MAIL[5] = alliance_analysis.EXTRACT_MAIL[5].replace(",",".")
    
            # Creating dataframe
            alliance_analysis.DF_EXTRACT_MAIL = pd.DataFrame(
                np.transpose(np.expand_dims(np.array(alliance_analysis.EXTRACT_MAIL),axis = 1)),
                columns = ["Export", "Import", "Shipment", "Container", "Header", "Pricing"])
    
            # Extracting information from dataframe
            alliance_analysis.alliance_info_extract()
    
            # Dealing with Exception with mail Sending
            try:
    
                # Object Creation to send information to LOGITUDE/CARGOWISE
                alliance_sending = Sending_object(
                    global_parameters.URL_LINE, global_parameters.URL_2_TOKEN,
                    global_parameters.URL_2_DOC, global_parameters.URL_2_DOC_FILLING,
                    global_parameters.ACCESS_KEY, global_parameters.PRIMARY_KEY)
                alliance_sending.CORE_MAIL = mail.CORE_MAIL
    
                # Creating PDF
                alliance_sending.pdf_creation(
                    global_parameters.GENERAL_FOLDER_PATH,
                    f"{alliance_analysis.COMPANY_RECEIVER}_{alliance_analysis.COMPANY_SENDER}_{alliance_analysis.INFO_SHIPMENT}")
    
                # Sending Payable lines into LOGITUDE
                if (global_parameters.LG_CNL_MODE and alliance_analysis.INFO_SHIPMENT[3:4].isnumeric()):
    
                    # JSON creation
                    alliance_sending.json_creation_alliance(
                        alliance_analysis.INFO_IMPORT, alliance_analysis.INFO_SHIPMENT, alliance_analysis.INFO_PRICING)
    
                    # Sending into FTP to send into LOGITUDE
                    if global_parameters.DEBUG_MODE == False:
                        
                        # FTP connexion
                        alliance_sending.ftp_connexion(global_parameters.FTP_HOST,
                                                       global_parameters.FTP_USER,
                                                       global_parameters.FTP_PASS)
                       
                        # Saving file into FTP folder
                        bio = io.BytesIO(f"{alliance_sending.JSON}".encode())
                        alliance_sending.ftp.storbinary(f'STOR {alliance_analysis.INFO_SHIPMENT}_CNL_ALLIANCE.json',
                                                       bio)
                       
                        # Removing PDF
                        alliance_sending.pdf_removal(
                            f"{alliance_analysis.COMPANY_RECEIVER}_{alliance_analysis.COMPANY_SENDER}_{alliance_analysis.INFO_SHIPMENT}")
    
                        # Completing shipment List
                        global_parameters.list_shipment_completion(alliance_analysis.INFO_SHIPMENT)
    
                        # Sending CONFIRMATION mail
                        alliance_sending.sent_confirmation_mail(
                            general_mail_connexion.SMTP_SSL,
                            global_parameters.SENDER_EMAIL,
                            global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                            alliance_analysis.INFO_SHIPMENT)
                    
                    # Saving JSON for checking
                    else:
                        
                        # Saving JSON file inside local folder
                        with open(f"{alliance_analysis.INFO_SHIPMENT}_CNL_ALLIANCE.json", "w") as f:
                            f.write(alliance_sending.JSON)
    
    
                # Sending Payable lines into CARGOWISE
                else:
    
                    # XML creation
                    alliance_sending.xml_alliance_cargowise(
                        global_parameters.DEBUG_MODE, alliance_analysis.COMPANY_RECEIVER,
                        alliance_analysis.INFO_IMPORT, alliance_analysis.INFO_SHIPMENT,
                        alliance_analysis.INFO_PRICING)
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
                                    f"{alliance_analysis.COMPANY_RECEIVER}_{alliance_analysis.COMPANY_SENDER}_{alliance_analysis.INFO_SHIPMENT}",
                                    alliance_analysis.COMPANY_SENDER,
                                    alliance_analysis.INFO_SHIPMENT)
    
                            # Completing shipment list
                            global_parameters.list_shipment_completion(
                                alliance_analysis.INFO_SHIPMENT)
    
                        # Removing PDF
                        alliance_sending.pdf_removal(
                            f"{alliance_analysis.COMPANY_RECEIVER}_{alliance_analysis.COMPANY_SENDER}_{alliance_analysis.INFO_SHIPMENT}")

            # Exception
            except Exception as e:
                Error.error_mail_return(
                    general_mail_connexion.IMAP_SSL, general_mail_connexion.SMTP_SSL,
                    global_parameters.SENDER_EMAIL, global_parameters.RECEIVER_EMAIL_CONFIRMATION,
                    mail.uid, alliance_analysis.COMPANY_SENDER, alliance_analysis.COMPANY_RECEIVER,
                    EXCEPTION = e, ERROR_TYPE = "SENDING", SHIPMENT = alliance_analysis.INFO_SHIPMENT)

        list_mail[i] = mail

    # End of Running program
    global_parameters.list_shipment_saving()

    # Exit
    return list_mail
