# -*- coding: utf-8 -*-
"""
Created on Fri Sep 13 10:05:41 2024

@author: ythiriet
"""

# Global librairies
import os
import sys
import requests
from fpdf import FPDF
import json
import xml.dom.minidom
from email.message import EmailMessage
from ftplib import FTP
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Local librairies
sys.path.append(f"{os.getcwd()}/Library")


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


# Class creation
class Sending_object():
    def __init__(self):

        self.CORE_MAIL = ""
        
        # LOGITUDE
        self.JSON = ""
        self.PDF = ""

        # CARGOWISE
        self.XML_CARGOWISE = ""
        self.URL_CARGOWISE = "https://nxcprdservices.wisegrid.net/eAdaptor"
        self.RESPONSE_CW = ""


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


    def json_creation_alliance(self, DATAFRAME_EXTRACT):

        # Difference case following Import/Export
        CHARGECODE = "OTPT"
        if DATAFRAME_EXTRACT["IMPORT"][0]:
            CHARGECODE = "DTPT"

        # Creating dictionnary
        Dict = {"ShipmentNumber": DATAFRAME_EXTRACT["SHIPMENT"][0],
              "Payables": [
                  {"ChargeType": CHARGECODE,
                  "Currency": "EUR",
                  "Vendor": "1003",
                  "UnitPrice": float(DATAFRAME_EXTRACT["PRICING"][0]),
                  "Quantity": 1
                  }]}

        # Converting dictionnary to JSON file
        self.JSON = json.dumps(Dict)


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


    def xml_alliance_cargowise(self, DEBUG_MODE, DATAFRAME_EXTRACT, COMPANY_RECEIVER):

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
        Key.appendChild(doc.createTextNode(DATAFRAME_EXTRACT["SHIPMENT"][0]))
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
        if DATAFRAME_EXTRACT["IMPORT"][0]:
            Code.appendChild(doc.createTextNode("DTPT"))
        else:
            Code.appendChild(doc.createTextNode("OTPT"))
        CHARGECODE.appendChild(Code)

        # Balise fermante

        # Element : Cost OS Amount
        CostOSAmount = doc.createElement("CostOSAmount")
        CostOSAmount.appendChild(doc.createTextNode(DATAFRAME_EXTRACT["PRICING"][0]))
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
                f"{DATAFRAME_EXTRACT['SHIPMENT'][0]}.xml", 'w'),
                indent ="",
                addindent ="    ",
                newl='\n')
    
    
    def xml_speedmove_cargowise(self, DEBUG_MODE, DATAFRAME_EXTRACT, COMPANY_RECEIVER):


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
        Key.appendChild(doc.createTextNode(DATAFRAME_EXTRACT["SHIPMENT"][0]))
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
        if DATAFRAME_EXTRACT["IMPORT"][0]:
            Code.appendChild(doc.createTextNode("DTPT"))
        else:
            Code.appendChild(doc.createTextNode("OTPT"))
        CHARGECODE.appendChild(Code)

        # Balise fermante

        # Element : Cost OS Amount
        CostOSAmount = doc.createElement("CostOSAmount")
        CostOSAmount.appendChild(doc.createTextNode(DATAFRAME_EXTRACT["PRICING"][0]))
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
        if DATAFRAME_EXTRACT["IMPORT"][0]:
            Code.appendChild(doc.createTextNode("DFUE"))
        else:
            Code.appendChild(doc.createTextNode("OFUE"))
        CHARGECODE.appendChild(Code)

        # Balise fermante

        # Element : CostOSAmount
        CostOSAmount = doc.createElement("CostOSAmount")
        if DATAFRAME_EXTRACT["MONTH_DATE"][0] == DATAFRAME_EXTRACT["MONTH_GASOLE_TAXE_2"][0]:
            CostOSAmount.appendChild(doc.createTextNode(DATAFRAME_EXTRACT["GASOLE_TAXE_2"][0]))
        else:
            CostOSAmount.appendChild(doc.createTextNode(DATAFRAME_EXTRACT["GASOLE_TAXE_1"][0]))
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
                f"{DATAFRAME_EXTRACT['SHIPMENT'][0]}.xml", 'w'),
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