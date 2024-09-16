# -*- coding: utf-8 -*-
"""
Created on Mon Feb  5 09:26:56 2024

@author: ythiriet
"""


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
sys.path.append(f"{os.getcwd()}/library")

# Object for mail Connexion
from Mail_connexion import Mail_connexion
# Object for mail Analysis
from Mail_analysis import Mail_analysis
# Object for Error Report
from Error_management import Error_management
# Object for Alliance PROLINAIR Analysis
from main_pro_alliance import main_pro_alliance
# Object for Alliance CNL Analysis
from main_cnl_alliance import main_cnl_alliance
# Object for Speedmove PROLINAIR Analysis
from main_pro_speedmove import main_pro_speedmove

from main_analysis_sending import main_analysis_sending


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

# Generic Parameters
class Parameters():
    def __init__(self):
        self.DEBUG_MODE = True
        self.LG_PRO_MODE = True
        self.LG_CNL_MODE = False
        self.LIST_SHIPMENT_MODE = False
        self.FOLDER_NAME = "Test2"
        # self.FOLDER_NAME = "INBOX"
        self.NB_CHARA_FOLDER = 11
        self.GENERAL_FOLDER_PATH = f"{os.getcwd()}"
        self.LIBRARY_FOLDER_PATH = ""
        self.SENDER_EMAIL = 
        self.RECEIVER_EMAIL = 
        self.RECEIVER_EMAIL_CONFIRMATION = 

        self.IMAP_HOST = 
        self.IMAP_USER = 
        self.IMAP_PASS = 
        self.IMAP_PORT = 
        self.SMTP_HOST = 
        self.SMTP_USER = 
        self.SMTP_PASS = 
        self.SMTP_PORT = 
        self.FTP_HOST = 
        self.FTP_USER = 
        self.FTP_PASS = 
        self.FTP_PORT = 

        self.LOGIN_CARGOWISE = 
        self.PASSWORD_CARGOWISE = 

        self.LIST_SHIPMENT_COMPLETED = []

        self.SUBJECTS = ["Subject:\r\n ", "Subject: "]
        self.SUBJECT_END = "\r\nThread"
        self.DATE = "Message-ID:"
        self.DATE_END = "ate:"

    # Change number of characters if specific name
    def chara_folder(self):
        if self.FOLDER_NAME == "INBOX":
            self.CHARA_FOLDER = 9
            self.FOLDER_NAME = "X"

    # Change Folder name if in Debug Mode
    def debug_mode_folder_path(self):
        if self.DEBUG_MODE:
            self.GENERAL_FOLDER_PATH = os.getcwd()

        self.LIBRARY_FOLDER_PATH = "./library/"

    # Adding the path for personal library download
    def debug_mode_path_append(self):
        if self.DEBUG_MODE == False:
            sys.path.append(self.LIBRARY_FOLDER_PATH)

    # Clearing variables only on local environment
    def local_clear(self):
        if self.DEBUG_MODE:
            from IPython import get_ipython
            get_ipython().magic('reset -sf')

    # Changing receiver email if in debug mode
    def debug_mode_mail(self):
        if self.DEBUG_MODE == False:
            self.RECEIVER_EMAIL = 

    # Loading List of shipment Completed
    def list_shipment_loading(self):
        if self.LIST_SHIPMENT_MODE:
            self.LIST_SHIPMENT_COMPLETED = np.zeros([0],dtype = object)
        else:
            with open("./list_shipment.pkl", "r"):
                self.LIST_SHIPMENT_COMPLETED = joblib.load("./list_shipment.pkl")

        # Turning array into dataframe
        self.dataframe_shipment_completed = pd.DataFrame(self.LIST_SHIPMENT_COMPLETED)


    # Comparing with list of shipment already completed
    def list_shipment_compare(self, SHIPMENT):
        PREVIOUS_SHIPMENT = False

        # Searching for current shipment into previous dataframe
        COMPARE = self.dataframe_shipment_completed.loc[
            (self.dataframe_shipment_completed[0] == SHIPMENT)]

        # Boolean result
        if COMPARE.shape[0] > 0:
            PREVIOUS_SHIPMENT = True

        return PREVIOUS_SHIPMENT

    # Saving shipment in the list
    def list_shipment_completion(self, SHIPMENT):
        self.LIST_SHIPMENT_COMPLETED = np.append(
            self.LIST_SHIPMENT_COMPLETED,
            np.zeros([1],
                     dtype= object),
            axis = 0)
        self.LIST_SHIPMENT_COMPLETED[-1] = SHIPMENT

    # Saving shipment list
    def list_shipment_saving(self):
        joblib.dump(self.LIST_SHIPMENT_COMPLETED, "list_shipment.pkl")


# Specific parameters following case analysis (PRO/CNL and ALLIANCE/SPEEDMOVE)
class Specific_analysis():
    def __init__(self):
        self.EXTRACT_MAIL = []
        self.DF_EXTRACT_MAIL = []
        self.MAIL_PROCESSED = False


    def alliance_pro_init(self):
        self.COMPANY_SENDER = "ALLIANCE"
        self.COMPANY_RECEIVER = "PROLINAIR"
        self.KEYWORD_TITLE = 'SUBJECT "PROLINAIR - Demande"'
        self.KEYWORD_DATA = [[0,"livraison",0],[0,"vement",0],
                             [1,"Donneur d’Ordre : ",10], [1,"Réf Alliance Logistics : ",8],
                             [1,"Objet : ",55], [2,[["Tarif :", "€ HT"],],0]]


    def alliance_cnl_init(self):
        self.COMPANY_SENDER = "ALLIANCE"
        self.COMPANY_RECEIVER = "CNL"
        self.KEYWORD_TITLE = 'SUBJECT "CN LOGISTICS FRANCE - Demande"'
        self.KEYWORD_MAIN_SEARCH = "confirmons"
        self.KEYWORD_MAIN_REMOVE = "<br>"
        self.KEYWORD_DATA = [[0,"livraison",0],[0,"vement",0],
                             [1,"Donneur d’Ordre : ",10], [1,"Réf Alliance Logistics : ",8],
                             [1,"Objet : ",55], [2,[["Tarif :", "€ HT"],],0]]
    
    
    def speedmove_init(self):
        self.COMPANY_SENDER = "SPEEDMOVE"
        self.COMPANY_RECEIVER = "PROLINAIR"
        self.KEYWORD_TITLE = 'SUBJECT "PROLINAIR - Demande"'
        self.KEYWORD_MAIN_SEARCH = "accusons"
        self.KEYWORD_MAIN_REMOVE = "<br>"
        # self.KEYWORD_DATA = [[0,"livraison",0],[0,"enlèvement",0],
        #                      [1,"Votre référence dossier est : ",10], [1,"Notre référence dossier est : ",6],
        #                      [1,"Objet : ",55], [2,[["montant de ce transport : ", "€"],["€", "+"]],0], [2,["Envoyé : ", "\nÀ :"],0]]
        self.KEYWORD_DATA = [[0,"livraison",0],[0,"vement",0],
                             [1,"Votre référence dossier est : ",10], [1,"Notre référence dossier est : ",6],
                             [1,"Objet : ",55], [2,[["€", "+ TAXE"],["montant de ce transport : ", "€"]],0], [2,[["Envoyé : ", "\nÀ :"],],0]]

#
# Keyword definition

# Execution time
START_TIME = time.time()


#
# General Parameters
global_parameters = Parameters()
global_parameters.list_shipment_loading()
global_parameters.chara_folder()
global_parameters.debug_mode_folder_path()
global_parameters.debug_mode_path_append()
global_parameters.local_clear()
global_parameters.debug_mode_mail()


#
# Mailbox Connexion
general_mail_connexion = Mail_connexion()
general_mail_connexion.Connexion_Test(global_parameters.IMAP_HOST, global_parameters.IMAP_PORT)
general_mail_connexion.Connexion_Creation(global_parameters.IMAP_HOST, global_parameters.IMAP_PORT, global_parameters.SMTP_HOST, global_parameters.SMTP_PORT)
general_mail_connexion.connexion_login(global_parameters.IMAP_USER, global_parameters.IMAP_PASS, global_parameters.SMTP_USER, global_parameters.SMTP_PASS)
general_mail_connexion.mail_folder_search(global_parameters.FOLDER_NAME, global_parameters.NB_CHARA_FOLDER)
general_mail_connexion.mail_folder_selection()



# Alliance for PROLINAIR
alliance_pro_analysis = Specific_analysis()
alliance_pro_analysis.alliance_pro_init()
list_mail_pro_alliance = main_analysis_sending(global_parameters, general_mail_connexion, alliance_pro_analysis)

# Alliance for CNL
alliance_cnl_analysis = Specific_analysis()
alliance_cnl_analysis.alliance_cnl_init()
list_mail_cnl_alliance = main_analysis_sending(global_parameters, general_mail_connexion, alliance_cnl_analysis)

# Speedmove for PROLINAIR
speedmove_pro_analysis = Specific_analysis()
speedmove_pro_analysis.speedmove_init()
list_mail_pro_speedmove = main_analysis_sending(global_parameters, general_mail_connexion, speedmove_pro_analysis)


# Inbox cleaning
if global_parameters.DEBUG_MODE == False:
    NB_mail_REMOVED = general_mail_connexion.inbox_cleaning()

# # Calculating the number of mails analysed
NB_MAIL_ANALYSED = (list_mail_pro_alliance.shape[0] +
                    list_mail_cnl_alliance.shape[0] +
                    list_mail_pro_speedmove.shape[0])

# Time execution
EXEC_TIME = time.time() - START_TIME

# Writing CONFIRMATION
body = "\nLes mails, issus d'Alliance et de Speed Move, ont été correctement traités \n"
body += f"Nombre de mails traités : {NB_MAIL_ANALYSED} mails \n"
body += f"Durée d'execution : {round(EXEC_TIME,2)} secondes"
print(body)

# CONFIRMATION report creation
Treatment_Report = open("Treatment_Report.txt", "a")
Treatment_Report.write(body)
Treatment_Report.close()

# Closing Connexion
general_mail_connexion.closing_connexion()

# Program Ending CONFIRMATION
print(f"<br> End of running program after {round(EXEC_TIME,2)} seconds")
