# -*- coding: utf-8 -*-
"""
Created on Mon Jun  3 14:46:50 2024

@author: ythiriet
"""

# Global librairies
import socket
import imaplib
import smtplib
from datetime import date
import sys


# Object for mail Connexion
class Mail_connexion:
    def __init__(self):
        self.list_uid = ""
        self.IMAP_SSL = None
        self.SMTP_SSL = None
        self.CONNEXION_OK = False
        self.folders = []
        self.folders_number = 0

    # Testing Connexion
    def Connexion_Test(self, IMAP_HOST, IMAP_PORT):

        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.CONNEXION_OK = s.connect_ex((IMAP_HOST,IMAP_PORT))
        s.close()

        # Printing results
        if self.CONNEXION_OK:
            print("Problem with mail Server Connexion" + "\n")
        else:
            print("mail Server Connexion OK" + "\n")


    # Establishing Connexion
    def Connexion_Creation(self, IMAP_HOST, IMAP_PORT, SMTP_HOST, SMTP_PORT):

        # Connexion to IMAP and SMTP server
        try:
            self.IMAP_SSL = imaplib.IMAP4_SSL(host = IMAP_HOST, port = IMAP_PORT)
            self.SMTP_SSL = smtplib.SMTP_SSL(host = SMTP_HOST, port = SMTP_PORT)

        # Connexion error
        except Exception as e:

            # Print an error report in the console
            print("ErrorType : {}, Error : {}".format(type(e).__name__, e))

            # Error report creation
            Error_Report = open("Error_Report.txt", "a")
            Error_Report.write("\n" + str(date.today()) + " // " + "Connexion Error")

            # Stopping the program
            sys.exit()


    # Connexion Login
    def connexion_login(self, IMAP_USER, IMAP_PASS, SMTP_USER, SMTP_PASS):

        # Login to server (SMTP and IMAP server)
        try:
            self.IMAP_SSL.login(IMAP_USER, IMAP_PASS)
            self.SMTP_SSL.login(SMTP_USER, SMTP_PASS)
        except:
            # Print an error report in the console
            print("Login Error")

            # Error report creation
            Error_Report = open("Error_Report.txt", "a")
            Error_Report.write("\n" + str(date.today()) + " // " + "Login Error")

            # Stopping the program
            sys.exit()


    # Folder Search
    def mail_folder_search(self, FOLDER_NAME, NB_CHARA_FOLDER):

        # Init
        Folder_Found = False

        try:
            # Listing the folders
            resp_code, self.folders = self.IMAP_SSL.list()
            # Searching for specific folders
            for i_folder, Folder in enumerate(self.folders):
                if Folder.decode("utf-8").split(")")[-1][NB_CHARA_FOLDER:] == FOLDER_NAME:
                    self.Folder_number = i_folder
                    Folder_Found = True

            if Folder_Found == False:
                self.Folder_number = 0

        except:

            # Print an error report in the console
            print("Error in listing the folder inside the mailbox")

            # Error report creation
            Error_Report = open("Error_Report.txt", "a")
            Error_Report.write("\n" + str(date.today()) + " // " + "Error for folder search")

            # Stopping the program
            sys.exit()


    # Selecting the mailbox folder
    def mail_folder_selection(self):

        try:
            # Selecting the wanted folder
            resp_code, mail_count = self.IMAP_SSL.select(
                mailbox = self.folders[self.Folder_number].decode("utf-8").split(")")[-1][5:])

        except:

            # Print an error report in the console
            print("Error selecting a specific mailbox")

            # Error report creation
            Error_Report = open("Error_Report.txt", "a")
            Error_Report.write("\n" + str(date.today()) + " // " + "Error selecting the wanted folder")

            # Stopping the program
            sys.exit()


    # Searching for mails with specific subject
    def mail_Search(self, KEYWORD_TITLE):

        resp_code, Data = self.IMAP_SSL.search(None, KEYWORD_TITLE)

        # Exit
        return Data


    # Inbox Cleaning
    def inbox_cleaning(self):

        # Selecting all mails
        for i_mail in range(len(self.Data_All[0].split())):

            # Getting the mail content
            latest_email_uid_All = self.Data_All[0].split()[i_mail]

            # Data stored into trash folder
            self.IMAP_SSL.store(latest_email_uid_All, "+FLAGS", "\\Deleted")

        # Removing mails analysed
        self.IMAP_SSL.expunge()

        # Returning the number of mails removed
        return len(self.Data_All[0].split())


    # Closing Connexion
    def closing_connexion(self):
        self.IMAP_SSL.close()
        self.IMAP_SSL.logout()
        self.SMTP_SSL.quit()