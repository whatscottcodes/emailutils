import win32com.client
from datetime import datetime
import os

class emailUtils:
    """
    Class for working with Outlook emails. Requires Outlook to be installed.
    """
    def __init__(self, user_email):
        """
        Initialized with user's email which should be the email address associated with the local Outlook install.
        
        Using the user's email address the user's folder and inbox folder are initialized.
        """
        self.user_email = user_email
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.user_folder = self.set_user_folder()
        self.inbox = self.set_inbox()
    
    def set_user_folder(self):
        """
        Get user's folder from Outlook
        """ 
    
        for folder in self.outlook.GetNamespace("MAPI").Folders:
            if str(folder) == self.user_email:
                return folder

    def set_inbox(self):
        """
        Get inbox folder from user's Outlook folder.
        """
        for folder in self.user_folder.Folders:
            if str(folder) == "Inbox":
                return folder

    def get_subfolder(self, subFolder):
        """
        Get specified sub folder from inbox.
        
        Helpful to create folder in the inbox and use rules to bin the emails you'll be looking to pull attachments from.
        """
        if subFolder is None:
            return self.inbox

        for folder in self.inbox.Folders:
            if str(folder) == subFolder:
                return folder
            
    def download_files(self, subj, subFolder, savePath, dateMatch = None, only_dl_today = False):
        """
        Downloaded attachment from file with matching subject.
        
        subj: string to match outlook subject to
        savePath: path to save the attachment files to
        dateMatch: email recieved date to match
            - if None and only_dl_date_match is True then defaults to today's date
            - if None and only_dl_date_match is False then downloads attachments from all subj matching emails
        only_dl_date_match: only downloads the extract from an email with a date matching dateMatch
        
        returns filePaths: file paths of downloaded extracts
        """
        
        fileEmails = self.get_subfolder(subFolder)
        
        subFolderMessages = fileEmails.Items
        filePaths = []
        
        if dateMatch is None:
            dateMatch = datetime.today().strftime("%y-%m-%d")
        else:
            dateMatch = dateMatch.strftime("%y-%m-%d")
            
        for i in range(0, len(subFolderMessages)):
            
            subject = str(subFolderMessages[i])
            email_date = subFolderMessages[i].SentOn.strftime("%y-%m-%d")
            if only_dl_today:
                if (subj in subject) & (email_date == dateMatch):
                    message = subFolderMessages[i]
                else:
                    message = None
                    continue
            else:
                if subj in subject:
                    message = subFolderMessages[i]
                else:
                    message = None
                    continue

            subFolderItemAttachments = message.Attachments
            nbrOfAttachmentInMessage = subFolderItemAttachments.Count
            
            for i in range(1, nbrOfAttachmentInMessage+1):
                attachment = subFolderItemAttachments.item(i)
                pathToFile = os.path.join(os.getcwd(), savePath, str(attachment))
                attachment.SaveAsFile(pathToFile)
                filePaths.append(pathToFile)

        return filePaths

    def get_matching_subjects(self, subj, subFolder, dateMatch = None, only_dl_today = False):
        """
        Downloaded attachment from file with matching subject.
        
        subj: string to match outlook subject to
        dateMatch: email recieved date to match
            - if None and only_dl_date_match is True then defaults to today's date
            - if None and only_dl_date_match is False then downloads attachments from all subj matching emails
        only_dl_date_match: only downloads the extract from an email with a date matching dateMatch
        
        returns emailSubjects: list of email subjects matching/containing subj
        """
        fileEmails = self.get_subfolder(subFolder)
        
        subFolderMessages = fileEmails.Items
        emailSubjects = []
        
        if dateMatch is None:
            dateMatch = datetime.today().strftime("%y-%m-%d")
        else:
            dateMatch = dateMatch.strftime("%y-%m-%d")
            
        for i in range(0, len(subFolderMessages)):
            subject = str(subFolderMessages[i])
            email_date = subFolderMessages[i].SentOn.strftime("%y-%m-%d")
            if only_dl_today:
                if (subj in subject) & (email_date == dateMatch):
                    emailSubjects.append(subject)
                else:
                    continue
            else:
                if subj in subject:
                    emailSubjects.append(subject)
                else:
                    continue
        return emailSubjects
    
    def get_email_bodies(self, subj, subFolder, dateMatch = None, only_dl_today = False):
        """
        Downloaded attachment from file with matching subject.
        
        subj: string to match outlook subject to
        dateMatch: email recieved date to match
            - if None and only_dl_date_match is True then defaults to today's date
            - if None and only_dl_date_match is False then downloads attachments from all subj matching emails
        only_dl_date_match: only downloads the extract from an email with a date matching dateMatch
        
        returns emailBodies: list of email bodies from emails with a subject matching/containing subj
        """
        fileEmails = self.get_subfolder(subFolder)
        
        subFolderMessages = fileEmails.Items
        emailBodies = []

        if dateMatch is None:
            dateMatch = datetime.today().strftime("%y-%m-%d")
        else:
            dateMatch = dateMatch.strftime("%y-%m-%d")
            
        for i in range(0, len(subFolderMessages)):
            subject = str(subFolderMessages[i])
            email_date = subFolderMessages[i].SentOn.strftime("%y-%m-%d")
            if only_dl_today:
                if (subj in subject) & (email_date == dateMatch):
                    message = subFolderMessages[i]
                else:
                    continue
            else:
                if subj in subject:
                    message = subFolderMessages[i]
                else:
                    continue
            emailBodies.append(message.body)            
            
        return emailBodies
    
    def send_email(self, subject, mailToList=None, incl_subj_date=False, body=None, htmlBody=None, attachmentPaths=None, ccList=None, sendFrom=None):
        """
        Uses win32 package to send outlook emails
        
        subject: subject of email
        mailToList: list of emails to send email to - if none sends to user's email.
        incl_subj_date: if true inserts today's date into the subject.
        body: body of the email - if none and there are attachments default to "See Attached."
        htmlBody: HTML body of the email to send.
        attachmentPaths: list of attachment paths to be attached to the email.
        ccList: list of emails to include in the CC.
        sendFrom: alternative email address to send from - if none send from user's email.
        """
    
        mail = self.outlook.CreateItem(0)
        
        if mailToList is None:
            mail.To = self.user_email
        else:
            mail.To = ";".join(mailToList)
        
        if sendFrom is not None:
            mail.SentOnBehalfOfName = di.mailFrom
        
        if ccList is not None:
            mail.CC = ";".join(ccList)        
        
        if incl_subj_date:
            today = datetime.today().strftime("%m-%d-%y")
            mail.Subject = f"{subject} {today}"
        else:
            mail.Subject = subject
        
        if (body is None) & len(attachmentPaths) > 0: 
            mail.Body = "See attached."
        else:
            mail.Body = "Message body"
    
        if htmlBody is not None:
            mail.HTMLBody = htmlBody

        for attachmentPath in attachmentPaths:
            mail.Attachments.Add(attachmentPath)
    
        mail.Send()

    def format_table(self, df, colFormats, headerFill = "#cc0000", headerFontColor = "#ffffff"):
        """
        Creates HTML tables of pandas dataframe for inclusion in emails.

        df: pandas dataframe
        colFormats: dictionary mapping column names to formating
        headerFill: fill color of the table's column names/headers - defaults to a CVS red.
        headerFontColor: color for column names/header font- defaults to white.

        returns
        html_table: the html for the table.
        """
        th_props = [
            ('text-align', 'center'),
            ('font-weight', 'bold'),
            ('color', headerFontColor),
            ('background-color', headerFill),
            ('border-collapse', 'collapse'),
            ('padding', '3px 5px'),
            ('border', '1px solid #000000')
            ]

        td_props = [
            ('text-align', 'center'),
            ('border', '1px solid #000000'),
            ('border-collapse', 'collapse'),
            ('padding', '3px 5px')
            ]
    
        styles = [
            dict(selector="th", props=th_props),
            dict(selector="td", props=td_props)
            ]

        table_html = df.style.format(colFormats).set_table_styles(styles).hide_index()
    
        return table_html    