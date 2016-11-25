'''
Created on Nov 12, 2016
'''
from openpyxl import load_workbook, Workbook
from sys import exit
import os
from Tkinter import Tk
from tkFileDialog import askdirectory
from win32com import client

EMAIL_FILE = 'emails.txt'
SUBJECT_FILE = 'subject.txt'
BODY_FILE = 'body.txt'

class SheetEmailer:
    
    def __init__(self):
        self.absolute_file_paths = []
        self.client_to_email = dict()
        self.all_client_numbers = set()
        self.subject = ''
        self.body = ''
        self._read_emails()
        self._read_subject()
        self._read_body()


        print 'Please choose a directory'
        Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
        directory = askdirectory()
        
        if directory:
            print 'Chosen directory: ' + directory
            self.workbook_directory = directory
        else:
            exit('Please choose a valid directory')

    def _read_emails(self):
        email_sources_path = os.path.abspath(EMAIL_FILE)
        print 'Loading email address from ' + email_sources_path
        try:
            email_file = open(EMAIL_FILE, 'r')
        except:
            exit(EMAIL_FILE + ' does not exist')

        for line in email_file:
            email_mapping = line.split('=')
            client_number = email_mapping[0]
            client_email = email_mapping[1]
            self.client_to_email[client_number] = client_email.strip().split(',')
            self.client_to_email[client_number] = [c.strip() for c in self.client_to_email[client_number]]

        print 'Email mapping ' + str(self.client_to_email)
        
    def _read_subject(self):
        subject_sources_path = os.path.abspath(SUBJECT_FILE)
        print 'Loading subject from ' + subject_sources_path
        try:
            with open(subject_sources_path, 'r') as content_file:
                self.subject = content_file.read()
        except:
            exit(subject_sources_path + ' does not exist')

        print 'Subject: ' + self.subject
        
    def _read_body(self):
        body_sources_path = os.path.abspath(BODY_FILE)
        print 'Loading body from ' + body_sources_path
        try:
            with open(body_sources_path, 'r') as content_file:
                self.body = content_file.read()
        except:
            exit(body_sources_path + ' does not exist')

        print 'Subject: ' + self.body
            
    def create_user_directories_and_files(self):
        for filename in os.listdir(self.workbook_directory):
            print 'Working on workbook: ' + filename
            workbook_name = filename.split('.')[0]
            full_workbook_path = os.path.join(self.workbook_directory, filename)
            try:
                loaded_workbook = load_workbook(filename = full_workbook_path)
            except:
                print "Unable to load workbook " + filename + ". Will continue to next workbook."
                continue
            
            
            for sheet_name in loaded_workbook.sheetnames:
                print 'Working on sheet: ' + sheet_name
                current_sheet = loaded_workbook[sheet_name]
                
                new_workbook, user_number = self._create_workbook_from_sheet(current_sheet)
                self.all_client_numbers.add(user_number)
                self._save_workbook(user_number, workbook_name, new_workbook)
            
    def _create_workbook_from_sheet(self, current_sheet):
        new_workbook = Workbook()
        active_sheet = new_workbook.active
        current_sheet_rows = current_sheet.rows

        user_number = current_sheet.oddFooter.left.text.strip()
         
        for row_num, row in enumerate(current_sheet_rows):
            for col_num, cell in enumerate(row):
                active_sheet.cell(column=col_num + 1, row=row_num + 1).value = cell.value
                
        return new_workbook, user_number
    
    def _save_workbook(self, user_number, workbook_name, new_workbook):
        self._create_directory_if_none_exists(user_number)
            
        new_file_path = os.path.join(user_number, user_number + '_' + workbook_name + '.xlsx')
        
        absolute_path_to_file = os.path.abspath(new_file_path)
    
        print "Writing " + absolute_path_to_file
        
        self.absolute_file_paths.append(absolute_path_to_file)
    
        new_workbook.save(new_file_path)
        
        
    def _create_directory_if_none_exists(self, user_number):
        absoulte_path_to_sheet = os.path.abspath(user_number)
        if not os.path.exists(user_number):
            print "Creating directory: " + absoulte_path_to_sheet
            os.makedirs(user_number)
        else:
            print "Directory already exists: " + absoulte_path_to_sheet
            
    def convert_all_to_pdf(self):
        print 'Starting excel'
        xlApp = client.Dispatch("Excel.Application")
        for path in self.absolute_file_paths:
            print 'Converting ' + path + ' to pdf'
            books = xlApp.Workbooks.Open(path)
            ws = books.Worksheets[0]
            ws.Visible = 1
            pdf_path = os.path.join(os.path.dirname(path), os.path.basename(path).split('.')[0] + '.pdf')
            ws.ExportAsFixedFormat(0, pdf_path)
            print 'Exported ' + pdf_path
        xlApp.quit()


    def save_email_to_draft(self):
        for user in self.all_client_numbers:
            absoulte_path_to_directory = os.path.abspath(user)
            if user in self.client_to_email:
                outlook = client.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = ','.join(self.client_to_email[user])
                mail.Subject = self.subject
                mail.HtmlBody = self.body
                for filename in os.listdir(absoulte_path_to_directory):
                    if filename.endswith(".pdf"):
                        path_to_pdf = os.path.join(absoulte_path_to_directory, filename)
                        print 'Adding pdf: ' + path_to_pdf
                        mail.Attachments.Add(Source=path_to_pdf )
                mail.Save()
            else:
                self._create_directory_if_none_exists('no_email')
                absoulte_path_to_no_email_directory = os.path.abspath('no_email')
                for filename in os.listdir(absoulte_path_to_directory):
                    os.rename(os.path.join(absoulte_path_to_directory, filename), os.path.join(absoulte_path_to_no_email_directory, filename))
                os.rmdir(absoulte_path_to_directory)
        outlook.quit()

if __name__ == "__main__":
    print 'Starting'
    sheet_emailer = SheetEmailer()
    sheet_emailer.create_user_directories_and_files()
    sheet_emailer.convert_all_to_pdf()
    sheet_emailer.save_email_to_draft()
    print 'Done'