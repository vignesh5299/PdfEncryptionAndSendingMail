import PyPDF2
import os
from datetime import datetime
# import pyPdf
import csv
import mysql.connector
import tkinter
from tkinter import filedialog
import  smtplib
from email.message import EmailMessage


class pdfEncyrptor:
    """This class will be used to encrypt the pdf files from particular folder and sending mail to user"""

    def getfilesFromFolder(self):
        """This function will perform fetching files from folder operation"""
        root = tkinter.Tk()
        root.withdraw()
        # folder = r"C:\Users\mohan.si\Desktop\Techneplus\project_techneplus"
        folder = filedialog.askdirectory()
        files = os.listdir(folder)
        print('files from folder')
        print(files)
        return files

    def encryption(self):
        """Here we are doing encryption of all the  pdf files and refactoring it by adding timeStamp"""
        lst1 = []
        already_encryptedpdf_lst = []
        pdf_files_list = []
        password_lst = []
        email_lst = []
        time_lst = []
        for pdf_files in self.getfilesFromFolder():
            if pdf_files.endswith(".pdf"):
                pdf_files_list.append(pdf_files)
                ouput_file_writer_name = pdf_files.split('.pdf')[0]
                password_lst.append(ouput_file_writer_name)
                if PyPDF2.PdfFileReader(open(pdf_files, "rb")).isEncrypted:
                    already_encryptedpdf_lst.append(pdf_files)
                    password_lst.pop()
                    continue
                pdf_in_file = open(pdf_files, 'rb')

                inputpdf = PyPDF2.PdfFileReader(pdf_in_file)
                pages_no = inputpdf.numPages
                usermail = input("enter the mail id of the enduser: ")
                email_lst.append(usermail)
                output = PyPDF2.PdfFileWriter()

                for i in range(pages_no):
                    inputpdf = PyPDF2.PdfFileReader(pdf_in_file)

                    output.addPage(inputpdf.getPage(i))
                    output.encrypt(ouput_file_writer_name)

                current_time = datetime.now().replace(microsecond=0)
                new_current_time = datetime.strftime(current_time, "%Y_%B_%d_%H_%M_%S")
                time_lst.append(new_current_time)
                output_file = ouput_file_writer_name + new_current_time + ".pdf"
                with open(output_file, "wb") as outputStream:
                    output.write(outputStream)
                    lst1.append(output_file)
                pdf_in_file.close()
        print("Pdf files from Folder")
        print(pdf_files_list)
        print("Already Encrypted files list")
        print(already_encryptedpdf_lst)
        print("Encrypted files List ")
        print(lst1)
        self.send_mail(lst1, password_lst, email_lst)
        return lst1, password_lst, time_lst, email_lst

    def getfileSize(self, output_file):
        file_size = os.path.getsize(output_file)
        return file_size

        #
        # file_size = os.stat(output_file)
        # print(file_size.st_size)
        #
        # write_file_obj = open(output_file)
        # write_file_obj.seek(0, os.SEEK_END)
        # file_size = write_file_obj.tell()
        # print(file_size)

    def sendRecord(self, lst_encrypted_pdf_files, password_lst, lst_file_size, time_lst, email_lst):

        with open("Pdf_File_Details.csv", "a", newline="") as file_obj:
            csv_file_obj = csv.writer(file_obj)

            csv_file_name = "Pdf_File_Details.csv"
            csv_file_size = os.stat(csv_file_name).st_size == 0

            if csv_file_size:
                csv_file_obj.writerow(["filename", "filepassword", "filesize", "timeofencryption", "usermail"])
            for i in range(len(lst_encrypted_pdf_files)):
                csv_file_obj.writerow([lst_encrypted_pdf_files[i], password_lst[i],
                                       lst_file_size[i], time_lst[i], email_lst[i]])

    def dbconnection(self, lst_encrypted_pdf_files, password_lst, lst_file_size, time_lst, email_lst):
        """This function is used to connect with dataBase to create table and insert record"""

        try:
            connection = mysql.connector.connect(user='Dhoni', password='dhoni07', host='localhost', port=3306, database='python')
            print("connection established")
            query1 = "CREATE TABLE IF NOT EXISTS PDFDETAILS (FILENAME varchar(255), FILEPASSWORD varchar(255), " \
                     "FILESIZE INT, TIME varchar(255), EMAIL varchar(255))"
            cursor = connection.cursor()
            cursor.execute(query1)
            print("Executed Query No 1")

            insert_query = """insert into pdfdetails (FILENAME, FILEPASSWORD, FILESIZE, TIME, EMAIL) values (%s, %s, %s, %s, %s)"""
            record_to_insert = [(lst_encrypted_pdf_files[0], password_lst[0], lst_file_size[0], time_lst[0], email_lst[0]), (lst_encrypted_pdf_files[1], password_lst[1], lst_file_size[1], time_lst[1], email_lst[1])]
            cursor.executemany(insert_query, record_to_insert)
            connection.commit()
            print("Executed Query No 2")
            print(cursor.rowcount, " Record inserted into PDFDETAILS:")

        except mysql.connector.Error as error:
            print("Failed to insert record into pdfdetails {}".format(error))

        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()
                print("connection Terminated")

    def send_mail(self, lst1, password_lst, email_lst):
        try:
            for i in range(len(email_lst)):
                msg = EmailMessage()
                msg["Subject"] = "This Mail is contains the pdfs Encrypted"
                msg['From'] = "vigneshpsv52@gmail.com"
                msg['To'] = email_lst[i]
                msg.set_content(f"The Password for below attached file is: {password_lst[i]}")
                print("Password attached")
                with open(lst1[i], 'rb') as f:
                    file_data = f.read()
                    file_name = f.name
                msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)
                with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                    smtp.login("vigneshpsv52@gmail.com", "passwordtakeitfromuserfile")
                    smtp.send_message(msg)
                    print("Mail sent successfully")

        except Exception as e:
            print(e)


    def maintain(self):
        lst_encrypted_pdf_files, password_lst, time_lst, email_lst = self.encryption()
        lst_file_size = []
        print("password list")
        print(password_lst)
        print("Time List")
        print(time_lst)
        print("Email list")
        print(email_lst)
        for file in lst_encrypted_pdf_files:
            lst_file_size.append(self.getfileSize(file))
        print("size of files list")
        print(lst_file_size)
        self.sendRecord(lst_encrypted_pdf_files, password_lst, lst_file_size, time_lst, email_lst)
        self.dbconnection(lst_encrypted_pdf_files, password_lst, lst_file_size, time_lst, email_lst)


if __name__ == "__main__":
    pen = pdfEncyrptor()
    pen.maintain()
