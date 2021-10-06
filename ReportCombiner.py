# For read email
import imaplib
import email
from email.header import decode_header
# For send email
import smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
# For directory
import os
# For date time
from datetime import date
from datetime import datetime
from dateutil import parser
# For excel and word application
import win32com.client
from PyPDF2 import PdfFileMerger
import re

""" 
	EmailReader object is to read the email from email server
	the initial arguments are email credentials including username and password, the directory to store the word and excel files
	the function __getFileName filtering filename and save as a file that is defined the sequence of combined file
	the function __subjectMatched will return True if subject is matched
	the function "readEmail" is reading email from outlook server to download the attachment with .docx and .xlsx extension in the current date
"""
class EmailReader:

	def __init__(self, _username, _passowrd, _directory):
		
		# Email Credentials
		self.username = _username
		self.password = _passowrd

		# Directory to download the files
		self.directory = _directory

		# Outlook server
		self.imap = imaplib.IMAP4_SSL("outlook.office365.com")

		# The sequence dictionary to decide what the sequence of file is. the letter will be located at the first of the file name 
		self.reportSeq = {"plant": "a",
						"qc": "b",
						"scheduling": "c",
						"shipping": "d"}
		
    # The accepted keyword when the program filtering the emails
		self.subjectKeywords = ["morning", "daily", "meeting", "report"]
	
	def __getFileName(self, attachment):
		
		for key in self.reportSeq:
			if key in re.sub("[- _)(]", "", attachment).lower():
				return self.reportSeq[key] + attachment
		return "z" + re.sub('[:/\,.><?"*|]', "", str(datetime.now())) + attachment

	def __subjectMatched(self, subject):

		for keyword in self.subjectKeywords:
			if keyword in subject:
				return True
		return False
		

	def readEmail(self):

		# Login to email and select to inbox
		self.imap.login(self.username, self.password)
		status, messages = self.imap.select("INBOX")
		# If status is "OK"
		try:
			# Get the messages count, today's date, and reading (Bool) to determine whether the while loop break
			messages = int(messages[0])
			today = date.today()
			reading = True

			# loop around each email if the email is received today
			while reading and messages > 0:
				# fetch the email message by ID
				res, msg = self.imap.fetch(str(messages), "(RFC822)")
				# Read the email message from the latest
				messages -= 1
				for response in msg:
					if isinstance(response, tuple):
						# parse a bytes email into a message object
						msg = email.message_from_bytes(response[1])

						# decode the email subject
						subject, encoding = decode_header(msg["Subject"])[0]
						if isinstance(subject, bytes):
							# if it's a bytes, decode to str
							subject = subject.decode(encoding)
						# decode email sender
						From, encoding = decode_header(msg.get("From"))[0]
						if isinstance(From, bytes):
							From = From.decode(encoding)
						# decode email sender
						mailDate, encoding = decode_header(msg.get("Date"))[0]
						if isinstance(mailDate, bytes):
							mailDate = mailDate.decode(encoding)
						# If date is today
						if parser.parse(mailDate).date() == today:
							# If subject is "meeting report"
							if self.__subjectMatched(subject.lower()):
								att_path = "No attachment found."
								for part in msg.walk():
									if part.get_content_maintype() == 'multipart':\
										continue
									if part.get('Content-Disposition') is None:\
										continue
									# Get the attachment file name
									
									filename = self.__getFileName(part.get_filename())

									att_path = os.path.join(self.directory, filename)
									if not os.path.isfile(att_path) and os.path.splitext(filename)[1].lower() in [".xlsx", ".xls", ".docx", ".doc"]:
										fp = open(att_path, 'wb')
										fp.write(part.get_payload(decode=True))
										fp.close()
						
						# If date is not today then set the reading to False to break the loop
						else:
							reading = False
					
			# close the connection and logout
			self.imap.close()
			self.imap.logout()
			return "Success"

		# Login error
		except Exception as e:
			self.imap.close()
			return e

""" 
	FileCombine object is to transfer the excels and words file in the directory to PDFs then combine into one PDF files
	the initial arguments are the directory to access the word and excel files
	the private function "__excelToPDF" is transferring excel file into PDF file
	the private function "__wordToPDF" is transferring word file into PDF file
	the private function "__pdfMerger" is combining all the transfered PDF file into one PDF file
	the function combine is calling all the private function
"""
class FileCombine:

	def __init__(self, _directory):

		# Set the directory to access the files
		self.directory = _directory
		
	# Transfer one excel file to PDF
	def __excelToPDF(self, filename, fileext):

		# Start the Excel application then set the visible to False
		o = win32com.client.Dispatch("Excel.Application")
		o.Visible = False
		
		# Set the Workbook path then open it
		wb_path = self.directory + filename + fileext
		wb = o.Workbooks.Open(wb_path)
		# Path (Include the file name) to transfer PDF file
		path_to_pdf = self.directory + filename + '.pdf'
		# Get the index list of worksheets number (Index starts from 1)
		ws_index_list= [i for i in range(1, len(wb.WorkSheets) + 1) if wb.WorkSheets[i - 1].Visible != 0] 
		try:
			# Loop around the worksheets
			for i in range(len(wb.WorkSheets)):

				#off-by-one so the user can start numbering the worksheets at 1

				ws = wb.Worksheets[i]

				ws.PageSetup.Zoom = False

				ws.PageSetup.FitToPagesTall = 1

				ws.PageSetup.FitToPagesWide = 1
			
			# Select all the sheets
			wb.WorkSheets(ws_index_list).Select()
			# Export to PDF file
			wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
		
		except Exception as e:
			print(e)
		
		finally:
			# Close and save with argument True
			wb.Close(True)
			# Remove the excel file
			if os.path.exists(self.directory + filename + fileext):
				os.remove(self.directory + filename + fileext)
			# Quit the application
			o.Quit()
		
	# Transfer one word file to PDF
	def __wordToPDF(self, filename, fileext):
		# 17 is the format number for word application to save as pdf file
		wdFormatPDF = 17
		# Start the Word Application then set the visible to False
		o = win32com.client.Dispatch("Word.Application")
		o.Visible = False
		
		# Get the word file path
		doc_path = self.directory + filename + fileext
		# Set the path (including file name) for pdf file
		path_to_pdf = self.directory + filename + '.pdf'
		
		# Open word document
		doc = o.Documents.Open(doc_path)
		# Save as PDF
		doc.SaveAs(path_to_pdf, FileFormat=wdFormatPDF)
		# Close word and save with argument True
		doc.Close(True)
		# Remove the word file
		if os.path.exists(self.directory + filename + fileext ):
			os.remove(self.directory + filename + fileext)
		# Quit word application
		o.Quit()

	# Merge all the PDF files (pdfs: list of pdf files under the directory)
	def __pdfMerger(self, pdfs):
		
		# Call the PdfFileMerger library
		merger = PdfFileMerger()

		# Append all the pdf file from pdfs
		for pdf in pdfs:
			merger.append(self.directory + pdf)
				
		# Save the PDF file as "reports2021-09-20.pdf" for date of 09/20/2021
		merger.write(self.directory + "reports\\meeting" + str(date.today()) + ".pdf")
		# Close the merger
		merger.close()

		for pdf in pdfs:
			if os.path.exists(self.directory + pdf):
				os.remove(self.directory + pdf)
	
	# this function is to be called by the object that do all the functions
	def combine(self):

		# Get all the file name from directory
		for filename in os.listdir(self.directory):

			# File name and file extension
			filename, fileext = os.path.splitext(filename)
			fileext = fileext.lower()

			# This function only work if file extension is .xlsx or .docx
			if fileext == ".xlsx" or fileext == ".xls":
				self.__excelToPDF(filename, fileext)
			elif fileext == ".docx" or fileext == ".doc":
				self.__wordToPDF(filename, fileext)
		
		# Get the PDF files list
		pdfs = [filename for filename in os.listdir(self.directory) if os.path.splitext(filename)[1] == ".pdf"]

		if pdfs:
			self.__pdfMerger(pdfs)
			return self.directory + "reports\\meeting" + str(date.today()) + ".pdf"
		else:
			return "No files combined."


"""
	EmailSender Object is used to send email with combined pdf file
	Private function __geSubject is applied to get subject
	send function is the function to send email
"""
class EmailSender:

	def __init__(self, _sender, _password, _receivers, _pdfFile):
		
		self.sender = _sender
		self.password = _password
		self.receivers = _receivers
		self.pdfFile = _pdfFile

	def __getSubject(self):
		return f"Meeting Report {date.today()}"
	
	def send(self):
		# Create a multipart message and set headers
		message = MIMEMultipart()
		message["From"] = self.sender
		message["To"] = self.receivers
		message["Subject"] = self.__getSubject()
		body = """\
Good morning all,

	There is the combined meeting report.
				
Thanks."""

		# Add body to email
		message.attach(MIMEText(body, "plain"))

		filename = self.pdfFile  # In same directory as script

		# Open PDF file in binary mode
		with open(filename, "rb") as attachment:
			# Add file as application/octet-stream
			# Email client can usually download this automatically as attachment
			part = MIMEBase("application", "octet-stream")
			part.set_payload(attachment.read())

		# Encode file in ASCII characters to send by email    
		encoders.encode_base64(part)

		# Add header as key/value pair to attachment part
		part.add_header(
			"Content-Disposition",
			f"attachment; filename= meeting{str(date.today())}.pdf",
		)

		# Add attachment to message and convert message to string
		message.attach(part)
		text = message.as_string()

		# Log in to server using secure context and send email
		context = ssl.create_default_context()
		with smtplib.SMTP("smtp.office365.com", 587) as server:
			server.ehlo()  # Can be omitted
			server.starttls(context=context)
			server.ehlo()  # Can be omitted
			server.login(self.sender, self.password)
			server.sendmail(self.sender, self.receivers.split(","), text)

if __name__ == "__main__":

	# Email Credentials
	_username = "YOUR EMAIL ACCOUNT@EMAIL.COM"
	_password = "YOUR PASSWORD"

	# Files directory
	_directory = "YOUR BASE DIRECTORY"

	# Get reports from email server
	getReportsFromEmail = EmailReader(_username, _password, _directory)
	readEmailStatus = getReportsFromEmail.readEmail()

	if readEmailStatus == "Success":

		print(f"Email Reading Successful at {datetime.now()}")
		# Combine all the word and excel files into PDF
		reportsCombine = FileCombine(_directory)
		pdfFile = reportsCombine.combine()
		if pdfFile != "No files combined.":
			print("Combine Successful!")
			# Send out PDF to selected people
			f = open(_directory + "receivers\\receiver_email.txt", "r")
			_receivers = f.readline()
			f.close()
			emailSender = EmailSender(_username, _password, _receivers, pdfFile)
			emailSender.send()
		else:
			print("There is no PDF file to be combined and send out")
	else:
		print(readEmailStatus)
