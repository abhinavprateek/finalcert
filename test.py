from fpdf import Template
import time
import smtplib, email
import xlrd
import sys
from email.mime.multipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders

license_no= raw_input("Enter license start series without license series : ")
wb_name = raw_input("Enter sheet name : ")
cc= raw_input("Enter cc address : ")
file_name=wb_name + ".xlsx"
print "Opened : " + file_name
workbook = xlrd.open_workbook(file_name)
worksheet=workbook.sheet_by_index(0)
num_row=worksheet.nrows-1
num_cell=worksheet.ncols-1
for rw in range(1, worksheet.nrows):
	name,email,course,lic = [data.value for data in worksheet.row(rw)]
	print name
	print email
	print course
