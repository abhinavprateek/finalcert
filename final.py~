from fpdf import Template
import xlwt
import time
import os
import smtplib, email
import xlrd
import sys
from email.mime.multipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders

#license_no= raw_input("Enter license start series without license series : ")
wb_name = raw_input("Enter sheet name : ")
cc_1= raw_input("Enter cc address1 : ")
cc_2= raw_input("Enter cc address2 : ")
file_name=wb_name + ".xlsx"
print "Opened : " + file_name
workbook = xlrd.open_workbook(file_name)
worksheet=workbook.sheet_by_index(0)
num_row=worksheet.nrows-1
num_cell=worksheet.ncols-1
for rw in range(1, worksheet.nrows):
	name,email,grade,course,lic_no = [data.value for data in worksheet.row(rw)]
	#print lic_no
	if not os.path.exists('/home/skillspeed/finalcert/cert/%s' % (course,)):
		os.makedirs('/home/skillspeed/finalcert/cert/%s' % (course,))
#this will define the ELEMENTS that will compose the template. 
	elements = [
	    { 'name': 'company_logo', 'type': 'I', 'x1': 68.0, 'y1': 9.0, 'x2': 150.0, 'y2': 38.0, 'font': None, 'size': 0.0, 'bold': 0, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': 'logo', 'priority': 2, },
		{ 'name': 'bg1', 'type': 'I', 'x1':44.0, 'y1': 38.0, 'x2': 180.0, 'y2': 50.0, 'font': None, 'size': 0.0, 'bold': 0, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': 'logo', 'priority': 2, },
		{ 'name': 'title', 'type': 'T', 'x1': 52.0, 'y1': 42, 'x2': 168.0, 'y2': 47, 'font': 'Times', 'size': 18.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0XFFFFFF, 'background': 0, 'align': 'I', 'text': '', 'priority': 3, },
		{ 'name': 'to', 'type': 'T', 'x1': 38.0, 'y1': 58, 'x2': 67.0, 'y2': 62, 'font': 'Times', 'size': 13.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 2, },
		{ 'name': 'ul1', 'type': 'T', 'x1': 69.5, 'y1': 61.7, 'x2': 200.0, 'y2': 61.7, 'font': 'Times', 'size': 13.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 2, },
		{ 'name': 'complete', 'type': 'T', 'x1': 14.0, 'y1': 73, 'x2': 76.4, 'y2': 72.4, 'font': 'Times', 'size': 13.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 2, },
		{ 'name': 'ul2', 'type': 'T', 'x1': 80, 'y1': 73, 'x2': 200.0, 'y2': 73, 'font': 'Times', 'size': 13.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 2, },
		{ 'name': 'grade', 'type': 'T', 'x1': 78.0, 'y1': 85, 'x2': 130, 'y2': 85, 'font': 'Times', 'size': 13.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 2, },
		{ 'name': 'date', 'type': 'T', 'x1': 70.0, 'y1': 99, 'x2': 140, 'y2': 99, 'font': 'Times', 'size': 13.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 3, },
		{ 'name': 'bg', 'type': 'I', 'x1': 0, 'y1': 100, 'x2': 216.0, 'y2': 154, 'font': None, 'size': 0.0, 'bold': 0, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': 'logo', 'priority': 2, },
		{ 'name': 'sign', 'type': 'T', 'x1': 145, 'y1': 137, 'x2': 197, 'y2': 142, 'font': 'Times', 'size': 10.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 3, },
		{ 'name': 'name', 'type': 'T', 'x1': 86, 'y1': 55, 'x2': 150, 'y2': 64, 'font': 'Times', 'size': 18.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 3, },
		{ 'name': 'course', 'type': 'T', 'x1': 86, 'y1': 68, 'x2': 150, 'y2': 74, 'font': 'Times', 'size': 18.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 3, },
		{ 'name': 'grade1', 'type': 'T', 'x1': 95, 'y1': 80, 'x2': 150, 'y2': 86, 'font': 'Times', 'size': 18.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 3, },
		{ 'name': 'on', 'type': 'T', 'x1': 97.5, 'y1': 95, 'x2': 150, 'y2': 100, 'font': 'Times', 'size': 18.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 3, },
		{ 'name': 'copywrite', 'type': 'T', 'x1': 74, 'y1': 140, 'x2': 150, 'y2': 145, 'font': 'Times', 'size': 8.0, 'bold': 1, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 3, },
		{ 'name': 'license', 'type': 'T', 'x1': 15, 'y1': 95, 'x2': 150, 'y2': 145, 'font': 'Times', 'size': 14.0, 'bold': 0, 'italic': 0, 'underline': 0, 'foreground': 0, 'background': 0, 'align': 'I', 'text': '', 'priority': 3, },
	  	 ]

	#here we instantiate the template and define the HEADER
	f = Template(format="A5",orientation ="L", elements=elements,
		     title="Sample Invoice")
	f.add_page()

	#we FILL some of the fields of the template with the information we want
	#note we access the elements treating the template instance as a "dict"
	f["company_logo"] = "/home/skillspeed/finalcert/index.png"
	f["bg1"] = "/home/skillspeed/finalcert/bg1.jpg"
	f["bg"] = "/home/skillspeed/finalcert/bg2.png"
	f["title"] = "THIS CERTIFICATE IS PRESENTED TO"
	f["to"] = "Mr./Ms./Mrs."
	f["sign"] = "Sanjay Verma, Founder & CEO"
	f["ul1"] = "____________________________________________________"
	f["ul2"] = "_______________________________________________"
	f["complete"] = "For Successfully completing the"
	f["copywrite"]= "@2016 Blue Camphor Technologies (P) Ltd"
	f["grade"] = "with  ________   Grade"
	f["date"] = "Awarded on  ______________"
	f["name"]= name
	f["course"]= course
	f["grade1"]= grade
	f["on"]= time.strftime("%d/%m/%Y")
	f["license"]= "License No - " + lic_no

	#and now we render the page
	f.render("/home/skillspeed/finalcert/cert/%s/%s.pdf" % (course,name,))
	
	course_url=course.replace(" ","%20")
	#print course_url
	lic_date=time.strftime("%Y%m")
	#print lic_date
	#email code
	fromaddr = "support@skillspeed.com"
	print name +" " + email +" " + course +" " + lic_no
	print
	toaddr = email
	name_array=name.split()
	print name_array[0]

	msg = MIMEMultipart('alternative')
	 
	msg['From'] = fromaddr
	msg['To'] = toaddr
	msg['cc']= cc_1 + "," + cc_2
	msg['Subject'] = "[Certificate]:" + course
	 
	body = """\
	<html>
		<head></head>
		<body>
		<p>
Dear  %s,
<br></br>
<br></br>
We thank you for attending our <b> %s Classes </b> & hope you enjoyed the sessions much as we enjoyed teaching you.
<br></br>
<br></br>
We congratulate you on attending both the days of our <b>%s Classes </b>, and therefore would like to present you with a complete <b> %s </b>.
<br></br>
<br></br>
Add your certificate to Linkedin
<div><a href="https://www.linkedin.com/profile/add?_ed=0_JhwrBa9BO0xNXajaEZH4q9ZriGQBiq56O8XQeptEb_xAD6iVbtTHBphjlBeRBwz4aSgvthvZk7wTBMS3S-m0L6A6mLjErM6PJiwMkk6nYZylU7__75hCVwJdOTZCAkdv&pfCertificationName=%s&pfLicenseNo=%s&pfCertStartDate=%s&trk=onsite_html" rel="nofollow" target="_blank"><img src="https://download.linkedin.com/desktop/add2profile/buttons/en_US.png" alt="LinkedIn Add to Profile button"></a></div>
<br></br>
<br></br>
If you would like to carry on in this path and learn more and secure a job, then please contact our team about Advance Courses at sales@skillspeed.com.
<br></br>
<br></br>
USA & Rest of World: +1-661-241-4796
<br></br>
India: +91-906-602-0904
<br></br>
<br></br>
<b>Support</b>
<br></br>
<br></br>
If you have any pending technical queries regarding practicals or VM, please write to support@skillspeed.com, we will do our best to assist you.
<br></br>
<br></br>
<b>Good Karma</b>
<br></br>
<br></br>
Please do spread the word about our sessions and like us on Facebook & leave a comment about your thoughts.
<br></br>
<div><a href="http://www.facebook.com/SkillspeedOnline"><img src="http://www.mail-signatures.com/articles/wp-content/themes/emailsignatures/images/facebook-35x35.gif"></a><a href="http://www.linkedin.com/company/skillspeed"><img src="http://www.mail-signatures.com/articles/wp-content/uploads/2014/08/linkedin.png" width="35" height="35"></a></div>
<br></br>
<b>Thanks & Regards,
<br></br>
Skillspeed Support Team
<br></br></b>
</p>
</body>
</html>
""" % (name_array[0],course, course, course, course_url,lic_no, lic_date)
	 
	msg.attach(MIMEText(body, 'html'))
	 
	filename = "certificate.pdf"
	attachment = open("/home/skillspeed/finalcert/cert/%s/%s.pdf" % (course,name,), "rb")
	 
	part = MIMEBase('application', 'octet-stream')
	part.set_payload((attachment).read())
	encoders.encode_base64(part)
	part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
	 
	msg.attach(part)
	toadd= [toaddr] + [cc_1] + [cc_2]
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.starttls()
	server.login(fromaddr, "skillspeed")
	text = msg.as_string()
	server.sendmail(fromaddr, toadd, text)
	server.quit()
