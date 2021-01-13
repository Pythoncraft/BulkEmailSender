import win32com.client as client
import os
import csv
from time import sleep


with open('TestSample.csv', 'r', newline='') as f:
	reader = csv.reader(f)
	name_list = [row for row in reader]

outlook = client.Dispatch("Outlook.Application") 
for email, name in name_list:
	message = outlook.CreateItem(0) # creating new message for each email in the loop
	message.Subject = name
	html_body = """
<div>
	ENTER YOUR TEXT BETWEEN THE TAGS
</div><br>
"""
	message.HTMLBody = html_body
	message.To = email
	# message.Send() ## uncomment if you want to send all the messages
	sleep(15) # delay after sending the message. in seconds

