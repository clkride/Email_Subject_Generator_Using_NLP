#!/usr/bin/env python
# coding: utf-8


from pathlib import Path


get_ipython().system('pip install pywin32')


# import libraries
import win32com.client
import pandas as pd



# create data folder
data_dir = Path.cwd() / "Output"
data_dir.mkdir(parents= True, exist_ok=True)



# Connect to outlook mailbox
outlook = win32com.client.dynamic.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Connect to folder
inbox_folder = outlook.GetDefaultFolder(6)

# Get messages
messages = inbox.Items



# Restrict messages received during last 1 year
received_dt = datetime.now() - timedelta(days=365)
received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")



# Create a dataframe to store message related details
my_df = pd.DataFrame(columns=['subject','body','categories','importance','message_class','recipient',
                              'sender_mail_type','sender_name','sensitivity','size','creation_time','bcc','cc','companies'])



# Get message details and store in a csv file
for message in messages:
  new_row = {'subject':message.Subject,'body':message.body,'categories':message.Categories,'importance':message.Importance,
             'message_class':message.MessageClass,'recipient':message.Recipients,'sender_mail_type':message.SenderEmailType,
             'sender_name':message.SenderName,'sensitivity':message.Sensitivity,'size':message.Size,'creation_time':message.CreationTime,
             'bcc':message.BCC,'cc':messgae.cc,'companies':message.Companies}
  my_df = my_df.append(new_row, ignore_index=True)


# Take a look at the data
my_df.head(10)







