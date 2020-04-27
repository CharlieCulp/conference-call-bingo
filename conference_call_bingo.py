# coding: utf-8
# =============================================================================
# Conference Call Bingo 
#
# Created by: Charlie Culp
# Creation date: 2/4/2020
# Python version: 3.7x 
#
# Change log
# Date         User Description
# 02/04/2020   CC   Initial release
# 02/20/2020   CC   Problems with regular Python running win32com so switched
#                   to Anaconda version and task scheduler now works. 
# 02/20/2020   CC   Added worksheet.center_horizontally() and worksheet.set_row(0, 97.5)
# =============================================================================

# In[11]:


import pandas as pd
import random
import datetime as dt  
import win32com.client as win32


path = ('C://Users/U037679/Documents/AnacondaProjects/conference-call-bingo/')

infile =  (path + 'BINGO_cc.xlsx')
outfile = (path + 'BINGO_latest.xlsx')

today = dt.datetime.today().strftime("%Y%m%d")
today_sheet_name = dt.datetime.today().strftime("%m-%d-%Y")

df = pd.read_excel(infile, sheet_name='list', )


# In[12]:


# print(df['Quotes'])


# In[13]:


random.shuffle(df['Quotes'])
# print (df['Quotes'])


# In[14]:


# Create a grid from the list

col0 = df['Quotes'][0:6]
col1 = df['Quotes'][6:11]
col2 = df['Quotes'][11:16]
col3 = df['Quotes'][16:21]
col4 = df['Quotes'][21:26]

df_bingo = pd.DataFrame(list(zip(col0, col1, col2, col3, col4)), columns = 
                  ['col0', 'col1', 'col2', 'col3', 'col4'])

# df_bingo


# In[15]:


# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(outfile, engine='xlsxwriter')
df_bingo.to_excel(writer, sheet_name = today_sheet_name, index=False, header=False)


# In[16]:


# Get the XlsxWriter objects from the dataframe writer object.
workbook = writer.book
worksheet = writer.sheets[today_sheet_name]


# In[17]:


# Format this biatch
cell_format = workbook.add_format()

cell_format.set_text_wrap()
cell_format.set_font_size(18)
cell_format.set_align('center')
cell_format.set_align('vcenter')
cell_format.set_border(1)

worksheet.set_row(0, 97.5) # test
worksheet.set_row(1, 97.5)
worksheet.set_row(2, 97.5)
worksheet.set_row(3, 97.5)
worksheet.set_row(4, 97.5)
worksheet.set_row(5, 97.5)

worksheet.set_page_view() # There's currently no way to remove the ruler from pageview
worksheet.hide_gridlines(2) # 0 = don't hide gridlines, 1 = hide printed gridlines only, 2 = hide screen and printed gridlines
worksheet.set_landscape()
worksheet.center_horizontally()
worksheet.set_margins(left=0.5, right=0.5, top=1.0, bottom=0.4)
worksheet.set_column('A:E', 23.0, cell_format) 
worksheet.print_area('A1:E5')

header = '&C&18&"Calibri,Bold"Conference Call Bingo!' + ' \n&12&A'
worksheet.set_header(header)


# In[18]:


writer.save()
writer.close()


# ### Mail this bad boy

# In[19]:


# Outlook method:

v = open(path + r'\email_list_just_me.txt', 'r')

# full list of co-workers, etc.:
# v = open(path + r'\email_list.txt', 'r')

contents = v.read()
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
nameList = contents.split('\n')
for name in nameList:
    if name != '':
        mail.Recipients.Add(name)

mail.Subject = "Here's the latest conference call bingo card"
mail.HTMLBody =  '<h3>The latest cards just arrived from HQ!</h3>' + \
                 'Attached is your card for the week.' + \
                 '<br><br>This is not an exhaustive list. There are ' + \
                 'more items on the list than can fit on one card, and ' + \
                 'the list keeps growing! So if you have suggestions, ' + \
                 'please forward them.' + \
                 '<br><br>Thanks!' + \
                 '<br><br>(This is an automated message.)'

attachment  = outfile
mail.Attachments.Add(attachment)

v.close()
mail.Send()


# In[38]:


# Outlook method 2
# https://stackoverflow.com/questions/6332577/send-outlook-email-via-python

# import win32com.client as win32
# outlook = win32.Dispatch('outlook.application')
# mail = outlook.CreateItem(0)
# mail.To = 'charles.culp@cvshealth.com'
# mail.Subject = 'Conference call bingo card - TEST2'
# # mail.Body = 'Attached is the card for the week. (This is an outomated message.)'
# mail.HTMLBody = '<h3>Conference Call Bingo!</h3>' + \
#                 'Attached is the card for the week.' + \
#                 '<br>(This is an outomated message.)'

# # To attach a file to the email (optional):
# # attachment  = "Path to the attachment"
# attachment  = outfile
# mail.Attachments.Add(attachment)

# mail.Send()

