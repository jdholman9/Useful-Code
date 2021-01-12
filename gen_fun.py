
# Date packages
from datetime import datetime
from datetime import timedelta
# Data handling packages
import pandas as pd
import numpy as np

import re, os
import pyperclip
# Packages for output to excel
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl
from openpyxl.styles import Color, PatternFill, Font, Border, Alignment, numbers

# Email stuff needed
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path


def trim_str(x_str):
    # 2 or more spaces to 1
    # All leading or trailing spaces removed
    trm_str = re.sub(' +', ' ', x_str)
    trm_str = re.sub('^ +', '', trm_str)
    trm_str = re.sub(' +$', '', trm_str)
    return trm_str


def add_grey_line_TF(df, column):
    # Adds bool column alternating based on given column values
    # If value above in given column is the same then output will be the same bool value (T/F)
    # If not then output will be the opposite
    # First value is True
    
    # Tip: Data frame should be in disired order before using this function
    # Example usage: True false column can be used on excel for formatting
    df.reset_index(drop=True, inplace=True)
    grey_line = [True]
    for i in range(1, len(df[column])):
        if df[column][i] == df[column][i-1]:
            grey_line.append(grey_line[i-1])
        else:
            grey_line.append(not grey_line[i-1])
    
    df.insert(0, 'grey_line', grey_line)
    
    return df


def time_coverage(times, tm_delta_min = 60):
    # Given list of times
    # output will be string showing covered times e.g. "12:13-15:10, 17:01-17:34"
    # times in order are close enough defined by tm_delta_min (number of minutes)
    pc_tm = np.unique(times)
    pc_tm = sorted(pd.to_datetime(pc_tm))
    
    tm_prd = [pc_tm[0].strftime('%H:%M')]
    if len(pc_tm) > 1:
        for i in range(1, len(pc_tm)-1):
            if (pc_tm[i] - pc_tm[i-1]) > timedelta(minutes = tm_delta_min):
                if i >= 2:
                    if (pc_tm[i-1] - pc_tm[i-2]) > timedelta(minutes = tm_delta_min):
                        tm_prd.append(', ')
                        tm_prd.append(pc_tm[i].strftime('%H:%M'))
                    else:
                        tm_prd.append('-')
                        tm_prd.append(pc_tm[i-1].strftime('%H:%M'))
                        tm_prd.append(', ')
                        tm_prd.append(pc_tm[i].strftime('%H:%M'))
                else:
                    tm_prd.append(', ')
                    tm_prd.append(pc_tm[i].strftime('%H:%M'))
                        
        
        tm_prd.append('-')
        tm_prd.append(pc_tm[len(pc_tm)-1].strftime('%H:%M'))    
    
    return ''.join(tm_prd) 


def comma_unique(items):
    # Input is pandas series of strings
    # Output is comma separated unique values in order
    items_unq = np.unique(items)
    items_unq = [unq_itm for unq_itm in items_unq if str(unq_itm) not in ['nan', 'None', '']]
    return ', '.join(items_unq)


# Send email
def send_email(email_sender, email_recipients, email_subject, email_message, 
               attachment_location = ''):
    #email_sender = 'Jacob Holman <jholman@sjrtd.com>'

    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = ','.join(email_recipients)
    msg['Subject'] = email_subject

    msg.attach(MIMEText(email_message, 'plain'))
    
    if attachment_location != '':
        filename = os.path.basename(attachment_location)
        attachment = open(attachment_location, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        msg.attach(part)
    
    try:
        server = smtplib.SMTP('erpsmtp')
        server.ehlo()
        text = msg.as_string()
        server.sendmail(email_sender, email_recipients, text)
        print('email sent')
        server.quit()
    except:
        input("SMPT server connection error " + chk_dt)
    return True


def format_xl(wb, sheet_nm, df, date_col, pcnt_col, cols_to_center, col_head_left, col_hide):
    ws = wb[sheet_nm]
    # Fill in Report sheet
    for r in dataframe_to_rows(df, index = False, header = True):
        ws.append(r)
    
    # Format headers
    for cell in ws[1]:
        cell.fill = PatternFill('solid', fgColor = '4F81BD')
        cell.font = Font(bold = True, size = 12, color = 'FFFFFF')
    
    # Format column as date
    for cell in ws[date_col]:
        cell.number_format = 'MM/DD/YYYY'
    
    # Format columns as percentage
    for col in pcnt_col:
        for cell in ws[col]:
            cell.number_format = numbers.FORMAT_PERCENTAGE
    
    # Center all values in specified columns
    for col in cols_to_center:
        for cell in ws[col]:
            cell.alignment = Alignment(horizontal = 'center')
    
    # Left align some column headers
    for col in col_head_left:
        ws[col][0].alignment = Alignment(horizontal = 'left')
    
    # Hide Columns
    for col in col_hide:
        ws.column_dimensions[col].hidden = True
    # Add Filter
    ws.auto_filter.ref = ws.dimensions
    
    return wb



