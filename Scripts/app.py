import streamlit as st
import pandas as pd
#import polars as pl
import numpy as np
import openpyxl
import warnings
warnings.filterwarnings("ignore")
#from functools import reduce

st.set_page_config('COS_AAP', page_icon="🏛️", layout='wide')
def title(url):
     st.markdown(f'<p style="color:#2f0d86;font-size:22px;border-radius:2%;"><br><br><br>{url}</p>', unsafe_allow_html=True)
def title_main(url):
     st.markdown(f'<h1 style="color:#230c6e;font-size:42px;border-radius:2%;"><br>{url}</h1>', unsafe_allow_html=True)

def success_df(html_str):
    html_str = f"""
        <p style='background-color:#baffc9;
        color: #313131;
        font-size: 15px;
        border-radius:5px;
        padding-left: 12px;
        padding-top: 10px;
        padding-bottom: 12px;
        line-height: 18px;
        border-color: #03396c;
        text-align: left;'>
        {html_str}</style>
        <br></p>"""
    st.markdown(html_str, unsafe_allow_html=True)

sidebar = st.sidebar
with sidebar:
    #st.title("FIS2 Container Status")
    title("DMS Inventory")
    st.write('## Current Status')
    status_inventory = st.radio(
        label='Select one',
        options=['AA', 'AV', 'AAP'],
        index=0
    )

def format_dataframe(df):
    df['DMS Gate in'] = df['DMS Gate in'].str.replace('00:00:00', '').str.strip()
    df['DMS Gate in'] = df['DMS Gate in' ].str.replace('.0', '').str.strip()
    df['2nd EOR'] = df['2nd EOR'].str.replace('\n', '')
    return df

def gate_out_status(df):
    df['DMS_GATE_OUT_STATUS'] = df['Container'].apply(
        lambda x: 'No GATEOUT' if x not in movementOut['Container'].tolist() else 'GATEOUT')
    return df

def replace_nan_with_dash(df, column_name):
    """
    Replaces NaN values in the specified column with a hyphen ('-').
    Args:
        df (pd.DataFrame): Input DataFrame.
        column_name (str): Name of the column to process.
    Returns:
        pd.DataFrame: DataFrame with NaN values replaced by '-' in the specified column.
    """
    df[column_name].fillna('-', inplace=True)
    return df

def sort_dataframe_col(df, col_name):
    df = df.sort_values(by=col_name)
    return df

def compare_intersect(x, y):
    return bool((len(frozenset(x).intersection(y))==len(x))) #compare 2 lists

title_main('DMS Inventory Status')

uploaded_file = st.file_uploader("Upload COS-AAP-GEN", type=['xlsx'])
if uploaded_file is None:
     st.write('Please upload a file')
elif uploaded_file:
    list_ws = ['DMS INVENTORY', 'REPAIR ESTIMATE', 'MOVMENT OUT', 'AUTH', 'FORMULA AAP']

    xl_uploaded = pd.ExcelFile(uploaded_file)
    lst_ws_uploaded = xl_uploaded.sheet_names
    if not compare_intersect(list_ws, lst_ws_uploaded):
        st.write('Please upload file with correct worksheets')
    elif compare_intersect(list_ws, lst_ws_uploaded):


        cols_dmsInventory = ['Container No.', 'Customer', 'Current Status', 'Rating']
        cols_repairEstimate = ['Container No', 'Customer', 'Surveyor Name', 'Total']
        cols_movementOut = ['Container No.', 'Customer', 'Status', 'Rating']
        cols_auth = ['Status', 'Eqpno', 'Approvaldate', 'approvalamount', 'Purpose', 'Remark']

        dmsInventory = pd.read_excel(uploaded_file, sheet_name=list_ws[0], engine='openpyxl')
        repairEstimate = pd.read_excel(uploaded_file, sheet_name=list_ws[1], engine='openpyxl')
        movementOut = pd.read_excel(uploaded_file, sheet_name=list_ws[2], engine='openpyxl')
        auth = pd.read_excel(uploaded_file, sheet_name=list_ws[3], engine='openpyxl')

        dmsInventory = dmsInventory[cols_dmsInventory]
        repairEstimate = repairEstimate[cols_repairEstimate]
        movementOut = movementOut[cols_movementOut]
        auth = auth[cols_auth]
        dmsInventory.rename(columns={'Container No.':'Container', 'Current Status': 'Container Current Status_Inventory'}, inplace=True)
        repairEstimate.rename(columns={'Container No':'Container', 'Total': 'DMS Repair Price','Surveyor Name': 'Repairer_Vendor'}, inplace=True)
        movementOut.rename(columns={'Container No.':'Container'}, inplace=True)
        auth.rename(columns={'Eqpno':'Container', 'Status': 'REPAIR COMPLETE STATUS-COSCO', 'Approvaldate': 'DMS Gate in', 'approvalamount': 'COSCO AUTH PRICE', 
                        'Purpose': 'REPAIR TYPE', 'Remark': '2nd EOR'}, inplace=True)
    
        mask_inventory = (dmsInventory['Customer'].isin(['COS'])) & (dmsInventory['Container Current Status_Inventory'].isin([status_inventory])) #2
        #mask_movement = (movementOut['Customer'].isin(['COS'])) & (movementOut['Status'].isin(['AV']))
        mask_repair = (repairEstimate['Customer'].isin(['COSCO SHIPPING LINES (SINGAPORE) PTE LTD'])) & (repairEstimate['Repairer_Vendor'].isin(['MD GLOBAL', 'Eastern Repairer'
                                                                                                                                  'MD Bala', 'MD Imran', 'MD Rasel']))
        assert mask_inventory.any() # sanity check that the mask is selecting something
        assert mask_repair.any()
        inventory_av = dmsInventory[mask_inventory]
        repairEstimate_md = repairEstimate[mask_repair]

        #replace 'COSCO SHIPPING LINES (SINGAPORE) PTE LTD' with 'COS'
        repairEstimate_md['Customer'] = 'COS'
        inventory_av['Customer'] = 'COS'

        InventoryRepair = pd.merge(inventory_av, repairEstimate_md, on='Container')
        InventoryRepair.drop(['Customer_y'], axis=1, inplace=True) #drop column 'Customer_y'

        InventoryRepairAuth_0 = pd.merge(InventoryRepair, auth, on='Container', how='inner')

        InventoryRepairAuth_0 = format_dataframe(InventoryRepairAuth_0)
        InventoryRepairAuth_0 = gate_out_status(InventoryRepairAuth_0) #Call the function
        replace_nan_with_dash(InventoryRepairAuth_0, 'Rating')
        InventoryRepairAuth = sort_dataframe_col(InventoryRepairAuth_0, 'Container') #sort col

        col_aap = ['Container', 'Container Current Status_Inventory', 'DMS Repair Price', 'COSCO AUTH PRICE', 'Repairer_Vendor', 
                'DMS_GATE_OUT_STATUS', 'REPAIR COMPLETE STATUS-COSCO', 'REPAIR TYPE', '2nd EOR', 'DMS Gate in', 'Rating']
        InventoryRepairAuth = InventoryRepairAuth[col_aap]

st.write('Click to get updated status')
if st.button("Get dataframe"):
    with st.spinner("Processing..."):
        st.dataframe(InventoryRepairAuth.reset_index(drop=True), use_container_width=True)
        success_df("Dataframe is ready 💸")

import smtplib, email, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import glob, re, os
from datetime import datetime
import timedelta

st.divider()
st.write("Please download csv before send email")
#password = "usec qyjx xfcd syhw"

def extract_file_name(s):
    base_name = os.path.basename(s)  # Get the base name
    file_name, _ = os.path.splitext(base_name)  # Split the extension
    return file_name+".csv"

def replace_date_with_status(file_name):
    # Use regex to find the date pattern and replace it with status_inventory
    new_file_name = re.sub(r'T\d{2}-\d{2}\_export', '_'+status_inventory, file_name)
    return new_file_name

email_receiver = st.text_input('To your email')
if st.button("Send email"):
    #email_sender = "sxk2929@gmail.com"
    email_sender = "john.tan@sh-cogent.com.sg"
    #subject = status_inventory+ " DMS Inventory Status"
    #body = "Updated status. This message is computer generated. "+ (datetime.today()+ timedelta(hours=9)).strftime("%Y%m%d %H:%M:%S")
    body = """
        <html>
        <head>
        <title>Dear User</title>
        </head>
        <body>
        <p style="color: blue;font-size:25px;">DMS Inventory updated.</strong><br></p>

        </body>
        </html>

        """+ InventoryRepairAuth.reset_index(drop=True).to_html() +"""
        <br>This message is computer generated. """+ (datetime.now()+ timedelta(hours=8)).strftime("%Y%m%d %H:%M:%S")
    
    password = st.secrets["password"]
    
    mailserver = smtplib.SMTP('smtp.office365.com',587)
    mailserver.ehlo()
    mailserver.starttls()
    mailserver.login(email_sender, password)
    
    try:
        if email_receiver is not None:
            try:
                rgx = r'^([^@]+)@[^@]+$'
                matchObj = re.search(rgx, email_receiver)
                if not matchObj is None:
                    usr = matchObj.group(1)
                    
            except:
                pass
            list_csv = glob.glob("C:/Users/"+usr+"/Downloads/*.csv")
            #list_csv = glob.glob("C:/Users/john.tan/Downloads/*.csv")
            #latest_csv = max(list_csv, key=os.path.getctime)
                        

        msg = MIMEMultipart()
        msg['From'] = email_sender
        msg['To'] = email_receiver
        msg['Subject'] = 'DMS Inventory Status ' +(datetime.today()+ timedelta(hours=9)).strftime("%Y%m%d %H:%M:%S")

        msg.attach(MIMEText(body, 'html'))
        #filename = latest_csv
        #filename = "C:/Users/john.tan/Downloads/2024-03-30_AV.csv"

        #with open(filename, 'rb') as attachment:
            #part = MIMEBase("application", "octet-stream")
            #part.set_payload(attachment.read())

        #file_name=extract_file_name(latest_csv)
        #fmt_file_name = (replace_date_with_status(file_name))
        fmt_file_name = "2024-03-30_AV.csv"
        #encoders.encode_base64(part)
        #part.add_header(
            #"Content-Disposition",
            #f"attachment; filename= {fmt_file_name}",
        #)

        #msg.attach(part)
        text = msg.as_string() + InventoryRepairAuth.reset_index(drop=True).to_html()

        #context = ssl.create_default_context() #login to secure server, 465 for ssl. smtplib.SMTP('smtp.office365.com',587)
        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.ehlo()
            server.starttls()
            server.login(email_sender, password)
            server.sendmail(email_sender, email_receiver, text)
            server.quit()

        st.success("Email sent successfully 💌 🚀")
    except Exception as e:
        st.error(f"Email not sent: {e}")

st.divider()
footer_html = """
    <div class="footer">
    <style>
        .footer {
            position: fixed;
            bottom: 0;
            left: 0;
            right: 0;
            background-color: #f0f2f6;
            padding: 10px 20px;
            text-align: center;
        }
        .footer a {
            color: #4a4a4a;
            text-decoration: none;
        }
        .footer a:hover {
            color: #3d3d3d;
            text-decoration: underline;
        }
    </style>
        All rights reserved @2024. Cogent Holdings IT Solutions.      
    </div>
"""
st.markdown(footer_html,unsafe_allow_html=True)

#https://realpython.com/python-send-email/#adding-attachments-using-the-email-package
#https://github.com/tonykipkemboi/streamlit-smtp-test/blob/main/streamlit_app.py
