from ast import In, Str
from operator import inv, truediv
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
import altair as alt
from datetime import date
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
import datetime
from datetime import time
import calendar
from dateutil import parser
from pandas.tseries.offsets import BDay
from dateutil import parser

############ CSS Format ######################

with open('streamlit.css') as modi:
    css = f'<style>{modi.read()}</style>'
    st.markdown(css, unsafe_allow_html=True)
###############################################

def format_dataframe_columns(df):
    formatted_df = df.copy()  # Create a copy of the DataFrame
    for column in formatted_df.columns:
        if formatted_df[column].dtype == 'float64':  # Check if column has a numeric type
            formatted_df[column] = formatted_df[column].apply(lambda x: '{:,.2f}'.format(x))
    return formatted_df
#########################################################
def formatted_display(label, value, unit):
    formatted_value = "<span style='color:green'>{:,.2f}</span>".format(value)  # Format value with comma separator and apply green color
    display_text = f"{formatted_value} {unit}"  # Combine formatted value and unit
    st.write(label, display_text, unsafe_allow_html=True)
st.image(logo_image, width=700)
st.header('ONE-SIM Sales Report 2023')
##################
db=pd.read_excel('Database-2022.xlsx')
#################
@st.cache(allow_output_mutation=True)
def load_data_from_drive():
    url="https://docs.google.com/spreadsheets/d/1sUH0WmtfrWbR8FrM33ljebSlW0BWE3w0/export?format=xlsx"
    data=pd.read_excel(url,header=0)
    return data
data = load_data_from_drive()
Invoices=data
Invoices['วันที่']=Invoices['วันที่'].astype(str)
################# T1 and T2 #############################
T1PlusT2=Invoices[Invoices['รหัสสินค้า'].astype(str).str.contains('MOLD')]
T1PlusT2=T1PlusT2[['วันที่','รหัสสินค้า','เลขที่','JOBCODE','มูลค่าสินค้า']]

##################################### Invoice No T0 ######################################################

Invoices[['วันที่','เลขที่','ลูกค้า','ชื่อสินค้า','JOBCODE']]=Invoices[['วันที่','เลขที่','ลูกค้า','ชื่อสินค้า','JOBCODE']].astype(str)
Invoices=Invoices[(~Invoices['รหัสสินค้า'].astype(str).str.contains('MOLD-T0'))]
# Invoices=Invoices[(~Invoices['รหัสสินค้า'].astype(str).str.contains('MOLD-T1'))]
Inv=Invoices[['วันที่','เลขที่','ลูกค้า','ชื่อสินค้า','จำนวน','มูลค่าสินค้า','รหัสสินค้า']]
Inv=Invoices[Invoices['เลขที่'].str.contains('IV')|Invoices['เลขที่'].str.contains('HS')]

###################### Total Invoice Check #################################################
Inv['มูลค่าสินค้า']= pd.to_numeric(Inv['มูลค่าสินค้า'], errors='coerce')
# Inv['มูลค่าสินค้า'].dropna(0, inplace=True)
	@@ -118,10 +121,44 @@ def load_data_from_drive():
st.write("Last update:", last_update)
st.write('Invoice Isuued Days:',COUNT)
###################### MASS ################################
TotalMASS=Invoices[Invoices['ลูกค้า'].str.contains('VALEO')|Invoices['ลูกค้า'].str.contains('แครทโค')|Invoices['ชื่อสินค้า'].str.contains('PACKING')|
Invoices['ลูกค้า'].str.contains('เซนทรัล')]
TotalMASS=TotalMASS[TotalMASS['วันที่'].between( ym_input, ym_input2)]
TotalMASS=pd.merge(TotalMASS,db[['Part_No','Mold-PM','Mold-DP']],left_on='รหัสสินค้า',right_on='Part_No',how='left')
####################### Steel Bush ################################
TotalSTB=Invoices[Invoices['ชื่อสินค้า'].str.contains('STEEL')]
TotalSTB = TotalSTB[TotalSTB['วันที่'].between( ym_input, ym_input2)]
	@@ -158,7 +195,7 @@ def format_dataframe_columns(df):
# Merge T1PlusT2 with the modified JOBCODE column
T1PlusT2 = pd.merge(T1PlusT2, ONESIM['JOBCODE'], on='JOBCODE', how='right')
# T1PlusT2=pd.merge(T1PlusT2,ONESIM['JOBCODE'],on='JOBCODE',how='right')
T1PlusT2=T1PlusT2[['รหัสสินค้า','JOBCODE','มูลค่าสินค้า']]
T1PlusT2=T1PlusT2.fillna(0)
T1PlusT2 = T1PlusT2.drop_duplicates()
T1PlusT2=T1PlusT2[T1PlusT2['มูลค่าสินค้า']!=0]
	@@ -172,32 +209,7 @@ def format_dataframe_columns(df):
    formatted_df = format_dataframe_columns(T1PlusT2)
    st.dataframe(formatted_df)
    ####################

formatted_display('Total Mold Deposit:',round(MoldT2,2),'B')
############# Display ##############
# st.write('---')
# #########################################################
# st.write('**ONE-SIM Sales Summarize**')
# TotalMoldPM=(TotalMASS['จำนวน']*TotalMASS['Mold-PM']).sum()
# TotalMoldDP=(TotalMASS['จำนวน']*TotalMASS['Mold-DP']).sum()
# TotalSaleCASH=SalesCash['มูลค่าสินค้า'].sum()
# TotalSalesMASS=(TotalMASS['มูลค่าสินค้า'].sum())-(TotalMoldPM+TotalMoldDP)
# TotalSaleSTB=TotalSTB['มูลค่าสินค้า'].sum()
# TotalSalesOTHER=TotalOTHER['มูลค่าสินค้า'].sum()
# TotalSalesMOLD=(TotalMOLD['มูลค่าสินค้า'].sum())+(MoldT2+TotalMoldPM+TotalMoldDP)
# TotalSales=(TotalSaleCASH+TotalSalesMASS+TotalSaleSTB+TotalSalesOTHER+TotalSalesMOLD)
# formatted_display('Total MASS BU Sales:',round(TotalSalesMASS,2),'B')
# formatted_display('Total Steel Bush Sales:',round(TotalSaleSTB,2),'B')
# formatted_display('Total Mold BU Sales:',round(TotalSalesMOLD,2),'B')
# formatted_display('Total Mold Deposit:',round(MoldT2,2),'B')
# formatted_display('Total Mold PM Internal Charged:',round(TotalMoldPM,2),'B')
# formatted_display('Total Mold DP Income:',round(TotalMoldDP,2),'B')
# formatted_display('Total Other Sales:',round(TotalSalesOTHER,2),'B')
# formatted_display('Total Cash:',round(TotalSaleCASH,2),'B')
# formatted_display('Total One-SIM Sales:',round(TotalSales+TotalSaleCASH,2),'B')
# DATASALES=[['One-SIM',TotalSales],['MASS',TotalSalesMASS],['Steel Bush',TotalSaleSTB],['Mold',TotalSalesMOLD],['Other',TotalSalesOTHER]]
# SUMSALES=pd.DataFrame(DATASALES,columns=['Items','AMT'])
# SUMSALES.set_index('Items',inplace=True)
# ############# Target ######################################################
# specify the start and end dates for the date range
start_date = ym_input
	@@ -208,7 +220,7 @@ def format_dataframe_columns(df):
business_days = date_range[date_range.weekday < 5]
# print the resulting number of business days
Days = len(business_days)-4
Target2023=pd.read_excel('Target-2023.xlsx')
Target2023=Target2023[Minput]
Target2023=(Target2023/Days)*COUNT
Target2023=list(Target2023)
	@@ -294,11 +306,38 @@ def format_dataframe_columns(df):
st.write('---')
##### SUMMARIZE SALRES ##################################
#########################################################
st.write('**ONE-SIM Sales Summarize**')
TotalMoldPM=(TotalMASS['จำนวน']*TotalMASS['Mold-PM']).sum()
TotalMoldDP=(TotalMASS['จำนวน']*TotalMASS['Mold-DP']).sum()
TotalSaleCASH=SalesCash['มูลค่าสินค้า'].sum()
TotalSalesMASS=((TotalMASS['มูลค่าสินค้า'].sum())-(TotalMoldPM+TotalMoldDP))+(MASSCNDNBL)
TotalSaleSTB=TotalSTB['มูลค่าสินค้า'].sum()
TotalSalesOTHER=TotalOTHER['มูลค่าสินค้า'].sum()
TotalSalesMOLD=((TotalMOLD['มูลค่าสินค้า'].sum())+(MoldT2+TotalMoldPM+TotalMoldDP))+(MOLDCNDNBL)
	@@ -323,8 +362,8 @@ def format_dataframe_columns(df):
num_months = (end_date.year - start_date.year) * 12 + end_date.month - start_date.month + 1

# Define the data for the bar chart
categories = ['One-SIM','MASS','Mold','Mold-T0','Steel Bush','Cash','Other']
values = [TotalSales+TotalSaleCASH, TotalSalesMASS,(TotalSalesMOLD-MoldT2),MoldT2,TotalSaleSTB,TotalSaleCASH,TotalSalesOTHER]
values2 =Target2023
# Use num_months as the monthly factor to multiply the values in values and values2
monthly_factor = num_months
	@@ -357,61 +396,124 @@ def format_dataframe_columns(df):

if st.button("Refresh data"):
    data = load_data_from_drive()
st.write('---')
####################################### ChecKing Fucntions #####################
st.write('---')
def format_dataframe_columns(df):
    # Format specific columns
    df['มูลค่าสินค้า'] = df['มูลค่าสินค้า'].apply(lambda x: '{:,.2f}'.format(x))
    return df
###########################
c1, c2 = st.columns(2)
with c1:

    st.write('**Checking MASS-BU Sales by AMT and Pcs**')
    # Get the user input for the 4-digit Part No
    PartNo = st.text_input('Input 4-digit Part No')

    # Find the matching 9-digit Part No in the DataFrame
    if len(PartNo) == 4:
        PartMASS = Invoices[['วันที่', 'รหัสสินค้า', 'จำนวน', 'มูลค่าสินค้า']]
        PartMASS = PartMASS[PartMASS['วันที่'].between( ym_input, ym_input2)]

        # Remove missing values from the 'รหัสสินค้า' column
        PartMASS = PartMASS.dropna(subset=['รหัสสินค้า'])

        # Find the matching rows using str.contains and the boolean mask
        mask = PartMASS['รหัสสินค้า'].str.contains(PartNo)
        matching_rows = PartMASS[mask]
        matching_rows=matching_rows.set_index('วันที่')
        matching_rows.index = pd.to_datetime(matching_rows.index).strftime('%Y-%m-%d')
        TTPCS=matching_rows['จำนวน'].sum()
        TTB=matching_rows['มูลค่าสินค้า'].sum()
        if len(matching_rows) > 0:
            ###################
            formatted_df = format_dataframe_columns(matching_rows)
            st.dataframe(formatted_df)
            ####################
            formatted_display('Total Pcs:',round(TTPCS,2),'Pcs')
            formatted_display('Total Sales:',round(TTB,2),'Pcs')
        else:
            st.write(f'No matching Part No found for "{PartNo}"')
    else:
        st.write('Please input a 4-digit Part No')
############## Check Product ############################################
with c2:
    st.write('**Checking Mold Sales by Product**')
    # Get the user input for the 4-digit Part No
    Product = st.selectbox('Select Product Type', ['MOLD','PART','REPAIR'])
    PartMold = Invoices[['วันที่', 'รหัสสินค้า','JOBCODE', 'มูลค่าสินค้า']]
    PartMold = PartMold[PartMold['วันที่'].between( ym_input, ym_input2)]
    PartMold=PartMold[PartMold['รหัสสินค้า'].str.contains(Product)]
    PartMold = PartMold.set_index('วันที่')
    PartMold.index = pd.to_datetime(PartMold.index).strftime('%Y-%m-%d')
    PartMold=PartMold[PartMold['มูลค่าสินค้า']>0]
    TTMold=PartMold['มูลค่าสินค้า'].sum()
    ###################
    formatted_df = format_dataframe_columns(PartMold)
    st.dataframe(formatted_df)
    ####################
    formatted_display('Total Mold Sales:',round(TTMold,2),'B')
