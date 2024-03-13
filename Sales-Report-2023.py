from ast import In, Str
# from operator import inv, truediv
import streamlit as st
import pandas as pd
import numpy as np
# import matplotlib.pyplot as plt
from PIL import Image
import altair as alt
from datetime import date
# import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
import datetime
from datetime import time
import calendar
from dateutil import parser
from pandas.tseries.offsets import BDay
from dateutil import parser
import plotly.graph_objs as go
################# BG #################################
# page_bg_img = f"""
# <style>
# [data-testid="stAppViewContainer"] > .main {{
# background-image: url("https://i.postimg.cc/4xgNnkfX/Untitled-design.png");
# background-size: cover;
# background-position: center center;
# background-repeat: no-repeat;
# background-attachment: local;
# }}
# [data-testid="stHeader"] {{
# background: rgba(1,2,3,4);
# }}
# </style>
# """

# st.markdown(page_bg_img, unsafe_allow_html=True)
# ############ CSS Format / Style ######################
# # with open('streamlit.css') as modi:
# #     css = f'<style>{modi.read()}</style>'
# #     st.markdown(css, unsafe_allow_html=True)
###############################################
def format_dataframe_columns(df):
    formatted_df = df.copy()  # Create a copy of the DataFrame
    for column in formatted_df.columns:
        if formatted_df[column].dtype == 'float64':  # Check if column has a numeric type
            formatted_df[column] = formatted_df[column].apply(lambda x: '{:,.2f}'.format(x))
    return formatted_df
#########################################################
def formatted_display(label, value, unit):
    formatted_value = "<span style='color:yellow'>{:,.2f}</span>".format(value)  # Format value with comma separator and apply green color
    display_text = f"{formatted_value} {unit}"  # Combine formatted value and unit
    st.write(label, display_text, unsafe_allow_html=True)
#######################################################################################
############ Logo ####################
logo_image = Image.open('SIM-022.jpg')
st.image(logo_image, width=700)
st.header('SIM Sales Report 2024')
################## Reas File ################
db=pd.read_excel('Database-2022.xlsx')
#################
MoldDP=pd.read_excel('Mold DP-2023.xlsx')
#################
Mold_PM=pd.read_excel('Mold-PM-List.xlsx')
############### 2024 #####################
@st.cache_data 
def load_data_from_drive():
    url="https://docs.google.com/spreadsheets/d/13GtjhI6mQJ055bNG5lxNYzDwlor5027n/export?format=xlsx"
    data2024=pd.read_excel(url,header=4)
    return data2024
data2024 = load_data_from_drive()
Invoices=data2024
# Invoices
############### Mold DP #####################
@st.cache_data 
def load_data_from_drive():
    Moldurl="https://docs.google.com/spreadsheets/d/1fA2OEz8pnLYUzDUFUOGy9ylL_X4RNd_L/export?format=xlsx"
    dataMold=pd.read_excel(Moldurl)
    return dataMold
dataMold = load_data_from_drive()
Mold_DP=dataMold
# Mold_DP
########### Menu Range ####################
y_map = {
    'Jan': '2024-01-01', 'Feb': '2024-02-01', 'Mar': '2024-03-01', 'Apr': '2024-04-01', 'May': '2024-05-01', 'Jun': '2024-06-01',
    'Jul': '2024-07-01', 'Aug': '2024-08-01', 'Sep': '2024-09-01', 'Oct': '2024-10-01', 'Nov': '2024-11-01', 'Dec': '2024-12-01'
}
y_map_range = {
    'Jan': '2024-01-31', 'Feb': '2024-02-29', 'Mar': '2024-03-31', 'Apr': '2024-04-30', 'May': '2024-05-31', 'Jun': '2024-06-30',
    'Jul': '2024-07-31', 'Aug': '2024-08-31', 'Sep': '2024-09-30', 'Oct': '2024-10-31', 'Nov': '2024-11-30', 'Dec': '2024-12-31'
}

# Streamlit sidebar for selecting start and end months
start_month = st.sidebar.selectbox('Select start month', list(y_map.keys()), index=0)
end_month = st.sidebar.selectbox('Select end month', list(y_map_range.keys()), index=0)

# Convert selected months to datetime objects
start_date = pd.to_datetime(y_map[start_month], errors='coerce')
end_date = pd.to_datetime(y_map_range[end_month], errors='coerce')

# Ensure 'วันที่' column is in datetime format
Invoices['วันที่'] = pd.to_datetime(Invoices['วันที่'], errors='coerce')

# Filter the DataFrame based on the date range
filtered = Invoices[
    (Invoices['วันที่'] >= start_date) &
    (Invoices['วันที่'] <= end_date)]
#########
############ BU Menu #####################################################
BU = st.sidebar.selectbox('Select BU',['MASS','Mold','One-SIM'] )
####################### Mass Info #########################################

TotalMASS = filtered[
    (Invoices['ลูกค้า'].str.contains('VALEO') |
    Invoices['ลูกค้า'].str.contains('แครทโค') |
    Invoices['ลูกค้า'].str.contains('อีเลคโทรลักซ์ ') |
    Invoices['ชื่อสินค้า'].str.contains('PACKING') |
    Invoices['ลูกค้า'].str.contains('เซนทรัล') |
    Invoices['ลูกค้า'].str.contains('โฮมเอ็ก') |
    Invoices['ลูกค้า'].str.contains('ศิริโกมล') |
    Invoices['ลูกค้า'].str.contains('โคชิน') |
    Invoices['รหัสสินค้า'].str.contains('SB')|
    Invoices['รหัสสินค้า'].str.contains('DENSE')) &
    (~Invoices['รหัสสินค้า'].astype(str).str.contains('M2') &
     ~Invoices['รหัสสินค้า'].astype(str).str.contains('MOLD') &
    ~Invoices['รหัสสินค้า'].astype(str).str.contains('SIM-P') &
    ~Invoices['เลขที่'].astype(str).str.contains('DR') &
    ~Invoices['เลขที่'].astype(str).str.contains('SR') &
    ~Invoices['เลขที่'].astype(str).str.contains('HS') &
    ~Invoices['รหัสสินค้า'].astype(str).str.contains('SIM-R'))
]
# TotalMASS
############################
MASSDNCN = filtered[
    (Invoices['ลูกค้า'].str.contains('VALEO') |
    Invoices['ลูกค้า'].str.contains('แครทโค') |
    Invoices['ลูกค้า'].str.contains('อีเลคโทรลักซ์ ') |
    Invoices['ชื่อสินค้า'].str.contains('PACKING') |
    Invoices['ลูกค้า'].str.contains('เซนทรัล') |
    Invoices['ลูกค้า'].str.contains('โฮมเอ็ก') |
    Invoices['ลูกค้า'].str.contains('ศิริ') |
    Invoices['ลูกค้า'].str.contains('โคชิน') |
    Invoices['ลูกค้า'].str.contains('ชนะชัย') |
    Invoices['รหัสสินค้า'].str.contains('SB')|
    Invoices['เลขที่'].astype(str).str.contains('DR') |
    Invoices['เลขที่'].astype(str).str.contains('SR') |
    Invoices['เลขที่'].astype(str).str.contains('HS') |
    Invoices['รหัสสินค้า'].str.contains('DENSE')) &
    (~Invoices['รหัสสินค้า'].astype(str).str.contains('M2') &
     ~Invoices['รหัสสินค้า'].astype(str).str.contains('MOLD') &
    ~Invoices['รหัสสินค้า'].astype(str).str.contains('SIM-P') &
    ~Invoices['รหัสสินค้า'].astype(str).str.contains('SIM-R'))
]
############# DN ###########
OtherSales1=MASSDNCN[MASSDNCN['เลขที่'].str.startswith('HS')]
# OtherSales1
OtherSales1=OtherSales1['มูลค่าสินค้า'].sum()
OtherSales2=MASSDNCN[MASSDNCN['ลูกค้า'].str.contains('ชนะชัย')]
# OtherSales2
OtherSales2=OtherSales2['มูลค่าสินค้า'].sum()
OtherSales=OtherSales1+OtherSales2
############# CN ###########
ChargeBack=MASSDNCN[MASSDNCN['เลขที่'].str.contains('SR')]
ChargeBack=ChargeBack['มูลค่าสินค้า'].sum()
TotalMASS['วันที่']=TotalMASS['วันที่'].astype(str)
SUMMASSP=TotalMASS['มูลค่าสินค้า'].sum()
TotalMASS=pd.merge(TotalMASS,Mold_PM,left_on='รหัสสินค้า',right_on='Part_No',how='left')
TotalMASS=pd.merge(TotalMASS,MoldDP[['Part_No','Mold-DP']],left_on='รหัสสินค้า',right_on='Part_No',how='left')
#################### DP ##############
TotalMASS['DP-Cost'] = TotalMASS['จำนวน'] * TotalMASS['Mold-DP']  # Direct multiplication
#################### PM ##############
TotalMASS['PM-Cost'] = TotalMASS['จำนวน'] * TotalMASS['Mold-PM']
#################### Steel Bush ###########
TotalMASS['รหัสสินค้า']=TotalMASS['รหัสสินค้า'].fillna('NoN')
STB=TotalMASS[TotalMASS['รหัสสินค้า'].str.contains('SB')|TotalMASS['รหัสสินค้า'].str.contains('DENSE')]
# STB
STB_AMT=STB['มูลค่าสินค้า'].sum()
################# Display ####################
TotalMASS=TotalMASS[['วันที่','ลูกค้า','รหัสสินค้า','จำนวน','มูลค่าสินค้า','Mold-DP','DP-Cost','Mold-PM','PM-Cost']]
TotalMASS.set_index('วันที่',inplace=True)
# TotalMASS=TotalMASS.groupby('เลขที่').agg({'ลูกค้า':'first','รหัสสินค้า':'first','จำนวน':'sum','มูลค่าสินค้า':'sum','Mold-DP':'mean','DP-Cost':'sum','Mold-PM':'mean','PM-Cost':'sum'})
############ SUM #############
TatalDP=TotalMASS['DP-Cost'].sum()
TatalPM=TotalMASS['PM-Cost'].sum()
TatalMASSSales=TotalMASS['มูลค่าสินค้า'].sum()

FinalMASSSales=TatalMASSSales-(TatalPM+TatalDP+ChargeBack)
########################
if BU=='MASS':
    st.write('MASS sales AMT')
    TotalMASS
    formatted_display('Total Sales-Steel Bush:',round(STB_AMT,2),'B')
    formatted_display('Total MASS-Sales Part:',round(TatalMASSSales,2),'B')
    formatted_display('Other Sales:',round(OtherSales,2),'B')
    TatalMASSSales=TatalMASSSales-STB_AMT
    MASS_BU=TatalMASSSales+STB_AMT+OtherSales
    formatted_display('Total Sales-MASS BU:',round(MASS_BU,2),'B')
    ################# DP Display #######
    formatted_display('Total Mold-DP:',round(TatalDP,2),'B')
    formatted_display('Total Mold-PM:',round(TatalPM,2),'B')
    formatted_display('Total Final Balance-Sales AMT:',round(FinalMASSSales,2),'B')
    ############ Mass Chart ##############################
    
    # Example data
    categories = ['MASS-BU','MASS-Parts','Steel Bush','Other','Mold-PM Cost', 'Mold-DP Cost','NG-Claim','Final Sales AMT']
    values = [MASS_BU,TatalMASSSales,STB_AMT,OtherSales, -TatalPM, -TatalDP,-ChargeBack,FinalMASSSales]

    # Format values with commas and two decimal places
    formatted_values = [f'{value:,.2f}' for value in values]

    # Create a Plotly figure with formatted value annotations
    fig = go.Figure()

    # Add bar trace
    fig.add_trace(go.Bar(x=categories, y=values, marker_color='#A5FF33', text=formatted_values, textposition='auto'))

    # Update layout
    fig.update_layout(
    title={
        'text': f"  MASS-Sales Metrics                                                     Selected date range: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}",
        'x': 0.5,'xanchor': 'center'},
    xaxis_title='',
    yaxis_title='',
    height=500,width=650,font=dict(size=12))
    st.plotly_chart(fig)
    st.write("---")

###################### Mold Info #########################################
TotalMold = filtered[
    (Invoices['รหัสสินค้า'].astype(str).str.contains('M2')|
    Invoices['รหัสสินค้า'].astype(str).str.contains('SIM-P')|
    Invoices['รหัสสินค้า'].astype(str).str.contains('SIM-R'))
]
MoldTO= filtered[
    (Invoices['รหัสสินค้า'].astype(str).str.contains('T0'))
]
# MoldTO[['วันที่','รหัสสินค้า','จำนวน','มูลค่าสินค้า','JOBCODE']]
MoldTOSales=MoldTO['มูลค่าสินค้า'].sum()
TotalMold['วันที่']=TotalMold['วันที่'].astype(str)
SUMMoldP=TotalMold['มูลค่าสินค้า'].sum()
TotalMold=pd.merge(TotalMold,db[['Part_No','Mold-PM']],left_on='รหัสสินค้า',right_on='Part_No',how='left')

TotalMold=TotalMold[['วันที่','ลูกค้า','รหัสสินค้า','จำนวน','มูลค่าสินค้า','JOBCODE']]
TotalMold.set_index('วันที่',inplace=True)
#############
TotalMoldUnit=TotalMold[TotalMold['รหัสสินค้า'].str.contains('M2')]
MoldSales=TotalMoldUnit['มูลค่าสินค้า'].sum()
MoldPM=TotalMASS['PM-Cost'].sum()
MoldDP=TotalMASS['DP-Cost'].sum()
TatalMoldSales=TotalMold['มูลค่าสินค้า'].sum()
############
filtered = Mold_DP[
(Mold_DP['Date'] >= start_date) &
(Mold_DP['Date'] <= end_date)]
Mold_DP=filtered
Mold_DP.set_index('Date',inplace=True)
MoldDPSales=Mold_DP['AMT'].sum()
GTT_MoldSales=TatalMoldSales+MoldDPSales+MoldPM
#################
if BU=='Mold':
    st.write('Mold sales AMT')
    TotalMold
    st.write('Mold DP AMT')
    #########
    Mold_DP
    ############ Mold 
    TotalMoldUnit=TotalMold[TotalMold['รหัสสินค้า'].str.contains('M2')]
    MoldSales=TotalMoldUnit['มูลค่าสินค้า'].sum()
    formatted_display('Total Mold Sales:',round(MoldSales,2),'B')
    formatted_display('Total Mold DP:',round(MoldDPSales,2),'B')
    ############ Part
    TotalPART=TotalMold[TotalMold['รหัสสินค้า'].str.contains('SIM-P')]
    TatalPARTSales=TotalPART['มูลค่าสินค้า'].sum()
    formatted_display('Total Part Sales:',round(TatalPARTSales,2),'B')
    ############ Repair
    TotalRep=TotalMold[TotalMold['รหัสสินค้า'].str.contains('SIM-R')]
    TatalRepSales=TotalRep['มูลค่าสินค้า'].sum()
    formatted_display('Total Repair Sales:',round(TatalRepSales,2),'B')
    ############ Mold PM
    MoldPM=TotalMASS['PM-Cost'].sum()
    formatted_display('TotalMold-PM Sales:',round(MoldPM,2),'B')
    ########### Mold BU SUM ##################
    TatalMoldSales=TotalMold['มูลค่าสินค้า'].sum()
    GTT_MoldSales=TatalMoldSales+MoldPM+MoldDPSales
    formatted_display('Total Mold BU Sales AMT:',round(GTT_MoldSales,2),'B')
    st.write('---')
    formatted_display('Note: Mold Deposit AMT:',round(MoldTOSales,2),'B')
    ############ Mold  Chart ##############################
    
    # Example data
    categories = ['TT Mold-BU Sales','Mold-Sales','Mold-DP','Part-Sales', 'Repair-Sales','Mold-PM']
    values = [GTT_MoldSales,MoldSales,MoldDPSales,TatalPARTSales, TatalRepSales,MoldPM]

    # Format values with commas and two decimal places
    formatted_values = [f'{value:,.2f}' for value in values]

    # Create a Plotly figure with formatted value annotations
    fig = go.Figure()

    # Add bar trace
    fig.add_trace(go.Bar(x=categories, y=values, marker_color='#F36B0D', text=formatted_values, textposition='auto'))

    # Update layout
    fig.update_layout(
    title={
        'text': f"  Mold-Sales Metrics                                                     Selected date range: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}",
        'x': 0.5,'xanchor': 'center'},
    xaxis_title='',
    yaxis_title='',
    height=500,width=650,font=dict(size=12))
    st.plotly_chart(fig)
    st.write("---")
###################### One-SIM Info #########################################
if BU=='One-SIM':
    st.write('One-SIM sales AMT')
    ############ One-SIM  Chart ##############################
    # Example data
    categories = ['TT One-SIM Sales','Mold-Sales','Mass-Sales']
    values = [(GTT_MoldSales+FinalMASSSales),GTT_MoldSales,FinalMASSSales]
    
    # Format values with commas and two decimal places
    formatted_values = [f'{value:,.2f}' for value in values]

    # Create a Plotly figure with formatted value annotations
    fig = go.Figure()

    # Add bar trace
    colors = ['#1990CC', '#F36B0D', '#A5FF33']
    fig.add_trace(go.Bar(x=categories, y=values, marker_color=colors, text=formatted_values, textposition='auto'))

    # Update layout
    fig.update_layout(
    title={
        'text': f" One-SIM Sales Metrics                                                     Selected date range: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}",
        'x': 0.5,'xanchor': 'center'},
    xaxis_title='',
    yaxis_title='',
    height=500,width=650,font=dict(size=12))
    st.plotly_chart(fig)
    st.write("---")
