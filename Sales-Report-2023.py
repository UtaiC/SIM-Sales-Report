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
############
Logo=Image.open('SIM-LOGO-02.jpg')
st.image(Logo,width=700)
st.markdown("<h2 style='text-align: center; color:#F1F0E5'>Sales Report 2023 </h2>",unsafe_allow_html=True)
# st.subheader("Sales 2023 Report:")
##################
db=pd.read_excel('Database-2022.xlsx')
#################
@st.cache(allow_output_mutation=True)
def load_data_from_drive():
    url="https://docs.google.com/spreadsheets/d/1sUH0WmtfrWbR8FrM33ljebSlW0BWE3w0/export?format=xlsx"
    data=pd.read_excel(url,header=4)
    return data
data = load_data_from_drive()
Invoices=data

# url="https://docs.google.com/spreadsheets/d/1sUH0WmtfrWbR8FrM33ljebSlW0BWE3w0/export?format=xlsx"
# Invoices=pd.read_excel(url,header=4)
Invoices[['วันที่','เลขที่','ลูกค้า','ชื่อสินค้า']]=Invoices[['วันที่','เลขที่','ลูกค้า','ชื่อสินค้า']].astype(str)
Invoices=Invoices[(~Invoices['รหัสสินค้า'].astype(str).str.contains('MOLD-T0'))]
Invoices=Invoices[(~Invoices['รหัสสินค้า'].astype(str).str.contains('MOLD-T1'))]
Inv=Invoices[['วันที่','เลขที่','ลูกค้า','ชื่อสินค้า','จำนวน','มูลค่าสินค้า','รหัสสินค้า']]
Inv=Invoices[Invoices['เลขที่'].str.contains('IV')|Invoices['เลขที่'].str.contains('HS')]
# TTCN=Invoices[Invoices['เลขที่'].str.contains('SR')]
# TTCN[['ลูกค้า','มูลค่าสินค้า']]
# TTDN=Invoices[Invoices['รหัสสินค้า'].str.contains('DR')]
# TTDN[['ลูกค้า','มูลค่าสินค้า']]
###################### Total Invoice Check #################################################
Inv['มูลค่าสินค้า']= pd.to_numeric(Inv['มูลค่าสินค้า'], errors='coerce')
Inv['มูลค่าสินค้า'].dropna(0, inplace=True)
TTSales=Inv['มูลค่าสินค้า']
TTSales=TTSales.sum()
################ Select ############################
MonthSelected=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

Minput=st.selectbox('MonthSelected',MonthSelected, key='unique-key-1')
MAPYM={'Month':['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']}
YMAP={'Jan':'2023-01-01','Feb':'2023-02-01','Mar':'2023-03-01','Apr':'2023-04-01','May':'2023-05-01','Jun':'2023-06-01','Jul':'2023-07-01',
'Aug':'2023-08-01','Sep':'2023-09-01','Oct':'2023-10-01','Nov':'2023-11-01','Dec':'2023-12-31-01'}
MAPYM=pd.DataFrame(MAPYM)
MAPYM['Year']=MAPYM['Month'].map(YMAP)
MAPYM=MAPYM[MAPYM['Month']==Minput]
Y=MAPYM['Year'].to_string(index=False)
YMInput=Y
Y
################ Range ############################
Minput2=st.selectbox('Range of MonthSelected',MonthSelected, key='unique-key-2')
MAPYM2={'Month':['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']}
YMAP2={'Jan':'2023-01-31','Feb':'2023-02-28','Mar':'2023-03-31','Apr':'2023-04-30','May':'2023-05-31','Jun':'2023-06-30','Jul':'2023-07-31',
'Aug':'2023-08-31','Sep':'2023-09-30','Oct':'2023-10-31','Nov':'2023-11-30','Dec':'2023-12-31'}
for key, value in YMAP2.items():
    YMAP2[key] = pd.to_datetime(value) + pd.Timedelta(days=1)
MAPYM2=pd.DataFrame(MAPYM2)
MAPYM2['Year']=MAPYM2['Month'].map(YMAP2)
MAPYM2=MAPYM2[MAPYM2['Month']==Minput2]
Y2=MAPYM2['Year'].to_string(index=False)
YMInput2=Y2
Y2
#########################################################
ym_Count = Y[:7]
DayCount=Invoices['วันที่']
DayCount=DayCount[DayCount.str.contains(ym_Count)]
DayCount=DayCount.drop_duplicates()
COUNT=DayCount.count()
############# Cash #############
SalesCash=Invoices[Invoices['ชื่อสินค้า'].str.contains('ขี้กลึงเหล็ก')|Invoices['ชื่อสินค้า'].str.contains('ขี้เตา')|Invoices['ชื่อสินค้า'].str.contains('เศษเหล็ก')
|Invoices['ชื่อสินค้า'].str.contains('ขี้กลึงอลูมิเนียม')]
# SalesCash=pd.read_excel("HS2022.xlsx",header=5)
# SalesCash['วันที่']=SalesCash['วันที่'].astype(str)
############################################
last_update = DayCount.max()
st.write("Last update:", last_update)
st.write('Invoice Isuued Days:',COUNT)
###################### MASS ################################
TotalMASS=Invoices[Invoices['ลูกค้า'].str.contains('VALEO')|Invoices['ลูกค้า'].str.contains('แครทโค')|Invoices['ชื่อสินค้า'].str.contains('PACKING')|
Invoices['ลูกค้า'].str.contains('เซนทรัล')]
TotalMASS=TotalMASS[TotalMASS['วันที่'].between( YMInput, YMInput2)]
TotalMASS=pd.merge(TotalMASS,db[['Part_No','Mold-PM','Mold-DP']],left_on='รหัสสินค้า',right_on='Part_No',how='left')
####################### Steel Bush ################################
TotalSTB=Invoices[Invoices['ชื่อสินค้า'].str.contains('STEEL')]
TotalSTB = TotalSTB[TotalSTB['วันที่'].between( YMInput, YMInput2)]
######################### Mold #########################################
TotalMOLD=Invoices[Invoices['รหัสสินค้า'].astype(str).str.contains('MOLD')|
Invoices['รหัสสินค้า'].astype(str).str.contains('PART')|
Invoices['รหัสสินค้า'].astype(str).str.contains('REPAIR')]
TotalMOLD=TotalMOLD[TotalMOLD['วันที่'].between( YMInput, YMInput2)]
######################## Other ###############################
TotalOTHER=Invoices[Invoices['ชื่อสินค้า'].str.contains('DENSE')|Invoices['ชื่อสินค้า'].str.contains('RTV')|Invoices['ชื่อสินค้า'].str.contains('ตู้')]
TotalOTHER=TotalOTHER[TotalOTHER['วันที่'].between( YMInput, YMInput2)]
####################### Cash ##################################
SalesCash=SalesCash[SalesCash['วันที่'].between( YMInput, YMInput2)]
SalesCash=SalesCash[['วันที่','มูลค่าสินค้า']]
############################################################
ONESIM=Inv[Inv['วันที่'].between( YMInput, YMInput2)]
ONESIM=ONESIM.groupby('ลูกค้า').agg({'มูลค่าสินค้า':'sum','ชื่อสินค้า':'first'})
ONESIM
#########################################################
TotalMoldPM=(TotalMASS['จำนวน']*TotalMASS['Mold-PM']).sum()
TotalMoldDP=(TotalMASS['จำนวน']*TotalMASS['Mold-DP']).sum()
TotalSaleCASH=SalesCash['มูลค่าสินค้า'].sum()
TotalSalesMASS=((TotalMASS['มูลค่าสินค้า'].sum())-(TotalMoldPM+TotalMoldDP))
TotalSaleSTB=TotalSTB['มูลค่าสินค้า'].sum()
TotalSalesOTHER=TotalOTHER['มูลค่าสินค้า'].sum()
TotalSalesMOLD=(TotalMOLD['มูลค่าสินค้า'].sum())+(TotalMoldPM+TotalMoldDP)
TotalSales=(TotalSaleCASH+TotalSalesMASS+TotalSaleSTB+TotalSalesOTHER+TotalSalesMOLD)
st.write('Total Sales Invoices:',round(TTSales,2))
st.write('Total MASS BU Sales:',round(TotalSalesMASS,2))
st.write('Total Steel Bush Sales:',round(TotalSaleSTB,2))
st.write('Total Mold BU Sales:',round(TotalSalesMOLD,2))
st.write('Total Mold PM Internal Charged :',round(TotalMoldPM,2))
st.write('Total Mold DP Income :',round(TotalMoldDP,2))
st.write('Total Other Sales:',round(TotalSalesOTHER,2))
st.write('Total Cash:',round(TotalSaleCASH,2))
st.write('Total One-SIM Sales:',round(TotalSales+TotalSaleCASH,2))
DATASALES=[['One-SIM',TotalSales],['MASS',TotalSalesMASS],['Steel Bush',TotalSaleSTB],['Mold',TotalSalesMOLD],['Other',TotalSalesOTHER]]
SUMSALES=pd.DataFrame(DATASALES,columns=['Items','AMT'])
SUMSALES.set_index('Items',inplace=True)
############# Target ######################################################
from pandas.tseries.offsets import BDay

# specify the start and end dates for the date range
start_date = Y
end_date = Y2

# create a pandas date range for the specified date range
date_range = pd.date_range(start=start_date, end=end_date)

# filter out non-business days using the BDay frequency
business_days = date_range[date_range.weekday < 5]

# print the resulting number of business days
Days = len(business_days)
Target2023=pd.read_excel('Target-2023.xlsx')
Target2023=Target2023[Minput]
Target2023=(Target2023/Days)*COUNT
Target2023=list(Target2023)
############################################################################
st.write('---')
st.write('**Credit Note Details-MASS**')
TotalMASSCN = TotalMASS[TotalMASS['เลขที่'].str.contains('SR')]
TotalMASSCN[['ลูกค้า', 'มูลค่าสินค้า']]
TotalMASSCN = TotalMASSCN[['ลูกค้า', 'มูลค่าสินค้า']].groupby('ลูกค้า').sum()
MASSCN = 0  # initialize MASSCN to 0
try:
    MASSCN = round(TotalMASSCN['มูลค่าสินค้า'].sum(),2)
except KeyError:
    pass  # do nothing if key error occurs
st.write('Total Credit Note Details-MASS:', MASSCN)
st.write('---')
st.write('**Credit Note Details-MOLD**')
TotalMoldCN = TotalMOLD[TotalMOLD['เลขที่'].str.contains('SR')]
TotalMoldCN[['ลูกค้า', 'มูลค่าสินค้า']]
TotalMoldCN = TotalMoldCN[['ลูกค้า', 'มูลค่าสินค้า']].groupby('ลูกค้า').sum()
MOLDCN = 0  # initialize MOLDCN to 0
try:
    MOLDCN = round(TotalMoldCN['มูลค่าสินค้า'].sum(),2)
except KeyError:
    pass  # do nothing if key error occurs
st.write('Total Credit Note Details-MOLD:', MOLDCN)
st.write('---')
##########################################################################
st.write('**Debit Note Details-MASS**')
TotalMASSDN = TotalMASS[TotalMASS['เลขที่'].str.contains('DR')]
TotalMASSDN[['ลูกค้า', 'มูลค่าสินค้า']]
TotalMASSDN = TotalMASSDN[['ลูกค้า', 'มูลค่าสินค้า']].groupby('ลูกค้า').sum()
MASSDN = 0  # initialize MASSDN to 0
try:
    MASSDN = round(TotalMASSDN['มูลค่าสินค้า'].sum(),2)
except KeyError:
    pass  # do nothing if key error occurs
st.write('Total Debit Note Details-MASS:', MASSDN)
st.write('---')
st.write('**Debit Note Details-MOLD**')
TotalMoldDN = TotalMOLD[TotalMOLD['เลขที่'].str.contains('DR')]
TotalMoldDN[['ลูกค้า', 'มูลค่าสินค้า']]
TotalMoldDN = TotalMoldDN[['ลูกค้า', 'มูลค่าสินค้า']].groupby('ลูกค้า').sum()
MOLDDN = 0  # initialize MOLDDN to 0
try:
    MOLDDN = round(TotalMoldDN['มูลค่าสินค้า'].sum(),2)
except KeyError:
    pass  # do nothing if key error occurs
st.write('Total Debit Note Details-MOLD:', MOLDDN)
st.write('---')
############################################################################
start_date = Minput
end_date = Minput2
start_date = parser.parse(start_date)
end_date = parser.parse(end_date)
num_months = (end_date.year - start_date.year) * 12 + end_date.month - start_date.month + 1

# Define the data for the bar chart
categories = ['One-SIM','MASS','Mold','Steel Bush','Cash','Other']
values = [TotalSales+TotalSaleCASH, TotalSalesMASS+(MASSCN+MASSDN),TotalSalesMOLD+(MOLDCN+MOLDDN),TotalSaleSTB,TotalSaleCASH,(TotalSalesOTHER)]
# values2 = [((19166666.66/(250/12))*COUNT),((13940823/(250/12))*COUNT),((954000/(250/12))*COUNT),((4350374.17/(250/12))*COUNT),0,0]
values2 =Target2023
# Use num_months as the monthly factor to multiply the values in values and values2
monthly_factor = num_months
for i in range(len(values2)):
    values2[i] = values2[i] * monthly_factor
##########################################################
today = datetime.datetime.today().strftime('%Y-%m-%d')
# Create a list of labels with the "k" suffix
# labels = [f"{value/1000:.0f}k" for value in values]
# labels = [f"{int(value)}" for value in values]
# labels = [f"{value:.2f}" for value in values]
labels = [f"{value:,.0f}" for value in values]
labels2 = [f"{value:,.0f}" for value in values2]
# Create the bar chart
trace1 = go.Bar(x=categories, y=values, name='Actual',text=labels, textposition='auto')
# trace2 = go.Bar(x=categories, y=values2, name='Target',text=labels2, textposition='auto')
trace2 = go.Scatter(x=categories, y=values2, name='Target', text=labels2, textposition='top center',line=dict(color='orange'))
############################
fig = go.Figure(data=[go.Bar(x=categories, y=values, text=labels, textposition='auto')])
fig = go.Figure(data=[go.Bar(x=categories, y=values2, text=labels2, textposition='auto')])
# Add a title and axis labels
data = [trace1, trace2]
# Create the figure object
fig = go.Figure(data=data)
# Add a title and axis labels
fig.update_layout(title_text='Chart of Sales by BU Items', xaxis_title='Category', yaxis_title='Value')
fig.add_annotation(go.Annotation(text=f"Report date: {last_update}", x=0.900, xref="paper", y=1, yref="paper"))
fig.update_layout(title_text='Sales-2023 Report by BU Items:', xaxis_title='Category', yaxis_title='Value')

# Show the plot
st.plotly_chart(fig)

if st.button("Refresh data"):
    data = load_data_from_drive()
