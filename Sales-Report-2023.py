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
#######################################################################################
logo_image = Image.open('SIM-022.jpg')
st.image(logo_image, width=700)
st.header('ONE-SIM Sales Report 2023')
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
TTSales=Inv['มูลค่าสินค้า']
TTSales=TTSales.sum()
###############################
col1, col2 = st.columns(2)
with col1:

    Minput = st.selectbox('Select month', ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], key='unique-key-1')

    map_ym = {'Month': ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']}
    y_map = {'Jan': '2023-01-01', 'Feb': '2023-02-01', 'Mar': '2023-03-01', 'Apr': '2023-04-01', 'May': '2023-05-01', 'Jun': '2023-06-01', 'Jul': '2023-07-01',
            'Aug': '2023-08-01', 'Sep': '2023-09-01', 'Oct': '2023-10-01', 'Nov': '2023-11-01', 'Dec': '2023-12-01'}

    map_ym = pd.DataFrame(map_ym)
    map_ym['Year'] = map_ym['Month'].map(y_map)
    map_ym = map_ym[map_ym['Month'] == Minput]

    y = map_ym['Year'].to_string(index=False)
    ym_input = y.strip()
    ym_input
    # Range
with col2:
    Minput2 = st.selectbox('Select range of months', ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], key='unique-key-2')

    map_ym_range = {'Month': ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']}
    y_map_range = {'Jan': '2023-01-31', 'Feb': '2023-02-28', 'Mar': '2023-03-31', 'Apr': '2023-04-30', 'May': '2023-05-31', 'Jun': '2023-06-30', 'Jul': '2023-07-31',
                'Aug': '2023-08-31', 'Sep': '2023-09-30', 'Oct': '2023-10-31', 'Nov': '2023-11-30', 'Dec': '2023-12-31'}

    for key, value in y_map_range.items():
        y_map_range[key] = pd.to_datetime(value) + pd.Timedelta(days=1)

    map_ym_range = pd.DataFrame(map_ym_range)
    map_ym_range['Year'] = map_ym_range['Month'].map(y_map_range)
    map_ym_range = map_ym_range[map_ym_range['Month'] == Minput2]

    y2 = map_ym_range['Year'].to_string(index=False)
    ym_input2 = y2.strip()
    ym_input2
#########################################################
ym_Count = ym_input[:7]
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
TotalMASS=TotalMASS[TotalMASS['วันที่'].between( ym_input, ym_input2)]
TotalMASS=pd.merge(TotalMASS,db[['Part_No','Mold-PM','Mold-DP']],left_on='รหัสสินค้า',right_on='Part_No',how='left')
####################### Steel Bush ################################
TotalSTB=Invoices[Invoices['ชื่อสินค้า'].str.contains('STEEL')]
TotalSTB = TotalSTB[TotalSTB['วันที่'].between( ym_input, ym_input2)]
######################### Mold #########################################
TotalMOLD=Invoices[Invoices['รหัสสินค้า'].astype(str).str.contains('MOLD')|
Invoices['รหัสสินค้า'].astype(str).str.contains('PART')|
Invoices['รหัสสินค้า'].astype(str).str.contains('REPAIR')]
TotalMOLD=TotalMOLD[TotalMOLD['วันที่'].between( ym_input, ym_input2)]
######################## Other ###############################
TotalOTHER=Invoices[Invoices['ชื่อสินค้า'].str.contains('DENSE')|Invoices['ชื่อสินค้า'].str.contains('RTV')|Invoices['ชื่อสินค้า'].str.contains('ตู้')]
TotalOTHER=TotalOTHER[TotalOTHER['วันที่'].between( ym_input, ym_input2)]
####################### Cash ##################################
SalesCash=SalesCash[SalesCash['วันที่'].between( ym_input, ym_input2)]
SalesCash=SalesCash[['วันที่','มูลค่าสินค้า']]
###########################################################
st.write('---')
st.write('**ONE-SIM Invoices and Sales AMT details**')
ONESIM=Inv[Inv['วันที่'].between( ym_input, ym_input2)]
ONESIM=ONESIM
ONESIM2=ONESIM.groupby('ลูกค้า').agg({'มูลค่าสินค้า':'sum'})
formatted_df = format_dataframe_columns(ONESIM2)
st.dataframe(formatted_df)
################ Mold T2 Added ################################
st.write('---')
st.write('**Mold Sales MOLD-T0 (Deposit)**')
def format_dataframe_columns(df):
    # Format specific columns
    df['มูลค่าสินค้า'] = df['มูลค่าสินค้า'].apply(lambda x: '{:,.2f}'.format(x))
    return df
T1PlusT2=T1PlusT2[T1PlusT2['รหัสสินค้า'].str.contains('MOLD-T0')]
# Extract the substring before the hyphen in JOBCODE column of ONESIM DataFrame
ONESIM['JOBCODE'] = ONESIM['JOBCODE'].str.split('-').str[0]
T1PlusT2['JOBCODE']=T1PlusT2['JOBCODE'].str.split('-').str[0]
# Merge T1PlusT2 with the modified JOBCODE column
T1PlusT2 = pd.merge(T1PlusT2, ONESIM['JOBCODE'], on='JOBCODE', how='right')
# T1PlusT2=pd.merge(T1PlusT2,ONESIM['JOBCODE'],on='JOBCODE',how='right')
T1PlusT2=T1PlusT2[['รหัสสินค้า','JOBCODE','มูลค่าสินค้า']]
T1PlusT2=T1PlusT2.fillna(0)
T1PlusT2 = T1PlusT2.drop_duplicates()
T1PlusT2=T1PlusT2[T1PlusT2['มูลค่าสินค้า']!=0]
T1PlusT2 = T1PlusT2.reset_index(drop=True)
T1PlusT2.index=T1PlusT2.index+1
MoldT2=T1PlusT2['มูลค่าสินค้า'].sum()
if T1PlusT2.empty:
    st.write('No Deposit')
else:
    ###################
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
end_date = ym_input2
# create a pandas date range for the specified date range
date_range = pd.date_range(start=start_date, end=end_date)
# filter out non-business days using the BDay frequencybbb
business_days = date_range[date_range.weekday < 5]
# print the resulting number of business days
Days = len(business_days)-4
Target2023=pd.read_excel('Target-2023.xlsx')
Target2023=Target2023[Minput]
Target2023=(Target2023/Days)*COUNT
Target2023=list(Target2023)
############################################################################
st.write('---')
st.write('**Credit Note Details-MASS**')
TotalMASSCN = TotalMASS[TotalMASS['เลขที่'].str.contains('SR')]

if TotalMASSCN.empty:
    st.write("No MASS Credit Note.")
else:
    MASSNOTE=TotalMASSCN[['ลูกค้า', 'มูลค่าสินค้า']]
    MASSNOTE=MASSNOTE.set_index('ลูกค้า')
    
    ###################
    formatted_df = format_dataframe_columns(MASSNOTE)
    st.dataframe(formatted_df)
    ####################

TotalMASSCN = TotalMASSCN[['ลูกค้า', 'มูลค่าสินค้า']].groupby('ลูกค้า').sum()
MASSCN = 0  # initialize MASSCN to 0
try:
    MASSCN = round(TotalMASSCN['มูลค่าสินค้า'].sum(),2)
except KeyError:
    pass  # do nothing if key error occurs
formatted_display('Total Credit Note Details-MASS:', MASSCN,'B')
st.write('---')
st.write('**Credit Note Details-MOLD**')
TotalMoldCN = TotalMOLD[TotalMOLD['เลขที่'].str.contains('SR')]
if TotalMoldCN.empty:
    st.write('No Mold Credit Note')
else:
    MOLDNOTE=(TotalMoldCN[['ลูกค้า', 'มูลค่าสินค้า']])
    MOLDNOTE=MOLDNOTE.set_index('ลูกค้า')
    ###################
    formatted_df = format_dataframe_columns(MOLDNOTE)
    st.dataframe(formatted_df)
    ####################
TotalMoldCN = TotalMoldCN[['ลูกค้า', 'มูลค่าสินค้า']].groupby('ลูกค้า').sum()
MOLDCN = 0  # initialize MOLDCN to 0
try:
    MOLDCN = round(TotalMoldCN['มูลค่าสินค้า'].sum(),2)
except KeyError:
    pass  # do nothing if key error occurs
formatted_display('Total Credit Note Details-MOLD:', MOLDCN,'B')
st.write('---')
##########################################################################
st.write('**Debit Note Details-MASS**')
TotalMASSDN = TotalMASS[TotalMASS['เลขที่'].str.contains('DR')]
if TotalMASSDN.empty:
    st.write('No Mold Debit Note')
else:
    st.dataframe(TotalMASSDN[['ลูกค้า', 'มูลค่าสินค้า']])

TotalMASSDN = TotalMASSDN[['ลูกค้า', 'มูลค่าสินค้า']].groupby('ลูกค้า').sum()
MASSDN = 0  # initialize MASSDN to 0
try:
    MASSDN = round(TotalMASSDN['มูลค่าสินค้า'].sum(),2)
except KeyError:
    pass  # do nothing if key error occurs
formatted_display('Total Debit Note Details-MASS:', MASSDN,'B')
st.write('---')
st.write('**Debit Note Details-MOLD**')
TotalMoldDN = TotalMOLD[TotalMOLD['เลขที่'].str.contains('DR')]
if TotalMoldDN.empty:
    st.write('No Mold Debit Note')
else:
    st.dataframe(TotalMoldDN[['ลูกค้า', 'มูลค่าสินค้า']])

TotalMoldDN = TotalMoldDN[['ลูกค้า', 'มูลค่าสินค้า']].groupby('ลูกค้า').sum()
MOLDDN = 0  # initialize MOLDDN to 0
try:
    MOLDDN = round(TotalMoldDN['มูลค่าสินค้า'].sum(),2)
except KeyError:
    pass  # do nothing if key error occurs
formatted_display('Total Debit Note Details-MOLD:', MOLDDN,'B')
st.write('---')
st.write('**Balance CN/DN**')
MASSCNDNBL=MASSCN+MASSDN
formatted_display('Balance MASS CN/DN:',MASSCNDNBL,'B')
MOLDCNDNBL=MOLDCN+MOLDDN
formatted_display('Balance MOLD CN/DN:',MOLDCNDNBL,'B')
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
TotalSales=(TotalSaleCASH+TotalSalesMASS+TotalSaleSTB+TotalSalesOTHER+TotalSalesMOLD)
formatted_display('Total MASS BU Sales:',round(TotalSalesMASS,2),'B')
formatted_display('Total Steel Bush Sales:',round(TotalSaleSTB,2),'B')
formatted_display('Total Mold BU Sales:',round(TotalSalesMOLD,2),'B')
formatted_display('Total Mold Deposit:',round(MoldT2,2),'B')
formatted_display('Total Mold PM Internal Charged:',round(TotalMoldPM,2),'B')
formatted_display('Total Mold DP Income:',round(TotalMoldDP,2),'B')
formatted_display('Total Other Sales:',round(TotalSalesOTHER,2),'B')
formatted_display('Total Cash:',round(TotalSaleCASH,2),'B')
formatted_display('Total One-SIM Sales:',round(TotalSales+TotalSaleCASH,2),'B')
DATASALES=[['One-SIM',TotalSales],['MASS',TotalSalesMASS],['Steel Bush',TotalSaleSTB],['Mold',TotalSalesMOLD],['Other',TotalSalesOTHER]]
SUMSALES=pd.DataFrame(DATASALES,columns=['Items','AMT'])
SUMSALES.set_index('Items',inplace=True)
############# Target ######################################################
start_date = Minput
end_date = Minput2
start_date = parser.parse(start_date)
end_date = parser.parse(end_date)
num_months = (end_date.year - start_date.year) * 12 + end_date.month - start_date.month + 1

# Define the data for the bar chart
categories = ['One-SIM','MASS','Mold','Mold-T0','Steel Bush','Cash','Other']
values = [TotalSales+TotalSaleCASH, TotalSalesMASS,(TotalSalesMOLD-MoldT2),MoldT2,TotalSaleSTB,TotalSaleCASH,TotalSalesOTHER]
values2 =Target2023
# Use num_months as the monthly factor to multiply the values in values and values2
monthly_factor = num_months
for i in range(len(values2)):
    values2[i] = values2[i] * monthly_factor
##########################################################
today = datetime.datetime.today().strftime('%Y-%m-%d')

labels = [f"{value:,.0f}" for value in values]
labels2 = [f"{value:,.0f}" for value in values2]

# Create the bar chart
trace1 = go.Bar(x=categories, y=values, name='Actual',text=labels, textposition='auto')
# trace2 = go.Bar(x=categories, y=values2, name='Target',text=labels2, textposition='auto')
trace2 = go.Scatter(x=categories, y=values2, name='Target', text=labels2, textposition='top center',line=dict(color='orange'))

############################
fig = go.Figure(data=[go.Bar(x=categories, y=values, text=labels, textposition='auto')])
# fig = go.Figure(data=[go.Bar(x=categories, y=values2, text=labels2, textposition='auto')])
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
