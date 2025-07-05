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
def try_parse_date(x):
    try:
        return pd.to_datetime(x).date()
    except:
        return x  # Keep as original string if not convertible
###############################################
def format_dataframe_columns(df):
    formatted_df = df.copy()  # Create a copy of the DataFrame
    for column in formatted_df.columns:
        if formatted_df[column].dtype == 'float64':  # Check if column has a numeric type
            formatted_df[column] = formatted_df[column].apply(lambda x: '{:,.2f}'.format(x))
    return formatted_df
#########################################################
def formatted_display0(label, value, unit):
    formatted_value = "<span style='color:yellow'>{:,.0f}</span>".format(value)  # Format value with comma separator and apply green color
    display_text = f"{formatted_value} {unit}"  # Combine formatted value and unit
    st.write(label, display_text, unsafe_allow_html=True)
#######################################################################################
def formatted_display(label, value, unit):
    formatted_value = "<span style='color:yellow'>{:,.2f}</span>".format(value)  # Format value with comma separator and apply green color
    display_text = f"{formatted_value} {unit}"  # Combine formatted value and unit
    st.write(label, display_text, unsafe_allow_html=True)
############ Logo ####################
logo_image = Image.open('SIM-022.jpg')
st.image(logo_image, width=700)
st.header('SIM Sales Report 2025')
################## Reas File ################
db=pd.read_excel('Database-2022.xlsx')
#################
MoldDP=pd.read_excel('Mold DP-2025.xlsx')
#################
Mold_PM=pd.read_excel('Mold-PM-List.xlsx')
#################
Sales_Target=pd.read_excel('Sales-2025-Target.xlsx')
Mold_Target=pd.read_excel('Sales-2025-Target-Mold.xlsx')
############### 2025 #####################
@st.cache_data 
def load_data_from_drive():
    url="https://docs.google.com/spreadsheets/d/1BG7w4vkBCCN6LTpl6gtUJAfye-8emUL8R-W5oPky_Js/export?format=xlsx"
    data2025=pd.read_excel(url,header=4)
    return data2025
data2025 = load_data_from_drive()
Invoices=data2025
# Invoices
############### Mold DP #####################
@st.cache_data 
def load_data_from_drive():
    Moldurl="https://docs.google.com/spreadsheets/d/1fA2OEz8pnLYUzDUFUOGy9ylL_X4RNd_L/export?format=xlsx"
    dataMold=pd.read_excel(Moldurl)
    return dataMold
dataMold = load_data_from_drive()
Mold_DP=dataMold
#########
import pandas as pd
import streamlit as st
from datetime import datetime

# Function to get the correct last day of February
def get_last_day_of_feb(year):
    return "29" if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else "28"

# Define year
# year = 2025  
year = st.sidebar.selectbox('Select Year',[2025])

########### Menu Range ####################
y_map = {month: f"{year}-{i:02d}-01" for i, month in enumerate(
    ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], start=1)}

y_map_range = {
    'Jan': f"{year}-01-31",
    'Feb': f"{year}-02-{get_last_day_of_feb(year)}",  # Dynamic last day of February
    'Mar': f"{year}-03-31",
    'Apr': f"{year}-04-30",
    'May': f"{year}-05-31",
    'Jun': f"{year}-06-30",
    'Jul': f"{year}-07-31",
    'Aug': f"{year}-08-31",
    'Sep': f"{year}-09-30",
    'Oct': f"{year}-10-31",
    'Nov': f"{year}-11-30",
    'Dec': f"{year}-12-31"
}

####################################
# Streamlit sidebar for selecting start and end months
start_month = st.sidebar.selectbox('Select start month', list(y_map.keys()), index=0)
end_month = st.sidebar.selectbox('Select end month', list(y_map_range.keys()), index=0)

# Convert selected months to datetime objects
start_date = pd.to_datetime(y_map[start_month])
end_date = pd.to_datetime(y_map_range[end_month])

# Ensure 'วันที่' column is in datetime format
Invoices['วันที่'] = pd.to_datetime(Invoices['วันที่'], errors='coerce')

# Filter the DataFrame based on the date range
filtered = Invoices[
    (Invoices['วันที่'] >= start_date) &
    (Invoices['วันที่'] <= end_date)
]

############ BU Menu #####################################################
BU = st.sidebar.selectbox('Select BU',['MASS','Mold','One-SIM'] )
####################### Mass Info #########################################

TotalMASS = filtered[
    (Invoices['ลูกค้า'].str.contains('VALEO AUTOMOT') |
    Invoices['ลูกค้า'].str.contains('แครทโค') |
    Invoices['ลูกค้า'].str.contains('อีเลคโทรลักซ์ ') |
    Invoices['ชื่อสินค้า'].str.contains('PACKING') |
    Invoices['ลูกค้า'].str.contains('เซนทรัล') |
    Invoices['ลูกค้า'].str.contains('โฮมเอ็ก') |
    Invoices['ลูกค้า'].str.contains('ศิริโกมล') |
    Invoices['ลูกค้า'].str.contains('โคชิน') |
    Invoices['ลูกค้า'].str.contains('คาวันญ่า') |
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
#################### Valeo  ###########
TotalMASS['รหัสสินค้า']=TotalMASS['รหัสสินค้า'].fillna('NoN')
VALT=TotalMASS[TotalMASS['ลูกค้า'].str.contains('VALEO')]
# STB
VALT_AMT=VALT['มูลค่าสินค้า'].sum()
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
# TotalMASS[['รหัสสินค้า','จำนวน','PM-Cost']]

TatalPM=TotalMASS['PM-Cost'].sum()
TatalMASSSales=TotalMASS['มูลค่าสินค้า'].sum()
TatalMASSSales=TatalMASSSales-TatalDP
FinalMASSSales=TatalMASSSales-(TatalPM)
########################
if BU=='MASS':
    st.write('MASS sales AMT')
    TotalMASS[['ลูกค้า','รหัสสินค้า','จำนวน','มูลค่าสินค้า']]
    formatted_display('Total Sales-Valeo:',round(VALT_AMT-TatalDP,2),'B')
    formatted_display('Total Sales-Steel Bush:',round(STB_AMT,2),'B')
    formatted_display('Total MASS-Sales Part:',round(FinalMASSSales,2),'B')
    formatted_display('Other Sales:',round(OtherSales,2),'B')
    TatalMASSSales=TatalMASSSales-STB_AMT
    MASS_BU=TatalMASSSales+STB_AMT+OtherSales
    formatted_display('Total Sales-MASS BU:',round(MASS_BU,2),'B')
    ################# DP Display #######
    formatted_display('Total Mold-DP:',round(TatalDP,2),'B')
    formatted_display('Total Mold-PM:',round(TatalPM,2),'B')
    formatted_display('Total Final Balance-Sales AMT:',round(FinalMASSSales,2),'B')
    ############ Mass Chart ##############################
    months_in_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    months_in_order_upper = [month.upper() for month in months_in_order]
    start_index = months_in_order.index(start_month)
    end_index = months_in_order.index(end_month)
    Sales_Target1 = Sales_Target[start_month].sum()
    Sales_Target2 = Sales_Target[end_month].sum()

    if Sales_Target[end_month].sum() == Sales_Target[start_month].sum():
        Sales_Target = Sales_Target1
    else:
        # Sales_Target = Sales_Target1 + Sales_Target2
        Sales_Target_Cumulative = sum(Sales_Target[months_in_order[i]] for i in range(start_index, end_index + 1))
        formatted_values = [f'{x:,.2f}' for x in Sales_Target_Cumulative]  # Converting to list
        formatted_values=pd.DataFrame(formatted_values)
        formatted_values=formatted_values[0].sum()
        # Check if formatted_values is a string, and if so, remove commas and convert it to a numeric type before rounding and converting to an integer
        if isinstance(formatted_values, str):
            formatted_values = formatted_values.replace(',', '')  # Remove commas from the string
            formatted_values = float(formatted_values)  # Convert string to float
        formatted_values = round(formatted_values)
        formatted_values = int(formatted_values)
        Sales_Target=formatted_values
        # Sales_Target
    ####################################################

    categories = ['MASS Sales Target','Total-MASS-Sales','MASS-Parts','Steel Bush','Other','Mold-PM Cost', 'Mold-DP Cost','NG-Claim']
    values = [Sales_Target,FinalMASSSales,TatalMASSSales,STB_AMT,OtherSales, -TatalPM, -TatalDP,-ChargeBack]

    # Format values with commas and two decimal places
    formatted_values = [f'{value:,.2f}' for value in values]

    # Create a Plotly figure with formatted value annotations
    fig = go.Figure()
    colors = ['#F36B0D','#A5FF33', '#A5FF33', '#A5FF33', '#A5FF33', '#A5FF33', '#A5FF33', '#A5FF33']
    # Add bar trace
    fig.add_trace(go.Bar(x=categories, y=values, marker_color=colors, text=formatted_values, textposition='auto'))

    # Update layout
    fig.update_layout(
    title={
        'text': f"  MASS-Sales Metrics                                                     Selected date range: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}",
        'x': 0.5,'xanchor': 'center'},
    xaxis_title='',
    yaxis_title='',
    height=500,width=650,font=dict(size=12))
    st.plotly_chart(fig)
    TargetAch=MASS_BU-Sales_Target
    formatted_display('Target and Actual Balance:',round(TargetAch,2),'B')
    st.write("---")
    ###########################################################
    #Checking PartMASS
    #############################################################
    st.write('**Checking MASS-BU Sales by AMT and Pcs**')
    ################# Add 2024 Data ####################

    PartMASS = Invoices[['วันที่', 'รหัสสินค้า', 'จำนวน', 'มูลค่าสินค้า']]
    PartNo = st.text_input('Input 4-digit Part No')
    # Find the matching 9-digit Part No in the DataFrame
    if len(PartNo) == 4:
        PartMASS = Invoices[['วันที่', 'รหัสสินค้า', 'จำนวน', 'มูลค่าสินค้า']]
        PartMASS = PartMASS[PartMASS['วันที่'].between(start_date, end_date)]
        # Remove missing values from the 'รหัสสินค้า' column
        PartMASS = PartMASS.dropna(subset=['รหัสสินค้า']) 
        # Find the matching rows using str.contains and the boolean mask
        mask = PartMASS['รหัสสินค้า'].str.contains(PartNo, na=False)
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
            formatted_display0('Total Pcs:',round(TTPCS,2),'Pcs')
            formatted_display('Total Sales:',round(TTB,2),'B')
        else:
            st.write(f'No matching Part No found for "{PartNo}"')
    else:
        st.write('Please input a 4-digit Part No')

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
Mold_DP['Date'] = pd.to_datetime(Mold_DP['Date'])
filtered = Mold_DP[
(Mold_DP['Date'] >= start_date) &
(Mold_DP['Date'] <= end_date)]
Mold_DP=filtered
Mold_DP.set_index('Date',inplace=True)
# Mold_DP
MoldDPSales=Mold_DP['AMT'].sum()
GTT_MoldSales=TatalMoldSales+MoldDPSales+MoldPM
#################
if BU=='Mold':

    st.write('Mold sales AMT')
    TotalMold
    st.write('Mold MRR AMT')
    ########
    Mold_DP
    TT_Mold_DP=Mold_DP['AMT'].sum()
    formatted_display('Total Mold-MRR:',round(TT_Mold_DP,2),'B')
    ############ Mass Chart ##############################
    months_in_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    start_index = months_in_order.index(start_month)
    end_index = months_in_order.index(end_month)
    Mold_Target1 = Mold_Target[start_month].sum()
    Mold_Target2 = Mold_Target[end_month].sum()

    if Mold_Target[end_month].sum() == Mold_Target[start_month].sum():
        Mold_Target = Mold_Target1
    else:
        # Mold_Target = Mold_Target1 + Mold_Target2
        Mold_Target_Cumulative = sum(Mold_Target[months_in_order[i]] for i in range(start_index, end_index + 1))
        formatted_values = [f'{x:,.2f}' for x in Mold_Target_Cumulative]  # Converting to list
        formatted_values=pd.DataFrame(formatted_values)
        formatted_values=formatted_values[0].sum()
        # Check if formatted_values is a string, and if so, remove commas and convert it to a numeric type before rounding and converting to an integer
        if isinstance(formatted_values, str):
            formatted_values = formatted_values.replace(',', '')  # Remove commas from the string
            formatted_values = float(formatted_values)  # Convert string to float
        formatted_values = round(formatted_values)
        formatted_values = int(formatted_values)
        Mold_Target=formatted_values
        # Mold_Target
    ###########################
    # Mold_DP
    ############ Mold 
    TotalMoldUnit=TotalMold[TotalMold['รหัสสินค้า'].str.contains('M2')]
    MoldSales=TotalMoldUnit['มูลค่าสินค้า'].sum()
    formatted_display('Total Mold Sales:',round(MoldSales,2),'B')
    # formatted_display('Total Mold DP:',round(MoldDPSales,2),'B')
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
    GTT_MoldSales=TatalMoldSales+MoldPM+TT_Mold_DP
    formatted_display('Total Mold BU Sales AMT:',round(GTT_MoldSales,2),'B')
    st.write('---')
    formatted_display('Note: Mold Deposit AMT:',round(MoldTOSales,2),'B')
    ############ Mold  Chart ##############################
    
    # Example data
    categories = ['Mold Sales Target','TT Mold-BU Sales','Mold-Sales','Mold-MRR','Part-Sales', 'Repair-Sales','Mold-PM']
    values = [Mold_Target,GTT_MoldSales,MoldSales,TT_Mold_DP,TatalPARTSales, TatalRepSales,MoldPM]

    # Format values with commas and two decimal places
    formatted_values = [f'{value:,.2f}' for value in values]

    colors = ['#F36B0D','#A5FF33', '#A5FF33', '#A5FF33', '#A5FF33', '#A5FF33']

    # Create a Plotly figure with formatted value annotations
    fig = go.Figure()

    # Add bar trace
    fig.add_trace(go.Bar(x=categories, y=values, marker_color=colors, text=formatted_values, textposition='auto'))

    # Update layout
    fig.update_layout(
    title={
        'text': f"  Mold-Sales Metrics                                                     Selected date range: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}",
        'x': 0.5,'xanchor': 'center'},
    xaxis_title='',
    yaxis_title='',
    height=500,width=650,font=dict(size=12))
    st.plotly_chart(fig)
    MOLD_BL=GTT_MoldSales-Mold_Target
    formatted_display('Target and Actual Balance:',round(MOLD_BL,2),'B')
    st.write("---")

###################### One-SIM Info #########################################
if BU=='One-SIM':
    st.write('One-SIM sales AMT')
    ############ One-SIM  Chart ##############################
    # Example data
    categories = ['One SIM Target','TT One-SIM Sales','Mold-Sales','Mass-Sales']
    SIM_Target=Sales_Target+Mold_Target
    #####################
    # Sales_Target,Mold_Target
    # SIM_Target
    ############ Mass Chart ##############################
    months_in_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    start_index = months_in_order.index(start_month)
    end_index = months_in_order.index(end_month)
    SIM_Target1 = SIM_Target[start_month].sum()
    SIM_Target2 = SIM_Target[end_month].sum()

    if SIM_Target[end_month].sum() == SIM_Target[start_month].sum():
        SIM_Target = SIM_Target1
    else:
        # SIM_Target = SIM_Target1 + SIM_Target2
        SIM_Target_Cumulative = sum(SIM_Target[months_in_order[i]] for i in range(start_index, end_index + 1))
        formatted_values = [f'{x:,.2f}' for x in SIM_Target_Cumulative]  # Converting to list
        formatted_values=pd.DataFrame(formatted_values)
        formatted_values=formatted_values[0].sum()
        # Check if formatted_values is a string, and if so, remove commas and convert it to a numeric type before rounding and converting to an integer
        if isinstance(formatted_values, str):
            formatted_values = formatted_values.replace(',', '')  # Remove commas from the string
            formatted_values = float(formatted_values)  # Convert string to float
        formatted_values = round(formatted_values)
        formatted_values = int(formatted_values)
        SIM_Target=formatted_values
        # SIM_Target
    ####################################################
    # SIM_Target=SIM_Target.sum()
    SIM_ACt=(GTT_MoldSales+FinalMASSSales)
    values = [SIM_Target,SIM_ACt,GTT_MoldSales,FinalMASSSales]
    
    # Format values with commas and two decimal places
    formatted_values = [f'{value:,.2f}' for value in values]

    # Create a Plotly figure with formatted value annotations
    fig = go.Figure()

    # Add bar trace
    colors = [ '#F36B0D', '#A5FF33', '#A5FF33', '#A5FF33']
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
    SIM_BL=SIM_ACt-SIM_Target
    formatted_display('Target and Actual Balance:',round(SIM_BL,2),'B')
    st.write("---")
############################## Mold Prospected
