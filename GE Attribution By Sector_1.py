# -*- coding: utf-8 -*-
"""
Created on Sat Apr  6 18:28:04 2024

@author: sowang
"""

import pyodbc
import pandas as pd
import numpy as np
import xlsxwriter #pip install XlsxWriter
import datetime
from pandas.tseries.offsets import BDay
import win32com.client
from datetime import date

import json
from streamlit_echarts import Map
from streamlit_echarts import JsCode
import streamlit as st
from streamlit_echarts import st_echarts
from pyecharts import options as opts
from pyecharts.charts import Bar,Line,Pie
from streamlit_echarts import st_pyecharts
import pandas as pd
import numpy as np
from millify import prettify,millify


#%%
# set up for wide page display
st.set_page_config(page_title="columns show",layout="wide")
#%%
st.title("GE Portfolio Attribution by Sector")
#%%

today = datetime.datetime.today()
effective_date = today - BDay(1)
effective_date = effective_date.strftime("%Y-%m-%d")

previous_date =  today-BDay(2)
previous_date = previous_date.strftime("%Y-%m-%d")

#%%
#INVESTMENTDB
INVESTMENTDB = {'servername': 'OMS7677',
      'database': 'DimensionalModel'}


# create the connection to investmentdb
conn = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=' + INVESTMENTDB['servername'] + ';DATABASE=' + INVESTMENTDB['database'] + ';Trusted_Connection=yes')

#%%

# Record of effective date
query1 = '''select * from [dbo].[Performance_Attribution_By_Economic_Sector_Program]
where effective_date =\'{0}\''''.format(effective_date)  
    
query1 = query1.replace('\n',' ')
# This is to remove space in sql script

df = pd.read_sql_query(query1,conn)

#st.table(df)

# Record of previous date
query_p = '''select * from [dbo].[Performance_Attribution_By_Economic_Sector_Program]
where effective_date =\'{0}\''''.format(previous_date)
  
query_p = query_p.replace('\n',' ')
# This is to remove space in sql script

df_p = pd.read_sql_query(query_p,conn)


#%%
comp_ID = df['Composite_ID'].unique()

#comp_ID

comp_selected = st.sidebar.selectbox("Please Select a Composite ID",comp_ID)

#comp_selected 
    
df_new=df.loc[df['Composite_ID']==comp_selected]    
#df_new
ptf_name = df_new['Composite_Name'].unique()

st.sidebar.write(f"Portfolio Name:{ptf_name}")  

# Get list of benchmark corresponding to selected portfolio   
bmk_list=df_new['Benchmark_ID'].unique()

#bmk_list

bmk_selected = st.sidebar.selectbox("Please Select a Benchmark ID",bmk_list)

df_new_2=df_new.loc[df_new['Benchmark_ID']==bmk_selected]
bmk_name = df_new_2['Benchmark_Name'].unique()

st.sidebar.write(f"Benchmark Name:{bmk_name}")  

#bmk_selected
#%%
# New table with selected Composite_ID and Benchmark_ID
excl=['Cash & FX','TOTAL','UNKNOWN']

df_selected=df.loc[(~df['GICS_Economic_Sector_Code'].isin(excl))&(df['Composite_ID']==comp_selected) & (df['Benchmark_ID']==bmk_selected)]


#df_selected['End_Market_Value_Round']=round(df_selected['End_Market_Value'])
#st.table(df_selected)

df_selected_grouped=df_selected.groupby(by='GICS_Economic_Sector_Name')['End_Market_Value'].sum()
#df_selected_grouped

mv=[]
for i in df_selected_grouped.items():
    mv.append({"value":i[1],"name":i[0]})
#mv
#%% For Top Line Data 
incl=['TOTAL']

df_effective_date=df.loc[(df['GICS_Economic_Sector_Code'].isin(incl))&(df['Composite_ID']==comp_selected) & (df['Benchmark_ID']==bmk_selected)]
df_previous_date=df_p.loc[(df_p['GICS_Economic_Sector_Code'].isin(incl))&(df_p['Composite_ID']==comp_selected) & (df_p['Benchmark_ID']==bmk_selected)]

#st.table(df_effective_date)
#st.table(df_previous_date)
Total_Market_Value=df_effective_date['End_Market_Value']
Total_Market_Value_Previous =df_previous_date['End_Market_Value']
Total_Market_Value_Change=Total_Market_Value.values - Total_Market_Value_Previous.values

MTD_Return_Portfolio=df_effective_date['MTD_Return_Portfolio']
MTD_Return_Portfolio_Change=df_effective_date['MTD_Return_Portfolio'].values-df_previous_date['MTD_Return_Portfolio'].values

QTD_Return_Portfolio=df_effective_date['QTD_Return_Portfolio']
QTD_Return_Portfolio_Change=df_effective_date['QTD_Return_Portfolio'].values-df_previous_date['QTD_Return_Portfolio'].values

YTD_Return_Portfolio=df_effective_date['YTD_Return_Portfolio']
YTD_Return_Portfolio_Change=df_effective_date['YTD_Return_Portfolio'].values-df_previous_date['YTD_Return_Portfolio'].values
#%% Top Line Data laoyout

col41,col42,col43,col44 =st.columns(4)

col41.metric(label="Portfolio Total Market Value",value=millify(Total_Market_Value,precision=3),delta=millify(Total_Market_Value_Change,precision=2))
#st.write(Total_Market_Value_Change.values)
col42.metric(label="MTD Return Portfolio",value="{}%".format(millify(MTD_Return_Portfolio,precision=2)),delta="{}%".format(millify(MTD_Return_Portfolio_Change,precision=2)))
#, delta="{}%".format(millify(MTD_Return_Portfolio_Change,precision=2))
col43.metric(label="QTD Return Portfolio",value="{}%".format(millify(QTD_Return_Portfolio,precision=2)), delta="{}%".format(millify(QTD_Return_Portfolio_Change,precision=2)))
#, delta="{}%".format(millify(QTD_Return_Portfolio_Change,precision=2))
col44.metric(label="YTD Return Portfolio",value="{}%".format(millify(YTD_Return_Portfolio,precision=2)), delta="{}%".format(millify(YTD_Return_Portfolio_Change,precision=2)))
#, delta="{}%".format(millify(YTD_Return_Portfolio_Change,precision=2))

#%%
def sector_pie_chart(data):
    option1 = {
  "title": {
    "text": 'Portfolio Market Value by Sector',
    "subtext": 'As of last reporting date',
    "left": 'center'
  },
  "tooltip": {
    "trigger": 'item'
  },
  #"legend": {
   # "orient": 'vertical',
    #"left": 'right',
  #},
  "series": [
    {
      "name": 'Access From',
      "type": 'pie',
      "radius": '60%',
      "data": data,
      "emphasis": {
        "itemStyle": {
          "shadowBlur": 10,
          "shadowOffsetX": 0,
          "shadowColor": 'rgba(0, 0, 0, 0.5)'
        }
      }
    }
  ]
}
    return option1


#%%
#st_echarts(sector_pie_chart(mv),height=400, width=700)

#%%
gics=df_selected['GICS_Economic_Sector_Name'].values.tolist()
weight_ptf=df_selected['End_Weight_Portfolio'].fillna(0).values.tolist()
weight_bmk=df_selected['End_Weight_Benchmark'].values.tolist()
weight_diff=df_selected['Daily_Weight_Variance'].values.tolist()
#%%
def weight_bar_chart(sector,weight1,weight2,weight3):
    option3 = {
        "legend":{},
        "tooltip": {},
      "xAxis": {
        "type": 'category',
        "data": sector,
        "axisLabel": {
       "interval": 0,
       "rotate": 45,
     },
        "splitLine": {
        "show": False
      }
      },
      "yAxis": {
        "type": 'value'
      },
      "series": [
        {
          "data": weight1,  # Up until now p1 is pd.series
          "type": 'bar',
          "name": "Porfolio Weight",  # if you include legend, you must define the name of your data series
          "showBackground": True,
        },
        {
          "data": weight2,  # Up until now p1 is pd.series
          "type": 'bar',
          "name": "Benchmark Weight",  # if you include legend, you must define the name of your data series
          "showBackground": True,
        },
        {
          "data": weight3,  # Up until now p1 is pd.series
          "type": 'bar',
          "name": "Weight Difference",  # if you include legend, you must define the name of your data series
          "showBackground": True,
        }
      ]
    }
    
    return option3
#%%
#st_echarts(weight_bar_chart(gics,weight_ptf,weight_bmk,weight_diff),height=400, width=800)
#%% YTD_Portfolio VS YTD_Benchmark Data Preparation
epoch_year=date.today().year
year_start=date(epoch_year,1,1)
#year_start
query2 = '''select * from [dbo].[Performance_Attribution_By_Economic_Sector_Program]
where GICS_Economic_Sector_Code='TOTAL' and effective_date>=\'{0}\' '''.format(year_start)
    
query2 = query2.replace('\n',' ')
# This is to remove space in sql script

df_2 = pd.read_sql_query(query2,conn)
df_2_selected=df_2.loc[(df_2['Composite_ID']==comp_selected) & (df_2['Benchmark_ID']==bmk_selected)]
#df_2_selected

ytd_ptf=(df_2_selected['YTD_Return_Portfolio'].fillna(0)).values.tolist()
ytd_bmk=(df_2_selected['YTD_Return_Benchmark'].fillna(0)).values.tolist()
date=(df_2_selected['Effective_Date']).dt.strftime("%Y-%m-%d").tolist()

#%%
def ytd_line_chart(date,d1,d2):
    option2 = {
  "title": {
    "text": 'YTD of Portfolio and Benchmark'
  },
  "tooltip": {},
  "legend": {
      "left": 'right',
      "orient": 'vertical'
      },
  "xAxis": {
    "type": 'category',
    "data": date
  },
  "yAxis": {
    "type": 'value',
    "axisLabel": {
      "formatter": '{value} %'
    }
  },
  "series": [
    {
      "data": d1,
      "name": "Portfolio YTD",
      "type": 'line',
      "encode":{"x":"time","y":"index"}
    },
    {
      "data": d2,
      "name": "Benchmark YTD",
      "type": 'line',
      "encode":{"x":"time","y":"index"}
    }
  ]
}
            
    return option2
#%%
#st_echarts(ytd_line_chart(date,ytd_ptf,ytd_bmk))
#%% # Alpha data preparation
alpha_total=(df_2_selected['YTD_Total_Attribution'].fillna(0)).values.tolist()
alpha_allocation=(df_2_selected['YTD_Allocation_Effect'].fillna(0)).values.tolist()
alpha_selection=(df_2_selected['YTD_Selection_and_Interaction_Effect'].fillna(0)).values.tolist()
#alpha_currency=(df_2_selected['YTD_Currency_Effect']).values.tolist()
#%%
def alpha_line_chart(date,d1,d2,d3):
    option3 = {
  "title": {
    "text": 'Alpha Attribution'
  },
  "tooltip": {},
  "legend": {
      "left": 'right',
      "orient": 'vertical'
      },
  "xAxis": {
    "type": 'category',
    "data": date
  },
  "yAxis": {
    "type": 'value',
    "axisLabel": {
      "formatter": '{value} %'
    }
  },
  "series": [
    {
      "data": d1,
      "name": "Total Alpha",
      "type": 'line',
      "encode":{"x":"time","y":"index"}
    },
    {
      "data": d2,
      "name": "Alpha from Allocation",
      "type": 'line',
      "encode":{"x":"time","y":"index"}
    },
    {
      "data": d3,
      "name": "Alpha from Selection and Interaction",
      "type": 'line',
      "encode":{"x":"time","y":"index"}
    }
  ]
}
            
    return option3
#%%
#st_echarts(alpha_line_chart(date, alpha_total, alpha_allocation, alpha_selection))

#%% Layout Display Set Up 1
c21,c22 =st.columns(2)

with c21:
    # Pie Chart Display
    option_pie=sector_pie_chart(mv)
    st_echarts(option_pie,height=350, width=600)
    
with c22:
    # Bar Chart Display
    option_bar=weight_bar_chart(gics,weight_ptf,weight_bmk,weight_diff)
    st_echarts(option_bar)

#%% Layout Display Set Up 2
c21,c22 =st.columns(2)

with c21:
    # Pie Chart Display
    option_line1=ytd_line_chart(date,ytd_ptf,ytd_bmk)
    st_echarts(option_line1,height=400)
    
with c22:
    # Bar Chart Display
    option_line2=alpha_line_chart(date, alpha_total, alpha_allocation, alpha_selection)
    st_echarts(option_line2,height=400)
