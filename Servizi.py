import pandas as pd
import streamlit as st
import plotly.express as px
from PIL import Image


st.set_page_config(page_title='Upspace Report', page_icon=":bar_chart:", layout='wide')
st.title(":bar_chart: Upspace Dashboard :rocket:")

st.markdown("##")
st.header("Seleziona il report dal menù a tendina :open_file_folder:")
st.sidebar.success("select report above.")


with st.expander("Report Servizi"):
       st.markdown("# Report Servizi :chart_with_upwards_trend:")
       
       st.markdown('##')
       st.markdown('##')
       
       ### --- LOAD DATAFRAME
       excel_file = 'Upspace.xlsx'
       sheet_name = 'Servizi'
       df = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='L:M', header=1, nrows=12)
       df1 = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='P:Q', header=1, nrows=12)
       df2 = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='T:U', header=1, nrows=12)
       df3 = pd.read_excel(excel_file, sheet_name=sheet_name, usecols='E:I', header=0, nrows=38)
       
       ### --- KPI
       sales = int(df3["Vendita"].sum())
       ore = int(df3["Ore lavorate"].sum())
       servizi = int(df3["Quantità"].sum())
       left_column, middle_column, right_column = st.columns(3)
       with left_column:
                 st.subheader('Totale servizi in Euro')
                 st.header(f'€ {sales:,}')
                 
       with middle_column:
                   st.subheader('Totale ore lavorate')
                   st.header(f' {ore:,}')
                   
       with right_column:
            st.subheader('Quantità servizi venduti')
            st.header(servizi)
            
       st.markdown("---")
            
       
       bar_chart = px.bar(df, x='Quantità', y='Servizio', title= 'Classsifica Servizi per quantità venduta', orientation="h", color_discrete_sequence=["#0083B8"]*len(df),
        template='plotly_white')
       
       
    
       
       bar_chart1 = px.bar(df1, x='Margine', y='Servizio.1', title= 'Classsifica Margine per servizi', orientation="h", color_discrete_sequence=["#0083B8"]*len(df),
        template='plotly_white')
       
       left_column2, right_column2 = st.columns(2)    
       left_column2.plotly_chart(bar_chart, use_container_width=True)
       right_column2.plotly_chart(bar_chart1, use_container_width=True)
       
       
       st.markdown("---")
       
       left_column1, middle_column1, right_column1 = st.columns(3)

       with left_column1:    
        st.subheader('Classifica servizi per quantità venduta')
        st.dataframe(df)


       with middle_column1:
        st.subheader('Classifica servizi in ordine di margine')
        st.dataframe(df1)
    


       with right_column1:
        st.subheader('Classifica ore lavorate per ogni servizio')
        st.dataframe(df2)

