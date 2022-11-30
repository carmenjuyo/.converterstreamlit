from datetime import datetime as dt
from streamlit_sortables import sort_items
from fuzzywuzzy import process
import pandas as pd
import streamlit as st
import numpy as np
import openpyxl

# Set up page config
st.set_page_config(page_title="Forecast converter - JUYO", page_icon=":arrows_clockwise:", layout="wide", initial_sidebar_state="collapsed")

# Hide streamlit footer
hide_default_format = """
       <style>
       footer {visibility: hidden;}
       </style>
       """
st.markdown(hide_default_format, unsafe_allow_html=True)

# Header of the page
with st.container():
    l_column, m_column, r_column = st.columns([2,4,1])
    with l_column:
        st.write("")
    with m_column:
        st.write(
            """
        # ♾️ Forecast / budget converter
        The process of converting the forecast file to the right format.
        """
        )
    with r_column:
        st.write("")

with st.container():

    st.write("---")
    disabled = 1

    left_column, right_column = st.columns(2)

    with left_column:

        st.header("Forecast File client")
        uploaded_file_CLIENT = st.file_uploader("Upload client file", type=".xlsx")

        use_example_file = st.checkbox(
            "Use example file", False, help="Use in-built example file to demo the app")

        if use_example_file:
            uploaded_file_CLIENT = 'Spier Budget Business Mix 2022I2023 (1).xlsx'

        if uploaded_file_CLIENT:
            
            df = pd.read_excel(uploaded_file_CLIENT)
            st.markdown("### Select wanted sheets for conversion.")
            tabs = pd.ExcelFile(uploaded_file_CLIENT).sheet_names

            cols = st.multiselect('select sheets:', tabs)
            
            try:
                st.write('You selected:', cols)

                # https://www.datacamp.com/tutorial/fuzzy-string-python
                str2Match = cols[0]
                strOptions = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Sept","Oct","Nov","Dec"]
                
                Ratios = process.extract(str2Match,strOptions)
                print(Ratios)

                highest = process.extractOne(str2Match,strOptions)
                print(highest)

                st.write(Ratios)
                st.write(highest)
            
            except:
                st.warning("No sheets selected.")

    with right_column:
        st.header("Format file JUYO")
        uploaded_file_JUYO = st.file_uploader("Upload JUYO file", type=".xlsx")

        use_example_file1 = st.checkbox(
            "Use example file1", False, help="Use in-built example file to demo the app")

        if use_example_file1:
            uploaded_file_JUYO = 'SPIER_1_MAJOR_DAILY (2).xlsx'

        if uploaded_file_JUYO:
            
            df1 = pd.read_excel(uploaded_file_JUYO)
            st.markdown("### Data preview")

            shape1 = df1.shape

            df1 = df1[[k for i, k in enumerate(df1.columns, 0) if i % 2 != 0]]

            shape = df1.shape

            st.write(shape[1], ' used of total columns: ', shape1[1])

            st.write(df1.columns)
            items = df1.columns.to_list()

            # https://github.com/ohtaman/streamlit-sortables
            sorted_items = sort_items(items, direction='vertical')

            lst = []
            x = sorted_items
            lst.append(x) 
            st.write(lst) 

st.write("---")