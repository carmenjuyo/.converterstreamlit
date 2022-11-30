from datetime import datetime
import time
from streamlit_sortables import sort_items
from streamlit_tags import st_tags
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

# The 2 different columns for the different files
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

                for col in cols:
                    str2Match = cols[col]
                    strOptions = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Sept","Oct","Nov","Dec"]
                    
                    highest = process.extractOne(str2Match,strOptions)
                    print(highest)
            
            except:
                with st.empty():
                    while len(cols) == 0:
                        st.subheader("⚠️ No sheets selected.")

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

st.write("---")

try:
    keywords = st_tags(
        label='# Enter segments in the order they appear in your Excel file:',
        text='Press enter to add more',
        suggestions=['leisure', 'Leisure', 'groups', 
                    'Groups', 'group', 'Group', 
                    'Business', 'business', 'corporate',
                    'Direct', 'direct', 'Indirect', 'indirect'
                    'individual', 'packages', 'complementary', 'house'
                    ],
        maxtags = shape[1],
        key='1')
    
    st.write(len(keywords), ' segments of ', shape[1], ' entered.')
    st.write(keywords)

    if len(keywords) == shape[1]:

        with st.container():

            st.write("---")

            left_column, right_column = st.columns(2)

            with left_column:
                # https://github.com/ohtaman/streamlit-sortables
                
                st.markdown("The segments on how the should be.")
                sorted_items1 = sort_items(df1.columns.to_list(), direction='vertical')

            with right_column:
                st.markdown("Reorder the segments is right order.")
                sorted_items2 = sort_items(keywords, direction='vertical')

        year1 = datetime.today().year

        year = st.select_slider(
            'Select starting year',
            options=range(year1 - 2, year1 + 2),value=year1)

        st.write(year)

        if year:
            terminology = st_tags(
            label='# Enter the terminology used in Excel file (e.g.: Rn & REV):',
            text='Press enter to add more',
            suggestions=['rn', 'RN', 'Rn', 
                        'Rev', 'REV', 'rev', 
                        'ADR', 'Adr', 'adr',
                        ],
            maxtags = 2,
            key='2')
        
            st.write(len(terminology), ' terminology of ', 2, ' entered.')
            st.write(terminology)

        if len(terminology) == 2:
            calculate = st.radio(
                "is the data stored as Rn & Rev or Rn & ADR?",
                ('Rn & Rev', 'Rn & ADR', 'No disered option'))

            if calculate == 'Rn & Rev':
                st.write('You selected Rn & Rev.')
            elif calculate == 'Rn & ADR':
                st.write('You selected Rn & ADR.')
            else:
                st.write("You didn't select anything, if disered combination isn't present, please contact JUYO.")

except:
    print('waiting...')