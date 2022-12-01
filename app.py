from datetime import datetime
import time
from streamlit_sortables import sort_items
from streamlit_tags import st_tags
from fuzzywuzzy import process
import pandas as pd
import streamlit as st
import numpy as np
import openpyxl

# //TODO install warnings.warn('Using slow pure-python SequenceMatcher. Install python-Levenshtein to remove this warning')

# Set up page config
st.set_page_config(page_title="Forecast converter - JUYO", page_icon=":arrows_clockwise:", layout="wide", initial_sidebar_state="collapsed")

# Hide streamlit footer; (header {visibility hidden;})
hide_default_format = """
       <style>
       footer {visibility: hidden;}
       </style>
       """
st.markdown(hide_default_format, unsafe_allow_html=True)

def run_process():
    with st.spinner('runnning process...'):

        print(f"starting process...")
        
        iSegments = []
        iMonths = []
        iTerm = []
        iDays = []
        iSort = []
        iData = []
        sMonths = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Sept","Oct","Nov","Dec"]

        for x in sorted_items2: iSegments.append(x)
        for x in cols: iMonths.append(x)
        for x in terminology: iTerm.append(x)
        for x in iMonths:
            str2Match = x
            strOptions = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Sept","Oct","Nov","Dec"]

            highest = process.extractOne(str2Match,strOptions)
            highest = highest
            iDays.append(highest)
        
        for x in keywords: iSort.append(x)

        print(f"segments: {iSegments}")
        print(f"sheets: {iMonths}")
        print(f"terminology: {iTerm}")
        print(f"months: {iDays[:3]}") # Use if in to see if month contains. https://stackabuse.com/python-check-if-string-contains-substring/
        print(f"sorting: {iSort}")

        for x in range(len(iMonths)):

            wb = openpyxl.load_workbook(uploaded_file_CLIENT, data_only=True)

            ws = wb[(iMonths[x])]

            for z in sMonths:
                try:
                    iDays[x].index(z)
                except ValueError:
                    pass
                else:
                    tMonth = z

            print(f'Current month: {tMonth}')
            if tMonth == 'Jan' or tMonth == 'Mar' or tMonth == 'May' or tMonth == 'Jul' or tMonth == 'Aug' or tMonth == 'Oct' or tMonth == 'Dec': mDay = 31
            elif tMonth == 'Apr' or tMonth == 'Jun' or tMonth == 'Sep' or tMonth == 'Sept' or tMonth == 'Nov': mDay = 30
            elif tMonth == 'Feb': mDay = 28
            else: mDay = 30
            
            if storage == 'Rows':

                for i in range(1, ws.max_row + 1):
                    for j in range(1, ws.max_column + 1):
                        if i == int(row_n):
                            if "Rms" == ws.cell(i,j).value:
                                #print(f"found {ws.cell(i,j)}")

                                for row in ws.iter_rows(min_row=i + 1,max_row=i + mDay, min_col=j, max_col=j):
                                    for cell in row:
                                        iData.append(cell.value)# = [cell.value]
                                        #print(cell.value)
                                iData.append("Rms")
            else:
                
                for i in range(1, ws.max_row + 1):
                    for j in range(1, ws.max_column + 1):
                        if j == int(row_n):
                            if "Rms" == ws.cell(i,j).value:
                                print(f"found {ws.cell(i,j)}")


            iData.append(iMonths[x])
            print(iData)
    st.success('Done!')
    

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

            st.write(cols)
            
            try:

                st.write('You selected:', cols)

                # https://www.datacamp.com/tutorial/fuzzy-string-python
                str2Match = cols[0]
                strOptions = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Sept","Oct","Nov","Dec"]
                
                Ratios = process.extract(str2Match,strOptions)

                highest = process.extractOne(str2Match,strOptions)

                st.write(Ratios)
                st.write(highest)
            
            except:
                with st.empty():
                    if len(cols) == 0:
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
                st.markdown("map the segments is right order.")
                sorted_items2 = sort_items(keywords, direction='vertical')

        year1 = datetime.today().year

        year = st.select_slider(
            'Select starting year',
            options=range(year1 - 2, year1 + 2),value=year1)

        st.write(year)

        if year:
            terminology = st_tags(
            label='# Enter the terminology used in Excel file (e.g.: Rms & ZAR):',
            text='Press enter to add more',
            suggestions=['rn', 'RN', 'Rn', 
                        'Rev', 'REV', 'rev', 
                        'ADR', 'Adr', 'adr',
                        ],
            maxtags = 2,
            key='2')

            st.warning('  Always first the terminology of RoomNights, then REV or ADR!!!', icon="⚠️")
            st.write(len(terminology), ' terminology of ', 2, ' entered.')
            st.write(terminology)

        if len(terminology) == 2:
            option = st.radio(
                "is the data stored as Rn & Rev or Rn & ADR?",
                ('No disered option', 'Rn & Rev', 'Rn & ADR'))

            if option == 'Rn & Rev':
                st.write('You selected Rn & Rev.')
            elif option == 'Rn & ADR':
                st.write('You selected Rn & ADR.')
            else:
                st.write("You didn't select anything, if disered combination isn't present, please contact JUYO.")

            storage = st.radio(
                "is the data stored in rows or in columns?",
                ('Rows', 'Columns'))

            if storage == 'Rows':
                row_n = st.text_input("in which row can the terminology be found?")
                #row_n = row_n + 1
            else:
                row_n = st.text_input("in which column can the terminology be found?")
                row_n = ord(row_n) - 96
                print(row_n)

except:
    print('waiting...')

if st.button("Start converting process."): run_process()