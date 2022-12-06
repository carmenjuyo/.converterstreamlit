from datetime import datetime
import time
from openpyxl.utils.cell import get_column_letter
from streamlit_sortables import sort_items
from streamlit_tags import st_tags
from fuzzywuzzy import process
import pandas as pd
import streamlit as st
import numpy as np
import openpyxl
from io import BytesIO
import xlsxwriter
import traceback
import json

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

        rn_c = 0
        rv_c = 0
        cRngn = []
        cRngv = []

        iSegments = []
        iMonths = []
        iTerm = []
        iDays = []
        iSort = []
        iDataRn = []
        iDataRnT = []
        iDataRv = []
        iDataRvT = []
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

        print(f"iSegments: {iSegments}")
        print(f"sheets (iMonths): {iMonths}")
        print(f"terminology (iTerm): {iTerm}")
        print(f"months (iDays): {iDays}")
        print(f"sorting (iSort): {iSort}")

        for x in range(len(iMonths)):

            wb = openpyxl.load_workbook(uploaded_file_CLIENT, data_only=True)

            ws = wb[(iMonths[x])]

            rn_c = 0
            rv_c = 0
            rn_c1 = 0
            rv_c1 = 0

            for z in sMonths:
                try:
                    iDays[x].index(z)
                except ValueError:
                    pass
                else:
                    tMonth = z
                    if x == 0: eMonth = z

            print(f'Current month: {tMonth}')
            if tMonth == 'Jan' or tMonth == 'Mar' or tMonth == 'May' or tMonth == 'Jul' or tMonth == 'Aug' or tMonth == 'Oct' or tMonth == 'Dec': mDay = 31
            elif tMonth == 'Apr' or tMonth == 'Jun' or tMonth == 'Sep' or tMonth == 'Sept' or tMonth == 'Nov': mDay = 30
            elif tMonth == 'Feb': mDay = 28
            else: mDay = 30

            if storage == 'Rows':

                for i in range(1, ws.max_row + 1):
                    for j in range(1, ws.max_column + 1):
                        if i == int(row_n):
                            if terminology[0] == ws.cell(i,j).value:
                                rn_c = rn_c + 1
                                if str(rn_c) in Skipper:
                                    print(f"{rn_c} is in list: {Skipper} so skip")
                                    break
                                else:
                                    pass
                                cRngn.append(f"{terminology[0]} {ws.cell(i,j)}")
                                for row in ws.iter_rows(min_row=i + 1,max_row=i + mDay, min_col=j, max_col=j):
                                    for cell in row:
                                        iDataRn.append(cell.value)
                                
                                iDataRnT.append(iDataRn)
                                iDataRn = []
                                
                            if rn_c == len(iSegments):
                                break
                
                if rn_c != len(iSegments):

                    st.error(f"""
                        ERROR! In total there are {len(iSegments)} entered. But {rn_c} segments were measured in the month / sheet: {iMonths[x]}.
                        See below an overview of the segments and their range that were succeeded:
                    """, icon="❌")
                    st.json(cRngn, expanded=True)
                    return

                for i in range(1, ws.max_row + 1):
                    for j in range(1, ws.max_column + 1):
                        if i == int(row_n):
                            if terminology[1] == ws.cell(i,j).value:
                                rv_c = rv_c + 1
                                if str(rn_c) in Skipper:
                                    print(f"{rn_c} is in list: {Skipper} so skip")
                                    break
                                else:
                                    pass
                                    cRngv.append(f"{terminology[0]} {ws.cell(i,j)}")
                                for row in ws.iter_rows(min_row=i + 1,max_row=i + mDay, min_col=j, max_col=j):
                                    for cell in row:
                                        iDataRv.append(round(cell.value,2))
                                
                                iDataRvT.append(iDataRv)
                                iDataRv = []

                            if rv_c == len(iSegments):
                                break

                if rv_c != len(iSegments):

                    st.error(f"""
                        ERROR! In total there are {len(iSegments)} entered. But {rv_c} segments were measured in the month / sheet: {iMonths[x]}.
                        See below an overview of the segments and their range that were succeeded:
                    """, icon="❌")
                    st.json(cRngv, expanded=True)
                    return
                    
            else:
                
                for i in range(1, ws.max_row + 1):
                    for j in range(1, ws.max_column + 1):
                        if j == int(row_n):
                            if terminology[0] == ws.cell(i,j).value:
                                rn_c = rn_c + 1
                                if str(rn_c) in Skipper:
                                    print(f"{rn_c} is in list: {Skipper} so skip")
                                    
                                else:
                                    rn_c1 = rn_c1 + 1
                                    print(f"{terminology[0]} {ws.cell(i,j)}")
                                    cRngn.append(f"{terminology[0]} {ws.cell(i,j)}")
                                    for column in ws.iter_cols(min_row=i,max_row=i, min_col=j + 1, max_col=j + mDay):
                                        for cell in column:
                                            iDataRn.append(cell.value)
                                    
                                    iDataRnT.append(iDataRn)
                                    iDataRn = []

                if rn_c1 != len(iSegments):
                    
                    st.error(f"""
                        ERROR!1 In total there are {len(iSegments)} entered. But {rn_c1} segments were measured in the month / sheet: {iMonths[x]}.
                        See below an overview of the segments and their range that were succeeded:
                    """, icon="❌")
                    st.json(cRngn, expanded=True)
                    return

                for i in range(1, ws.max_row + 1):
                    for j in range(1, ws.max_column + 1):
                        if j == int(row_n):
                            if terminology[1] == ws.cell(i,j).value:
                                rv_c = rv_c + 1
                                if str(rv_c) in Skipper1:
                                    print(f"{rv_c} is in list: {Skipper1} so skip")

                                else:
                                    rv_c1 = rv_c1 + 1
                                    cRngv.append(f"{terminology[1]} {ws.cell(i,j)}")
                                    print(f"{terminology[1]} {ws.cell(i,j)}")
                                    for column in ws.iter_cols(min_row=i,max_row=i, min_col=j + 1, max_col=j + mDay):
                                        for cell in column:
                                            iDataRv.append((cell.value))
                                    
                                    iDataRvT.append(iDataRv)
                                    iDataRv = []

                if rv_c1 != len(iSegments):
                    
                    st.error(f"""
                        ERROR!2 In total there are {len(iSegments)} entered. But {rv_c1} segments were measured in the month / sheet: {iMonths[x]}.
                        See below an overview of the segments and their range that were succeeded:
                    """, icon="❌")
                    st.json(cRngv, expanded=True)
                    
        print("---")

        a = 0
        for x in range(len(iMonths)):
            
            for i in range(len(iSegments)):

                strFind = iSegments[i]

                for y in range(len(iSort)):

                    strStored = iSort[y]

                    print(f"L: {(iSegments[i])} | R: {(iSort[y])}")

                    if strFind == strStored:
                        
                        if i == y:
                            print(f"-- Level & Name match: {strFind} = {i}")
                        
                        else:
                            print(f"- Name match: {strFind} : {i} - {y}")
                            print(f"- reordering...")
                            
                            arrTemp = iDataRnT[y + a]
                            arrTempV = iDataRvT[y + a]

                            iDataRnT[y + a] = iDataRnT[i + a]
                            iDataRvT[y + a] = iDataRvT[i + a]

                            iDataRnT[i + a] = arrTemp
                            iDataRvT[i + a] = arrTempV

                            temp = iSort[i]

                            iSort[i] = iSegments[i]
                            iSort[y] = temp
            
            if x == 0:
                a = len(iSort)
                print(f"a = {a}")
            else:
                a = a + len(iSort)
                print(f"a = {a}")

            iSort.clear()
            for z in keywords: iSort.append(z)

        st.json(iDataRnT, expanded=False)

        d =[]

        for x in range(len(iDataRnT)):
            d.append(iDataRnT[x])
            d.append(iDataRvT[x])

        df2 = pd.DataFrame(data=d)

        df2 = df2.T

        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title='test'

        y = 1
        for x in range(len(all_columns)):
            sheet.cell(row=1, column=y).value=all_columns[x]
            y =y + 1

        x = 2 # COLUMN
        y = 2 # ROW
        t = 2 # ROW
        s = 1 # SEGMENTS
    
        for i in range(len(iDataRnT)):

            for z in range(len(iDataRnT[i])):

                sheet.cell(row=y, column=x).value=iDataRnT[i][z]
                sheet.cell(row=y, column=x+1).value=iDataRvT[i][z]
                y = y + 1

            if s == len(iSegments):
                x = 2
                s = 1
                t = 2 + (len(iDataRnT[i]))
            else:
                s = s + 1
                x = x + 2
                y = t 

        datetime_object = datetime.strptime(eMonth, "%b")
        month_number = datetime_object.month
        print(month_number)

        datelist = pd.date_range(datetime(year, month_number, 1), periods=sheet.max_row - 1).to_pydatetime().tolist()
        
        i = 2
        for x in range(len(datelist)):
            sheet.cell(row=i, column=1).value=datelist[x]
            i = i + 1

        for cell in sheet["A"]:
            cell.number_format = "YYYY/MM/DD"
        
        wb.save('NameFile.xlsx')

        df3 = pd.read_excel('NameFile.xlsx')

        st.dataframe(df3)

    st.success('Process ran!')

    with open("NameFile.xlsx", "rb") as file:
        st.download_button(
            label="click me to download excel",
            data=file,
            file_name=uploaded_file_JUYO,
            mime="application/octet-stream"
            )
    

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

        use_example_file_R = st.checkbox(
            "Use example fileR", False, help="Use in-built example file to demo the app")
        use_example_file_C = st.checkbox(
            "Use example fileC", False, help="Use in-built example file to demo the app")

        if use_example_file_R:
            uploaded_file_CLIENT = 'Spier Budget Business Mix 2022I2023 (1).xlsx'
        if use_example_file_C:
            uploaded_file_CLIENT = 'CHLEP - Forecast 2022-23.xlsx'

        if uploaded_file_CLIENT:

            df = pd.read_excel(uploaded_file_CLIENT)
            st.markdown("### Select wanted sheets for conversion.")
            tabs = pd.ExcelFile(uploaded_file_CLIENT).sheet_names

            cols = st.multiselect('select sheets:', tabs)

            st.write(cols)
            
            try:

                st.write('You selected:', cols)

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

        use_example_file_R1 = st.checkbox(
            "Use example fileR1", False, help="Use in-built example file to demo the app")
        
        use_example_file_C1 = st.checkbox(
            "Use example fileC1", False, help="Use in-built example file to demo the app")

        if use_example_file_R1:
            uploaded_file_JUYO = 'SPIER_1_MAJOR_DAILY (2).xlsx'

        if use_example_file_C1:
            uploaded_file_JUYO = 'MTELP_2_MINOR_DAILY (3).xlsx'

        if uploaded_file_JUYO:
            
            df1 = pd.read_excel(uploaded_file_JUYO)

            all_columns = df1.columns
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

    if len(keywords) == shape[1] - 1:
        st.warning('Do you miss 1 segment that is placed on the end of the JUYO segments? Read this 👇')
        st.write("""
            If you miss a segments in your segments file, and it is placed on the end of the JUYO segments file.

            You can check this checkbox true, and the last segments  will be filled in with zero's.
            """)
        e_segments = st.checkbox('Extra (empty!) segment on last place?')

    if len(keywords) == shape[1] or e_segments:

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
                "is the data stored as Rn & Rev or Rn & ADR? (DOES NOT WORK CURRENTLY)",
                ('No disered option', 'Rn & Rev', 'Rn & ADR'))

            if option == 'Rn & Rev':
                st.write('You selected Rn & Rev.')
            elif option == 'Rn & ADR':
                st.write('You selected Rn & ADR.')
            else:
                st.write("You didn't select anything, if disered combination isn't present, please contact JUYO.")

            Skipper = st_tags(
                label='# TEST!!! to skip terminology:',
                text='Press enter to add more',
                suggestions=['1', '2', '3', 
                            '4', '5', '6'],
                maxtags = 5,
                key='3')

            Skipper1 = st_tags(
                label='# TEST!!! to skip terminology: Colums',
                text='Press enter to add more',
                suggestions=['1', '2', '3', 
                            '4', '5', '6'],
                maxtags = 5,
                key='4')

            storage = st.radio(
                "is the data stored in rows or in columns?",
                ('Rows', 'Columns'))

            if storage == 'Rows':
                row_n = st.text_input("in which row can the terminology be found? (press enter when ready)")
                #row_n = row_n + 1
                if row_n:
                    if st.checkbox("want to store the input for future reference?"):
                        if st.button("Start converting process."): run_process()
            else:
                row_n = st.text_input("in which column can the terminology be found?")
                row_n = ord(row_n) - 96
                if row_n:
                    if st.checkbox("want to store the input for future reference? (press enter when ready)"):
                        if st.button("Start converting process."): run_process()
                print(row_n)

except Exception:
    traceback.print_exc()