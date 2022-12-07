from datetime import datetime
from streamlit_sortables import sort_items
from streamlit_tags import st_tags
from fuzzywuzzy import process
from fuzzywuzzy import fuzz
import pandas as pd
import streamlit as st
import openpyxl
import traceback
import json

# Set up page config
st.set_page_config(page_title="Forecast converter - JUYO", page_icon=":arrows_clockwise:", layout="wide")#, initial_sidebar_state="collapsed")
#header {visibility: hidden;}
hide_default_format = """
       <style>
       footer {visibility: hidden;}
       
       </style>
       """
st.markdown(hide_default_format, unsafe_allow_html=True)

def check_segments():
    
    juyo_segments = []
    client_segments = []
    segments_score = []

    for x in sorted_items1: juyo_segments.append(x.lower())
    for x in sorted_items2: client_segments.append(x.lower())

    for x in range(len(client_segments)):

        str2Match = juyo_segments[x]
        strOptions = client_segments[x]

        highest =  fuzz.ratio(str2Match,strOptions)
        
        print(highest)


def save_storage():

    print('Start storage...')

    data_storage = {}

    iSegments = []
    iTerm = []
    iSort = []
    iSkipper = []
    iSkipper1 = []
    iDataSt = []
    iLoc = []

    for x in sorted_items2: iSegments.append(x)
    for x in terminology: iTerm.append(x)
    for x in keywords: iSort.append(x)
    for x in Skipper: iSkipper.append(x)
    for x in Skipper1: iSkipper1.append(x)

    if option == 'Rn & Rev':
        iDataSt.append('Rev')
    elif option == 'Rn & ADR':
        iDataSt.append('ADR')

    if storage == 'Rows':
        iLoc.append('Rows')
        iLoc.append(row_n)
    else:
        iLoc.append('Columns')
        iLoc.append(row_n)

    data_storage['iSegments'] = iSegments
    data_storage['iTerm'] = iTerm
    data_storage['iSort'] = iSort
    data_storage['iSkipper'] = iSkipper
    data_storage['iSkipper1'] = iSkipper1
    data_storage['iDataSt'] = iDataSt
    data_storage['iLoc'] = iLoc

    json_string = json.dumps(data_storage)

    with st.sidebar:
        st.write("""
            # Here you can download your input for future use.
             Check if the data is complete and then download the json file.
             """)

        st.json(json_string, expanded=False)

        st.download_button(
            label="Download JSON storage file.",
            file_name=f"{uploaded_file_JUYO[:5]}_data_JUYO.json",
            mime="application/json",
            data=json_string,
        ) 

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
        json_storage = []
        row_n_storage = []
        Skipper_s = []
        Skipper_s1 = []
        Storage_1 = []
        sMonths = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Sept","Oct","Nov","Dec"]

        if stro == 'No':
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

            row_n_storage.append(row_n)
            json_storage.append(storage)

            for x in Skipper: Skipper_s.append(x)
            Skipper_s = [round(float(i)) for i in Skipper_s]
            for x in Skipper1: Skipper_s1.append(x)
            Skipper_s1 = [round(float(i)) for i in Skipper_s1]

        elif stro == 'Yes': 

            for x in data_json['iSegments'].dropna(): iSegments.append(x)
            for x in cols: iMonths.append(x)
            for x in data_json['iTerm'].dropna(): iTerm.append(x)
            for x in iMonths:
                str2Match = x
                strOptions = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Sept","Oct","Nov","Dec"]

                highest = process.extractOne(str2Match,strOptions)
                highest = highest
                iDays.append(highest)
            
            for x in data_json['iSort'].dropna(): iSort.append(x)

            # //TODO Kan inprincipe weg. Omdat iLoc[0] altijd nu aangeeft was het is.
            if data_json['iLoc'].iloc[0] == "Rows":

                json_storage.append(data_json['iLoc'].iloc[0])
                row_n_storage.append(data_json['iLoc'].iloc[1])

            elif data_json['iLoc'].iloc[0] == 'Columns':
                
                json_storage.append(data_json['iLoc'].iloc[0])
                row_n_storage.append(data_json['iLoc'].iloc[1])

            else:
                st.warning('Error at iLoc process')
                return
            
            for x in data_json['iSkipper'].dropna(): Skipper_s.append(x)
            Skipper_s = [round(float(i)) for i in Skipper_s]
            for x in data_json['iSkipper1'].dropna(): Skipper_s1.append(x)
            Skipper_s1 = [round(float(i)) for i in Skipper_s1]

        else:
            return

        print(f"iSegments: {iSegments}")
        print(f"sheets (iMonths): {iMonths}")
        print(f"terminology (iTerm): {iTerm}")
        print(f"months (iDays): {iDays}")
        print(f"sorting (iSort): {iSort}")
        
        wb = openpyxl.load_workbook(uploaded_file_CLIENT, data_only=True)
        
        for x in range(len(iMonths)):
            
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

            if json_storage[0] == 'Rows':

                for i in range(1, ws.max_row + 1):
                    for j in range(1, ws.max_column + 1):
                        if i == int(row_n_storage[0]):
                            if iTerm[0] == ws.cell(i,j).value:
                                rn_c = rn_c + 1
                                if int(rn_c) in Skipper_s:
                                    print(f"{rn_c} is in list: {Skipper_s} so skip")
                                    
                                else:
                                    rn_c1 = rn_c1 + 1
                                    cRngn.append(f"{iTerm[0]} {ws.cell(i,j)}")
                                    for row in ws.iter_rows(min_row=i + 1,max_row=i + mDay, min_col=j, max_col=j):
                                        for cell in row:
                                            iDataRn.append(cell.value)
                                    
                                    iDataRnT.append(iDataRn)
                                    iDataRn = []
                
                if rn_c1 != len(iSegments):

                    st.error(f"""
                        ##### ERROR for: {iTerm[0]}. In total {len(iSegments)} segmets entered. But {rn_c} segments were measured in the month / sheet: {iMonths[x]}.
                        See below an overview of the segments and their range that were succeeded:
                    """, 
                    icon="❌")
                    st.json(cRngn, expanded=True)
                    return

                for i in range(1, ws.max_row + 1):
                    for j in range(1, ws.max_column + 1):
                        if i == int(row_n_storage[0]):
                            if iTerm[1] == ws.cell(i,j).value:
                                rv_c = rv_c + 1
                                if int(rv_c) in Skipper_s1:
                                    print(f"{rn_c} is in list: {Skipper_s1} so skip")
                                    
                                else:
                                    rv_c1 = rv_c1 + 1
                                    cRngv.append(f"{iTerm[0]} {ws.cell(i,j)}")
                                    for row in ws.iter_rows(min_row=i + 1,max_row=i + mDay, min_col=j, max_col=j):
                                        for cell in row:
                                            iDataRv.append(round(cell.value,2))
                                    
                                    iDataRvT.append(iDataRv)
                                    iDataRv = []

                if rv_c1 != len(iSegments):

                    st.error(f"""
                        ##### ERROR for: {iTerm[1]}. In total {len(iSegments)} segmets entered. But {rv_c} segments were measured in the month / sheet: {iMonths[x]}.
                        See below an overview of the segments and their range that were succeeded:
                    """, 
                    icon="❌")
                    st.json(cRngv, expanded=True)
                    return
                    
            else:
                
                for i in range(1, ws.max_row + 1):
                    for j in range(1, ws.max_column + 1):
                        if j == int(row_n_storage[0]):
                            if iTerm[0] == ws.cell(i,j).value:
                                rn_c = rn_c + 1
                                if int(rn_c) in Skipper_s:
                                    print(f"{rn_c} is in list: {Skipper_s} so skip")
                                    
                                else:
                                    rn_c1 = rn_c1 + 1
                                    cRngn.append(f"{iTerm[0]} {ws.cell(i,j)}")
                                    for column in ws.iter_cols(min_row=i,max_row=i, min_col=j + 1, max_col=j + mDay):
                                        for cell in column:
                                            iDataRn.append(cell.value)
                                    
                                    iDataRnT.append(iDataRn)
                                    iDataRn = []

                if rn_c1 != len(iSegments):
                    
                    st.error(f"""
                        ##### ERROR for: {iTerm[0]}. In total {len(iSegments)} segmets entered. But {rn_c} segments were measured in the month / sheet: {iMonths[x]}.
                        See below an overview of the segments and their range that were succeeded:
                    """, 
                    icon="❌")
                    st.json(cRngn, expanded=True)
                    return

                for i in range(1, ws.max_row + 1):
                    for j in range(1, ws.max_column + 1):
                        if j == int(row_n_storage[0]):
                            if iTerm[1] == ws.cell(i,j).value:
                                rv_c = rv_c + 1
                                if int(rv_c) in Skipper_s1:
                                    print(f"{rv_c} is in list: {Skipper_s1} so skip")

                                else:
                                    rv_c1 = rv_c1 + 1
                                    cRngv.append(f"{iTerm[1]} {ws.cell(i,j)}")
                                    for column in ws.iter_cols(min_row=i,max_row=i, min_col=j + 1, max_col=j + mDay):
                                        for cell in column:
                                            iDataRv.append((cell.value))
                                    
                                    iDataRvT.append(iDataRv)
                                    iDataRv = []

                if rv_c1 != len(iSegments):
                    
                    st.error(f"""
                        ##### ERROR for: {iTerm[1]}. In total {len(iSegments)} segmets entered. But {rv_c} segments were measured in the month / sheet: {iMonths[x]}.
                        See below an overview of the segments and their range that were succeeded:
                    """, 
                    icon="❌")
                    st.json(cRngv, expanded=True)
                    return
                    
        print("---")

        a = 0
        for x in range(len(iMonths)):

            for i in range(len(iSegments)):

                strFind = iSegments[i]

                for y in range(len(iSort)):

                    strStored = iSort[y]

                    #print(f"L: {(iSegments[i])} | R: {(iSort[y])}")

                    if strFind == strStored:
                        
                        if i == y:
                            #print(f"-- Level & Name match: {strFind} = {i}")
                            pass
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
            else:
                a = a + len(iSort)

            iSort.clear()
            if stro == 'No':
                for z in keywords: iSort.append(z)
            elif stro == 'Yes':
                for z in data_json['iSort'].dropna(): iSort.append(z)

        st.json(iDataRnT, expanded=False)

        d =[]

        for x in range(len(iDataRnT)):
            d.append(iDataRnT[x])
            d.append(iDataRvT[x])

        df2 = pd.DataFrame(data=d)

        df2 = df2.T

        st.dataframe(df2)

        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title='test'

        y = 1
        for x in range(len(all_columns)):
            sheet.cell(row=1, column=y).value=all_columns[x]
            y = y + 1

        x = 2 # COLUMN
        y = 2 # ROW
        t = 2 # ROW
        s = 1 # SEGMENTS
        
        # //FIXME Hier gaat iets mis wanneer er meerdere maanden zijn!
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
    l_column, m_column, r_column = st.columns([1.6,2.5,1])
    with l_column:
        st.write("")
    with m_column:
        st.write(
            """
        # ♾️Forecast / budget converter
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

        use_example_file_R = st.checkbox(
            "Use example fileR", False, help="Use in-built example file to demo the app")
        use_example_file_C = st.checkbox(
            "Use example fileC", False, help="Use in-built example file to demo the app")

        if use_example_file_R:
            uploaded_file_CLIENT = 'Spier Budget Business Mix 2022I2023 (1).xlsx'
        if use_example_file_C:
            uploaded_file_CLIENT = 'CHLEP - Forecast 2022-23.xlsx'

        if uploaded_file_CLIENT:

            st.markdown("### Select wanted sheets for conversion.")

            tabs = pd.ExcelFile(uploaded_file_CLIENT).sheet_names

            cols = st.multiselect('Select sheets in the chronological order of the months:', tabs)

            st.write('You selected:', cols) 

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
            st.markdown("### Segments in Juyo")

            shape1 = df1.shape

            df1 = df1[[k for i, k in enumerate(df1.columns, 0) if i % 2 != 0]]

            shape = df1.shape

            st.write(shape[1], ' segments detected in Juyo file')
            
            st.write(df1.columns.to_list())
        
    if uploaded_file_JUYO:

        st.write("---")

        st.write("## Reload JSON file from previous input?")
        stro = st.radio(
            label="-",
            options=("No", "Yes"),
            horizontal=True,
            index=0
            )

        if stro == 'No':
            try:
                st.write("---")

                keywords = st_tags(
                    label="""
                        # Enter your segments in the exact order they appear in your Excel file:
                        So go to your sheet to where the segments with data is stored, 
                        and then go from left to right or top to bottom and then fill in all your segments 
                        here in that order.
                            """,
                    
                    text='Press enter to add more',
                    suggestions=['leisure', 'Leisure', 'groups', 
                                'Groups', 'group', 'Group', 
                                'Business', 'business', 'corporate',
                                'Direct', 'direct', 'Indirect', 'indirect'
                                'individual', 'packages', 'complementary', 'house'
                                ],
                    maxtags = shape[1],
                    key='1')

                if len(keywords) == shape[1] - 1:
                    st.warning('Do you miss 1 segment that is placed on the end of the JUYO segments? Read this 👇', icon="⚠️")
                    st.write("""
                        If you are 1 segment short, and that segment happens to be last, there is a possibility.
                        When the missing segment is at the last, this process can recognize it and leave it empty, automatically becoming zero in JUYO. Click the checkbox if so.
                        """)
                    e_segments = st.checkbox('Extra (empty!) segment on last place?')
                    ("---")

                st.write(len(keywords), ' segments of ', shape[1], ' entered.')
                st.write(keywords)

                if len(keywords) == shape[1] or e_segments:

                    with st.container():

                        st.write("---")

                        left_column, right_column = st.columns(2)

                        with left_column:
                            # https://github.com/ohtaman/streamlit-sortables
                            
                            st.markdown("### Segments in the order they should be.")
                            sorted_items1 = sort_items(df1.columns.to_list(), direction='vertical')

                        with right_column:
                            st.markdown("### Map the segments so that they match the segments on the left!")
                            sorted_items2 = sort_items(keywords, direction='vertical')

                        if st.button('checksegments'): check_segments()

                        st.write('## Select starting year of first sheet.')
                        year = st.select_slider(
                            label=".",
                            options=range(datetime.today().year - 2, datetime.today().year + 3),value=datetime.today().year)

                        st.write(year)

                        terminology = st_tags(
                            label="""
                                # Enter the terminology used in Excel file for roomnights and revenue:
                                For example; roonnights = Rms, Rn, etc. Revenue = Rev, Rvu, etc.
                                """,
                            text='Press enter to add more',
                            suggestions=['rn', 'RN', 'Rn', 
                                        'Rev', 'REV', 'rev', 
                                        'ADR', 'Adr', 'adr',
                                        ],
                            maxtags = 2,
                            key='2')

                        st.warning(f"""
                            ##### Always put the terminology of roomnights as first!, then revenue or ADR! 
                            {len(terminology)} terminology of {2} entered.
                            """, 
                            icon="⚠️")

                        if len(terminology) == 2:
                            st.write("")
                            st.write("##### is the data stored as Rn & Rev or Rn & ADR? (DOES NOT WORK CURRENTLY)")
                            option = st.radio(
                                label=".",
                                options=('No disered option', 'Rn & Rev', 'Rn & ADR'))

                            if option == 'Rn & Rev':
                                st.write('You selected Rn & Rev.')
                            elif option == 'Rn & ADR':
                                st.write('You selected Rn & ADR.')
                            else:
                                st.write("You didn't select anything, if disered combination isn't present, please contact JUYO.")

                            skip_term = st.checkbox('Want to skip terminology on certain places?')

                            if skip_term:

                                l_c, r_c = st.columns(2)

                                with l_c:
                                    Skipper = st_tags(
                                        label='### Skip a terminology in order (Rms)',
                                        text='Press enter to add more',
                                        suggestions=['1', '2', '3', 
                                                    '4', '5', '6'],
                                        maxtags = 5,
                                        key='3')

                                with r_c:
                                    Skipper1 = st_tags(
                                        label='### Skip a terminology in order (Rev)',
                                        text='Press enter to add more',
                                        suggestions=['1', '2', '3', 
                                                    '4', '5', '6'],
                                        maxtags = 5,
                                        key='4')

                            else:
                                Skipper = []
                                Skipper1 = []
                            
                            storage = st.radio(
                                label=" is the data stored in rows or in columns?",
                                options=('Rows', 'Columns'))
                        
                            if storage == 'Rows':
                                row_n = st.text_input("in which row can the terminology be found?")
                                if row_n:
                                    if st.checkbox("want to store the input for future reference?"):
                                        if st.button('store data'): 
                                            save_storage()
                                            if st.button("Start converting process."): run_process()
                                    if st.button("Start converting process."): run_process()
                            else:
                                row_n = st.text_input("in which column can the terminology be found?")
                                row_n = ord(row_n) - 96
                                if row_n:
                                    if st.checkbox("want to store the input for future reference?"):
                                        if st.button('stora date '): 
                                            save_storage()
                                            if st.button("Start converting process."): run_process()
                                    if st.button("Start converting process."): run_process()
                                print(row_n)

            except Exception:
                traceback.print_exc()

        elif stro == 'Yes':

            uploaded_file_JSON = st.file_uploader("Upload json file", type=".json")

            if uploaded_file_JSON:
                data_json = pd.read_json(uploaded_file_JSON, orient='index')
                data_json = data_json.transpose()
                data_json.dropna()

                st.write('## Select starting year of first sheet.')

                year = st.select_slider(
                    label=".",
                    options=range(datetime.today().year - 2, datetime.today().year + 3),value=datetime.today().year)

                if st.button("Start converting process."): run_process()

        else:
            st.write('please select a option.')