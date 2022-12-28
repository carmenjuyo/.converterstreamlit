import sys, os
from datetime import datetime
import json
from streamlit_sortables import sort_items
from streamlit_tags import st_tags
from fuzzywuzzy import process
import pandas as pd
import streamlit as st
import openpyxl
import traceback
import gspread
import random
import string
from PIL import Image
import streamlit_analytics

# Explanation of script

# This script is build with the help of Streamlit (https://streamlit.io/)

# The script is set-up with 3 different functions and at the button of the script the page lay-out is defined.

# In this script, the store_input() function is used to store the input of the clients in some form of data structure, such as a dictionary. 
# The run_credentials() function is used to check the validity of the client's input, such as by verifying their credentials or checking for required fields. 
# The run_process() function is used to run the entire process, such as by performing the converting process on the client's input.

# At the bottom of the script, the page layout is defined using the st.container() and st.button() functions from Streamlit.
# When the client clicks the button, the input is stored using the store_input() function, the credentials are checked using the run_credentials() function, 
# and the process is run using the run_process() function. The results of running the process are then displayed on the page.

# Overall, this script provides a step-by-step procedure for clients to complete in order to input their data, 
# have their credentials verified, and see the results of running the process.

# Set up page config

im = Image.open('images/JUYOcon.ico')
st.set_page_config(page_title="Forecast converter - JUYO", page_icon=im, layout="wide")

hide_default_format = """
       <style>
       footer {visibility: hidden;}
       header {visibility: hidden;}
       </style>
       """
st.markdown(hide_default_format, unsafe_allow_html=True)

st.experimental_get_query_params()

if str(st.experimental_user.email) == st.secrets["user_email_lh"] or st.experimental_user == st.secrets["user_email_dev"]:
    st.experimental_set_query_params(analytics=st.secrets["query_params_settings"]["setting_1"])
else:
    st.experimental_get_query_params()

def run_credentials():
    credentials = {
            "type": st.secrets["gcp_service_account"]["type"],
            "project_id": st.secrets["gcp_service_account"]["project_id"],
            "private_key_id": st.secrets["gcp_service_account"]["private_key_id"],
            "private_key": st.secrets["gcp_service_account"]["private_key"],
            "client_email": st.secrets["gcp_service_account"]["client_email"],
            "client_id": st.secrets["gcp_service_account"]["client_id"],
            "auth_uri": st.secrets["gcp_service_account"]["auth_uri"],
            "token_uri": st.secrets["gcp_service_account"]["token_uri"],
            "auth_provider_x509_cert_url": st.secrets["gcp_service_account"]["auth_provider_x509_cert_url"],
            "client_x509_cert_url": st.secrets["gcp_service_account"]["client_x509_cert_url"]
        }
    return(credentials)

def save_storage():
    with st.sidebar:
        with st.spinner('Storing data...'):
            print('Start storage...')

            data_storage = {}
            gs_storage = []

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

            if option == 'Roomnights & Revenue':
                iDataSt.append('Rev')
            elif option == 'Roomnights & ADR':
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
            data_storage['iStepper'] = iSkipper1
            data_storage['iDataSt'] = iDataSt
            data_storage['iLoc'] = iLoc

            json_string = json.dumps(data_storage)

            credentials = run_credentials()

            gc = gspread.service_account_from_dict(credentials)

            sh = gc.open(st.secrets["private_gsheets_url"])

            for key, values in data_storage.items():
                gs_storage.append(key)
                if(isinstance(values, list)):
                    for value in values:
                        gs_storage.append(value)
                else:
                    gs_storage.append(value)
            
            a=len(sh.sheet1.row_values(1)) + 65
            letter = chr(a)
            
            start_letter = letter
            end_letter = letter
            start_row = 2
            end_row = start_row + len(gs_storage) - 1
            range1 = "%s%d:%s%d" % (start_letter, start_row, end_letter, end_row)
            
            cell_list = sh.sheet1.range(range1)

            sh.sheet1.update_cells(cell_list)

            result_str = ''.join(random.choice(string.ascii_letters) for i in range(8))

            print(f'{result_str}')

            z = 0
            for x in range(len(gs_storage)):
                cell_list[z].value = gs_storage[x]
                z = z + 1
            sh.sheet1.update_cells(cell_list)

            a = a - 64
            sh.sheet1.update_cell(1, a, f'{result_str}')

            st.write("""
                # Here you can check your input for future use.
                Save the generated password for later purposes.
                """)

            st.json(json_string, expanded=False)

        st.write(f'''
            #### Here is your password: 
            {result_str}
            ##### Save it well.
            ''')
        st.success('Input saved!')

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
        iSort_t = []
        iDataSt = []
        iDataRn = []
        iDataRnT = []
        iDataRv = []
        iDataRvT = []
        json_storage = []
        row_n_storage = []
        Skipper_s = []
        Skipper_s1 = []
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

            iDataSt = iData_choice

            row_n_storage.append(row_n)
            json_storage.append(storage)

            for x in Skipper: Skipper_s.append(x)
            Skipper_s = [round(float(i)) for i in Skipper_s]
            
            for x in Skipper1: Skipper_s1.append(x)
            Skipper_s1 = [round(float(i)) for i in Skipper_s1]

        elif stro == 'Yes': 

            for x in range(len(values_list)):
                if iSegments_l[0] < x < iTerm_l[0]:
                    iSegments.append(values_list[x])

                elif iTerm_l[0] < x < iSort_l[0]:
                    iTerm.append(values_list[x])

                elif iSort_l[0] < x < iSkipper_l[0]:
                    iSort_t.append(values_list[x])

                    iSort.append(values_list[x])
                elif iSkipper_l[0] < x < iStepper_l[0]:
                    Skipper_s.append(values_list[x])

                elif iStepper_l[0] < x < iDataSt_l[0]:
                    Skipper_s1.append(values_list[x])

                elif iDataSt_l[0] < x < iLoc_l[0]:
                    iDataSt = (values_list[x])

                elif iLoc_l[0] < x:
                    json_storage.append(values_list[x])
                    iLoc_l.append(values_list[x])

            row_n_storage.append(json_storage[1])

            Skipper_s = [round(float(i)) for i in Skipper_s]
            Skipper_s1 = [round(float(i)) for i in Skipper_s1]

            for x in cols: iMonths.append(x)
            
            for x in iMonths:
                str2Match = x
                strOptions = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Sept","Oct","Nov","Dec"]

                highest = process.extractOne(str2Match,strOptions)
                highest = highest
                iDays.append(highest)

        print(f"iSegments: {iSegments}")
        print(f"sheets (iMonths): {iMonths}")
        print(f"terminology (iTerm): {iTerm}")
        print(f"months (iDays): {iDays}")
        print(f"sorting (iSort): {iSort}")

        if len(iTerm) == 0:
            st.error('Err1: No terminology filled in. Press "enter" after typing the terminology!', icon='❌')
            return
        
        wb = openpyxl.load_workbook(uploaded_file_CLIENT, data_only=True)

        try:    
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
                                        pass
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
                            ##### Err2: ERROR for: {iTerm[0]}. In total {len(iSegments)} segmets entered. But {rn_c} segments were measured in the month / sheet: {iMonths[x]}.
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
                                        pass
                                        
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
                            ##### Err3: ERROR for: {iTerm[1]}. In total {len(iSegments)} segmets entered. But {rv_c} segments were measured in the month / sheet: {iMonths[x]}.
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
                                        pass
                                        
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
                                        pass

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
                            ##### ERROR for: {iTerm[1]}. In total {len(iSegments)} segments entered. But {rv_c} segments were measured in the month / sheet: {iMonths[x]}.
                            See below an overview of the segments and their range that were succeeded:
                        """, 
                        icon="❌")
                        st.json(cRngv, expanded=True)
                        return
           
            a = 0
            for x in range(len(iMonths)):
                
                for i in range(len(iSegments)):

                    strFind = iSegments[i]

                    for y in range(len(iSort)):

                        strStored = iSort[y]

                        print(f"L: {(iSegments[i])} | R: {(iSort[y])}")

                        if strFind == strStored:
                            
                            if i == y:
                                #print(f"-- Level & Name match: {strFind} = {i}")
                                pass
                            else:
                                #print(f"- Name match: {strFind} : {i} - {y}")
                                #print(f"- reordering...")
                                
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
                    for z in iSort_t: iSort.append(z)

            d =[]

            for x in range(len(iDataRnT)):
                d.append(iDataRnT[x])
                d.append(iDataRvT[x])

            df2 = pd.DataFrame(data=d)

            df2 = df2.T
            
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.title='sheet0'

            y = 1
            for x in range(len(all_columns)):
                sheet.cell(row=1, column=y).value=all_columns[x]
                y = y + 1

            x = 2 # COLUMN
            y = 2 # ROW
            t = 2 # ROW
            s = 1 # SEGMENTS

            for i in range(len(iDataRnT)):

                for z in range(len(iDataRnT[i])):

                    sheet.cell(row=y, column=x).value=iDataRnT[i][z]
                    if iDataSt == "ADR":
                        sheet.cell(row=y, column=x+1).value=(iDataRnT[i][z] * iDataRvT[i][z])
                    else:
                        sheet.cell(row=y, column=x+1).value=iDataRvT[i][z]
                    y = y + 1

                if s == len(iSegments):
                    x = 2
                    s = 1
                    if i <= len(iSegments):
                        t = 2 + (len(iDataRnT[i]))
                    else:
                        t = t + (len(iDataRnT[i]))
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
        
        except Exception as e:

            traceback.print_exc()

            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]

            credentials = run_credentials()

            gc = gspread.service_account_from_dict(credentials)

            sh = gc.open(st.secrets["private_gsheets_url_log"])
            a=len(sh.sheet1.col_values(1))

            sh.sheet1.update_cell(a + 1, 1, f'Err4: {exc_type}; {exc_obj}; ({str(e)}), line: {exc_tb.tb_lineno}, in {fname} {datetime.utcnow()}')
            st.error(f'Err4: {exc_type}; {exc_obj}; ({str(e)}), line: {exc_tb.tb_lineno}, in {fname} | UTC: {datetime.utcnow()}')
            return

    st.success('Process ran!')

    with open("NameFile.xlsx", "rb") as file:
        st.download_button(
            label="click me to download excel",
            data=file,
            file_name=f'{uploaded_file_JUYO.name}',
            mime="application/octet-stream"
            )

# Header of the page
with st.container():
    l_column, m_column, r_column = st.columns([3,5,1])
    with l_column:
        st.write("")
    with m_column:
        st.write(
            f"""
        # ♾️Forecast / budget converter
        ######  The process of converting the forecast file to the right format.
        """
        )
    with r_column:
        st.write("")
        imagejuyo = Image.open('images/JUYO3413_Logo_Gris1.png')
        st.image(imagejuyo)

with streamlit_analytics.track():        
    # Here will start the step-by-step process for data input.
    with st.container():

        st.write("---")
        disabled = 1

        left_column, right_column = st.columns(2)

        with left_column:

            st.header("Forecast File client")
            uploaded_file_CLIENT = st.file_uploader("Upload client file", type=".xlsx")

            if uploaded_file_CLIENT:
                
                if uploaded_file_CLIENT.name[-5:] == '.xlsb':
                    print('xslb')
                    st.error('.XLSB file format (Excel Binary File Format) is not supported by python or its liberies. Please convert the file to an .XLSX file format.')

                st.markdown("### Select wanted sheets in chronological order for conversion.")

                tabs = pd.ExcelFile(uploaded_file_CLIENT).sheet_names

                cols = st.multiselect('Select sheets in the **chronological** order of the months:', tabs)
                st.write('press enter after each sheet.')
                st.write('You selected:', len(cols)) 

        with right_column:
            st.header("Format file JUYO")
            uploaded_file_JUYO = st.file_uploader("Upload JUYO file", type=".xlsx")

            if uploaded_file_JUYO:
                
                df1 = pd.read_excel(uploaded_file_JUYO)

                all_columns = df1.columns

                shape1 = df1.shape

                df1 = df1[[k for i, k in enumerate(df1.columns, 0) if i % 2 != 0]]

                shape = df1.shape

                st.write(shape[1], ' segments detected in Juyo file')
            
        if uploaded_file_JUYO:

            st.write("---")

            # //TODO add explanation for why password and what it will return.
            st.write("## Enter password to retreive previous data stored:")
            st.write("When the password is entered, it will retreive the data stored, so that you won't have to enter all the input")
            stro = st.radio(
                label="-",
                options=("Yes", "No"),
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
                    
                    expander =  st.expander('See explantion of why the segments needs to be in chronological order segments :')
                    expander.write('''
                        ### Important.
                        Because everyone's forecast/ budget file looks different, it is difficult to determine exactly where the segments are.
                        
                        The reason why you need to put the segments in **chronological order** here is, because the script will look only at the index of the postition
                        of where the segments are entered here and not the name of the segment..

                        Soon you will have to map the segments in the **correct order**, but not here, because the script needs to know where the segments are in your forecast/ budget file.

                        !! If you do **NOT** enter your segments in chronological order (left to right, or top to bottom) here, the data will **not** be correct in the end.
                    ''')

                    if len(keywords) == shape[1] - 1:
                        st.warning('Are you 1 segment short and it is at the last place of the segments? Read this 👇', icon="⚠️")
                        st.write("""
                            If you are 1 segment short, and that segment happens to be last, there is a possibility.
                            When the missing segment is at the last, this process can recognize it and leave it empty, automatically becoming zero in JUYO. Click the checkbox if so.
                            """)
                        e_segments = st.checkbox('Extra (empty!) segment on last place?')

                    st.write(len(keywords), ' segments of ', shape[1], ' entered.')

                    if len(keywords) == shape[1] or e_segments:

                        with st.container():

                            st.write("---")

                            left_column, right_column = st.columns(2)

                            with left_column:
                                # https://github.com/ohtaman/streamlit-sortables
                                
                                st.markdown("### Segments in correct the order.")
                                sorted_items1 = sort_items(df1.columns.to_list(), direction='vertical')

                            with right_column:
                                st.markdown("### Map the segments so they exact match on the left!")
                                sorted_items2 = sort_items(keywords, direction='vertical')

                            st.write('## Select starting year of first sheet.')
                            year = st.select_slider(
                                label='## Select starting year of first sheet.',
                                options=range(datetime.today().year - 2, datetime.today().year + 3),value=datetime.today().year)

                            st.write(year)

                            st.write('---')

                            terminology = []

                            st.write("### Is the data stored as *roomnights and revenue* or *roomnights and ADR*")
                            option = st.radio(
                                    label="## .",
                                    options=('Roomnights & Revenue', 'Roomnights & ADR'))

                            if option == 'Roomnights & Revenue':
                                term = "revenue"
                                iData_choice = "Rev"
                            else: 
                                term = 'ADR'
                                iData_choice = "ADR"
                                print(iData_choice)
                            
                            st.write(f"""
                                # Enter the terminology used in Excel file for room nights and {term}:
                                For example; roonnights = Rms, Rn, etc. Revenue = Rev, Rvu, etc. (press ENTER when filled in!)
                                """)

                            terminologyR = st_tags(
                                label='Enter the terminology of **room nights:**',
                                text='Press enter to add more',
                                suggestions=['rn', 'RN', 'Rn', 
                                            'Rev', 'REV', 'rev', 
                                            'ADR', 'Adr', 'adr',
                                            ],
                                maxtags = 1,
                                key='t1')

                            terminologyR1 = st_tags(
                                label=f'Enter the terminology of **{term}:**',
                                text='Press enter to add more',
                                suggestions=['rn', 'RN', 'Rn', 
                                            'Rev', 'REV', 'rev', 
                                            'ADR', 'Adr', 'adr',
                                            ],
                                maxtags = 1,
                                key='t2')

                            terminology = terminologyR + terminologyR1

                            storage = st.radio(
                                label=f'Are the headers of the terminology room nights and {term} stored in a row or in a column?',
                                options=('Rows', 'Columns'))

                            with st.expander('Click for more explantion'):
                                st.write(f'Here you can see with what me mean with row or column. E.g.:')
                                image = Image.open('images/voorbeeld_excel.png')
                                st.image(image)
                                st.write('''In this picture you can see that the terminology of room nights and revenue is stored in column 'a'.
                                        So than you will choose *'Columns'* and enter 'a' in the inputbox. If the terminology is stored in a row, then you
                                        will choose 'Rows' and press the row wanted, e.g.; '6'.
                                ''')
                        
                            if storage == 'Rows':
                                row_n = st.text_input(f"in which row can the roomnigts and {term} be found?")
                                if row_n:
                                    st.write(f'The terminology can ben found in row {row_n}')
                            
                            elif storage == 'Columns':
                                row_n = st.text_input(f"In which column can the roomnigts and {term} be found? (use lowercase letters)")
                                row_n = ord(row_n) - 96
                                if row_n:
                                    st.write(f'The terminology can be found in column {row_n}')
                            
                            st.write('### Skip terminology on certain places?')
                            skip_term = st.checkbox('Want to skip terminology on certain places?')

                            with st.expander('Click for more explantion'):
                                st.write(f'You have just indicated the terminology (roomnights and {term}) is in {storage} {row_n}. It may happen that on that same {storage} the terminology (roomnights and {term}) of totals are included, but should not be included. E.g.:')
                                image = Image.open('images/voorbeeld_excel.png')
                                st.image(image)
                                st.write('''Here you can see that in column A, there are minor 2 segments. But the terminology Rms and REV are used 3 times.
                                        So if you want to skip the terminology, you have to click the check box and indicate on which places you want to skip the terminology.''')
                                        
                                st.write('''If you want to skip an item, you must write down the index of that item in the correct order that it needs to be skipped. 
                                        This is because the order in which you skip items is important. For example, if you have a list of segments that you need to go through and you want to skip the second item, 
                                        you would need to write down the index "2" to indicate that this is the item you want to skip. This ensures that the system knows which item to skip and in what order to skip it. 
                                        Additionally, it's important to write down the index in the correct order because the system may not be able to understand or process your request if the index is out of order. 
                                        So if you want to skip an item, be sure to write down the correct index in the correct order.
                                ''')

                                st.write('If you already seen the error box, it is good to know that you need to add +1 to the index, as the list is starting from zero (0)')

                            if skip_term:

                                l_c, r_c = st.columns(2) 

                                with l_c:
                                    Skipper = st_tags(
                                        label=f'### Skip a terminology in order ({terminologyR})',
                                        text='Press enter to add more',
                                        suggestions=['1', '2', '3', 
                                                    '4', '5', '6'],
                                        maxtags = 5,
                                        key='3')

                                with r_c:
                                    Skipper1 = st_tags(
                                        label=f'### Skip a terminology in order ({terminologyR1})',
                                        text='Press enter to add more',
                                        suggestions=['1', '2', '3', 
                                                    '4', '5', '6'],
                                        maxtags = 5,
                                        key='4')

                            else:
                                Skipper = []
                                Skipper1 = []
                            
                            st.write('---')
                            st.write('If this is your first time, it is better to first start the converting process instead of storing the data.')
                            st.write('You will then have a check that the output data is correct')

                            if st.button('store data', key="store"): 
                                save_storage()

                            if st.button("Start converting process.", key="run4"):                        
                                run_process()


                except Exception:
                    traceback.print_exc()

            elif stro == 'Yes':

                key_s = st.text_input("Enter key")

                try: 
                    credentials = run_credentials()

                    gc = gspread.service_account_from_dict(credentials)

                    sh = gc.open(st.secrets["private_gsheets_url"])
                                        
                except Exception:
                    traceback.print_exc()

                try:
                    if key_s == "":
                        pass
                    else:
                        cell = sh.sheet1.findall(key_s)
                        
                        loc = str(cell[0])
                        
                        values_list = sh.sheet1.col_values(loc[9:10])
                        
                        iSegments_l = [i for i, s in enumerate(values_list) if 'iSegments' in s]
                        iTerm_l = [i for i, s in enumerate(values_list) if 'iTerm' in s]
                        iSort_l = [i for i, s in enumerate(values_list) if 'iSort' in s]
                        iSkipper_l = [i for i, s in enumerate(values_list) if 'iSkipper' in s]
                        iStepper_l = [i for i, s in enumerate(values_list) if 'iStepper' in s]
                        iDataSt_l = [i for i, s in enumerate(values_list) if 'iDataSt' in s]
                        iLoc_l = [i for i, s in enumerate(values_list) if 'iLoc' in s]

                        sh_log = gc.open(st.secrets["private_gsheets_url_log"])
                        a=len(sh_log.sheet1.col_values(1))
                        sh_log.sheet1.update_cell(a + 1, 1, f'Key={key_s} used at: {datetime.utcnow()}')

                        st.success(f'password: {values_list[0]} succesfull', icon='✅')

                except:
                    st.write('❌ No match found!')

                st.write('## Select starting year of first sheet.')

                year = st.select_slider(
                    label="# .",
                    options=range(datetime.today().year - 2, datetime.today().year + 3),value=datetime.today().year,
                    label_visibility='collapsed'
                    )

                if st.button("Start converting process.", key="run1"): run_process()

            else:
                st.write('please select a option.')