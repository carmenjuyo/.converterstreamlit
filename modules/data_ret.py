
from datetime import datetime
import streamlit as st
import gspread

from modules.data_process import Gscon

class Gsret:

    def retreive_data(key_s):

        iSegments = []
        iTerm = []
        iSort = []
        iSort_t = []
        iDataSt = []
        storage = []
        iLoc = []
        Skipper_s = []
        Skipper_s1 = []
        
        with st.spinner('Retreiving data...'):
            try:
                credentials = Gscon.run_credentials()

                gc = gspread.service_account_from_dict(credentials)

                sh = gc.open(st.secrets["private_gsheets_url"])

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

                #st.success(f'password: {values_list[0]} succesfull', icon='✅')

            except:
                st.write('❌ No match found!, please try again.')
                return
            
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
                    storage.append(values_list[x])
                    iLoc.append(values_list[x])
                
            
            result_list = {
                'iSegments:': iSegments,
                'iTerm': iTerm,
                'iSort': iSort,
                'iSkipper': Skipper_s,
                'iStepper': Skipper_s1,
                'iDataSt': iDataSt,
                'iLoc': [storage[0], iLoc[1]]
            }

            #st.json(result_list, expanded=False)

            return(result_list)