import streamlit as st
import gspread

import random
import string

# With combination of lower and upper case
result_str = ''.join(random.choice(string.ascii_letters) for i in range(8))
# print random string
print(result_str)

# credentials = {
#   "type": st.secrets["gcp_service_account"]["type"],
#   "project_id": st.secrets["gcp_service_account"]["project_id"],
#   "private_key_id": st.secrets["gcp_service_account"]["private_key_id"],
#   "private_key": st.secrets["gcp_service_account"]["private_key"],
#   "client_email": st.secrets["gcp_service_account"]["client_email"],
#   "client_id": st.secrets["gcp_service_account"]["client_id"],
#   "auth_uri": st.secrets["gcp_service_account"]["auth_uri"],
#   "token_uri": st.secrets["gcp_service_account"]["token_uri"],
#   "auth_provider_x509_cert_url": st.secrets["gcp_service_account"]["auth_provider_x509_cert_url"],
#   "client_x509_cert_url": st.secrets["gcp_service_account"]["client_x509_cert_url"]
# }

# gc = gspread.service_account_from_dict(credentials)

# sh = gc.open("juyo_Db_Private")

# print(sh.sheet1.get('A1'))

# val = sh.sheet1.cell(1, 2).value

# print(val)

# sh.sheet1.update_cell(1, 2, 'Bingo?')