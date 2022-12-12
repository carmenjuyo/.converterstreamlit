import streamlit as st
import gspread

import random
import string

# With combination of lower and upper case
result_str = ''.join(random.choice(string.ascii_letters) for i in range(8))

# print random string
print(result_str)

for x, y in range(10):
    print(x, y)