#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import os
from expenses_functions import *


def main():
    # Let's start!
    year = choosing_interface("year")
    month = choosing_interface("month",year)

    main_df = pd.read_excel(EXPENSES_PATH, sheet_name="Data")
    file_path = os.path.join(f"Years\{year}", month)
    full_new_df = edit_month_file(file_path)

    if ask_to_add(main_df, full_new_df):  # If the user decided to add the expenses.
        add_file(EXPENSES_PATH, full_new_df)
        print("Done")
        

if __name__ == "__main__":
    main()

