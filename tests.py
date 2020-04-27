"""
Program test script to be used for updates and bug fixes.
"""

import os
import re
import glob
import zipfile
import fnmatch
import datetime
from time import sleep
from zipfile import ZipFile
import pandas as pd
from pandas import DataFrame

def create_timestamp():
    """
    Creates a timestamp in DB format.
    """
    today = datetime.date.today()
    year = today.year
    month = today.month
    day = today.day
    
    timestamp = f"{str(year)}-{str(month)}-{str(day)}"
    
    return timestamp

def main():
    """
    Main guts of the script.
    """
    try:
        if not os.path.exists(archive_dir):
            os.makedirs(archive_dir)
            if not os.path.exists(data_dir):
                os.makedirs(data_dir)

        if fnmatch.fnmatch(data_file, "*.xlsx") == True:
            os.remove(glob.glob(data_dir + "*.xlsx"))

    finally:
        for data_file in os.listdir(data_dir):
            if fnmatch.fnmatch(data_file, "*.zip") == True:
                with ZipFile(data_dir + data_file, "r") as zip_obj:
                    zip_obj.extractall(data_dir)
                os.remove(data_dir + data_file)

        xl_list = glob.glob(data_dir + "*.xlsx")
        cs_list = []

        for xl_file in xl_list:
            workbook = pd.ExcelFile(xl_file)
            sheet_list = workbook.sheet_names
        
            for (index, sheet) in enumerate(sheet_list):
                cs_list.append(workbook.parse(index))    

        df_cs_data = pd.concat(cs_list)            
        df_cs_data.to_excel(f"{archive_dir}CS-Sales-Change-{create_timestamp()}.xlsx")

        exit()

if __name__ == "__main__":

    data_dir = ".\\excel\\data\\"
    archive_dir = ".\\excel\\archive\\"

    main()
