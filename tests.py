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
                cs_list.append(workbook.parse(index, skiprows=1, header=None))

        txt_list = glob.glob(data_dir + "*.txt")

        for txt_file in txt_list:
            if fnmatch.fnmatch(txt_file.lower(), "*mic*.txt") == True:
                df_mic = pd.read_csv(txt_file, sep="|", skiprows=1, header=None)
                df_mic.columns = [3, 24, 1, 25]
                df_mic.drop(columns=[1], inplace=True)

        df_cs_data = pd.concat(cs_list)
        df_cs_data = df_cs_data.merge(df_mic, how="left", on=3)
        df_cs_data.drop_duplicates(inplace=True)

        df_cs_data[5] = ["Rusty Ames" if x == "Rusty Amees" else x for x in df_cs_data[5]]
        df_cs_data[3] = df_cs_data[3].map('{:0>6}'.format)
        df_cs_data[6] = df_cs_data[6].map('{:0>6}'.format)
        df_cs_data[31] = df_cs_data[13] / df_cs_data[12]
        df_cs_data[31] = df_cs_data[31].apply(lambda x: round(x, 1))

        df_cs_data = df_cs_data.reindex(df_cs_data.columns.tolist() + [26, 27, 28, 29, 30], axis=1)
        df_cs_data = df_cs_data[[0, 1, 2, 3, 4, 5, 24, 22, 23, 6, 7, 8, 25, 9, 10, 11, 26, 27, 12, 13, 31, 14, 28, 15, 16, 17, 18, 19, 20, 21, 29, 30]]
        df_cs_data.columns = ['GL', 
                            'Location Name', 
                            'Customer Code', 
                            'Tops Code', 
                            'Buyer Code', 
                            'Category Business Manager', 
                            'Category', 
                            'Private Label Flag', 
                            'Vendor Name', 
                            'C&S Code', 
                            'Item Description', 
                            'Size', 
                            'Brand', 
                            'UPC - Vendor', 
                            'UPC - Case', 
                            'UPC - Item', 
                            'WTD Category Unit Lift %', 
                            'WTD Item Unit Lift %', 
                            'Weekly Turn (Forecast)', 
                            'BOH',
                            'Inventory Weeks on Hand',
                            'Total On Order', 
                            'OOS Yesterday', 
                            'Next PO Due Date', 
                            'Next PO Appt Date', 
                            'Next PO Qty', 
                            'Next Biceps PO#', 
                            'Lead Time', 
                            'Current Week Bookings', 
                            'Future Bookings', 
                            'Item Key', 
                            'Manufacturer Status']

        df_cs_data.to_excel(f"{archive_dir}CS-Sales-Change-{create_timestamp()}.xlsx", index=False)

        exit()

if __name__ == "__main__":
    data_dir = ".\\excel\\data\\"
    archive_dir = ".\\excel\\archive\\"

    main()
