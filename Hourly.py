##############################################################################################################################
# File name : hourly.py                                                                                                      #
# Created on : 02-11-2018                                                                                                    #
# Author : Deepak Chandella  <deepak.chandella@orange.com>                                                                   #
# Usage : This file is used to parse access data and                                                                         #
#          get the hourly presence of employee                                                                               #
#                                                                                                                            #
# Run : python Hourly.py -filename <excel_file_to_read> -read_sheet <sheets_to_read> -write_sheet <sheet_to_write>           #
#                       -start_date <2018-08-01> -end_date <2018-08-31>                                                      #
#                                                                                                                            #
#                                                                                                                            #
##############################################################################################################################


import pandas as pd
import os,argparse
import openpyxl


def arguments():
    parser = argparse.ArgumentParser(description="Command line arguments to process data hourly")
    parser.add_argument('-filename', required=True, help='filename to read data')
    parser.add_argument('-read_sheet', required=True, help='sheet names to read data')
    parser.add_argument('-write_sheet', required=True, help='sheet name to write data')
    parser.add_argument('-start_date', required=True, help='Start date for hourly calucation')
    parser.add_argument('-end_date', required=True, help='End date for hourly calucation')
    args = parser.parse_args()
    return args

def calender_date_df(start_date, end_date):
    df = pd.DataFrame(
        {'Hours': pd.date_range(start_date, end_date, freq='1H', closed='left')}
     )
    return df


def load_write_file(filepath,filename):
    book = openpyxl.load_workbook(filepath)
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    return writer,book

def write_file(my_df,sheet):
    my_df.to_excel(writer, sheet, index=False)
    writer.save()


if __name__ == "__main__":
    arg = arguments()
    filepath = os.path.join(os.path.dirname(os.path.abspath(__file__)),arg.filename).replace("\\","\\")
    writer,book = load_write_file(filepath, arg.filename)
    
    d = pd.read_excel(filepath, arg.read_sheet)
    df = pd.DataFrame(d)
    
    my_df = calender_date_df(arg.start_date, arg.end_date)
    
    df['Intime'] = df['Intime'].apply(lambda dt: dt.replace(minute=0, second=0))
    df['Outtime'] = df['Outtime'].apply(lambda dt: dt.replace(minute=0, second=0))

    df = df[df['Intime'].notnull()]

    for idx in set(df["Badge"]):
        print "Processing User : {}".format(idx)
        emp_data =  df[df["Badge"] == idx]
        for item in set(emp_data["Date"].dt.date):
            in_time = emp_data["Intime"][emp_data["Date"] == item]
            out_time = emp_data["Outtime"][emp_data["Date"] == item]

            in_time = in_time.values[:1]
            out_time = out_time.values[:1]
            
            my_df.loc[(my_df["Hours"].values >= in_time) & (my_df["Hours"].values <= out_time), idx] = 1

    #write to file            
    write_file(my_df,arg.write_sheet)
