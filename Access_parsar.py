##############################################################################################################################
# File name : Access_parsar.py                                                                                               #
# Created on : 21-08-2018                                                                                                    #
# Author : Deepak Chandella                                                                                                  #
# Usage : This file is used to parse access data and get below fields                                                        #
#         <Date> <Badge_ID> <Intime> <Outtime> <Totaltime> <Effectivetime>                                                   #
#         <Breaks> <Breaktime> <tail gatings>                                                                                #
#                                                                                                                            #
# Run : python Access_parsar.py -filename <excel_file_to_read> -read_sheet <sheet_to_read> -write_sheet <sheet_to_write>     #
#                                                                                                                            #
##############################################################################################################################


import pandas as pd
import numpy as np
import openpyxl
import os,argparse
from collections import OrderedDict


def arguments():
    parser = argparse.ArgumentParser(description="Command line arguments to process data")
    parser.add_argument('-filename', required=True, help='filename to read data')
    parser.add_argument('-read_sheet', required=True, help='sheet name to read data')
    parser.add_argument('-write_sheet', required=True, help='sheet name to write data')
    args = parser.parse_args()
    return args


def calender_date_df():
    Column_names = ["Date","Badge","Intime","Outtime","Totaltime","Effectivetime","Breaks","Breaktime","Tail"]
    my_df = pd.DataFrame(columns=Column_names)
    return my_df


def load_write_file(filepath,filename):
    book = openpyxl.load_workbook(filepath)
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    return writer,book

def write_file(val,my_df,sheet):
    headers = ["Date","Badge","Intime","Outtime","Totaltime","Effectivetime","Breaks","Breaktime","Tail"]
    newpd.to_excel(writer, sheet, startrow=0, startcol=0, index=False, header=headers)
    writer.save()


def get_effective_hours(data):
    effective = 0
    for k,v in data.items():
        diff = v - k
        effective += diff
    return effective

def get_breaks_time(data):
    btime = 0
    key = data.keys()[1:]
    val = data.values()[:len(data.values())-1]
    for i in range(len(key)):
        diff = key[i] - val[i]
        btime += diff
    return btime
    

def parse_location(arr,gate):

    data = OrderedDict()
    in_flag = False
    out_flag = False
    tail_count = 0
    while len(gate) != 0:

        if ("entry" in gate[len(gate)-1].lower()) and (in_flag == True):
            tail_count+=1
            intime = arr[len(gate)-1]
            in_flag = True
            gate = np.delete(gate, len(gate)-1)
        elif ("entry" in gate[len(gate)-1].lower()) and ((out_flag == False) and (in_flag == False)):
            intime = arr[len(gate)-1]
            in_flag = True
            gate = np.delete(gate, len(gate)-1)
        elif ("exit" in gate[len(gate)-1].lower()) and (in_flag == False):
            tail_count+=1
            gate = np.delete(gate, len(gate)-1)
        elif ("exit" in gate[len(gate)-1].lower()) and ((in_flag == True) and (out_flag == False)):
            outtime = arr[len(gate)-1]
            out_flag = True
            gate = np.delete(gate, len(gate)-1)
            
        if in_flag and out_flag:
            data.update({intime:outtime})
            in_flag = False
            out_flag = False
            
    return data,tail_count

def parse_data(emp_data,id1,my_df):
    if not pd.isnull(id1):
        print "Processing data for user - {}".format(id1)
    for item in set(emp_data["Date/Time"].dt.date):
        arr = np.array(emp_data["Date/Time"][emp_data["Date/Time"].dt.date == item])
        gate = np.array(emp_data["Location"][emp_data["Date/Time"].dt.date == item])
        
        data,tail = parse_location(arr,gate)

        if len(data) == 0:
            my_df = my_df.append({'Date': item, 'Badge': id1, 'Tail': tail}, ignore_index=True)
        else:
            intime = data.keys()[0]
            outtime = data.values()[len(data)-1]
            total_time = outtime - intime
            effective = get_effective_hours(data)
            breaktime = get_breaks_time(data)
            breaks = len(data)-1

            my_df = my_df.append({'Date': item, 'Badge': id1, 'Intime': intime, 'Outtime': outtime, 'Totaltime': total_time, 'Effectivetime': effective, 'Breaks': breaks, 'Breaktime': breaktime, 'Tail': tail}, ignore_index=True)

    return my_df



if __name__ == "__main__":
    arg = arguments()
    filepath = os.path.join(os.path.dirname(os.path.abspath(__file__)),arg.filename).replace("\\","\\")

    writer,book = load_write_file(filepath,arg.filename)
    d = pd.read_excel(filepath, arg.read_sheet)
    df = pd.DataFrame(d)
    
    my_df = calender_date_df()
    li = []
    for id1 in set(df["Badge ID"]):
        emp_data =  df[df["Badge ID"] == id1]
        new_df = parse_data(emp_data,id1,my_df)
        li.append(new_df)
        
    newpd = pd.concat(li)
    sort_val = newpd.sort_values(by="Date", ascending=True)

    ###writing df to file    
    write_file(sort_val,my_df,arg.write_sheet)
    
    








