##################################################################################################
# File name : converter.py                                                                       #
# Created on : 02-11-2018                                                                        #
# Author : Deepak Chandella                                                                      #
# Usage : This file is used to convert the xls file to xlsx.                                     #
#         This file also perform the filtering of data. for example.                             #
#         - removing empty lines                                                                 #
#         - concatinate the date with time                                                       #
# Requirements : pip install pyexcel-cli pyexcel-xls pyexcel-xlsx                                # 
#                                                                                                #
# Run : python converter.py -input_file <xls_excel_file> -output_file <xlsx_excel_file>          #
#                                                                                                #
##################################################################################################

import os,argparse,sys
from subprocess import call


def arguments():
    parser = argparse.ArgumentParser(description="Command line arguments to convert xls to xlsx")
    parser.add_argument('-input_file', required=True, help='input xls file name to convert')
    parser.add_argument('-output_file', required=True, help='output xlsx file name')
    args = parser.parse_args()
    return args

def create_cmd(arg):
    cmd = "pyexcel transcode " + arg.input_file + " " + arg.output_file
    return cmd

def execute_cmd(cmd):
    try:
        code = call(cmd)
    except (OSError,CalledProcessError) as e:
        print "ERROR : {}".format(e)
        sys.exit()
    return code

if __name__ == "__main__":
    arg = arguments()
    cmd = create_cmd(arg)
    status = execute_cmd(cmd)
    if status == 0:
        print "xls file is successfully converted to xlsx file"
    
    
