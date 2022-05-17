import encodings
import os
import sys
import openpyxl
import argparse

from pandas import DataFrame
from dbfread import DBF
from itertools import islice

def vprint(msg):
    """prints only when verbosity is switched on

    Args:
        msg (all): message to print
    """
    if args.verbose:
        print(msg)

def load_data(filename):
    """loads the data from either an excel file or a .dbf file and puts them into a
    pandas DataFrame

    Args:
        filename (str): absolute path to the file, given via args

    Returns:
        DataFrame: pandas DataFrame containing all the data
    """
    if filename.endswith(('.xlsx','.xlsm','.xltx','.xltm')):
        try:
            wb = openpyxl.load_workbook(filename=filename)
            ws = wb[wb.sheetnames[0]]
            data = ws.values
            cols = next(data)[1:]
            data = list(data)
            data = (islice(r, 1, None) for r in data)
            df = DataFrame(data, columns=cols)
            
            return df
        except openpyxl.utils.exceptions.InvalidFileException as e:
            print(e)
            sys.exit(1)
    elif filename.endswith(('.dbf')):
        try:
            data = DBF(filename, encoding='utf-8')
            df = DataFrame(iter(data))
            
            return df
        except dbfread.exceptions.DBFNotFound as e:
            print(e)
            sys.exit(1)
    else:
        print("Your supplied file is not an excel (.xlsx, .xlsm, .xltx, .xltm) or a .DBF file! We are currently not able to parse any other files")
        sys.exit(1)

formatter = lambda prog: argparse.ArgumentDefaultsHelpFormatter(prog,max_help_position=50, width=120)
my_parser = argparse.ArgumentParser(
    add_help=False,
    prog='python3 main.py',
    formatter_class=formatter)

my_parser.add_argument(
    '-h',
    '--help',
    action='store_true',
    help='displaying this help message',
    default=False)

my_parser.add_argument(
    '-v',
    '--verbose',
    action='store_true',
    help='verbose output',
    default=False)

input_group = my_parser.add_argument_group("### INPUTS ###")

input_group.add_argument(
    '-f',
    '--file',
    action='store',
    type=str,
    metavar="$fn",
    help='absolute path to the file',
    default=None)

try:
    args, unknown = my_parser.parse_known_args()
except Exception:
    print("Something went wrong with the parser. Please consult an exorcist.")
    exit()

if len(unknown) != 0:
    for ukarg in unknown:
        print("%s is not a known argument!" % ukarg)
    sys.exit(1)

if args.help:
    ascii = """
       _           _                                         
      (_)         | |                                        
 _ __  ___   _____| |    _ __  _ __ ___ _ __  _ __ ___   ___ 
| '_ \| \ \ / / _ \ |   | '_ \| '__/ _ \ '_ \| '__/ _ \ / __|
| | | | |\ V /  __/ |   | |_) | | |  __/ |_) | | | (_) | (__ 
|_| |_|_| \_/ \___|_|   | .__/|_|  \___| .__/|_|  \___/ \___|
                  ______| |            | |                   
                 |______|_|            |_|                   
"""
    help_text = my_parser.format_help()
    print(ascii)
    print(help_text)
    sys.exit(0)

if args.file == None:
    print("please provide a file! exiting...")
    sys.exit(1)

#!# GETTING VARIABLES FROM ARGPARSE #!#
fn = args.file

vprint("loading file...")

data = load_data(filename=fn)