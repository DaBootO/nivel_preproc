import sys
import openpyxl
import argparse
import datetime
import numpy as np

from pandas import DataFrame
from dbfread import DBF
from itertools import islice
from tqdm import tqdm
from yaspin import yaspin

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

def convert_date(date):
    try:
        return date.to_pydatetime()
    except AttributeError:
        if isinstance(date, datetime.date):
            return date
        else:
            print("date: %s could not be converted to a datetime.date object" % date)
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
    metavar="$ifn",
    help='absolute path to the input file',
    default=None)

output_group = my_parser.add_argument_group("### OUTPUTS ###")

output_group.add_argument(
    '-o',
    '--output',
    action='store',
    type=str,
    metavar="$ofn",
    help='absolute path to the output file',
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
ofn = args.output
#!# GETTING VARIABLES FROM ARGPARSE #!#

# sorry if all this spinning bullshit is too much, never used it and saw an opportunity here
with yaspin(text="loading file...") as sp:
    data = load_data(filename=fn)
    sp.write("> finished loading file")

# convert all dates to the needed format -> D_YYYYmmdd
dates_with_repetitions = [convert_date(d).strftime('D_%Y%m%d') for d in data['Datum']]

# switch out the Datum column with the new format -> easier to compare later
data['Datum'] = dates_with_repetitions

# list of all dates on which a nivellement measurement was done -> no repetitions and sorted from earliest to latest
singular_date_list = sorted(list(dict.fromkeys(dates_with_repetitions)))

# we will define an empty dict to put all the needed data into
# the dict will be of following structure:
# {NAME: X_koord, Y_koord, dict of dates the nivellement measurement was undertaken with their Z_koord} e.g.
# dict = {
    # PUNKTNAME_1: [X_1, Y_1, {DATE_1: Z_at_d1, DATE_3: Z_at_d3, DATE_9: Z_at_d9}],     
    # PUNKTNAME_2: [X_2, Y_2, {DATE_2: Z_at_d2, DATE_3: Z_at_d3, DATE_5: Z_at_d5}],
    # .     
    # .     
    # .     
    # PUNKTNAME_N: [X_N, Y_N, {DATE_1: Z_at_d1, DATE_2: Z_at_d2, DATE_3: Z_at_d3}]
    #}
base_set = {}

with tqdm(total=len(data)) as pbar:
    for index, row in data.iterrows():
        NR = row['Punktname_']
        X_koord = row['X']
        Y_koord = row['Y']
        Z_koord = row['Z']
        date = row['Datum']
        if NR not in base_set.keys():
            base_set[NR] = [X_koord, Y_koord, {date: Z_koord}]
        else:
            base_set[NR][2][date] = Z_koord
        pbar.update(1)

# the base set still is not in the correct format
# we still need to do the following things:
#1 relative Z_koords (absolute ones are given) -> subtract Z_koord of earliest date from the others
#!# watch out! commas are needed, after this step no math can be done with the 'floats' anymore
#!# coordinates will be coverted to strings with comma as separator
#2 compare the dates to all possible dates and add non-existing ones -> Z_koord will be taken from an earlier date

for NR in base_set:
    #1 relative Z_koords
    # convert coords to string and use comma
    base_set[NR][0] = str("{:.3f}".format(base_set[NR][0])).replace('.',',')
    base_set[NR][1] = str("{:.3f}".format(base_set[NR][1])).replace('.',',')
    
    # generate sorted date list -> needed for getting the baseline height
    sorted_dates = sorted(base_set[NR][2])
    baseline = base_set[NR][2][sorted_dates[0]]
    
    # subtract baseline from every date -> will make first occurence a zero height (makes sense)
    for date in sorted_dates:
        base_set[NR][2][date] = "{:.3f}".format(base_set[NR][2][date] - baseline).replace('.',',')
    
    #2 compare dates and fill up
    dt_firstdate = datetime.datetime.strptime(sorted_dates[0], 'D_%Y%m%d')
    for checkdate in singular_date_list:
        dt_checkdate = datetime.datetime.strptime(checkdate, 'D_%Y%m%d')
        
        # if checkdate is earlier than the first date it has to be 0,000
        # and it is clear that we do not have to go through the other dates
        if dt_checkdate < dt_firstdate:
            base_set[NR][2][checkdate] = '0,000'
            continue
        
        for date in sorted_dates:
            # we only have to work with the dates not present in the date dict
            # as we are changing the date dict we are using the sorted_dates from before
            if checkdate not in sorted_dates:
                
                # as we are using the sorted lists we can compare their indexes
                checkdate_index = singular_date_list.index(checkdate)
                date_index = singular_date_list.index(date)
                
                # this little algorithm checks if the checkdate is smaller than the date
                # if it is -> it will set checkdate to the height of date
                # because we use sorted lists we are going up from the earliest date
                # each checkdate is compared to every date but because we are going from earliest to latest
                # we are overwriting earlier 'wrong' data
                if checkdate_index > date_index:
                    base_set[NR][2][checkdate] = base_set[NR][2][date]

# put it all back into a DataFrame (tbh could have done it earlier - mb sry)
output_df = DataFrame()

output_df['NR'] = list(base_set)
output_df['X'] = [base_set[key][0] for key in base_set]
output_df['Y'] = [base_set[key][1] for key in base_set]

for date in singular_date_list:
    # nested list comprehension with further if clauses checking and taking data from dict is always funny
    output_df[date] = [base_set[key][2][date_val] for key in base_set for date_val in base_set[key][2].keys() if date_val == date]

# output to .txt file
fmt_list = ['%s']*len(output_df.columns)
header_txt = ' '.join(list(output_df.columns.values))

# if there was no output filename given revert to a default
if ofn == None:
    ofn = 'nivel_preproc.txt'

np.savetxt(ofn, output_df.values, fmt=fmt_list, header=header_txt)