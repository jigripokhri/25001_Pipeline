# Script to reformat data of Project 25000

import os
import pandas as pd
from datetime import datetime
from sys import exit

# Define input and output folder paths

input_folder = r'R:\400 Services Lab\Projects\25000\01 Administration\04 Data Transfer\01 Formal Data Transfer\02 Data Transfer\Soham\INPUT'
output_folder = r'R:\400 Services Lab\Projects\25000\01 Administration\04 Data Transfer\01 Formal Data Transfer\02 Data Transfer\Soham\OUTPUT'
sdf_file_path = r'R:\400 Services Lab\Projects\25000\01 Administration\06 Documents for Testing\25000 Sample Data (PM001-F03_03).xlsx'


# To be updated; waiting for sponsor confirmation
visit_values = {
    'Screening': 'Visit 2',
    'Visit 1': 'Visit 2',
    'Visit 2': 'Visit 2',
    'Visit 3': 'Visit 3',
    'Visit 6': 'Visit 6',
    'Visit 11': 'Visit 11',
    'Visit 12': 'Visit 12',
    'Visit 16': 'Visit 16',
    'Visit 17': 'Visit 17',
    'Visit 18': 'Visit 18',
    'Visit 19': 'Visit 19',
    'Visit 20': 'Visit 20',
    'Visit 21': 'Visit 21',
    'Visit 22': 'Visit 22',
    'Visit 23': 'Visit 23',
    'Visit 24': 'Visit 24',
    }


# reformat date strings
# input: date string of the format %d-%b-%y or other non-datestring
# output: date string of the format %d %b %Y or other non-datestring
# todo: account for possibility that string given as argument is not a date string
def reformat_date_string(value):
    if type(value) == 'str':
        value = datetime.strftime(
            datetime.strptime(value, '%Y-%m-%d %H:%M:%S'),
            value,
            '%d %b %Y')
    else:
        value = value.strftime('%d %b %Y')
    return value


# reformat time strings
# input: time string of the format %H:%M:%S or other non-timestring
# output: time string of the format '%H:%M' or other non-timestring
# todo: account for possibility that string given as argument is not a time string
def reformat_time_string(value):
    if type(value) == 'str':
        value = datetime.strftime(
            datetime.strptime(value, '%H:%M:%S'),
            '%H:%M')
    else:
        value = value.strftime('%H:%M')
    return value


# adjust visit value
# input: str
# output: str value of key value pair of dictionary visit values or 'Unscheduled'
def visit_alteration(value):
    if value in visit_values.keys():
        return visit_values[value]
    else:
        return 'Unscheduled'


# adjust row of populated fraem df_input, if call value is NMD
def NMD_call_adjustment(row):
    if row['Call'] == 'NMD':
        index_values = list(row.index.values)
        columns_to_be_emptied = index_values[index_values.index('CDS Change'):]
        for column in columns_to_be_emptied:
            row[column] = ''
        row['Call'] = 'NMD'
        row['CDS Change'] = 'NMD'
        row['AA Change'] = 'NMD'
    return row


# Check if the input folder exists
if not os.path.exists(input_folder):
    print(f"Input folder '{input_folder}' does not exist.")
    exit()

# Check if the output folder exists, and create it if it doesn't
if not os.path.exists(output_folder):
    os.makedirs(output_folder)
    print(f"Output folder '{output_folder}' has been created.")

# Check if the SDF file exists
if not os.path.exists(sdf_file_path):
    print(f"SDF file '{sdf_file_path}' does not exist.")
    exit()

# Get a list of Excel files in the input folder
excel_files = [file for file in os.listdir(input_folder) if file.endswith('.xlsx')]

# Get a list of sheet names in the first excel folder
sheet_names = pd.ExcelFile(os.path.join(input_folder, excel_files[0])).sheet_names

# Define the date in the format YYYYMMDD
date_today = datetime.now().strftime('%Y%m%d')

# Function to remove spaces from a string
def remove_spaces(text):
    return text.replace(" ", "")

# Read sdf file
df_sdf = pd.read_excel(sdf_file_path, sheet_name='SampleDataFile')
# rename relevant columns both for merging and customer specifications
df_sdf.rename(columns={
    'InosticsID': 'Sample ID',
    'Visit': 'VISIT',
    'Collection Date': 'CTDNADT',
    'Collection Time': 'CTDNATM',
    'Sample Comment': 'Comment'
    }, inplace=True)
# drop non relevant columns
df_sdf.drop(['No.', 'ReportDate', 'SampleID', 'SpecimenID', 'Subject DOB', 'Received Date', 'Specimen Type', 'Report Comment', 'Site', 'COR Filename', 'CLIA Report Filename', 'Physician', 'Office/Hospital', 'PhysStreetAddress', 'PhysCityStateZip', 'PhysPhone', 'Test ID', 'Study'], axis=1, inplace=True)


for sheet_name in sheet_names:
    # Remove spaces from the sheet name
    sheet_name_without_spaces = remove_spaces(sheet_name)
    output_file_name = f"Sysmex_ctDNAProductionData_{date_today}_{sheet_name_without_spaces}"
    output_file_path = os.path.join(output_folder, output_file_name)
    # Read the existing output file if it exists
    if os.path.exists(output_file_path):
        df_output = pd.read_excel(output_file_path+'.xlsx')
    else:
        df_output = pd.DataFrame()
    # Loop through the Excel files and append data from the input files to the corresponding output files
    for excel_file in excel_files:
        input_file_path = os.path.join(input_folder, excel_file)
        xls = pd.ExcelFile(input_file_path)
        df_input = pd.read_excel(xls, sheet_name)
        
        # Add a static column "PROTOCOL" with the value "ELI-002-001" at the beginning
        df_input.insert(0, 'PROTOCOL', 'ELI-002-001')

        # merge With relevant columns from sdf file
        df_input = df_input.merge(df_sdf, how='left')
        # move columns into position
        visit_column = df_input.pop('VISIT')
        CTDNADT_column = df_input.pop('CTDNADT')
        CTDNATM_column = df_input.pop('CTDNATM')
        df_input.insert(1, 'VISIT', visit_column)
        df_input.insert(2, 'CTDNADT', CTDNADT_column)
        df_input.insert(3, 'CTDNATM', CTDNATM_column)

        # Adjust date and time formats as per customer specifications
        df_input['CTDNADT'] = df_input['CTDNADT'].apply(reformat_date_string)
        df_input['CTDNATM'] = df_input['CTDNATM'].apply(reformat_time_string)

        # alter Visit enumeration
        df_input['VISIT'] = df_input['VISIT'].apply(visit_alteration)

        # adjust values for NMD calls
        if 'Call' in df_input.columns:
            df_input = df_input.apply(
                NMD_call_adjustment,
                axis=1
                )

        # Append the data from the input file to the existing output file
        df_output = pd.concat([df_output, df_input], ignore_index=True)
        
        # Save the combined DataFrame to the output file
        df_output.to_excel(output_file_path+'.xlsx', index=False)
        df_output.to_csv(output_file_path+'.csv', index=False, sep=',')
        print(f"Data from Sheet '{sheet_name}' in '{excel_file}' appended to '{output_file_path}'.")

print("Data separation, renaming, and appending complete.")
