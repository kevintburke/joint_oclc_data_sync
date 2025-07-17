import pandas as pd
from pandas import DataFrame, merge, ExcelWriter
import datetime
import constants
import os

def create_reports():
    #Find doublecheck file
    doublecheck = None
    for root, dirs, files in os.walk("."):
        if constants.DOUBLECHECK_FILE in files:
            doublecheck = os.path.join(root, constants.DOUBLECHECK_FILE)

    if not AlmaExport:
        print("Could not find doublecheck file")
        return

    #Read Excel file created with Alma Analytics 'OCLC-doublecheck'
    AlmaExport = pd.read_excel(doublecheck, dtype=str)
    AlmaExport.columns =["Network Id","OCLC Control Number (035a)","OCLC Control Number (035z)","Bibliographic Lifecycle", "Institution Name"]

    #Create a dataframe for 'OCLC-doublecheck'
    df1 = pd.DataFrame(AlmaExport, columns= ['Network Id','OCLC Control Number (035a)','OCLC Control Number (035z)','Institution Name'])

    #comparison_file
    comparison_file = constants.COMPARISON_FILE

    #Read the DIFF tab of the comparison_IZ file created with NZ_Script_1
    Import = pd.read_excel(comparison_file, sheet_name='DIFF', dtype=str)
    Import.columns = ['JobID', 'Network Id', 'Existing 035a', 'Incoming 035a', 'Action']

    # Dataframe for do not change list
    values = pd.read_excel(constants.DO_NOT_CHANGE, sheet_name='Do_not_change', dtype=str)
    values.columns = ["Network Id"]
    values_df = pd.DataFrame(values, columns=['Network Id'])
    values_df['Network Id'] = values_df['Network Id'].astype(str)
    print(values_df)

    # Create a dataframe from the comparison_IZ file, setting the format as text for the MMS ID.
    df2 = pd.DataFrame(Import, columns=['Network Id', 'Existing 035a', 'Incoming 035a', 'Action'])
    df2['Network Id'] = df2['Network Id'].astype(str)
    print(df2)

    # Drop rows in df2 that have Network Id in values_df and create do_not_change dataframe
    do_not_change = df2[df2['Network Id'].isin(values_df['Network Id'])]
    df2 = df2[~df2['Network Id'].isin(values_df['Network Id'])]

    print(do_not_change)
    print(df2)

    # Filter rows where "Action" is not "match" and move them to a new dataframe
    non_match_df = df2[df2['Action'] != 'match']
    print(non_match_df)

    # Keep rows where "Action" is "match" in the original dataframe
    match_df = df2[df2['Action'] == 'match']
    print(match_df)

    # Merge the Alma Analytic and match set
    match = match_df.merge(df1, on='Network Id', how='inner')
    print(match)

    # Merge the Alma Analytic report and non-match set
    review = non_match_df.merge(df1, on='Network Id', how='inner')
    print(review)

    diff = match[match['Incoming 035a'] != match['OCLC Control Number (035a)']]
    updated = match[match['Incoming 035a'] == match['OCLC Control Number (035a)']]
    print(diff)
    print(updated)

    # Get the current date
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")

    #Create the analysis file
    filename = os.path.join(constants.OUTPUT_FOLDER, f"{constants.OCLC_IDENTIFIER_REPORT_FILE}{current_date}.xlsx")
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    do_not_change.to_excel(writer, sheet_name='do_not_change')
    diff.to_excel(writer, sheet_name='update by job')
    updated.to_excel(writer, sheet_name='updated records')
    review.to_excel(writer, sheet_name='records for review')
    writer.close()

    print(f"File saved as {filename}")

    # Read the Excel file
    df = pd.read_excel(filename, sheet_name='update by job', dtype=str)

    # Create a text file for column '001' with header 'MMS ID'
    df['Network Id'].to_csv(constants.MMSID_FOR_SET_FILE, header=['MMS Id'], index=False)

    # Create a text file with OCLC numbers for deduplication 
    df['Network Id'].to_csv(constants.UPDATED_NUMBERS_FILE, header=['Incoming 035a'], index=False)

    # Create a new DataFrame for 'for_import_to_NZ' with columns '001' and '035 $a'
    for_import_to_NZ = df[['Network Id', 'Incoming 035a']]
    for_import_to_NZ.columns = ['001','035 $a']

    with pd.ExcelWriter(constants.IMPORT_TO_NZ_FILE, engine='xlsxwriter') as writer:
        for_import_to_NZ.to_excel(writer, index=False, sheet_name='for_import_to_NZ')

    # Get the xlsxwriter workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['for_import_to_NZ']

    # Set the text format for the entire sheet
    text_fmt = workbook.add_format({'num_format': '@'})
    worksheet.set_column('A:XFD', None, text_fmt)  # Set the entire sheet to text format

if __name__ == "__main__":
    create_reports()