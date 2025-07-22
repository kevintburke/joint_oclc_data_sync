import pandas as pd
from tkinter import messagebox, filedialog
import constants
import json
import os

NZID = "5151"
CEFID = "5153"
UWOID = "5163"

#Merge Bib Processing Reports and dedupe using NZ ID, then split into NZ and IZ reports
def merge_reports():
    #Prompt user for .txt files and open both

    with open(constants.INSTITUTIONS_JSON_FILE) as institutions_json:
        institutions = json.load(institutions_json)

    columns=["JobID", "Network Id", "Existing 035a", "Incoming 035a", "Action"]
    merged = pd.DataFrame(columns=columns)

    for file in os.listdir(constants.INPUT_FOLDER):
        with open(os.path.join(constants.INPUT_FOLDER, file)) as csv:
            df = pd.read_csv(csv, sep="|", header=None, dtype=str, names=columns)

        merged = pd.merge(merged, df, how="outer")
        merged.drop_duplicates(subset=["Network Id"], keep="first", inplace = True)
        merged["Network Id"] = merged["Network Id"].astype(str)

    if merged.empty:
        raise Exception("Merged DataFrame is empty. Make sure BibProcessing files are in Inputs directory.")
    for library in institutions:
        id = library["id"]
        code = library["code"]

        bibprocess = merged[merged['Network Id'].str.endswith(id)]
        bibprocess_file = os.path.join(constants.OUTPUT_FOLDER, f"{code}{constants.BIBPROCESS_IZ_FILE_NAME}")
        bibprocess.to_csv(bibprocess_file, mode = "w", index = False)
        print(f"{code} records saved to {code}{constants.BIBPROCESS_IZ_FILE_NAME}")

    nz = merged[merged['Network Id'].str.endswith(NZID)]
    nz.to_csv(constants.BIBPROCESS_MERGED_FILE, mode = "w", index = False)
    print(f"NZ records saved to {constants.BIBPROCESS_MERGED_FILE}")


def merge_reports_old():
    messagebox.showinfo(title=None, message='Please select the Bib Processing Reports to load and compare.')
    fn1 = filedialog.askopenfilename()
    f1 = open(fn1, 'r')
    fn2 = filedialog.askopenfilename()
    f2 = open(fn2, 'r')

    #Create dataframes with labelled columns (copied from UWO code)
    df1 = pd.read_csv(f1, sep="|", header=None, dtype=str)
    df1.columns = ["JobID", "Network Id", "Existing 035a", "Incoming 035a", "Action"]
    df1 = pd.DataFrame(df1, columns = ["JobID", "Network Id", "Existing 035a", "Incoming 035a", "Action"])

    df2 = pd.read_csv(f2, sep="|", header=None, dtype=str)
    df2.columns = ["JobID", "Network Id", "Existing 035a", "Incoming 035a", "Action"]
    df2 = pd.DataFrame(df2, columns = ["JobID", "Network Id", "Existing 035a", "Incoming 035a", "Action"])

    #Merge frames and drop duplicates by NZ ID
    merged = pd.merge(df1, df2, how="outer")
    merged.drop_duplicates(subset=["Network Id"], keep="first", inplace = True)
    merged["Network Id"] = merged["Network Id"].astype(str)
    #Read 035a as str

    #Split off IZ records (MMS ID 18 digits, not ending in 5151)
    #Save CEF records to csv
    CEF = merged[merged['Network Id'].str.endswith(CEFID)]
    CEF.to_csv("CEFbibprocess.csv", mode = "w", index = False)
    print("CEF records saved to CEFbibprocess.csv")

    #Save UWO records to csv
    UWO = merged[merged['Network Id'].str.endswith(UWOID)]
    UWO.to_csv("UWObibprocess.csv", mode = "w", index = False)
    print("UWO records saved to UWObibprocess.csv")

    #Save NZ records to csv
    NZ = merged[merged['Network Id'].str.endswith(NZID)]
    NZ.to_csv("bibprocessmerged.csv", mode = "w", index = False)
    print("NZ records saved to bibprocessmerged.csv")


def compare_OCLC(): #Copied from UWO code
    #Read the BIB processing report. Add the filepath to the txt file
    data = pd.read_csv(constants.BIBPROCESS_MERGED_FILE)

    #make a dataframe from the BIB processing report and set the format as text for columns b and c
    df = pd.DataFrame(data, columns= ['JobID', 'Network Id', 'Existing 035a', 'Incoming 035a', 'Action'])
    df['Network Id']= df['Network Id'].astype(str)

    #Define the sheets that we will have in the comparison_fileIZ
    DIFF = df[df['Existing 035a'] != df['Incoming 035a']]
    SAME = df[df['Existing 035a'] == df['Incoming 035a']]

    print(DIFF)
    print(SAME)

    #Create the comparison_fileIZ file. Add the file path to the Excel file
    writer = pd.ExcelWriter(constants.COMPARISON_FILE, engine='xlsxwriter')
    DIFF.to_excel(writer, index=False, sheet_name='DIFF')
    SAME.to_excel(writer, index=False, sheet_name='SAME')
    workbook = writer.book
    worksheet = writer.sheets['DIFF']
    text_fmt = workbook.add_format({'num_format': '@'})
    worksheet.set_column('B:B',20, text_fmt)
    writer.close()

def main():
    merge_reports()
    return
    # compare_OCLC()

if __name__ == '__main__':
    main()
