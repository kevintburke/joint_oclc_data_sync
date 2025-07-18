import os

# Directories
INPUT_FOLDER = "./Inputs"
OUTPUT_FOLDER = "./Outputs"

# Files
INSTITUTIONS_JSON_FILE = "institutions.json"
DO_NOT_CHANGE_FILE = "Do_not_change.xlsx"
BIBPROCESS_MERGED_FILE = os.path.join(OUTPUT_FOLDER, "bibprocessmerged.csv")
COMPARISON_FILE = os.path.join(OUTPUT_FOLDER, "comparison_file_IZ.xlsx")
IMPORT_TO_NZ_FILE = os.path.join(OUTPUT_FOLDER, "for_import_to_NZ.xlsx")
UPDATED_NUMBERS_FILE = os.path.join(OUTPUT_FOLDER, "oclc_numbers_updated.txt")
MMSID_FOR_SET_FILE = os.path.join(OUTPUT_FOLDER, "a_to_z_MMSid_for_set.txt")

#File Names (These constants are either only a section of a file name or the file's location is unknown)
INPUT_CSV_FILE_NAME = "testfile.txt"
BIBPROCESS_IZ_FILE_NAME = "bibprocess.csv"
DOUBLECHECK_FILE = "OCLC-doublecheck.xlsx"
OCLC_IDENTIFIER_REPORT_FILE = "NZ-OCLC-Identifier-report_"
