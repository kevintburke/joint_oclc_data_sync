import NZ_Script_1
import OCLC_doublecheck
import NZ_Script_2

def main():
    print("Running NZ_Script_1...")
    NZ_Script_1.merge_reports()
    NZ_Script_1.compare_OCLC()

    print("Obtaining OCLC doublecheck data...")
    OCLC_doublecheck_df = OCLC_doublecheck.get_OCLC_doublecheck()
    OCLC_doublecheck.write_doublecheck_to_excel(OCLC_doublecheck_df)

    print("Running NZ_Script_2...")
    NZ_Script_2.create_reports()

if __name__ == "__main__":
    main()