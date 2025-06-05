import pandas as pd
import requests
import io
import api
from urllib.parse import quote


NUMBER_OF_REQUESTS = 11
API_KEY = api.get_api_key()

def get_OCLC_doublecheck():
    # This function is not used in the current script, but can be defined if needed.
    
    doublecheck_df = pd.DataFrame(columns=["Network Id", "OCLC Control Number (035a)", "OCLC Control Number (035z)", "Bibliographic Lifecycle", "Institution Name"])
    comparison_df = pd.read_excel("comparison_file_IZ.xlsx", sheet_name='DIFF', dtype=str)
    comparison_df.columns = ['JobID', 'Network Id', 'Existing 035a', 'Incoming 035a', 'Action']
    network_ids = comparison_df['Network Id']
    print(network_ids)
    for i in range(0, len(network_ids), NUMBER_OF_REQUESTS):
        filter_xml = """
        <sawx:expr xsi:type="sawx:comparison" op="greaterOrEqual"
        xmlns:saw="com.siebel.analytics.web/report/v1.1"
        xmlns:sawx="com.siebel.analytics.web/expression/v1.1"
        xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
        xmlns:xsd="http://www.w3.org/2001/XMLSchema"
        >

        <sawx:expr xsi:type="sawx:list" op="in">
            <sawx:expr xsi:type="sawx:sqlExpression">"Bibliographic Details"."Network Id"</sawx:expr>
        """
        for network_id in network_ids[i:i + NUMBER_OF_REQUESTS]:
            filter_xml += f'<sawx:expr xsi:type="xsd:string">{str(network_id)}</sawx:expr>'

        filter_xml += """
        </sawx:expr>
        <sawx:expr xsi:type="sawx:comparison" op="equal">
            <sawx:expr xsi:type="sawx:sqlExpression">"Bibliographic Details"."Bibliographic
                Lifecycle"</sawx:expr>
            <sawx:expr xsi:type="xsd:string">In Repository</sawx:expr>
        </sawx:expr>

        </sawx:expr>
        """

        filter_str = " ".join(filter_xml.split())
        encoded_filter = quote(filter_str)
        print(f"Encoded filter: {encoded_filter}")

        result = requests.get(f"https://api-ca.hosted.exlibrisgroup.com/almaws/v1/analytics/reports?path=%2Fshared%2FUTON+Network+01OCUL_NETWORK%2FReports%2FOCLC+Identifiers%2FOCLC-doublecheck&col_names=true&filter={encoded_filter}&apikey={API_KEY}")
        if result.status_code != 200:
            print(f"Error fetching XML: {result.status_code} - {result.reason}")
            exit()
        content = result.content.decode('utf-8')

        contenet_df = pd.read_xml(io.StringIO(content), xpath='.//rs:Row', namespaces={'rs': 'urn:schemas-microsoft-com:xml-analysis:rowset'}, dtype=str)
        contenet_df.columns = ["Column0", "Bibliographic Lifecycle", "Network Id", "OCLC Control Number (035a)", "OCLC Control Number (035z)", "Institution Name"]
        contenet_df.drop(columns=["Column0"], inplace=True)
        contenet_df = contenet_df.reindex(["Network Id", "OCLC Control Number (035a)", "OCLC Control Number (035z)", "Bibliographic Lifecycle", "Institution Name"], axis=1)
        doublecheck_df = pd.concat([doublecheck_df, contenet_df], ignore_index=True)

    return doublecheck_df

def write_doublecheck_to_excel(doublecheck_df):
    try:
        writer = pd.ExcelWriter('OCLC-doublecheck.xlsx', engine='xlsxwriter')
        doublecheck_df.to_excel(writer, index=False)
        writer.close()
    except Exception as e:
        print(f"Error writing to Excel file: {e}")


def main():
    doublecheck_df = get_OCLC_doublecheck()
    if doublecheck_df.empty:
        print("No data found in OCLC doublecheck.")
    
    print("Doublecheck DataFrame:")
    print(doublecheck_df)
    write_doublecheck_to_excel(doublecheck_df)
    print("OCLC doublecheck data written to OCLC-doublecheck_test.xlsx")


if __name__ == "__main__":
    main()
    print("Script completed successfully.")