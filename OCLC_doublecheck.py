from math import ceil
import pandas as pd
import requests
import io
import api
import constants
import os
from urllib.parse import quote

NETWORK_IDS_PER_REQUEST = 5
API_KEY = api.get_api_key()

def get_OCLC_doublecheck():
    """
    Fetches OCLC doublecheck data from Alma Analytics and returns it as a DataFrame.
    """
    if not API_KEY:
        print("API key is not set. Please set the NZ_API_KEY environment variable.")
        return
    if NETWORK_IDS_PER_REQUEST <= 0:
        print("NETWORK_IDS_PER_REQUEST must be greater than 0.")
        return

    # Get network IDs from comparision_file_IZ.xlsx
    doublecheck_df = pd.DataFrame(columns=["Network Id", "OCLC Control Number (035a)", "OCLC Control Number (035z)", "Bibliographic Lifecycle", "Institution Name"])
    comparison_df = pd.read_excel(constants.COMPARISON_FILE, sheet_name='DIFF', dtype=str)
    comparison_df.columns = ['JobID', 'Network Id', 'Existing 035a', 'Incoming 035a', 'Action']
    network_ids = comparison_df['Network Id']

    # Send NETWORK_IDS_PER_REQUEST requests to the API to fetch OCLC doublecheck data
    number_of_requests = ceil(len(network_ids) / NETWORK_IDS_PER_REQUEST)
    request_count = 1
    for i in range(0, len(network_ids), NETWORK_IDS_PER_REQUEST):

        print(f"Processing request {request_count} of {number_of_requests}")

        # Construct the filter XML for the API request
        xml_filter = """
        <sawx:expr xsi:type="sawx:comparison" op="greaterOrEqual"
        xmlns:saw="com.siebel.analytics.web/report/v1.1"
        xmlns:sawx="com.siebel.analytics.web/expression/v1.1"
        xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
        xmlns:xsd="http://www.w3.org/2001/XMLSchema"
        >

        <sawx:expr xsi:type="sawx:list" op="in">
            <sawx:expr xsi:type="sawx:sqlExpression">"Bibliographic Details"."Network Id"</sawx:expr>
        """
        for network_id in network_ids[i:i + NETWORK_IDS_PER_REQUEST]:
            xml_filter += f'<sawx:expr xsi:type="xsd:string">{str(network_id)}</sawx:expr>'
        xml_filter += """
        </sawx:expr>
        <sawx:expr xsi:type="sawx:comparison" op="equal">
            <sawx:expr xsi:type="sawx:sqlExpression">"Bibliographic Details"."Bibliographic Lifecycle"</sawx:expr>
            <sawx:expr xsi:type="xsd:string">In Repository</sawx:expr>
        </sawx:expr>

        </sawx:expr>
        """

        # URL encode the filter
        filter_str = " ".join(xml_filter.split())
        encoded_filter = quote(filter_str)
        url = f"https://api-ca.hosted.exlibrisgroup.com/almaws/v1/analytics/reports?path=%2Fshared%2FUTON+Network+01OCUL_NETWORK%2FReports%2FOCLC+Identifiers%2FOCLC-doublecheck&col_names=true&filter={encoded_filter}&apikey={API_KEY}"
        print(f"URL: {url}")
        # Make the API request
        result = requests.get(url)
        if result.status_code != 200:
            print(f"Error fetching XML: {result.status_code} - {result.text}")
            exit()
        content = result.content.decode('utf-8')
        print("Successfully obtained data from API.")

        # Parse the XML content into a DataFrame and merge it with the existing DataFrame
        print("Adding data to DataFrame...")
        content_df = pd.read_xml(io.StringIO(content), xpath='.//rs:Row', namespaces={'rs': 'urn:schemas-microsoft-com:xml-analysis:rowset'}, dtype=str)

        # Ensure all expected columns exist, even if missing in some rows
        for col in ["Column0", "Column1", "Column2", "Column3", "Column4", "Column5"]:
            if col not in content_df.columns:
                content_df[col] = None  # or pd.NA

        # Reorder columns to expected order
        content_df = content_df[["Column0", "Column1", "Column2", "Column3", "Column4", "Column5"]]

        content_df.columns = ["Column0", "Bibliographic Lifecycle", "Network Id", "OCLC Control Number (035a)", "OCLC Control Number (035z)", "Institution Name"]
        content_df.drop(columns=["Column0"], inplace=True)
        content_df = content_df.reindex(["Network Id", "OCLC Control Number (035a)", "OCLC Control Number (035z)", "Bibliographic Lifecycle", "Institution Name"], axis=1)
        doublecheck_df = pd.concat([doublecheck_df, content_df], ignore_index=True)

        request_count += 1

        print("Successfully added new data to DataFrame.")
    print("Request processing complete.")
    print("Doublecheck DataFrame: ", doublecheck_df)
    return doublecheck_df

def write_doublecheck_to_excel(doublecheck_df):
    """
    Writes the OCLC doublecheck DataFrame to an Excel file.
    """

    if (doublecheck_df.empty):
        return

    try:
        doublecheck_file = os.path.join(constants.OUTPUT_FOLDER, constants.DOUBLECHECK_FILE)
        writer = pd.ExcelWriter(doublecheck_file, engine='xlsxwriter')
        doublecheck_df.to_excel(writer, index=False)
        writer.close()
    except Exception as e:
        print(f"Error writing to Excel file: {e}")


def main():
    doublecheck_df = get_OCLC_doublecheck()
    if doublecheck_df.empty:
        print("No data found in OCLC doublecheck.")

    write_doublecheck_to_excel(doublecheck_df)
    print("OCLC doublecheck data written to OCLC-doublecheck_test.xlsx")


if __name__ == "__main__":
    main()
    print("Script completed successfully.")