# joint_oclc_data_sync
Joint processing of OCLC data sync for Collaborative Futures institutions.

These scripts are currently used to batch update the OCLC numbers for NZ records published to OCLC through the weekly publishing jobs at Western and Carleton. They require the following inputs:

- 2 bib processing reports (one from Western, one from Carleton), named mergetest1.txt and mergetest2.txt
- 1 do-not-change file listing records to skip in the batch update process
- 1 OCLC-doublecheck Analytics report (NOTE: requires the output of NZ script 1 to generate)

The scripts return the following outputs:

- CEFbibprocess - a file listing Carleton's IZ records from the publishing job
- UWObibprocess - a file listing Western's IZ records from the publishing job
- Comparison_file_IZ - used to filter the OCLC-doublecheck analytics report
- a_to_z_MMSid_for_set - used to create an itemized set of records for batch updating in the NZ
- for_import_to_NZ - used to import updated OCLC #s to records in itemized set
- NZ-OCLC-Identifier-report-<date> - lists all records updated and actions taken
- Various placeholder files (useful for troubleshooting, but not required for workflow)

The current proposed workflow for this process follows:

_Data Sync Report Processing_
- Collect bib processing reports from each institution and combine using script (NZ Script 1.py)

-- Combines reports matching and deduping on NZ ID

-- Outputs reports of IZ record divided by institution and shared records

-- Shared records divided by whether the incoming OCLC # matches the existing one (if any) or not
- Collect report from Analytics with existing 035 $a and $z, institutions with holdings
- Run results of NZ Script 1, “don’t update” list, and Analytics through NZ Script 2.py

-- Eliminates records marked as “don’t update” from batch process list

-- Separates records by action type

--- See OCLC documentation for action types

--- Records with “match” action selected for batch processing; other records require manual review

-- Sets aside records with already updated OCLC #s in 035 $a

_Batch Updating Records in NZ_
- Using outputs of NZ Script 2, create an itemized set in the NZ Alma Instance of titles requiring OCLC # updates

-- Uses output a_to_z_MMSid_for_set.txt
- Run job “Move OCLC 035 $a to $z” on itemized set to move all OCLC #s in 035 $a to $z
- Using import profile OCLC number (035) import - From XLS, import new OCLC #s to add to records in itemized set

-- Uses output for_import_to_NZ.xlsx

_These scripts were copied and adapted from ones developed by Erin Bourgard and available at https://github.com/ernieejo/OCLC-imports-for-Alma-Network_
