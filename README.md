# GeL_Reports v1.6

This script is used to attach a cover page to the summary of findings downloaded from GeL. It is used for negative cases that do not require a Geneworks report.

Prior to running the script, the GeL case must be entered into Moka (using form `88005_GEL_InsertPatient`) and the summary of findings PDF must be saved in `P:\Bioinformatics\GeL\technical_reports`.

The script can be used stand alone, or for negative negative cases can be run from Moka using form `88009_GEL_neg_neg_list`.

This script performs the following steps:
* Optionally downloads summary of findings PDF 
* Populates the cover page using patient demographics from Moka and Geneworks (An appropriate pre-defined summary is included whihc is based on the Moka result code)
* Updates checker and test status in Moka
* Attaches the report to the NGS test in Moka
* Enters a charge into Geneworks
* Pre-populates an Outlook email with report attached ready to be checked and sent to clinician.

## Usage

Requirements:

* ODBC connection to Moka and Geneworks
* Python 2.7
* Python packages:
    * pyodbc
    * pdfkit
    * PyPDF2
    * jinja2

Using the python installation at `S:\Genetics_Data2\Array\Software\Python\python.exe` will satisfy the above requirements.

The script is called by passing any number of NGS test IDs as input arguments.

If there is an issue with a case, an error message will be printed to terminal and the script will simply skip to the next case. This is to prevent the whole batch failing when there's an issue with one individual case. It's therefore important to check the output to avoid cases being missed.

```
usage: gel_cover_report.py [-h] -n NGSTestID [NGSTestID ...]
                           [--download_summary]

Creates cover page for GeL results and attaches to report provided by GeL

optional arguments:
  -h, --help            show this help message and exit
  -n NGSTestID [NGSTestID ...]
                        Moka NGSTestID from NGSTest table
  --download_summary    Optional flag to download summary of findings
                        automatically from CIP-API to
                        P:\Bioinformatics\GeL\technical_reports
```
