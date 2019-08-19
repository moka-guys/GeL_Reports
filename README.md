# GeL_Reports v1.7

The `gel_cover_report.py` script is used to attach a cover page to the summary of findings downloaded from GeL. It is used for cases that do not require a Geneworks report.

Prior to running the script, the GeL case must be entered into Moka (using form `88005_GEL_InsertPatient` or the automed negneg scripts). If not using the --submit_exit_q and --download_summary flags, the case must be closed through the interpretation portal and the summary of findings PDF must be saved in `P:\Bioinformatics\GeL\technical_reports`.

The script can be used stand alone, or run from Moka.

This script performs the following steps:
* Optionally submits negneg summary of findings and exit questionnaire
* Optionally downloads summary of findings PDF
* Populates the cover page using patient demographics from Moka and Geneworks (An appropriate pre-defined summary is included whihc is based on the Moka result code)
* Updates checker and test status in Moka
* Attaches the report to the NGS test in Moka
* Enters a charge into Geneworks
* Pre-populates an Outlook email with report attached ready to be checked and sent to clinician.

Note the final email step can be run on it's own using the `generate_email.py` script. (This is useful for when cover report generation and email sending need to happen at different times.)

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

### `gel_cover_report.py`

The script is called by passing any number of NGS test IDs as input arguments.

If there is an issue with a case, an error message will be printed to terminal and the script will simply skip to the next case. This is to prevent the whole batch failing when there's an issue with one individual case. It's therefore important to check the output to avoid cases being missed.

By default the script will check that the patient's DoB and NHS number in labkey and Geneworks match. If they don't (or are missing) you
can use the `-skip_labkey` flag to skip this step, however you must manually check that patient details are correct (in case the 100K participant ID is recorded incorrectly in Geneworks).

By default this script will not work for any case where automated reporting has been blocked in Moka (indicated by a non-zero value in the BlockAutomatedReporting in the dbo.NGSTest table). To override this, you can use the `--ignore_block` flag.

```
usage: gel_cover_report.py [-h] -n NGSTestID [NGSTestID ...] [--skip_labkey]
                           [--ignore_block] [--submit_exit_q]
                           [--download_summary]

Creates cover page for GeL results and attaches to report provided by GeL

optional arguments:
  -h, --help            show this help message and exit
  -n NGSTestID [NGSTestID ...]
                        Moka NGSTestID from NGSTest table
  --skip_labkey         Optional flag to skip the check that DOB and NHS
                        number in LIMS match labkey before reporting.
  --ignore_block        Optional flag to allow reporting of blocked cases.
  --submit_exit_q       Optional flag to submit a negneg clinical report and
                        exit questionnaire automatically to CIP-API
  --download_summary    Optional flag to download summary of findings
                        automatically from CIP-API to
                        P:\Bioinformatics\GeL\technical_reports
```

### `generate_email.py`

This script can be used standalone to populate an Outlook email with supplied values.

```
usage: generate_email.py [-h] -t TO -s SUBJECT -b BODY
                         [-a ATTACHMENTS [ATTACHMENTS ...]]

Generates an email and opens in Outlook

optional arguments:
  -h, --help            show this help message and exit
  -t TO, --to TO        Recipient email address
  -s SUBJECT, --subject SUBJECT
                        Subject line for email
  -b BODY, --body BODY  Email body
  -a ATTACHMENTS [ATTACHMENTS ...], --attachments ATTACHMENTS [ATTACHMENTS ...]
                        File paths for attachments
```
