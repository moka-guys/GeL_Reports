"""
v1.6 - AB 2019/01/22
Requirements:
    ODBC connection to Moka
    Python 2.7
    pyodbc
    pdfkit
    PyPDF2
    jinja2

usage: gel_cover_report.py [-h] -n NGSTestID [NGSTestID ...]
                           [--download_summary]

Creates cover page for GeL results and attaches to report provided by GeL

optional arguments:
  -h, --help            show this help message and exit
  -n NGSTestID [NGSTestID ...]
                        Moka NGSTestID from NGSTest table
  --download_summary    Optional flag to download summary of findings
                        automatically from CIP-API
"""
import sys
import os
import io
import re
import argparse
import datetime
import fnmatch
from ConfigParser import ConfigParser
import win32com.client as win32
import pyodbc
import pdfkit
from PyPDF2 import PdfFileMerger
from jinja2 import Environment, FileSystemLoader
from ssh_run_summary_findings import SummaryFindings_SSH

# Read config file
config = ConfigParser()
config.read(r"F:\Moka\Files\Software\100K\config.ini")

def process_arguments():
    """
    Uses argparse module to define and handle command line input arguments and help menu
    """
    # Create ArgumentParser object. Description message will be displayed as part of help message if script is run with -h flag
    parser = argparse.ArgumentParser(description='Creates cover page for GeL results and attaches to report provided by GeL')
    # Define the arguments that will be taken. nargs='+' allows multiple NGSTestIDs from NGSTest table in Moka can be passed as arguments.
    parser.add_argument('-n', metavar='NGSTestID', required=True, type=int, nargs='+', help='Moka NGSTestID from NGSTest table')
    parser.add_argument('--download_summary', action='store_true', help='Optional flag to download summary of findings automatically from CIP-API')
    # Return the arguments
    return parser.parse_args()

def generate_email(to_address, subject, body, attachment):
    '''
    Populates an Outlook email and opens in separate window
    '''
    # Create Outlook message object
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    # Set email attributes
    mail.To = to_address
    mail.Subject = subject
    mail.HtmlBody = body
    # Attach file
    mail.Attachments.Add(Source=attachment)
    # Open the email in outlook. False argument prevents the Outlook window from blocking the script 
    mail.Display(False)

class GeLGeneworksCharge(object):
    def __init__(self):
        # PRU
        self.pru = None
        # Geneworks test ID for adding charge
        self.test_id = None
        # Geneworks specimen ID for adding charge
        self.specimen_id = None

    def get_test_details(self, pru):
        """
        Retrieves required Geneworks details for entering charge
        """
        # Store the PRU for use by other functions
        self.pru = pru
        # establish pyodbc connection to Moka
        cnxn = pyodbc.connect('DRIVER={{SQL Server}}; SERVER={server}; DATABASE={database};'.format(
            server=config.get("MOKA", "SERVER"),
            database=config.get("MOKA", "DATABASE")
            ), 
            autocommit=True
        )
        # return cursor to execute query
        cursor = cnxn.cursor()
        # query to retrieve the test ID and specimen number for the test. These are required for the stored procedure that adds the charge.
        # DisorderID 60 is 'Whole Genome Sequencing' 
        # TestDescriptionID 18 is 'GeL'
        # There can be multiple specimens/tests per patient with these Disorder and TestDescriptionIDs, (where the omics samples that are stored have also had these tests added)
        # Therefore need to select the sample that has actually had DNA extracted, hence the inner join to the dnanumberlinked view.
        test_details_sql = (
            'SELECT "gwv-testlinked".TestID, "gwv-specimenlinked".SpecimenTrustID '
            'FROM ("gwv-dnatestrequestlinked" INNER JOIN (("gwv-testlinked" INNER JOIN "gwv-specimenlinked" ON "gwv-testlinked".SpecimenID = "gwv-specimenlinked".SpecimenID) '
            'INNER JOIN "gwv-patientlinked" ON "gwv-specimenlinked".PatientID = "gwv-patientlinked".PatientID) ON "gwv-dnatestrequestlinked".TestID = "gwv-testlinked".TestID) '
            'INNER JOIN "gwv-dnanumberlinked" ON "gwv-specimenlinked".SpecimenID = "gwv-dnanumberlinked".SpecimenID '
            'WHERE "gwv-patientlinked".PatientTrustID = \'{pru}\' AND "gwv-dnatestrequestlinked".DisorderID = 60 AND "gwv-dnatestrequestlinked".TestDescriptionID = 18'
        ).format(pru=pru)
        # Exexute the query
        rows = cursor.execute(test_details_sql).fetchall()
        # Check that there is only 1 record returned and set the test id and spec id attributes. If a different number were returned, leave as None and print error message.
        if len(rows) == 1:
            self.test_id = rows[0].TestID
            self.specimen_id = rows[0].SpecimenTrustID
        elif len(rows) == 0:
            print "ERROR: Unable to find GeL test in Geneworks for {pru}.".format(pru=pru)
        elif len(rows) > 1:
            print "ERROR: Multiple GeL tests found in Geneworks for {pru}.".format(pru=pru)

    def insert_charge(self):
        """
        Inserts GeL charge to GeneWorks for given PRU
        """
        if self.test_id and self.specimen_id:
            # establish pyodbc connection to Geneworks
            cnxn = pyodbc.connect('DRIVER={{SQL Server}}; SERVER={server}; DATABASE={database}; UID={user}; PWD={password};'.format(
                server=config.get("GENEWORKS", "SERVER"),
                database=config.get("GENEWORKS", "DATABASE"),
                user=config.get("GENEWORKS", "USER"),
                password=config.get("GENEWORKS", "PASSWORD")
                ), 
            autocommit=True
            )
            # return cursor to execute query
            cursor = cnxn.cursor()
            # Add in sql for spInsertLabReportCostDetail
            # Necessary to use SET NOCOUNT ON to prevent 'pyodbc.ProgrammingError: No results.  Previous SQL was not a query.' error, see:
            # https://stackoverflow.com/questions/7753830/mssql2008-pyodbc-previous-sql-was-not-a-query
            charge_sql = (
                'SET NOCOUNT ON; '
                'DECLARE @return_value int, @RecordNo int, @timestamp datetime; '
                'set @timestamp = getdate(); '
                'EXEC @return_value = [dbo].[spInsertLabReportCostDetail] '
                '@SpecimenNo = \'{specimen_id}\', '
                '@CostTypeID = NULL, '
                '@Destination = NULL, '
                '@DateReported = @timestamp, '
                '@TestType = N\'GEL NEG\', '
                '@Cost = 150, '
                '@DNAUnitsSample = NULL, '
                '@DNAUnitsAnalysis = NULL, '
                '@DNAUnitsReport = NULL, '
                '@DNAUnitsTotal = 1, '
                '@CytoUnits = NULL, '
                '@EnteredByID = 888, ' # This is the user 'moka' in geneworks
                '@PCRFee = NULL, '
                '@TestID = {test_id}, '
                '@RecordNo = @RecordNo OUTPUT; '
                'SELECT @RecordNo as record_no;'
            ).format(
                specimen_id=self.specimen_id, 
                test_id=self.test_id
                )            
            try:
                # Insert the charge and capture the returned record number
                rows = cursor.execute(charge_sql).fetchall()
                # Check that one and only one record has been updated: 
                # Print error message if anything other than 1 record number is returned from the query
                if len(rows) != 1:
                    print "ERROR: When inserting charge for {pru}, {records} were updated.".format(pru=self.pru, records=len(rows))
            except:
                # Error message if theres an error inserting charge
                print "ERROR: Encountered error when inserting charge for {pru}.".format(pru=self.pru)

        else:
            print "ERROR: Unable to insert charge for {pru}. Test ID or Specimen ID not determined. Please add manually.".format(pru=self.pru)
            
        

class MokaQueryExecuter(object):
    def __init__(self):
        # establish pyodbc connection to Moka
        cnxn = pyodbc.connect('DRIVER={{SQL Server}}; SERVER={server}; DATABASE={database};'.format(
            server=config.get("MOKA", "SERVER"),
            database=config.get("MOKA", "DATABASE")
            ), 
            autocommit=True
        )
        # return cursor to execute query
        self.cursor = cnxn.cursor()

    def execute_query(self, sql):
        """
        Executes a supplied SQL query
        """
        self.cursor.execute(sql)

    def get_data(self, ngs_test_id):
        """
        Takes a Moka NSGTestID as input.
        Pulls out details from Moka needed to populate the cover page. 
        """
        data_sql = (
            'SELECT NGSTest.NGSTestID, NGSTest.InternalPatientID, NGSTest.ResultCode, Checker.Name AS clinician_name, Checker.ReportEmail, Item_Address.Item AS clinician_address, '
            '"gwv-patientlinked".FirstName, "gwv-patientlinked".LastName, "gwv-patientlinked".DoB, "gwv-patientlinked".Gender, "gwv-patientlinked".NHSNo, '
            '"gwv-patientlinked".PatientTrustID, NGSTest.GELProbandID, NGSTest.IRID '
            'FROM (((NGSTest INNER JOIN Patients ON NGSTest.InternalPatientID = Patients.InternalPatientID) '
            'INNER JOIN "gwv-patientlinked" ON "gwv-patientlinked".PatientTrustID = Patients.PatientID) INNER JOIN Checker ON NGSTest.BookBy = Checker.Check1ID) '
            'INNER JOIN Item AS Item_Address ON Checker.Address = Item_Address.ItemID '
            'WHERE NGSTestID = {ngs_test_id};'
            ).format(ngs_test_id=ngs_test_id)
        # Execute the query to get patient data
        row = self.cursor.execute(data_sql).fetchone()
        # If results have been returned from the query...
        if row:
            # Populate data dictionaries with values returned by query
            data = {
                'clinician': row.clinician_name,
                'clinician_report_email': row.ReportEmail,
                'clinician_address': row.clinician_address,
                'internal_patient_id': row.InternalPatientID,
                'patient_name': '{first_name} {last_name}'.format(first_name=row.FirstName, last_name=row.LastName),
                'sex': row.Gender,
                'DOB': row.DoB.strftime(r'%d/%m/%Y'), # Extract date from datetime field in format dd/mm/yyyy
                'NHSNumber': row.NHSNo,
                'PRU': row.PatientTrustID,
                'GELID': row.GELProbandID,
                'IRID': row.IRID,
                'date_reported': datetime.datetime.now().strftime(r'%d/%m/%Y'), # Current date in format dd/mm/yyyy
                'result_code': row.ResultCode
            }
            # If None has been returned for gender (because there isn't one in geneworks) change value to 'Unknown'
            if not data['sex']: 
                data['sex'] = 'Unknown'
            return data

class GelReportGenerator(object):
    def __init__(self, path_to_wkhtmltopdf):
        # path to wkhtmltopdf executable used by pdfkit
        self.path_to_wkhtmltopdf = path_to_wkhtmltopdf
        # Attribute to hold the in-memory cover file
        self.cover_pdf = None

    def create_cover_pdf(self, data, template):
        """
        Populate html template with data and store as pdf
        """
        # specify the folder containing the html template for cover report 
        html_template_dir = Environment(loader=FileSystemLoader(os.path.dirname(template)))
        # specify which html template to use
        html_template = html_template_dir.get_template(os.path.basename(template))
        # populate the template with values from data dictionary
        cover_html = html_template.render(data)
        # Specify path to wkhtmltopdf executable (used by pdfkit)
        pdfkit_config = pdfkit.configuration(wkhtmltopdf=self.path_to_wkhtmltopdf)
        # Specify options. 'quiet' turns off verbose stdout when writing to pdf.
        pdfkit_options = {'quiet':''}
        # Convert html to PDF. Set output_path to False so that it returns a byte string rather than writing out to file.
        cover_pdf = pdfkit.from_string(cover_html, output_path=False, configuration=pdfkit_config, options=pdfkit_options)
        # Read the byte string into an in memory file-like object
        self.cover_pdf = io.BytesIO(cover_pdf)

    def pdf_merge(self, output_file, *pdfs):
        """
        Takes multiple PDF filepaths and merges into one document.
        Outputs to filepath specified in merged_report 
        """
        # Create PdfFileMerger object
        merger = PdfFileMerger()
        # Concatenate the PDFs together. PDFs output from wkhtmltopdf break it if import_bookmarks is set to True (see https://github.com/mstamy2/PyPDF2/issues/193)
        [merger.append(pdf, import_bookmarks=False) for pdf in pdfs]
        # Write out the merged PDF report
        with open(output_file, 'wb') as merged_report:
            merger.write(merged_report)

def main():
    # Output folder for combined reports
    gel_report_output_folder = r'\\gstt.local\apps\Moka\Files\ngs\{year}\{month}'.format(
        year=datetime.datetime.now().year,
        month=datetime.datetime.now().month
        )
    # Get command line arguments
    args = process_arguments()
    # Create MokaQueryExecuter object
    moka = MokaQueryExecuter()
    # Loop through each Moka NGStestID supplied as an argument
    for ngs_test_id in args.n:
        # Get data for cover page from Moka.
        data = moka.get_data(ngs_test_id)
        # If no data are returned, print an error message
        if not data:
            print 'ERROR: No results returned from Moka data query for NGSTestID {ngs_test_id}. Check there are records in all inner joined tables (eg clinician address in checker table)'.format(ngs_test_id=ngs_test_id)
        # identify any missing values. if any of the values in the dict are Null (None)
        elif None in data.values():
            # loop through each key
            for field in data:
                # if the value is none
                if not data[field]:
                    # print the missing field.
                    print "ERROR: No {field} value in Moka for NGSTestID {ngs_test_id}".format(field=field, ngs_test_id=ngs_test_id)
        # Check that interpretation request ID is in expected format
        elif not re.search("^\d*-\d*$", data['IRID']):
            print "ERROR: Interpretation request ID {irid} does not match pattern <id>-<version> for NGSTestID {ngs_test_id}".format(ngs_test_id=ngs_test_id, irid=data['IRID'])
        # Otherwise continue...
        else:
            # If download_summary flag is used, call script to download the summary of findings report from CIP-API
            if args.download_summary:
                ir_id = data['IRID'].split("-")[0]
                ir_version = data['IRID'].split("-")[1]
                try:
                    SummaryFindings_SSH(
                        ir_id=ir_id,
                        ir_version=ir_version,
                        output_path=r"\\gstt.local\shared\Genetics\Bioinformatics\GeL\technical_reports\ClinicalReport_{ir_id}-{ir_version}-1.pdf".format(ir_id=ir_id, ir_version=ir_version),
                        header="{patient_name}    DoB {DOB}    PRU {PRU}    NHS {NHSNumber}".format(**data),
                        SSH_config=r"\\gstt.local\apps\Moka\Files\Software\100K\ssh_credentials.txt"
                        )
                except Exception as e:
                    print "ERROR: Encountered following error when downloading summary of findings for NGSTestID {ngs_test_id}: {error}".format(ngs_test_id=ngs_test_id, error=e)
                    continue
            # Set summary of findings text based on result code If result code is Negative (1) or Negative Negative (1189679668)
            if data['result_code'] in [1, 1189679668]:
                # 1 = Negative, 1189679668 = NegNeg
                data['summary_of_findings'] = (
                    'Whole genome sequencing has been completed by Genomics England and the primary analysis has not identified any underlying genetic cause of the clinical presentation.'
                )
            elif data['result_code'] in [1189679670]:
                # 1189679670 = Negative but with a previously reported variant i.e. No new findings
                data['summary_of_findings'] = (
                    'Whole genome sequencing has been completed by Genomics England; please see the genome interpretation section for details of previously reported variant(s).'
                )                
            else:
                # If result code not known, print error and skip to the next case
                print 'ERROR: Unknown result code for NGSTestID {ngs_test_id}.'.format(ngs_test_id=ngs_test_id)
                continue
            # Create GelReportGenerator object
            g = GelReportGenerator(path_to_wkhtmltopdf=r'\\gstt.local\shared\Genetics_Data2\Array\Software\wkhtmltopdf\bin\wkhtmltopdf.exe')            
            # Create the cover pdf
            g.create_cover_pdf(data, r'\\gstt.local\apps\Moka\Files\Software\100K\gel_cover_report_template.html')
            # Specify the path to the folder containing the technical reports downloaded from the interpretation portal
            gel_original_report_folder = r'\\gstt.local\shared\Genetics\Bioinformatics\GeL\technical_reports'
            # create a search pattern to identify the correct HTML report. Use single character wildcard as the verison of the report is not known
            gel_original_report_search_name = "ClinicalReport_{ir_id}-?.pdf".format(ir_id=data['IRID'])
            # Specify the output path for the combined report, based on the GeL participant ID and the interpretation request ID retrieved from Moka
            gel_combined_report = r'{gel_report_output_folder}\{pru}_{proband_id}_{ir_id}_{date}.pdf'.format(
                    gel_report_output_folder=gel_report_output_folder,
                    pru=data['PRU'].replace(':', '_'),
                    date=datetime.datetime.now().strftime(r'%y%m%d'),
                    proband_id=data['GELID'],
                    ir_id=data['IRID']
                    )
            # create an empty list to hold all the reports which match the search pattern
            list_of_html_reports = []
            # populate this list with the results of os.listdir which match the search term (created above). 
            list_of_html_reports = fnmatch.filter(os.listdir(gel_original_report_folder), gel_original_report_search_name)
            # if there is more than one report for this case
            if len(list_of_html_reports) > 1:
                # print error message
                print 'ERROR: Multiple ({file_count}) versions of the HTML report exist for IR-ID {ir_id}. Ensure only the correct version exists in {gel_original_report_folder}.'.format(file_count=len(list_of_html_reports), ir_id=data['IRID'], gel_original_report_folder=gel_original_report_folder)
            # if the original GeL report is not found, 
            elif len(list_of_html_reports) < 1:
                # print an error message
                print 'ERROR: Original GeL report not found for IR-ID {ir_id}. Please ensure it has been saved as PDF with the following filepath: {gel_original_report}'.format(gel_original_report=os.path.join(gel_original_report_folder, gel_original_report_search_name), ir_id=data['IRID'])
            else:
                # If only one report found create the name of the report using the file identified using the wildcard
                gel_original_report = os.path.join(gel_original_report_folder, list_of_html_reports[0])
                # Attach the GeL report to the cover page and output to the output path specified above.
                g.pdf_merge(gel_combined_report, g.cover_pdf, gel_original_report)
                # Store report filepath as an NGSTestFile in Moka
                ngstestfile_insert_sql = (
                    "INSERT INTO NGSTestFile (NGSTestID, Description, NGSTestFile, DateAdded) "
                    "VALUES ({ngs_test_id},  '100k Results', '{gel_combined_report}', '{today_date}');"
                    ).format(
                        ngs_test_id=ngs_test_id,
                        gel_combined_report=gel_combined_report,
                        today_date=datetime.datetime.now().strftime(r'%Y%m%d %H:%M:%S %p')
                        )
                moka.execute_query(ngstestfile_insert_sql)
                # Update the status for NGSTest
                ngstest_update_sql = (
                    "UPDATE n SET n.Check1ID = c.Check1ID, n.Check1Date = '{today_date}', n.StatusID = 1202218814 "
                    "FROM NGSTest AS n, Checker AS c WHERE c.UserName = '{username}' AND n.NGSTestID = {ngs_test_id};"
                    ).format(
                        today_date=datetime.datetime.now().strftime(r'%Y%m%d %H:%M:%S %p'), 
                        username=os.getenv('username'), 
                        ngs_test_id=ngs_test_id
                        )
                moka.execute_query(ngstest_update_sql)
                # Record in patient log
                patientlog_insert_sql = (
                    "INSERT INTO PatientLog (InternalPatientID, LogEntry, Date, Login, PCName) "
                    "VALUES ({internal_patient_id},  'NGS: 100k results letter generated for interpretation request {IRID}', '{today_date}', '{username}', '{computer}');"
                    ).format(
                        internal_patient_id=data['internal_patient_id'],
                        IRID=data['IRID'],
                        today_date=datetime.datetime.now().strftime(r'%Y%m%d %H:%M:%S %p'),
                        username=os.getenv('username'),
                        computer=os.getenv('computername')
                        )
                moka.execute_query(patientlog_insert_sql)				
                # Create email body
                email_subject = "100,000 Genomes Project Result"
                email_body = (
                    '<body style="font-family:Calibri,sans-serif;">'
                    '<b>100,000 Genomes Project result from the Genetics Laboratory at Viapath - Guy\'s Hospital</b><br><br>'
                    'PLEASE DO NOT REPLY TO THIS EMAIL ADDRESS WITH ENQUIRIES ABOUT REPORTS<br>'
                    'FOR ALL ENQUIRIES PLEASE CONTACT THE LABORATORY USING <a href="mailto:DNADutyScientist@viapath.co.uk">DNADutyScientist@viapath.co.uk</a><br><br>'
                    'Kind regards<br>'
                    'Genetics Laboratory<br>'
                    '5th Floor, Tower Wing<br>'
                    'Guy\'s Hospital<br>'
                    'London, SE1 9RT<br>'
                    'United Kingdom<br><br>'
                    'Tel: + 44 (0) 207 188 1709'
                    '</body>'
                    )
                # Populate an outlook email addressed to clinican with results attached 
                generate_email(data['clinician_report_email'], email_subject, email_body, gel_combined_report)
                # Insert charge to Geneworks
                g = GeLGeneworksCharge()
                g.get_test_details(data['PRU'])
                g.insert_charge()
            # Print output location of reports
            print '\nGenerated reports can be found in: {gel_report_output_folder}'.format(gel_report_output_folder=gel_report_output_folder)
        

if __name__ == '__main__':
    main()
