"""
v1.3 - AJ 2018/06/12 - use ReportEmail field from checker table.
v1.2 - AB 2018/04/04
Requirements:
	ODBC connection to Moka
	Python 2.7
	pyodbc
	pdfkit
	PyPDF2
	jinja2

usage: gel_cover_report.py [-h] -n NGSTestID [NGSTestID ...]

Creates cover page for GeL results and attaches to report provided by GeL

optional arguments:
  -h, --help            show this help message and exit
  -n NGSTestID [NGSTestID ...]
                        Moka NGSTestID from NGSTest table
"""
import sys
import os
import io
import argparse
import datetime
import fnmatch
import win32com.client as win32
import pyodbc
import pdfkit
from PyPDF2 import PdfFileMerger
from jinja2 import Environment, FileSystemLoader

def process_arguments():
	"""
	Uses argparse module to define and handle command line input arguments and help menu
	"""
	# Create ArgumentParser object. Description message will be displayed as part of help message if script is run with -h flag
	parser = argparse.ArgumentParser(description='Creates cover page for GeL results and attaches to report provided by GeL')
	# Define the arguments that will be taken. nargs='+' allows multiple NGSTestIDs from NGSTest table in Moka can be passed as arguments.
	parser.add_argument('-n', metavar='NGSTestID', required=True, type=int, nargs='+', help='Moka NGSTestID from NGSTest table')
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

class MokaQueryExecuter(object):
	def __init__(self):
		# establish pyodbc connection to Moka
		cnxn = pyodbc.connect('DRIVER={SQL Server}; SERVER=GSTTV-MOKA; DATABASE=mokadata;', autocommit=True)
		# return cursor to execute query
		self.cursor = cnxn.cursor()

	def execute_query(self, sql):
		"""
		Executes a supplied SQL query
		"""
		self.cursor.execute(sql)

	def get_demographics(self, ngs_test_id):
		"""
		Takes a Moka NSGTestID as input.
		Pulls out details from Moka needed to populate the cover page. 
		"""
		demographics_sql = (
			'SELECT NGSTest.NGSTestID, NGSTest.InternalPatientID, Checker.Name AS clinician_name, Checker.ReportEmail, Item_Address.Item AS clinician_address, '
			'"gwv-patientlinked".FirstName, "gwv-patientlinked".LastName, "gwv-patientlinked".DoB, "gwv-patientlinked".Gender, "gwv-patientlinked".NHSNo, '
			'"gwv-patientlinked".PatientTrustID, NGSTest.GELProbandID, NGSTest.IRID '
			'FROM (((NGSTest INNER JOIN Patients ON NGSTest.InternalPatientID = Patients.InternalPatientID) '
			'INNER JOIN "gwv-patientlinked" ON "gwv-patientlinked".PatientTrustID = Patients.PatientID) INNER JOIN Checker ON NGSTest.BookBy = Checker.Check1ID) '
			'INNER JOIN Item AS Item_Address ON Checker.Address = Item_Address.ItemID '
			'WHERE NGSTestID = {ngs_test_id};'
			).format(ngs_test_id=ngs_test_id)
		# Execute the query to get patient demographics
		row = self.cursor.execute(demographics_sql).fetchone()
		# If results have been returned from the query...
		if row:
			# Populate demographics dictionaries with values returned by query
			demographics = {
				'clinician': row.clinician_name,
				'clinician_email': row.ReportEmail,
				'clinician_address': row.clinician_address,
				'internal_patient_id': row.InternalPatientID,
				'patient_name': '{first_name} {last_name}'.format(first_name=row.FirstName, last_name=row.LastName),
				'sex': row.Gender,
				'DOB': row.DoB.strftime(r'%d/%m/%Y'), # Extract date from datetime field in format dd/mm/yyyy
				'NHSNumber': row.NHSNo,
				'PRU': row.PatientTrustID,
				'GELID': row.GELProbandID,
				'IRID': row.IRID,
				'date_reported': datetime.datetime.now().strftime(r'%d/%m/%Y') # Current date in format dd/mm/yyyy
			}
			# If None has been returned for gender (because there isn't one in geneworks) change value to 'Unknown'
			if not demographics['sex']: 
				demographics['sex'] = 'Unknown'
			return demographics

class GelReportGenerator(object):
	def __init__(self, path_to_wkhtmltopdf):
		# path to wkhtmltopdf executable used by pdfkit
		self.path_to_wkhtmltopdf = path_to_wkhtmltopdf
		# Attribute to hold the in-memory cover file
		self.cover_pdf = None

	def create_cover_pdf(self, demographics, template):
		"""
		Populate html template with demographics and store as pdf
		"""
		# specify the folder containing the html template for cover report 
		html_template_dir = Environment(loader=FileSystemLoader(os.path.dirname(template)))
		# specify which html template to use
		html_template = html_template_dir.get_template(os.path.basename(template))
		# populate the template with values from demographics dictionary
		cover_html = html_template.render(demographics)
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
		# Get demographics for cover page from Moka.
		demographics = moka.get_demographics(ngs_test_id)
		# If no demographics are returned, print an error message
		if not demographics:
			print 'ERROR: No results returned from Moka demographics query for NGSTestID {ngs_test_id}. Check there are records in all inner joined tables (eg clinician address in checker table)'.format(ngs_test_id=ngs_test_id)
		# Otherwise continue...
		else:
			# Create GelReportGenerator object
			g = GelReportGenerator(path_to_wkhtmltopdf=r'\\gstt.local\shared\Genetics_Data2\Array\Software\wkhtmltopdf\bin\wkhtmltopdf.exe')
			# Create the cover pdf
			g.create_cover_pdf(demographics, r'\\gstt.local\apps\Moka\Files\Software\100K\gel_cover_report_template.html')
			# Specify the path to the folder containing the technical reports downloaded from the interpretation portal
			gel_original_report_folder = r'\\gstt.local\shared\Genetics\Bioinformatics\GeL\technical_reports'
			# create a search pattern to identify the correct HTML report. Use single character wildcard as the verison of the report is not known
			gel_original_report_search_name = "ClinicalReport_{ir_id}-?.pdf".format(ir_id=demographics['IRID'])
			# Specify the output path for the combined report, based on the GeL participant ID and the interpretation request ID retrieved from Moka
			gel_combined_report = r'{gel_report_output_folder}\{pru}_{proband_id}_{ir_id}_{date}.pdf'.format(
					gel_report_output_folder=gel_report_output_folder,
					pru=demographics['PRU'].replace(':', '_'),
					date=datetime.datetime.now().strftime(r'%y%m%d'),
					proband_id=demographics['GELID'],
					ir_id=demographics['IRID']
					)
			# create an empty list to hold all the reports which match the search pattern
			list_of_html_reports = []
			# populate this list with the results of os.listdir which match the search term (created above). 
			list_of_html_reports = fnmatch.filter(os.listdir(gel_original_report_folder), gel_original_report_search_name)
			# if there is more than one report for this case
			if len(list_of_html_reports) > 1:
				# print error message
				print 'ERROR: Multiple ({file_count}) versions of the HTML report exist for IR-ID {ir_id}. Ensure only the correct version exists in {gel_original_report_folder}.'.format(file_count=len(list_of_html_reports), ir_id=demographics['IRID'], gel_original_report_folder=gel_original_report_folder)
			# if the original GeL report is not found, 
			elif len(list_of_html_reports) < 1:
				# print an error message
				print 'ERROR: Original GeL report not found for IR-ID {ir_id}. Please ensure it has been saved as PDF with the following filepath: {gel_original_report}'.format(gel_original_report=os.path.join(gel_original_report_folder, gel_original_report_search_name), ir_id=demographics['IRID'])
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
						internal_patient_id=demographics['internal_patient_id'],
						IRID=demographics['IRID'],
						today_date=datetime.datetime.now().strftime(r'%Y%m%d %H:%M:%S %p'),
						username=os.getenv('username'),
						computer=os.getenv('computername')
						)
				moka.execute_query(patientlog_insert_sql)				
				# Create email body
				email_subject = "100,000 Genomes Project Report for {PRU}".format(PRU=demographics['PRU'])
				email_body = (
					'<body style="font-family:Calibri,sans-serif;">'
					'Please find a 100,000 Genomes Project report attached.<br><br>'
					'Best wishes,<br>'
					'Wook'
					'</body>'
					)
				# Populate an outlook email addressed to clinican with results attached 
				generate_email(demographics['clinician_email'], email_subject, email_body, gel_combined_report)
	# Print output location of reports
	print '\nGenerated reports can be found in: {gel_report_output_folder}'.format(gel_report_output_folder=gel_report_output_folder)
		

if __name__ == '__main__':
	main()
