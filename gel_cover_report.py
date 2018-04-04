"""
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
	# Define the arguments that will be taken. Multiple NGSTestIDs from NGSTest table in Moka can be passed as arguments.
	parser.add_argument('-n', metavar='NGSTestID', required=True, type=int, nargs='+', help='Moka NGSTestID from NGSTest table')
	# Return the arguments
	return parser.parse_args()

def get_moka_demographics(ngs_test_id):
	"""
	Takes a Moka NSGTestID as input.
	Pulls out details from Moka needed to populate the cover page. 
	"""
	# establish pyodbc connection to Moka
	cnxn = pyodbc.connect('DRIVER={SQL Server}; SERVER=GSTTV-MOKA; DATABASE=mokadata;')
	# cursor to execute query
	cursor = cnxn.cursor()
	# SQL to pull out required demographics from Moka for given NGS test ID
	demographics_sql = (
		'SELECT NGSTest.NGSTestID, Item_Title.Item AS clinician_title, Checker.Name AS clinician_name, Item_Address.Item AS clinician_address, '
		'"gwv-patientlinked".FirstName, "gwv-patientlinked".LastName, "gwv-patientlinked".DoB, gw_GenderTable.Gender, "gwv-patientlinked".NHSNo, '
		'"gwv-patientlinked".PatientTrustID, NGSTest.GELProbandID, NGSTest.IRID '
		'FROM (((((NGSTest INNER JOIN Patients ON NGSTest.InternalPatientID = Patients.InternalPatientID) '
		'INNER JOIN "gwv-patientlinked" ON "gwv-patientlinked".PatientTrustID = Patients.PatientID) INNER JOIN Checker ON NGSTest.BookBy = Checker.Check1ID) '
		'LEFT JOIN Item AS Item_Title ON Checker.Title = Item_Title.ItemID) INNER JOIN Item AS Item_Address ON Checker.Address = Item_Address.ItemID) '
		'LEFT JOIN gw_GenderTable ON "gwv-patientlinked".GenderID = gw_GenderTable.GenderID '
		'WHERE NGSTestID = {ngs_test_id};'
		).format(ngs_test_id=ngs_test_id)
	# Execute the query to get patient demographics
	row = cursor.execute(demographics_sql).fetchone()
	# If results have been returned from the query...
	if row:
		# Populate demographics dictionaries with values returned by query
		demographics = {
			'clinician': row.clinician_name,
			'clinician_address': row.clinician_address,
			'patient_name': '{first_name} {last_name}'.format(first_name=row.FirstName, last_name=row.LastName),
			'sex': row.Gender,
			'DOB': row.DoB.strftime(r'%d/%m/%Y'), # Extract date from datetime field in format dd/mm/yyyy
			'NHSNumber': row.NHSNo,
			'PRU': row.PatientTrustID,
			'GELID': row.GELProbandID,
			'IRID': row.IRID,
			'date_reported': datetime.datetime.now().strftime(r'%d/%m/%Y') # Current date in format dd/mm/yyyy
		}
		# If clinician has a title (e.g. Dr.), update the clinician name to include it
		if row.clinician_title:
			demographics['clinician'] = '{title} {name}'.format(title=row.clinician_title, name=demographics['clinician'])
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
	# Get command line arguments
	args = process_arguments()
	# Loop through each Moka NGStestID supplied as an argument
	for ngstestid in args.ngstestid:
		# Get demographics for cover page from Moka.
		demographics = get_moka_demographics(ngstestid)
		# If no demographics are returned, print an error message
		if not demographics:
			print 'ERROR: No results returned from Moka query for NGSTestID {ngs_test_id}. Check there are records in all inner joined tables (eg clinician stated, clinician address in checker table,),'.format(ngs_test_id=ngs_test_id)
		# Otherwise continue...
		else:
			# Create GelReportGenerator object
			g = GelReportGenerator(path_to_wkhtmltopdf=r'\\gstt.local\shared\Genetics_Data2\Array\Software\wkhtmltopdf\bin\wkhtmltopdf.exe')
			# Create the cover pdf
			g.create_cover_pdf(demographics, r'\\gstt.local\apps\Moka\Files\Software\100K\gel_cover_report_template.html')
			# Specify the path to the folder containing the technical reports downloaded from the interpretation portal
			gel_original_report_folder = r'\\gstt.local\shared\Genetics\Bioinformatics\GeL\technical_reports\\'
			# create a search pattern to identify the correct HTML report. Use single character wildcard as the verison of the report is not known
			gel_original_report_search_name = "ClinicalReport_{ir_id}-?.pdf".format(ir_id=demographics['IRID'])
			# Specify the output path for the combined report, based on the GeL participant ID and the interpretation request ID retrieved from Moka
			gel_combined_report = r'\\gstt.local\shared\Genetics\Bioinformatics\GeL\reports_to_send\{pru}_{proband_id}_{ir_id}_{date}.pdf'.format(
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
				# exit with a statement explaining why
				print 'ERROR: Multiple ({file_count}) versions of the HTML report exist for IR-ID {ir_id}. Ensure only the correct version exists in S:\Genetics\Bioinformatics\GeL\\technical_reports.'.format(file_count = len(list_of_html_reports), ir_id = demographics['IRID'])
			elif len(list_of_html_reports) < 1 :
				# If the original GeL report is not found, print an error message
				print 'ERROR: Original GeL report not found. Please ensure it has been saved as PDF with the following filepath: {gel_original_report}'.format(gel_original_report = os.path.join(gel_original_report_folder, gel_original_report_search_name))
			else:
				# if only one report found create the name of the report using the file identified using the wildcard
				gel_original_report = os.path.join(gel_original_report_folder, list_of_html_reports[0])
				# Attach the GeL report to the cover page and output to the output path specified above.
				g.pdf_merge(gel_combined_report, g.cover_pdf, gel_original_report)
				# Print output location of file
				print 'Report has been generated: {gel_combined_report}'.format(gel_combined_report=gel_combined_report)
		

if __name__ == '__main__':
	main()
