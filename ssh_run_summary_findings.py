#!/usr/bin/env python2
"""
Requirements:
    Python 2.7
    paramiko

usage: ssh_run_summary_findings.py [-h] --ir_id IR_ID --ir_version IR_VERSION
                                   -o OUTPUT_FILE [--header HEADER]

Downloads summary of findings for given interpretation request

optional arguments:
  -h, --help            show this help message and exit
  --ir_id IR_ID         Interpretation request ID
  --ir_version IR_VERSION
                        Interpretation request version
  -o OUTPUT_FILE, --output_file OUTPUT_FILE
                        Output PDF
  --header HEADER       Text for header of report
"""
import os
import sys
import argparse
from ConfigParser import ConfigParser
import paramiko

# Read config file
config = ConfigParser()
config.read(r"F:\Moka\Files\Software\100K\config.ini")

class SummaryFindings_SSH():
    '''
    Call summary_findings.py on the Viapath GENAPP01 server via ssh and transfer the PDF.
    '''
    def __init__(self, ir_id, ir_version, output_path, header, SSH_config):
        self.ir_id = ir_id
        self.ir_version = ir_version
        self.output_path_local = output_path
        self.output_path_server = "/home/mokaguys/Documents/100K_summary_findings_pdf/{}".format(os.path.basename(self.output_path_local))
        self.header = header
        self.ssh_host = config.get("GENAPP01", "SERVER")
        self.ssh_user = config.get("GENAPP01", "USER")
        self.ssh_pwd = config.get("GENAPP01", "PASSWORD")
        self.transferred_bytes = None
        self.total_bytes = None
        self.download_summary_findings()
        self.copy_summary_findings()
    
    def download_summary_findings(self):
        """Call summary_findings.py on the server with input details.
        """
        # Set up paramiko SSH client
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(self.ssh_host, username=self.ssh_user, password=self.ssh_pwd)
        command = "/home/mokaguys/miniconda2/envs/jellypy_py3/bin/python /home/mokaguys/Apps/100K_summary_findings_pdf/summary_findings.py --ir_id {ir_id} --ir_version {ir_version} -o {output_path}".format(
                ir_id=self.ir_id,
                ir_version=self.ir_version,
                output_path=self.output_path_server
            )
        if self.header:
            command += " --header '{header}'".format(header=self.header)
        # Execute command to download summary of findings on the server
        stdin, stdout, stderr = client.exec_command(command)
        # Call .read() on paramiko StderrFile object 
        # (Just printing the stderr instead of calling .read() was causing this issue:
        # https://stackoverflow.com/questions/20951690/paramiko-ssh-exec-command-to-collect-output-says-open-channel-in-response)        
        stderr = stderr.read()
        client.close()
        # If an error was encountered, print error message and exit
        if stderr:
            sys.exit(stderr)

    def check_sftp_progress(self, transferred, total):
        """
        Callback function for recording bytes transferred and total bytes 
        """
        self.transferred_bytes = transferred
        self.total_bytes = total

    def copy_summary_findings(self):
        """
        Copies summary of findings pdf from the server to local directory via SFTP
        """
        # Connect to server using IP from config file and port 22 
        transport = paramiko.Transport((self.ssh_host, 22))
        transport.connect(username=self.ssh_user, password=self.ssh_pwd)
        sftp = paramiko.SFTPClient.from_transport(transport)
        # Copy file from server. Use callback function to record transferred and total bytes.
        sftp.get(remotepath=self.output_path_server, localpath=self.output_path_local, callback=self.check_sftp_progress)
        sftp.close()
        transport.close()
        # Error if not all bytes have been transferred
        if self.transferred_bytes != self.total_bytes:
            sys.exit("Incomplete file transfer. {transferred} out of {total} bytes".format(
                    transferred=self.transferred_bytes,
                    total=self.total_bytes
                )
            )
    
def main():
    # Define and capture arguments.
    parser = argparse.ArgumentParser(description='Downloads summary of findings for given interpretation request')
    parser.add_argument('--ir_id', required=True, help='Interpretation request ID')
    parser.add_argument('--ir_version', required=True, help='Interpretation request version')
    parser.add_argument('-o', '--output_file', required=True, help='Output PDF')
    parser.add_argument('--header', required=False, help='Text for header of report')
    parsed_args = parser.parse_args()
    s = SummaryFindings_SSH(
        ir_id=parsed_args.ir_id,
        ir_version=parsed_args.ir_version,
        output_path=parsed_args.output_file,
        header=parsed_args.header,
        SSH_config=r"F:\Moka\Files\Software\100K_dev\ssh_credentials.txt"
        )

if __name__ == '__main__':
    main()
