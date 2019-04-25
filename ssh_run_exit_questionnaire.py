#!/usr/bin/env python2
"""
Requirements:
    Python 2.7
    paramiko

usage: ssh_run_exit_questionnaire.py [-h] --ir_id IR_ID --ir_version
                                     IR_VERSION --user USER

Submits a negneg clinical report and exit questionnaire for given
interpretation request

optional arguments:
  -h, --help            show this help message and exit
  --ir_id IR_ID         Interpretation request ID
  --ir_version IR_VERSION
                        Interpretation request version
  --user USER           cip-api username
"""
import os
import sys
import argparse
import datetime
from ConfigParser import ConfigParser
import paramiko

# Read config file (must be called config.ini and stored in same directory as script)
config = ConfigParser()
config.read(os.path.join(os.path.dirname(os.path.realpath(__file__)), "config.ini"))

class ExitQuestionnaire_SSH():
    '''
    Call summary_findings.py on the Viapath GENAPP01 server via ssh and transfer the PDF.
    '''
    def __init__(self, ir_id, ir_version, user):
        self.ir_id = ir_id
        self.ir_version = ir_version
        self.user = user
        self.ssh_host = config.get("GENAPP01", "SERVER")
        self.ssh_user = config.get("GENAPP01", "USER")
        self.ssh_pwd = config.get("GENAPP01", "PASSWORD")
        self.submit_exit_questionnaire()
    
    def submit_exit_questionnaire(self):
        """Call summary_findings.py on the server with input details.
        """
        # Set up paramiko SSH client
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(self.ssh_host, username=self.ssh_user, password=self.ssh_pwd)
        command = "/home/mokaguys/miniconda2/envs/jellypy_py3/bin/python /home/mokaguys/Apps/100K_exit_questionnaire/exit_questionnaire.py -i {ir_id}-{ir_version} -r {user} -d {date}".format(
                ir_id=self.ir_id,
                ir_version=self.ir_version,
                user=self.user,
                date=datetime.datetime.now().strftime(r'%Y-%m-%d')
            )
        # Execute command to submit clinical report and exit questionnaire on the server
        stdin, stdout, stderr = client.exec_command(command)
        # Call .read() on paramiko StderrFile object 
        # (Just printing the stderr instead of calling .read() was causing this issue:
        # https://stackoverflow.com/questions/20951690/paramiko-ssh-exec-command-to-collect-output-says-open-channel-in-response)        
        stderr = stderr.read()
        client.close()
        # If an error was encountered, print error message and exit
        if stderr:
            sys.exit(stderr)
    
def main():
    # Define and capture arguments.
    parser = argparse.ArgumentParser(description='Submits a negneg clinical report and exit questionnaire for given interpretation request')
    parser.add_argument('--ir_id', required=True, help='Interpretation request ID')
    parser.add_argument('--ir_version', required=True, help='Interpretation request version')
    parser.add_argument('--user', required=True, help='cip-api username')
    parsed_args = parser.parse_args()
    s = ExitQuestionnaire_SSH(
        ir_id=parsed_args.ir_id,
        ir_version=parsed_args.ir_version,
        user=parsed_args.user
        )

if __name__ == '__main__':
    main()
