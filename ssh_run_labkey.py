#!/usr/bin/env python2
"""
ssh_run_labkey.py

Call LabKey.py on the Viapath GENAPP server via ssh. Requires a config file with SSH credentials.

Usage:
    ssh_run_labkey.py -i participant_id -c config_file
"""
import os
import argparse
from ConfigParser import ConfigParser
import paramiko

# Read config file (must be called config.ini and stored in same directory as script)
config = ConfigParser()
config.read(os.path.join(os.path.dirname(os.path.realpath(__file__)), "config.ini"))

class LabKey_SSH():
    '''Call LabKey.py on the Viapath GENAPP01 server via ssh.

    Args:
        participant_id: A GEL participant ID
    Attributes:
        name: Patient Name
        dob: Patient date of birth in the format "DAY/MONTH/YEAR"
        nhsid: Patient nhs id
    Methods:
        call_labkey_api(): Calls API on GENAPP using input details
    '''
    def __init__(self, participant_id):
        self.participant_id = participant_id
        self.ssh_host = config.get("GENAPP01", "SERVER")
        self.ssh_user = config.get("GENAPP01", "USER")
        self.ssh_pwd = config.get("GENAPP01", "PASSWORD")
        self.raw_string = self.call_labkey_api()
        # Remove newlines, flanking quote characters and separate into a list of variables
        self.name, self.dob, self.nhsid = self.raw_string.rstrip('\n').strip("'").split(",")
    
    def call_labkey_api(self):
        """Call LabKey.py on the server with input details.
        Returns:
            A string form the stdout of the LabKey script - contains patient details.
        """
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(self.ssh_host, username=self.ssh_user, password=self.ssh_pwd)
        # Send command
        stdin, stdout, stderr = client.exec_command(
            "/home/mokaguys/miniconda2/envs/jellypy_py3/bin/python /home/mokaguys/Apps/100k_check_labkey/LabKey.py -i {}".format(
                self.participant_id
            )
        )
        # If an error was encountered, print error message and exit
        stderr = stderr.read()
        stdout = stdout.read()
        client.close()
        if stderr:
            sys.exit(stderr)
        return stdout
        
    def __str__(self):
        return ",".join([self.name, self.dob, self.nhsid])
    
def main():
    # Call LabKey script on GENAPP via SSH and print patient details to std_out
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--pid', required=True, help="A Genomics England participant ID")
    parsed_args = parser.parse_args()

    # Get patient data and print
    patient_data = LabKey_SSH(parsed_args.pid)
    print(patient_data)

if __name__ == '__main__':
    main()
