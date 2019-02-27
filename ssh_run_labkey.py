#!/usr/bin/env python2
"""
ssh_run_labkey.py

Call LabKey.py on the Viapath GENAPP server via ssh. Requires a config file with SSH credentials.

Usage:
    ssh_run_labkey.py -i participant_id -c config_file
"""
import argparse
import paramiko
import pprint

class LabKey_SSH():
    '''Call LabKey.py on the Viapath GENAPP01 server via ssh.

    Args:
        participant_id: A GEL participant ID
        SSH_config: The name of a file containing comma-separated SSH details (host,username,password)
    Attributes:
        name: Patient Name
        dob: Patient date of birth in the format "DAY/MONTH/YEAR"
        nhsid: Patient nhs id
    Methods:
        read_ssh_config(config_file): Reads SSH details from config file
        call_labkey_api(): Calls API on GENAPP using input details
    '''
    def __init__(self, participant_id, SSH_config):
        self.participant_id = participant_id
        self.ssh_host, self.ssh_user, self.ssh_pwd = self.read_ssh_config(SSH_config)
        self.raw_string = self.call_labkey_api()
        # Remove newlines, flanking quote characters and separate into a list of variables
        self.name, self.dob, self.nhsid = self.raw_string.rstrip('\n').strip("'").split(",")

    
    def read_ssh_config(self, config_file):
        """Read ssh details from config file.
        Returns:
            A list of ssh config details.
        """
        with open(config_file) as f:
            # Take the first non-commented (#) line from the config file and clean newline character
            config = [ line.rstrip('\n') for line in f.readlines() if not line.startswith('#')][0]
        return config.split(',')
    
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
            "/home/mokaguys/Documents/LabKey/jellypy_labkey/LabKey.py -i {}".format(
                self.participant_id
            )
        )
        if stdout:
            return stdout.read()
        else:
            pprint.pprint(stderr.read())
            raise IOError('Calling LabKey failed. See stderr trace.')
        
    def __str__(self):
        return ",".join([self.name, self.dob, self.nhsid])
    
def main():
    # Call LabKey script on GENAPP via SSH and print patient details to std_out
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--pid', required=True, help="A Genomics England participant ID")
    parser.add_argument('-c', '--config', required=True, help="Path to a config file with GENAPP ssh credentials. \
        Host, username and password are separated by commas on a single line.")
    parsed_args = parser.parse_args()

    # Get patient data and print
    patient_data = LabKey_SSH(parsed_args.pid, parsed_args.config)
    print(patient_data)

if __name__ == '__main__':
    main()