import subprocess
#from subprocess import check_call
import time
import sys
import os

def start_script(script_name, pause):

    subprocess.check_call('Y:& cd Start Menu\Programs\The Boeing Company & ' +
                           script_name + '.appref-ms', shell=True)
    time.sleep(pause)


def start_script_local(script_name):

    work_path_folder = os.getcwd()
    #subprocess.check_call('cd ' + work_path_folder + ' & ' + script_name + '.appref-ms', shell=True)
    subprocess.check_call('cd ' + work_path_folder + ' & ' + script_name + '.exe', shell=True)
    #time.sleep(pause)
    
if __name__ == "__main__":

    start_script_local('GetPointCoordinates')
