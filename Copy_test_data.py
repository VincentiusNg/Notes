# This script is to copy input data from folder ..\test_data\RDM and ..\test_data\RPS 
# to test environment directory.

import shutil, os

script_location = os.getcwd();
cur_dir = script_location;

while True:
    if ('trunk' in os.listdir(cur_dir)) == True:
        trunk_folder = cur_dir + '\\trunk'
        break
    else:
        old_dir = cur_dir
        cur_dir = os.path.dirname(cur_dir)
        if old_dir == cur_dir:
            correct_pos = False
            print('Wrong location! Please put this script into trunk folder on repository')
            input()

testdata_folder = trunk_folder + '\\test_data'
RDM_workspace = trunk_folder + '\\source'
RPS_workspace = trunk_folder + '\\source_rps'

for folderName, subfolders, filenames in os.walk(testdata_folder):
    for filename in filenames:
        if ('DMAA' in filename) == True:
            if ('input' in filename) == True and ('.bin' in filename) == True :
                shutil.copy(folderName + '\\' + filename, RPS_workspace + '\\env\\work_rrl_test\\input')
                print('Copy ' + filename + ' to RPS workspace')
            
        else:
            if (('_input_' in filename) and ('.bin' in filename)) == True:
                shutil.copy(folderName + '\\' + filename, RDM_workspace + '\\env\\work_rrl_test\\input')
                print('Copy ' + filename + ' to RDM workspace')
            elif (('_output_' in filename) and ('.bin' in filename)) == True:
                shutil.copy(folderName + '\\' + filename, RDM_workspace + '\\env\\work_rrl_test\\expected_output')
                print('Copy ' + filename + ' to RDM workspace')
