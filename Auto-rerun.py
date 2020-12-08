"""
2020.12.02 @Jane Chen
1. Find the excel and the latest sheet.
2. Catch (0,1) to (3,-) data.
3. Remove the prefix. ex: Scan\
4. Dispatch the Case from VM type
5. Generate the rerun_<VM type>.txt in the same path
"""
import numpy as np
import pandas as pd
# Please modify below element
file_path = ""
input_excel = ""
output_file_name_prefix = ""
case_type_filter = 'Scan'
catch_cols = [0, 2]
# Read the latest sheet of target file name
row_data = pd.read_excel(file_path+input_excel, sheet_name=-1, usecols=catch_cols)
array_data = row_data.to_numpy()
type_filter = len(str(case_type_filter))+1
vm_name = ['ADMIN', 'WIN', 'IP']
cases = [[]for j in range(3)]
case = []
# Deal cases name to target case array
for i in range(0, len(array_data)):
    # Fail content trans to VM type(ADMIN, WIN, IP)
    vm_type = (str(array_data[i][1]).split("-"))[0]
    # Case trans to one to one and remove the prefix
    if str(array_data[i][0])[5:] != "":
        case.append((str(array_data[i][0])[type_filter:]).strip(" ")+",")
    else:
        case.append(case[i-1])
    # Dispatch the case from vm_type
    if vm_type == vm_name[0]:
        cases[0].append(case[i])
    elif vm_type == vm_name[1]:
        cases[1].append(case[i])
    elif vm_type == vm_name[2]:
        cases[2].append(case[i])
    else:
        print("The case %s not match any type" % vm_type)
        pass
# Generate the _rerun.txt in the same path
for j in range(0, len(vm_name)):
    # If case type is Null will pass
    if not cases[j]:
        pass
    else:
        txt_file_name = output_file_name_prefix+vm_name[j]+"_rerun.txt"
        independent_cases = list(set(cases[j]))
        independent_cases[-1] = str(independent_cases[-1]).strip(",")
        np.savetxt(file_path+txt_file_name, independent_cases, fmt='%s', newline='\n')
        print(txt_file_name+" is saved...")
