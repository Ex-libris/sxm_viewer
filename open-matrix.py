import access2thematrix
mtrx_data = access2thematrix.MtrxData()
data_file = r'C:\Users\chem-ptch0510\Downloads\matrix_sergio\20230201-232229_STM_AtomManipulation--145_1.Z_mtrx'
traces, message = mtrx_data.open(data_file)
print(message)
