import pandas as pd
import numpy as np
from simpledbf import Dbf5
import os, time, sys
import shutil
import glob

path_input = os.path.dirname(os.path.realpath(__file__))        # path_input = 'C:\\Footways_and_cycleways\\'

# path_output = os.path.join(path_input, "excel")                 # path_output = 'C:\\Footways_and_cycleways\\Excel\\'

def dbf_to_excel(path=path_input):
    path_input = path
    path_output = os.path.join(path_input, "excel") 
    # delete output path
    if os.path.exists(path_output):
        shutil.rmtree(path_output, ignore_errors=True)
        time.sleep(1)
    # create output folder
    if not os.path.exists(path_output):
        os.makedirs(path_output)

    # Convert DBFs to CSVs
    for file in os.listdir(path_input):
        if file.endswith(".dbf"):
            dbf = Dbf5(os.path.join(path_input, file))                          # dbf = Dbf5(f"{path_input}/{file}")
            dbf.to_csv(os.path.join(path_output, f"{file.split('.')[0]}.csv"))  # dbf.to_csv(f"{path_output}/{file.split('.')[0]}.csv")

    # Optional - Read from CSVs to DFs and check headers
    headers_list = []
    writer = pd.ExcelWriter(os.path.join(path_output, "CSVs_combined.xlsx"))    # writer = pd.ExcelWriter(f"{path_output}\\CSVs_combined.xlsx")
    csv_files = glob.glob(path_output+'/*.csv')

    dfs=[]

    for file in csv_files:                                                     # for file in os.listdir(path_output):
        df = pd.read_csv(file, low_memory=False) # df = pd.read_csv(f"{path_output}\{file.split('.')[0]}.csv", low_memory=False)
        df = df.loc[:, (df != 0).any(axis=0)] # this gets rid of any columns with 0 only
        df.replace('*', np.nan, inplace=True)
        df = df.dropna(axis='columns', how='all')


        # df = df.reindex(sorted(df.columns), axis='columns') 
        df['ASSET_LEN'] = None # add ASSET_LEN to columns 
        df.columns = map(str.upper, df.columns) # capitalise all the columns name 
        columns_names = list(df.columns)
        df['ASSET_LEN'] = df['ENDCHAIN'] - df['STCHAIN']
        swap_col1 = columns_names.index('ENDCHAIN') + 1
        swap_col2 = columns_names.index('ASSET_LEN')
        df['FROM'] = os.path.basename(file).split(".")[0]
        columns_names[swap_col1], columns_names[swap_col2] = columns_names[swap_col2], columns_names[swap_col1]
        columns_names.insert(0, 'FROM') # put where data is from on the first column

        df = df[columns_names]

        dfs.append(df)


        # if 'EndChain' in columns_names:
        #     df['ASSET_LEN'] = df.EndChain - df.StChain
        #     swap_col1 = columns_names.index('EndChain') + 1
        #     swap_col2 = columns_names.index('ASSET_LEN')
        #     columns_names[swap_col1], columns_names[swap_col2] = columns_names[swap_col2], columns_names[swap_col1]
        #     df = df[columns_names]
        # else:
        #     df['ASSET_LEN'] = df.ENDCHAIN - df.STCHAIN
        #     swap_col1 = columns_names.index('ENDCHAIN') + 1
        #     swap_col2 = columns_names.index('ASSET_LEN')
        #     columns_names[swap_col1], columns_names[swap_col2] = columns_names[swap_col2], columns_names[swap_col1]
        #     df = df[columns_names]

        # df.to_excel(writer, sheet_name=os.path.basename(file).split(".")[0][:30], index=False)
    combined = pd.concat(dfs, ignore_index=True)
    combined.to_excel(writer, sheet_name='csv_combined', index=False)
    writer.save()
    csv_path = os.path.join(path_output, 'CSV_combined.csv')
    combined.to_csv(csv_path, index=False)


if __name__ == "__main__":
    dbf_to_excel()