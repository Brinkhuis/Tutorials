import sys
import pandas as pd
import numpy as np
import datetime

# save_xls function to save multiple sheets to one excel file
def save_xls(list_dfs, xls_path):
    writer = pd.ExcelWriter(xls_path)
    for sheet, df in list_dfs:
        df.to_excel(writer, '{}'.format(sheet), index=False)
    writer.save()

def main():
    # set number of sheets
    sheets = int(sys.argv[1])
 
    # generate data
    data = pd.DataFrame({'region': [region for region in ['North', 'East', 'South', 'West']] * sheets,
                         'Q1': np.random.randint(1000, 2000, 4 * sheets),
                         'Q2': np.random.randint(1000, 2000, 4 * sheets),
                         'Q3': np.random.randint(1000, 2000, 4 * sheets),
                         'Q4': np.random.randint(1000, 2000, 4 * sheets),
                         'year': [year for year in range(datetime.datetime.now().year - sheets,
                                                         datetime.datetime.now().year) for _ in range(4)]})

    # split datadrame in dataframes per sheets
    df_list = list()
    for year in data.year.unique():
        df_list.append((year, data.loc[data.year == year, 'region':'Q4']))

    # save dataframes to excel
    save_xls(df_list, sys.argv[2])

    # print message
    print('Excel file saved to disk with', sheets, 'sheets.')

# This is the standard boilerplate that calls the main() function
if __name__ == '__main__':
   main()