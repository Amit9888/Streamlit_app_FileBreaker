import pandas as pd

import streamlit as st
import win32com.client
import time
from xlsxwriter.utility import xl_rowcol_to_cell
import pandas as pd
import logging
import os
from datetime import date
outlook = win32com.client.Dispatch("Excel.Application",pythoncom.CoInitialize())


st.subheader("Dataset")
data_file = st.file_uploader("Upload excel",type=["xlsx"])

if data_file is not None:
    df=pd.read_excel(data_file,sheet_name='Regular GS',skiprows=[0])
    df_2=pd.read_excel(data_file,sheet_name='mailing list')
    st.write(df_2)
    
    df_2=df_2[['to','CC','Emp Name','Emp Code']]
    column_list=df.columns
    filtered_dataFrame=pd.DataFrame()
    base_dataframe=pd.DataFrame()

    parent_dir = "C:/data_sharing_files"

    mapped_employeeId=df_2["Emp Code"].values
    Names_df_2=df_2['Emp Name'].values
    to_df_2=df_2['to'].values
    final_CC_df_2=df_2['CC'].values
    print(column_list,"the columnlist")

    
    if "Emp Code" in column_list:
        print("Goes here")
        unique_values=df["Emp Code"].unique()
        
        for single_value in unique_values:
            
        
            filtered_dataFrame=df[df["Emp Code"]==single_value]
            sales_manager_name=filtered_dataFrame["Emp Name"]
            final_name=sales_manager_name.iloc[0]
            final_name=final_name.lower()
        
            if not os.path.exists(parent_dir):
        

                os.makedirs(parent_dir)
                
            writer_orig = pd.ExcelWriter(f'{parent_dir}/Employee_Code_No-{single_value}_{final_name}_report.xlsx', engine='xlsxwriter')
                
            filtered_dataFrame.to_excel(writer_orig,sheet_name='report')
            writer_orig.save()
                
                
            writer = pd.ExcelWriter(f'{parent_dir}/Employee_Code_No-{single_value}_{final_name}_report.xlsx', engine='xlsxwriter')
            filtered_dataFrame.to_excel(writer, index=False, sheet_name='report')

                # Get access to the workbook and sheet
            workbook = writer.book
            worksheet = writer.sheets['report']
                # cell_format = workbook.add_format({'align':'center','border':1})
            border_center = workbook.add_format({'align': 'center','border':1})
                
                # Add a header format.
            header_format = workbook.add_format({
                                'bold': True,
                                'align':'center',
                                'fg_color':'#D7E4BC',
                                'border': 1})
                
                
            cols_count=0
            rows_count=0
                
                
        
                # Write the column headers with the defined format.
            column_valuees_array=filtered_dataFrame.columns.values
            for col_num, value in enumerate(column_valuees_array):
                worksheet.write(0, col_num , value, header_format)
                cols_count+=1
                    
        
            number_of_rows = len(filtered_dataFrame.index)
                
            
            
                
            def str_time_conversion(x):
                return x.strftime('%B %d, %Y')

                
            filtered_dataFrame['EMP DOJ']=filtered_dataFrame['EMP DOJ'].apply(str_time_conversion)
                
                    
                    
            for col_num, value in enumerate(filtered_dataFrame.iloc[0].values):
                percentage_symbol = workbook.add_format({'num_format': '0.00%','align': 'center','border':1})
                try:
                    if col_num in [49,48,47,46,45,44,34,35,36,37,33,32,31,44,19,20,16]:
                        worksheet.write(1, col_num,value,percentage_symbol)
                    else:
                        worksheet.write(1, col_num , value, border_center)

                except:                
                    pass
                    
            for i, col in enumerate(filtered_dataFrame.columns):
                column_len = max(filtered_dataFrame[col].astype(str).str.len().max(), len(col) + 2)
                worksheet.set_column(i, i, column_len)
                    
            worksheet.hide_gridlines(2)
                    
            
                
                # worksheet.set_column(0,cols_count, 30)
            writer.save()
                
            print("File break completed")
                
        else:
            print("this code is exceuted")
            logging.error('The Required column has not been found in the excel file Please check the file')
            
        
                
        
        counter=0
