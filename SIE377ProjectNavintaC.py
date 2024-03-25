import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
import xlsxwriter
import openpyxl

application_window=tk.Tk()

my_filetype=[('excel files', '*.xlsx')]

filename = filedialog.askopenfilename(parent=application_window,
                                    initialdir=os.getcwd(),
                                    title="Please select the IMF Inflation Data file:",
                                    filetypes=my_filetype)

def main():
    # Reading the excel file selected
    df1 = pd.read_excel(filename, sheet_name='Overview')
    df1 = df1.drop(columns=['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 3', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9'])
    df1 = df1.drop([0, 1, 2,3,4,10,11,12,13,14,15,16,17])
    df2 = pd.read_excel(filename, sheet_name='Data', usecols="B:G", skiprows=3)
    df3 = pd.read_excel(filename, sheet_name='Data', usecols="C:G", skiprows=3)
    
    excelWriter = pd.ExcelWriter(r"C:\\Final project\\processed_imf_inflation_data.xlsx", engine='xlsxwriter')
    workbook=excelWriter.book

    values1 = df1.loc[[6,7,8,9,18,19,20], 'Unnamed: 2'].values
    values2 = df1.loc[5, 'Unnamed: 4']

    df1.to_excel(excelWriter,sheet_name="Overview", index=False)
    df2.to_excel(excelWriter,sheet_name="Data",  index=False)
    
    worksheet3=workbook.add_worksheet("Statistics")
    worksheet2 = excelWriter.sheets["Data"]
    worksheet1 = excelWriter.sheets['Overview']

    for i in range(7):
        worksheet1.write(f'B{i+1}', values1[i])
    worksheet1.write('B8', values2)
    worksheet1.write('A9', ' ')

    worksheet1.set_column(0, df2.shape[1], 30)
    worksheet1.set_default_row(20)
    worksheet2.set_column(0, df2.shape[1], 30)
    worksheet2.set_default_row(20)
    worksheet3.set_column(0, df2.shape[1], 30)
    worksheet3.set_default_row(20)

    bold = workbook.add_format({'bold':True})

    worksheet1.write('A1', 'Source', bold)
    worksheet1.write('A2', 'Conducted by', bold)
    worksheet1.write('A3', 'Survey Period', bold)
    worksheet1.write('A4', 'Published By', bold)
    worksheet1.write('A5', 'Published Date', bold)
    worksheet1.write('A6', 'Publication Date', bold)
    worksheet1.write('A7', 'Original Source', bold)
    worksheet1.write('A8', 'Description', bold)

    average = df3.mean()
    minimum = df3.min()
    maximum = df3.max()

    data3_columns = ['Statistics by Country','United Kingdom', 'France','Germany','Italy','Spain']
    worksheet3.write_row('A1', data3_columns, bold)

    data3_rows=['Mean','Min','Max']
    worksheet3.write_column('A2', data3_rows, bold)

    worksheet3.write_row('B2', average)
    worksheet3.write_row('B3', minimum)
    worksheet3.write_row('B4', maximum)

    excelWriter.save()
    workbook.close()

    
    

if __name__=="__main__":
    main()  
