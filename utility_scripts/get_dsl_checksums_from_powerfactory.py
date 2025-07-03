# -*- coding: utf-8 -*-
"""
Created on Fri Jan 31 12:44:25 2025

@author: PRW
"""

import powerfactory
import pandas as pd

app = powerfactory.GetApplication()
script = app.GetCurrentScript()
excel_output_path = script.GetAttribute('excel_output_path')

app.ClearOutputWindow()
app.PrintInfo('Get all encrypted DSL model types in the project, and write their checksums to an Excel Spreadsheet')
#app.PrintInfo(f'Input Parameter: {excel_output_path}')
app.PrintPlain('')

############
#BlkDef Data
############
oBlkDefs = app.GetCalcRelevantObjects('*.BlkDef')

#Display DSL object and Checksum
for oBlkDef in oBlkDefs:
  for line in oBlkDef.sAddEquat:
    if line == '001! Encrypted model; Editing not possible.':
      checkSum = oBlkDef.GetCheckSum()  # Dummy read to ensure the checksum is updated
      app.PrintInfo(f'{oBlkDef}, Checksum = {oBlkDef.cCheckSum}')

#Make a DataFrame with the DSL model type name and Checksum
dfChecksums = pd.DataFrame([[oBlkDef.loc_name,
                             str(oBlkDef.cCheckSum).strip('\'[]')] for oBlkDef in oBlkDefs for line in oBlkDef.sAddEquat if line == '001! Encrypted model; Editing not possible.'],
                             columns=['DSL model type name',
                                      'Checksum'])
app.PrintPlain('')
app.PrintPlain(dfChecksums)

#Write DataFrame to Excel
with pd.ExcelWriter(excel_output_path, mode = "w", engine = "openpyxl") as project_data_writer:
  dfChecksums.to_excel(project_data_writer, sheet_name = 'DSL Encryption Data')

app.PrintPlain('')
app.PrintInfo(f'Output written to \'{excel_output_path}\'')