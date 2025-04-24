# -*- coding: utf-8 -*-
"""
Created on Mon Dec 9 12:24:34 2024

@author: PRWenerginet
"""

import powerfactory
import pandas as pd
import re
from numpy import sqrt

def get_typtrx(typ_id):
  regex = r'(.*)(\.TypTr.\.*)'
  match = re.search(regex, re.split(r'\\', typ_id)[-1])
  return match.group(1)
  
app = powerfactory.GetApplication()
script = app.GetCurrentScript()
excel_output_path = script.GetAttribute('excel_output_path')

app.ClearOutputWindow()
app.PrintInfo('Get all cables, transformers and transformer types in the project, and write it to an Excel Spreadsheet')
app.PrintInfo(f'Input Parameter: {excel_output_path}')

############
#ElmLne Data
############
oElmLnes = app.GetCalcRelevantObjects('*.ElmLne')
  
dfCables = pd.DataFrame([[oElmLne.loc_name,
                          oElmLne.nlnum,
                          oElmLne.dline,
                          oElmLne.GetType().rline,
                          oElmLne.GetType().lline,
                          oElmLne.GetType().xline,
                          oElmLne.GetType().rline0,
                          oElmLne.GetType().lline0,
                          oElmLne.GetType().xline0,
                          oElmLne.GetType().cline,
                          oElmLne.GetType().bline,
                          oElmLne.GetType().cline0,
                          oElmLne.GetType().bline0] for oElmLne in oElmLnes if oElmLne.GetType().IsCable()],
                        columns=['Cable name',
                                 'par lines',
                                 'length [km]',
                                 'R\'[Ohm/km]',
                                 'L\'[mH/km]',
                                 'X\'[Ohm/km]',
                                 'R0\'[Ohm/km]',
                                 'L0\'[mH/km]',
                                 'X0\'[Ohm/km]',
                                 'C\'[uF/km]',
                                 'B\'[uS/km]',
                                 'C0\'[uF/km]',
                                 'B0\'[uS/km]'])
                                 
#app.PrintInfo(dfCables)

#Calculate PSCAD equivalent values
dfCables['Eq. R\' [Ohm/km]'] = dfCables['R\'[Ohm/km]']/dfCables['par lines']
dfCables['Eq. X\' [Ohm/km]'] = dfCables['X\'[Ohm/km]']/dfCables['par lines']
dfCables['Eq. Shunt X\' [MOhm*km]'] = dfCables['par lines']/dfCables['B\'[uS/km]']
dfCables['Eq. R0\' [Ohm/km]'] = dfCables['R0\'[Ohm/km]']/dfCables['par lines']
dfCables['Eq. X0\' [Ohm/km]'] = dfCables['X0\'[Ohm/km]']/dfCables['par lines']
dfCables['Eq. Shunt X0\' [MOhm*km]'] = dfCables['par lines']/dfCables['B0\'[uS/km]']

############
#ElmTr2 Data
############
oElmTr2s = app.GetCalcRelevantObjects('*.ElmTr2')
 
dfElmTr2s = pd.DataFrame([[oElmTr2.loc_name,
                           get_typtrx(str(oElmTr2.typ_id)),
                           oElmTr2.ntnum,
                           oElmTr2.typ_id.vecgrp,
                           oElmTr2.typ_id.strn,
                           oElmTr2.typ_id.utrn_h,
                           oElmTr2.typ_id.utrn_l,
                           oElmTr2.typ_id.uktr,
                           oElmTr2.typ_id.pcutr,
                           oElmTr2.typ_id.uk0tr,
                           oElmTr2.typ_id.curmg,
                           oElmTr2.typ_id.pfe] for oElmTr2 in oElmTr2s],
                         columns=['Transformer name',
                                  'Transformer type',
                                  'Number of par. trfrs.',
                                  'Vector Grouping',
                                  'Sn [MVA]',
                                  'Un_HV [kV]',
                                  'Un_LV [kV]',
                                  'uk [%]',
                                  'P_Cu [kW]',
                                  'uk0 [%]',
                                  'I_NL [%]',
                                  'P_NL [kW]']) 
  
#Calculate PSCAD equivalent values
dfElmTr2s['P_Cu [pu]'] = dfElmTr2s['P_Cu [kW]']/(dfElmTr2s['Sn [MVA]']*1000)
dfElmTr2s['P_NL [pu]'] = dfElmTr2s['P_NL [kW]']/(dfElmTr2s['Sn [MVA]']*1000)
dfElmTr2s['X1_leak [pu]'] = dfElmTr2s['uk [%]']/100
dfElmTr2s['I1_mag [%]'] = sqrt(dfElmTr2s['I_NL [%]']**2-(dfElmTr2s['P_NL [pu]']/100)**2) #Ignoring P_Cu at no load

############
#ElmTr3 Data
############
oElmTr3s = app.GetCalcRelevantObjects('*.ElmTr3')

ELMTR3_EXIST = True if len(oElmTr3s)>0 else False

if ELMTR3_EXIST:
  dfElmTr3s = pd.DataFrame([[oElmTr3.loc_name,
                             get_typtrx(str(oElmTr3.typ_id)),
                             oElmTr3.nt3nm,
                             f'{oElmTr3.typ_id.tr3cn_h}{oElmTr3.typ_id.nt3ag_h:.0f}{oElmTr3.typ_id.tr3cn_m}{oElmTr3.typ_id.nt3ag_m:.0f}{oElmTr3.typ_id.tr3cn_l}{oElmTr3.typ_id.nt3ag_l:.0f}',
                             oElmTr3.typ_id.strn3_h,
                             oElmTr3.typ_id.strn3_m,
                             oElmTr3.typ_id.strn3_l,
                             oElmTr3.typ_id.utrn3_h,
                             oElmTr3.typ_id.utrn3_m,
                             oElmTr3.typ_id.utrn3_l,
                             oElmTr3.typ_id.uktr3_h,
                             oElmTr3.typ_id.uktr3_m,
                             oElmTr3.typ_id.uktr3_l,
                             oElmTr3.typ_id.pcut3_h,
                             oElmTr3.typ_id.pcut3_m,
                             oElmTr3.typ_id.pcut3_l,
                             oElmTr3.typ_id.uk0hm,
                             oElmTr3.typ_id.uk0ml,
                             oElmTr3.typ_id.uk0hl,
                             oElmTr3.typ_id.curm3,
                             oElmTr3.typ_id.pfe] for oElmTr3 in oElmTr3s],
                           columns=['Transformer name',
                                    'Transformer type',
                                    'Number of par. trfrs.',
                                    'Vector Grouping',
                                    'Sn_HV [MVA]',
                                    'Sn_MV [MVA]',
                                    'Sn_LV [MVA]',
                                    'Un_HV [kV]',
                                    'Un_MV [kV]',
                                    'Un_LV [kV]',
                                    'uk (HV-MV) [%]',
                                    'uk (MV-LV) [%]',
                                    'uk (LV-HV) [%]',
                                    'P_Cu (HV-MV) [kW]',
                                    'P_Cu (MV-LV) [kW]',
                                    'P_Cu (LV-HV) [kW]',
                                    'uk0 (HV-MV) [%]',
                                    'uk0 (MV-LV) [%]',
                                    'uk0 (LV-HV) [%]',
                                    'I_NL [%]',
                                    'P_NL [kW]']) 
  
  #Calculate PSCAD equivalent values
  dfElmTr3s['P_Cu (HV-MV) [pu]'] = dfElmTr3s['P_Cu (HV-MV) [kW]']/(dfElmTr3s['Sn_HV [MVA]']*1000)
  dfElmTr3s['P_Cu (MV-LV) [pu]'] = dfElmTr3s['P_Cu (MV-LV) [kW]']/(dfElmTr3s['Sn_HV [MVA]']*1000)
  dfElmTr3s['P_Cu (LV-HV) [pu]'] = dfElmTr3s['P_Cu (LV-HV) [kW]']/(dfElmTr3s['Sn_HV [MVA]']*1000)
  dfElmTr3s['P_NL [pu]'] = dfElmTr3s['P_NL [kW]']/(dfElmTr3s['Sn_HV [MVA]']*1000)
  dfElmTr3s['X1_leak (HV-MV) [pu]'] = dfElmTr3s['uk (HV-MV) [%]']/100
  dfElmTr3s['X1_leak (MV-LV) [pu]'] = dfElmTr3s['uk (MV-LV) [%]']/100
  dfElmTr3s['X1_leak (LV-HV) [pu]'] = dfElmTr3s['uk (LV-HV) [%]']/100
  dfElmTr3s['I1_mag [%]'] = sqrt(dfElmTr3s['I_NL [%]']**2-(dfElmTr3s['P_NL [pu]']/100)**2) #Ignoring P_Cu at no load

#app.PrintInfo(dfCables)
#app.PrintInfo(dfElmTr2s)
#app.PrintInfo(dfElmTr3s)

#Write Each DataFrame to a separate Excel Sheet
with pd.ExcelWriter(excel_output_path, mode = "w", engine = "openpyxl") as project_data_writer:
  dfCables.to_excel(project_data_writer, sheet_name = 'PowerFactory Cable Data')
  dfElmTr2s.to_excel(project_data_writer, sheet_name = 'PowerFactory ElmTr2 Data')
  if ELMTR3_EXIST:
    dfElmTr3s.to_excel(project_data_writer, sheet_name = 'PowerFactory ElmTr3 Data')

app.PrintInfo(f'Output written to \'{excel_output_path}\'')
