# -*- coding: utf-8 -*-
"""
Created on Fri Dec 13 12:41:52 2024

@author: PRW
"""

import powerfactory
import pandas as pd
import re
from types import NoneType

app = powerfactory.GetApplication()
app.ClearOutputWindow()

script = app.GetCurrentScript()
excel_output_path = script.GetAttribute('excel_output_path')

def get_type(typ_id):
  regex = r'(.*)(\.Typ.*)(\<.*)'
  match = re.search(regex, re.split(r'\\', typ_id)[-1])
  return match.group(2)
  
application_dict = {0 : 'Main Protection',
                    1 : 'Backup Protection'}

outserv_dict = {0 : 'No',
                1 : 'Yes'}   

#General
idir_dict = {0 : 'None',
             1 : 'Forward',
             2 : 'Reverse'}

#TypToc
toc_atype_dict = {'3ph'  : 'Phase Current (3ph)',
                  '1ph'  : 'Phase Current (1ph)',
                  '3I0'  : 'Earth Current (3*I0)',
                  'S3I0' : 'Sensitive Earth Current (3*I0)',
                  'I0'   : 'Zero Sequence Current (I0)',
                  'I2'   : 'Negative Sequence Current (I2)',
                  '3I2'  : '3*Negative Sequence Current (3*I2)',
                  'phA'  : 'Phase A Current',
                  'phB'  : 'Phase B Current',
                  'phC'  : 'Phase C Current',
                  'th'   : 'Thermal image (3ph)',
                  'th1p' : 'Thermal image (1ph)',
                  'd3m'  : '3ph (other)',
                  'd1m'  : '1ph (other)'}
              
#TypChar
char_atype_dict = {'d3m'  : '3ph (other)',
                   'd1m'  : '1ph (other)',
                   'P'    : 'Active power (P)',
                   'Q'    : 'Reactive power (Q)',
                   'S'    : 'Apparent power (S)',
                   'V'    : 'Voltage',
                   'f'    : 'Frequency (f)',
                   'dfdt' : 'RoCoF (df/dt)',
                   'Vf'   : 'Volt-per-Hertz (V/Hz)'}
              
#TypUlim               
ulim_ifunc_dict = {0 : 'Undervoltage',
                   1 : 'Overvoltage',
                   2 : 'Phase Shift' }                  

#TypFrq
freq_ifunc_dict = {0 : 'Instantaneous',
                   1 : 'Gradient',
                   2 : 'Gradient Digital' }                  

slot_name_ignore_list = ['Voltage Transformer', 'Measurement', 'fMeasurement', 'Output Logic', 'Clock', 'Phase Measurement Device PLL-Type', 'Sample and Hold', 'Moving Average Filter', 'Filter']
slot_type_ignore_list = ['.TypVt', '.TypCt', '.TypMeasure', '.TypFmeas', '.TypLogic', '.TypLogdip']

oElmRelays = app.GetCalcRelevantObjects('*.ElmRelay')
oRelChar = app.GetCalcRelevantObjects('*.RelChar')
app.PrintInfo(oElmRelays)
app.PrintInfo(oRelChar)
relays = []
for oElmRelay in oElmRelays:
  relay =[oElmRelay.loc_name,
          application_dict[oElmRelay.application],
          outserv_dict[oElmRelay.outserv],     
          ]
  slots = oElmRelay.pdiselm
  for slot in slots:
    try:
      if type(slot.typ_id) is NoneType:
        if slot.loc_name not in slot_name_ignore_list:
          relay.append(slot.loc_name)
      else:
        slot_type = get_type(str(slot.typ_id))
        if slot_type not in slot_type_ignore_list and slot.loc_name not in slot_name_ignore_list:
          relay.append(slot.loc_name)
    except:
      if slot.loc_name not in slot_name_ignore_list:
        relay.append(slot.loc_name)
    
  relays.append(relay)

dfRelays = pd.DataFrame(relays)
dfRelays.rename(columns={0 : 'Relay',
                         1 : 'Application',
                         2 : 'Out of Service'}, inplace=True)

with pd.ExcelWriter(excel_output_path, mode = "w", engine = "openpyxl") as project_data_writer:
  dfRelays.to_excel(project_data_writer, sheet_name = 'Relay Data')

for oElmRelay in oElmRelays:
  relay_info = []
  for slot in oElmRelay.pdiselm:
    try:
      if type(slot.typ_id) is not NoneType:
        slot_type = get_type(str(slot.typ_id))
        if slot_type not in slot_type_ignore_list:
          if slot_type == '.TypChar': 
            iec_symb = slot.typ_id.sfiec
            ansi_symb = slot.typ_id.sfansi
            relay_type = char_atype_dict[slot.typ_id.atype]
            direct = idir_dict[slot.idir]
            charact = slot.pcharac.loc_name
            thresh = slot.Ipset
            time_set = slot.Tpset
          elif slot_type == '.TypToc':
            iec_symb = slot.typ_id.sfiec
            ansi_symb = slot.typ_id.sfansi
            relay_type = toc_atype_dict[slot.typ_id.atype]
            direct = idir_dict[slot.idir]
            charact = slot.pcharac.loc_name
            thresh = slot.Ipset
            time_set = slot.Tpset
          elif slot_type == '.TypIoc':
            iec_symb = slot.typ_id.sfiec
            ansi_symb = slot.typ_id.sfansi
            relay_type = toc_atype_dict[slot.typ_id.atype]
            direct = idir_dict[slot.idir]
            charact = 'Definite Time'
            thresh = slot.Ipset
            time_set = slot.Tset
          elif slot_type == '.TypUlim': # Over voltage protection
            iec_symb = slot.typ_id.sfiec
            ansi_symb = slot.typ_id.sfansi
            relay_type = ulim_ifunc_dict[slot.typ_id.ifunc]
            direct = 'None'
            charact = 'Definite Time'
            thresh = slot.Usetr # [sec.V per phase] or Uset [p.u.] or cUpset [pri.V per phase]
            time_set = slot.Tdel
          elif slot_type == '.TypFrq':
            relay_type = freq_ifunc_dict[slot.typ_id.itype]
            direct = 'None'
            charact = 'Definite Time'
            if relay_type == 'Instantaneous':
              iec_symb = 'f'
              ansi_symb = '81'
              thresh = slot.Fset
              if thresh >= 0:
                iec_symb = iec_symb + '>'
              else:
                iec_symb = iec_symb + '<'
              time_set = slot.Tdel
            else:
              iec_symb = 'df/dt'
              ansi_symb = '81R'
              thresh = slot.dFset
              if thresh >= 0:
                iec_symb = iec_symb + '>'
              else:
                iec_symb = iec_symb + '<'
              time_set = slot.Tdel
          else:
            app.PrintInfo('Slot Type not defined')
            exit(0) 
  
          slot_info = [slot.loc_name,
                       iec_symb,
                       ansi_symb,
                       outserv_dict[slot.outserv],
                       relay_type,
                       direct,
                       charact,
                       thresh,
                       time_set]
          relay_info.append(slot_info)

          app.PrintInfo(f'{slot}, Out of Service = {outserv_dict[slot.outserv]}, Relay Type = {relay_type}, Direction = {direct}, Characterist = {charact}, Threshold = {thresh:.2f}, Time Setting = {time_set:.2f}''')     
    except:
      continue
  dfRelay = pd.DataFrame(relay_info, columns = ['Name',
                                                'IEC Symbol',
                                                'ANSI Number',
                                                'Out of Service',
                                                'Relay Type',
                                                'Tripping Direction',
                                                'Characteristic',
                                                'Threshold Value',
                                                'Time Value'])   
                                                             
  #dfRelay.round({'Threshold Value':2, 'Time Value' :3})      
                     
  with pd.ExcelWriter(excel_output_path, mode = "a", if_sheet_exists = 'replace', engine = "openpyxl") as project_data_writer:
    dfRelay.to_excel(project_data_writer, sheet_name = oElmRelay.loc_name)
                    
app.PrintInfo(f'Output written to \'{excel_output_path}\'')
  