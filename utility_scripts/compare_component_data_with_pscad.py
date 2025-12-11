#!/usr/bin/env python3
from __future__ import annotations 
import mhi.pscad
import pandas as pd
import openpyxl
import re
from math import pi

#PSCAD Project Name
pscad_project_name = 'Solbakken'
    
#Read PowerFactory component data from the following Excel file
excel_path = r'E:\Users\<username>\Solbakken_PowerFactory_Component_Data.xlsx'

#SET_PARMS = True implies that all the parameters from PowerFactory will be taken as default and over write the PSCAD parameters
SET_PARAMS = False

#Extract floating point value and unit type from the returned parameter value
def parse_PSCAD_value(PSCAD_value):
    #print(PSCAD_value, type(PSCAD_value))
    if type(PSCAD_value) == float or type(PSCAD_value) == int:
        value = PSCAD_value
        unit = ''
    else: #if type(PSCAD_value) == str or type(PSCAD_value) == 'mhi.pscad.unit.Value':
        regex = r'([+-]?\d*\.?\d+(?:[eE][+-]?\d+)?)\s*\[(.*)\]'
        output = re.search(regex, str(PSCAD_value)) #force str for 'mhi.pscad.unit.Value'
        try:
            value = float(output.group(1))
            unit = output.group(2)
        except:
            value = float(PSCAD_value)
            unit = ''
    return (value, unit)

#Test if there are any 3-Winding Transformers in the Excel Workbook (work in progress)
ELMTR3_EXISTS = True if 'PowerFactory ElmTr3 Data' in openpyxl.load_workbook(excel_path, read_only=True).sheetnames else False    

#Read in all existing Worksheets
powerfactory_cable_data_df = pd.read_excel(excel_path, 'PowerFactory Cable Data')
powerfactory_trfr2_data_df = pd.read_excel(excel_path, 'PowerFactory ElmTr2 Data')
if ELMTR3_EXISTS: powerfactory_trfr3_data_df = pd.read_excel(excel_path, 'PowerFactory ElmTr3 Data')

#Delete the 'index' which is equal to the first Unnamed column
powerfactory_cable_data_df.drop('Unnamed: 0', axis=1, inplace=True)
powerfactory_trfr2_data_df.drop('Unnamed: 0', axis=1, inplace=True)
if ELMTR3_EXISTS: powerfactory_trfr3_data_df.drop('Unnamed: 0', axis=1, inplace=True)

#Get the cable and transformer names
cable_names = powerfactory_cable_data_df['Cable name'].tolist()
trfr2_names = powerfactory_trfr2_data_df['Transformer name'].tolist()
if ELMTR3_EXISTS: trfr3_names = powerfactory_trfr3_data_df['Transformer name'].tolist()

with mhi.pscad.application() as pscad:
    #Get a list of cases
    #cases = pscad.cases()    
    #project = pscad.project(cases[0].name)
    project = pscad.project(pscad_project_name)
    project.focus()
    main = project.canvas('Main')

    #Get all the cable component
    cable_components = []
    for cable_name in cable_names:
        cable_components.append(project.find(cable_name))

    #Read all the parameters for each cable component
    cables_list = []
    for cable_component in cable_components:
        if cable_component is None: continue
        params_dict = cable_component.parameters()
        if len(params_dict) == 1: continue
        #print(params_dict) #DEBUG
        if SET_PARAMS:
            #Set the PSCAD parameters equal to the PowerFactory parameters
            i = cable_names.index(params_dict['Name'])
            params_dict['PU'] = 'R_XL_XC_OHM_'
            params_dict['len'] = powerfactory_cable_data_df.iloc[i]['length [km]']
            params_dict['Rp'] = powerfactory_cable_data_df.iloc[i]['Eq. R\' [Ohm/km]']
            params_dict['Xp'] = powerfactory_cable_data_df.iloc[i]['Eq. X\' [Ohm/km]']
            params_dict['Bp'] = powerfactory_cable_data_df.iloc[i]['Eq. Shunt X\' [MOhm*km]']
            params_dict['Rz'] = powerfactory_cable_data_df.iloc[i]['Eq. R0\' [Ohm/km]']
            params_dict['Xz'] = powerfactory_cable_data_df.iloc[i]['Eq. X0\' [Ohm/km]']
            params_dict['Bz'] = powerfactory_cable_data_df.iloc[i]['Eq. Shunt X0\' [MOhm*km]']
            params_dict.pop('Rp2') #Only used if 'PU': 'R_L_C_OHM_H_UF_'         
            params_dict.pop('Lp')  #Only used if 'PU': 'R_L_C_OHM_H_UF_'
            params_dict.pop('Cp')  #Only used if 'PU': 'R_L_C_OHM_H_UF_'
            params_dict.pop('Rz2') #Only used if 'PU': 'R_L_C_OHM_H_UF_'
            params_dict.pop('Lz')  #Only used if 'PU': 'R_L_C_OHM_H_UF_'
            params_dict.pop('Cz')  #Only used if 'PU': 'R_L_C_OHM_H_UF_'
            cable_component.parameters(parameters = params_dict)
            params_dict = cable_component.parameters() #Read parameter values again
            
        #Extract PSCAD parameters
        length = parse_PSCAD_value(params_dict['len'])[0]/1000 if parse_PSCAD_value(params_dict['len'])[1] == 'm' else parse_PSCAD_value(params_dict['len'])[0]
        #If impedance, admittance data where entered as: 'R_XL_XC_OHM_'
        if params_dict['PU'] == 'R_XL_XC_OHM_':
            if parse_PSCAD_value(params_dict['Rp'])[1] == 'ohm/m' or parse_PSCAD_value(params_dict['Rp'])[1] == '': #'ohm/m' default
                Rp = parse_PSCAD_value(params_dict['Rp'])[0]*1000
            elif parse_PSCAD_value(params_dict['Rp'])[1] == 'ohm/km':
                Rp = parse_PSCAD_value(params_dict['Rp'])[0]
            else:
                print(f"Unknown unit for positive sequence series resistance: {parse_PSCAD_value(params_dict['Rp'])[1]}")
                exit(0)
            if parse_PSCAD_value(params_dict['Xp'])[1] == 'ohm/m' or parse_PSCAD_value(params_dict['Xp'])[1] == '': #'ohm/m' default
                Xp = parse_PSCAD_value(params_dict['Xp'])[0]*1000
            elif parse_PSCAD_value(params_dict['Xp'])[1] == 'ohm/km':
                Xp = parse_PSCAD_value(params_dict['Xp'])[0]
            else:     
                print(f"Unknown unit for positive sequence series inductance: {parse_PSCAD_value(params_dict['Xp'])[1]}")
                exit(0)
            if parse_PSCAD_value(params_dict['Bp'])[1] == 'Mohm*m' or parse_PSCAD_value(params_dict['Bp'])[1] == '': #'Mohm*m' default
                Bp = parse_PSCAD_value(params_dict['Bp'])[0]/1000
            elif parse_PSCAD_value(params_dict['Bp'])[1] == 'Mohm*km':
                Bp = parse_PSCAD_value(params_dict['Bp'])[0]
            else:
                print(f"Unknown unit for positive sequence shunt capacitance: {parse_PSCAD_value(params_dict['Bp'])[1]}")
                exit(0)
            #If zero sequence data is entered, then extract the zero sequence parameters
            if params_dict['Estim'] == 'ENTER':
                if parse_PSCAD_value(params_dict['Rz'])[1] == 'ohm/m' or parse_PSCAD_value(params_dict['Rz'])[1] == '': #'ohm/m' default
                    Rz = parse_PSCAD_value(params_dict['Rz'])[0]*1000
                elif parse_PSCAD_value(params_dict['Rz'])[1] == 'ohm/km':
                    Rz = parse_PSCAD_value(params_dict['Rz'])[0]
                else:
                    print(f"Unknown unit for zero sequence series resistance: {parse_PSCAD_value(params_dict['Rz'])[1]}")
                    exit(0)
                if parse_PSCAD_value(params_dict['Xz'])[1] == 'ohm/m' or parse_PSCAD_value(params_dict['Xz'])[1] == '': #'ohm/m' default
                    Xz = parse_PSCAD_value(params_dict['Xz'])[0]*1000
                elif parse_PSCAD_value(params_dict['Xz'])[1] == 'ohm/km':
                    Xz = parse_PSCAD_value(params_dict['Xz'])[0]
                else:
                    print(f"Unknown unit for zero sequence series inductance: {parse_PSCAD_value(params_dict['Xz'])[1]}")
                    exit(0)
                if parse_PSCAD_value(params_dict['Bz'])[1] == 'Mohm*m' or parse_PSCAD_value(params_dict['Bz'])[1] == '': #'Mohm*m' default
                    Bz = parse_PSCAD_value(params_dict['Bz'])[0]/1000
                elif parse_PSCAD_value(params_dict['Bz'])[1] == 'Mohm*km':
                    Bz = parse_PSCAD_value(params_dict['Bz'])[0]
                else:
                    print(f"Unknown unit for zero sequence shunt capacitance: {parse_PSCAD_value(params_dict['Bz'])[1]}")
                    exit(0)
            else:
                Rz = Xz = Bz = 0
        #If impedance, admittance data where entered as: 'R_L_C_OHM_H_UF_'
        elif params_dict['PU'] == 'R_L_C_OHM_H_UF_':
            f = parse_PSCAD_value(params_dict['F'])[0]
            if parse_PSCAD_value(params_dict['Rp2'])[1] == 'ohm/m' or parse_PSCAD_value(params_dict['Rp2'])[1] == '': #'ohm/m' default
                Rp = parse_PSCAD_value(params_dict['Rp2'])[0]*1000
            elif parse_PSCAD_value(params_dict['Rp2'])[1] == 'ohm/km':
                Rp = parse_PSCAD_value(params_dict['Rp2'])[0]
            else:
                print(f"Unknown unit for positive sequence series resistance: {parse_PSCAD_value(params_dict['Rp2'])[1]}")
                exit(0)            
            if parse_PSCAD_value(params_dict['Lp'])[1] == 'mH/m' or parse_PSCAD_value(params_dict['Lp'])[1] == 'H/km':
                Xp = 2*pi*f*parse_PSCAD_value(params_dict['Lp'])[0]
            elif parse_PSCAD_value(params_dict['Lp'])[1] == 'mH/km':
                Xp = 2*pi*f*parse_PSCAD_value(params_dict['Lp'])[0]/1000
            elif parse_PSCAD_value(params_dict['Lp'])[1] == '': #'H/m' default
                Xp = 2*pi*f*parse_PSCAD_value(params_dict['Lp'])[0]*1000
            else:
                print(f"Unknown unit for positive sequence series inductance: {parse_PSCAD_value(params_dict['Lp'])[1]}")
                exit(0)
            if parse_PSCAD_value(params_dict['Cp'])[1] == 'uF/m' or parse_PSCAD_value(params_dict['Cp'])[1] == '': #'uF/m' default              
                Bp = 1/(2*pi*f*parse_PSCAD_value(params_dict['Cp']*1000)[0]) 
            elif parse_PSCAD_value(params_dict['Cp'])[1] == 'uF/km':
                Bp = 1/(2*pi*f*parse_PSCAD_value(params_dict['Cp'])[0])
            else:
                print(f"Unknown unit for positive sequence shunt capacitance: {parse_PSCAD_value(params_dict['Cp'])[1]}")
                exit(0)                
            #If zero sequence data is entered, then extract the zero sequence parameters               
            if params_dict['Estim'] == 'ENTER':
                if parse_PSCAD_value(params_dict['Rz2'])[1] == 'ohm/m' or parse_PSCAD_value(params_dict['Rz2'])[1] == '': #'ohm/m' default
                    Rz = parse_PSCAD_value(params_dict['Rz2'])[0]*1000 
                elif parse_PSCAD_value(params_dict['Rz2'])[1] == 'ohm/km':
                    Rz = parse_PSCAD_value(params_dict['Rz2'])[0]
                else:
                    print(f"Unknown unit for zero sequence series resistance: {parse_PSCAD_value(params_dict['Rz2'])[1]}")
                    exit(0)            
                if parse_PSCAD_value(params_dict['Lz'])[1] == 'mH/m':
                    Xz = 2*pi*f*parse_PSCAD_value(params_dict['Lz'])[0]
                elif parse_PSCAD_value(params_dict['Lz'])[1] == 'mH/km':
                    Xz = 2*pi*f*parse_PSCAD_value(params_dict['Lz'])[0]/1000
                elif parse_PSCAD_value(params_dict['Lz'])[1] == '': #'H/m' default
                    Xz = 2*pi*f*parse_PSCAD_value(params_dict['Lz'])[0]*1000
                else:
                    print(f"Unknown unit for zero sequence series inductance: {parse_PSCAD_value(params_dict['Lz'])[1]}")
                    exit(0)                
                if parse_PSCAD_value(params_dict['Cz'])[1] == 'uF/m' or parse_PSCAD_value(params_dict['Cz'])[1] == '':  #'uF/m' default
                    Bz = 1/(2*pi*f*parse_PSCAD_value(params_dict['Cz']*1000)[0])
                elif parse_PSCAD_value(params_dict['Cz'])[1] == 'uF/km':
                    Bz = 1/(2*pi*f*parse_PSCAD_value(params_dict['Cz'])[0])
                else:
                    print(f"Unknown unit for zero sequence shunt capacitance: {parse_PSCAD_value(params_dict['Cz'])[1]}")
                    exit(0)                                    
            else:
                Rz = Xz = Bz = 0
        else:
            print(f"Impedance and admittance data: {params_dict['PU']}, not yet supported")

        cable_row = [params_dict['Name'], length, Rp, Xp, Bp, Rz, Xz, Bz]

        cables_list.append(cable_row)

    #Write all the cable data to a DataFrame
    pscad_cables_data_df = pd.DataFrame(cables_list, columns = ['Cable name',
                                                                'length [km]',
                                                                'Eq. R\' [Ohm/km]',
                                                                'Eq. X\' [Ohm/km]',
                                                                'Eq. Shunt X\' [MOhm*km]',
                                                                'Eq. R0\' [Ohm/km]',
                                                                'Eq. X0\' [Ohm/km]',
                                                                'Eq. Shunt X0\' [MOhm*km]']) 
    #Get all the 2 winding transformer components
    trfr2_components = []
    for trfr2_name in trfr2_names:
        trfr2_components.append(project.find(trfr2_name))

    #Read all the parameters for each transformer component   
    trfrs2_list = []
    for trfr2_component in trfr2_components:
        if trfr2_component is None: continue
        params_dict = trfr2_component.parameters()
        #print(params_dict) #DEBUG
        
        #Check the transformer component model definition
        component_def = trfr2_component.defn_name[1]
        #print(component_def) #DEBUG
        
        #Duality based 3 phase 2 winding transformer
        if component_def == 'db_xfmr_3p2w':
            if SET_PARAMS:
                #Set the PSCAD parameters equal to the PowerFactory parameters
                i = trfr2_names.index(params_dict['Name'])
                params_dict['Tmva'] = powerfactory_trfr2_data_df.iloc[i]['Sn [MVA]']
                params_dict['f_'] = 50.0
                params_dict['V1LL'] = powerfactory_trfr2_data_df.iloc[i]['Un_HV [kV]']
                params_dict['V2LL'] = powerfactory_trfr2_data_df.iloc[i]['Un_LV [kV]']
                params_dict['TCuL_'] = powerfactory_trfr2_data_df.iloc[i]['P_Cu [pu]']
                params_dict['CoreEddyLoss_'] = powerfactory_trfr2_data_df.iloc[i]['P_NL [pu]']
                params_dict['Xl_'] = powerfactory_trfr2_data_df.iloc[i]['X1_leak [pu]']
                params_dict['Iexc_'] = powerfactory_trfr2_data_df.iloc[i]['I1_mag [%]']
                params_dict.pop('kVPerTurn_') #Remove this non-writable parameter
                trfr2_component.parameters(parameters = params_dict)
                params_dict = trfr2_component.parameters() #Read parameter values again

            #Extract PSCAD parameters
            W1 = 'Y' if params_dict['W1'] == 'Y' else 'D'
            W2 = 'y' if params_dict['W2'] == 'Y' else 'd'
            vec_group = W1+W2
            if parse_PSCAD_value(params_dict['CoreEddyLoss_'])[1] == 'pu' or parse_PSCAD_value(params_dict['CoreEddyLoss_'])[1] == '':  #'pu' default              
                CoreEddyLoss = parse_PSCAD_value(params_dict['CoreEddyLoss_'])[0]
            elif parse_PSCAD_value(params_dict['CoreEddyLoss_'])[1] == '%':
                CoreEddyLoss = parse_PSCAD_value(params_dict['CoreEddyLoss_'])[0]/100
            else:
                print(f"Unknown unit for Eddy Current Core Losses: {parse_PSCAD_value(params_dict['CoreEddyLoss_'])[1]}")
                exit(0)                                                
            trfr_row = [params_dict['Name'],
                        params_dict[vec_group],
                        parse_PSCAD_value(params_dict['Tmva'])[0],
                        parse_PSCAD_value(params_dict['V1LL'])[0],
                        parse_PSCAD_value(params_dict['V2LL'])[0],
                        parse_PSCAD_value(params_dict['TCuL_'])[0],
                        CoreEddyLoss ,
                        parse_PSCAD_value(params_dict['Xl_'])[0],
                        parse_PSCAD_value(params_dict['Iexc_'])[0],
                        'YES']
        
        #3 Phase 2 Winding Transformer
        elif component_def == 'xfmr-3p2w':
            if SET_PARAMS:
                #Set the PSCAD parameters equal to the PowerFactory parameters
                i = trfr2_names.index(params_dict['Name'])
                params_dict['Tmva'] = powerfactory_trfr2_data_df.iloc[i]['Sn [MVA]']
                params_dict['f'] = 50.0
                params_dict['V1'] = powerfactory_trfr2_data_df.iloc[i]['Un_HV [kV]']
                params_dict['V2'] = powerfactory_trfr2_data_df.iloc[i]['Un_LV [kV]']
                params_dict['CuL'] = powerfactory_trfr2_data_df.iloc[i]['P_Cu [pu]']
                params_dict['NLL'] = powerfactory_trfr2_data_df.iloc[i]['P_NL [pu]']
                params_dict['Xl'] = powerfactory_trfr2_data_df.iloc[i]['X1_leak [pu]']
                params_dict['Im1'] = powerfactory_trfr2_data_df.iloc[i]['I1_mag [%]']
                params_dict['Enab'] = 'YES' #Enable Saturation
                trfr2_component.parameters(parameters = params_dict)
                params_dict = trfr2_component.parameters() #Read parameter values again

            #Extract PSCAD parameters
            W1 = 'Y' if params_dict['YD1'] == 'Y' else 'D'
            W2 = 'y' if params_dict['YD2'] == 'Y' else 'd'
            D_lead_lag = params_dict['Lead']
            if D_lead_lag == 'LAGS':
                if W1 == 'D' and W2 == 'y':
                    hour = '11'
                elif W1 == 'Y' and W2 == 'd':
                    hour = '1'
                else:
                    hour = '0'
            else:
                if W1 == 'D' and W2 == 'y':
                    hour = '1'
                elif W1 == 'Y' and W2 == 'd':
                    hour = '11'
                else:
                    hour = '0'
            if parse_PSCAD_value(params_dict['NLL'])[1] == 'pu' or parse_PSCAD_value(params_dict['NLL'])[1] == '':  #'pu' default              
                NLL = parse_PSCAD_value(params_dict['NLL'])[0]
            elif parse_PSCAD_value(params_dict['NLL'])[1] == '%':
                NLL = parse_PSCAD_value(params_dict['NLL'])[0]/100
            else:
                print(f"Unknown unit for No Load Losses (NLL): {parse_PSCAD_value(params_dict['NLL'])[1]}")
                exit(0)                                                                    
            trfr_row = [params_dict['Name'],
                        W1+W2+hour,
                        parse_PSCAD_value(params_dict['Tmva'])[0],
                        parse_PSCAD_value(params_dict['V1'])[0],
                        parse_PSCAD_value(params_dict['V2'])[0],
                        parse_PSCAD_value(params_dict['CuL'])[0],
                        NLL,
                        parse_PSCAD_value(params_dict['Xl'])[0],
                        parse_PSCAD_value(params_dict['Im1'])[0],
                        params_dict['Enab']]

        else:
            print(f'2 Winding transformer model {component_def} not defined!')
            exit(0)
               
        trfrs2_list.append(trfr_row)
        
    pscad_trfr2_data_df = pd.DataFrame(trfrs2_list, columns = ['Transformer name',
                                                               'Vector Grouping',
                                                               'Sn [MVA]',
                                                               'Un_HV [kV]',
                                                               'Un_LV [kV]',
                                                               'P_Cu [pu]',
                                                               'P_NL [pu]',
                                                               'X1_leak [pu]',
                                                               'I1_mag [%]',
                                                               'Saturation Enabled'])

    if ELMTR3_EXISTS: 
        #Get all the 3 winding transformer components
        trfr3_components = []
        for trfr3_name in trfr3_names:
            trfr3_components.append(project.find(trfr3_name))
        
        
        #Read all the parameters for each transformer component   
        trfrs3_list = []
        for trfr3_component in trfr3_components:
            if trfr3_component is None: continue
            params_dict = trfr3_component.parameters()

            #Check the transformer component model definition
            component_def = trfr3_component.defn_name[1]
            
            #Duality based 3 phase 2 winding transformer
            #3 Phase 3 Winding Transformer
            if component_def == 'xfmr-3p3w2':
                if SET_PARAMS:
                    #Set the PSCAD parameters equal to the PowerFactory parameters
                    i = trfr3_names.index(params_dict['Name'])
                    params_dict['Tmva'] = powerfactory_trfr3_data_df.iloc[i]['Sn_HV [MVA]']
                    params_dict['f'] = 50.0
                    params_dict['V1'] = powerfactory_trfr3_data_df.iloc[i]['Un_HV [kV]']
                    params_dict['V2'] = powerfactory_trfr3_data_df.iloc[i]['Un_MV [kV]']
                    params_dict['V3'] = powerfactory_trfr3_data_df.iloc[i]['Un_LV [kV]']
                    params_dict['CuL12'] = powerfactory_trfr3_data_df.iloc[i]['P_Cu (HV-MV) [pu]']
                    params_dict['CuL23'] = powerfactory_trfr3_data_df.iloc[i]['P_Cu (MV-LV) [pu]']
                    params_dict['CuL13'] = powerfactory_trfr3_data_df.iloc[i]['P_Cu (LV-HV) [pu]']
                    params_dict['NLL'] = powerfactory_trfr3_data_df.iloc[i]['P_NL [pu]']
                    params_dict['Xl12'] = powerfactory_trfr3_data_df.iloc[i]['X1_leak (HV-MV) [pu]']
                    params_dict['Xl23'] = powerfactory_trfr3_data_df.iloc[i]['X1_leak (MV-LV) [pu]']
                    params_dict['Xl13'] = powerfactory_trfr3_data_df.iloc[i]['X1_leak (LV-HV) [pu]']
                    params_dict['Im1'] = powerfactory_trfr3_data_df.iloc[i]['I1_mag [%]']
                    params_dict['Enab'] = 'YES' #Enable Saturation
                    trfr3_component.parameters(parameters = params_dict)
                    params_dict = trfr3_component.parameters() #Read parameter values again

                #Extract PSCAD parameters
                W1 = 'Y' if params_dict['YD1'] == 'Y' else 'D'
                W2 = 'y' if params_dict['YD2'] == 'Y' else 'd'
                W3 = 'y' if params_dict['YD3'] == 'Y' else 'd'
                D_lead_lag = params_dict['Lead']
                if D_lead_lag == 'LAGS':
                    if W1 == 'D' and W2 == 'y' and W3 == 'y':
                        hour = '11'
                    elif W1 == 'Y' and W2 == 'd' and W3 == 'd':
                        hour = '1'
                    else:
                        hour = '0'
                else:
                    if W1 == 'D' and W2 == 'y' and W3 == 'y':
                        hour = '1'
                    elif W1 == 'Y' and W2 == 'd' and W3 == 'd':
                        hour = '11'
                    else:
                        hour = '0'
                    
                trfr_row = [params_dict['Name'],
                            W1+W2+hour+W3+hour,
                            params_dict['Tmva'],
                            params_dict['V1'],
                            params_dict['V2'],
                            params_dict['V3'],
                            params_dict['CuL12'],
                            params_dict['CuL23'],
                            params_dict['CuL13'],
                            params_dict['NLL'],
                            params_dict['Xl12'],
                            params_dict['Xl23'],
                            params_dict['Xl13'],
                            params_dict['Im1'],
                            params_dict['Enab']]            
            else:
                print('3 Winding transformer model {component_def} not defined!')
                exit(0)
                   
            trfrs3_list.append(trfr_row)
            
        pscad_trfr3_data_df = pd.DataFrame(trfrs3_list, columns = ['Transformer name',
                                                                   'Vector Grouping',
                                                                   'Sn_HV [MVA]',
                                                                   'Un_HV [kV]',
                                                                   'Un_MV [kV]',
                                                                   'Un_LV [kV]',
                                                                   'P_Cu (HV-MV) [pu]',
                                                                   'P_Cu (MV-LV) [pu]',
                                                                   'P_Cu (LV-HV) [pu]',
                                                                   'P_NL [pu]',
                                                                   'X1_leak (HV-MV) [pu]',
                                                                   'X1_leak (MV-LV) [pu]',
                                                                   'X1_leak (LV-HV) [pu]',
                                                                   'I1_mag [%]',
                                                                   'Saturation Enabled'])
    
#Combine the PowerFactory and PSCAD data into one DataFrame
cable_data_df = pd.concat([powerfactory_cable_data_df, pscad_cables_data_df], ignore_index = True)
trfr2_data_df = pd.concat([powerfactory_trfr2_data_df, pscad_trfr2_data_df], ignore_index = True)
if ELMTR3_EXISTS: trfr3_data_df = pd.concat([powerfactory_trfr3_data_df, pscad_trfr3_data_df], ignore_index = True)

#Compare the PowerFactory Cable data with the PSCAD Cable data
cables = [] 
for cable_name in  cable_names:
    cable = cable_data_df.index[cable_data_df['Cable name'] == cable_name].tolist()
    cables.append(cable)

perc_diff = []
for i, cable in enumerate(cables):
    if len(cable)>1: #check to see if there are two cables to compare
        name = cable_names[i]
        dLen = abs((cable_data_df.iloc[cable[0]]['length [km]']-cable_data_df.iloc[cable[1]]['length [km]']))/cable_data_df.iloc[cable[0]]['length [km]']
        dRp  = abs((cable_data_df.iloc[cable[0]]['Eq. R\' [Ohm/km]']-cable_data_df.iloc[cable[1]]['Eq. R\' [Ohm/km]']))/cable_data_df.iloc[cable[0]]['Eq. R\' [Ohm/km]']
        dXp  = abs((cable_data_df.iloc[cable[0]]['Eq. X\' [Ohm/km]']-cable_data_df.iloc[cable[1]]['Eq. X\' [Ohm/km]']))/cable_data_df.iloc[cable[0]]['Eq. X\' [Ohm/km]']
        dBp  = abs((cable_data_df.iloc[cable[0]]['Eq. Shunt X\' [MOhm*km]']-cable_data_df.iloc[cable[1]]['Eq. Shunt X\' [MOhm*km]']))/cable_data_df.iloc[cable[0]]['Eq. Shunt X\' [MOhm*km]']
        dRz  = abs((cable_data_df.iloc[cable[0]]['Eq. R0\' [Ohm/km]']-cable_data_df.iloc[cable[1]]['Eq. R0\' [Ohm/km]']))/cable_data_df.iloc[cable[0]]['Eq. R0\' [Ohm/km]']
        dXz  = abs((cable_data_df.iloc[cable[0]]['Eq. X0\' [Ohm/km]']-cable_data_df.iloc[cable[1]]['Eq. X0\' [Ohm/km]']))/cable_data_df.iloc[cable[0]]['Eq. X0\' [Ohm/km]']
        dBz  = abs((cable_data_df.iloc[cable[0]]['Eq. Shunt X0\' [MOhm*km]']-cable_data_df.iloc[cable[1]]['Eq. Shunt X0\' [MOhm*km]']))/cable_data_df.iloc[cable[0]]['Eq. Shunt X0\' [MOhm*km]']
        perc_diff.append([name, dLen, dRp, dXp, dBp, dRz, dXz, dBz])

cable_perc_diff_df = pd.DataFrame(perc_diff, columns = ['Cable name',
                                                        'length [km]',
                                                        'Eq. R\' [Ohm/km]',
                                                        'Eq. X\' [Ohm/km]',
                                                        'Eq. Shunt X\' [MOhm*km]',
                                                        'Eq. R0\' [Ohm/km]',
                                                        'Eq. X0\' [Ohm/km]',
                                                        'Eq. Shunt X0\' [MOhm*km]'])

cable_comparative_data_df = pd.concat([powerfactory_cable_data_df, pscad_cables_data_df, cable_perc_diff_df], keys=['PF', 'PSCAD', '% Diff'])

#Compare the PowerFactory ElmTr2 data with the PSCAD Trfr2 data
trfr2s = [] 
for trfr2_name in  trfr2_names:
    trfr2 = trfr2_data_df.index[trfr2_data_df['Transformer name'] == trfr2_name].tolist()
    trfr2s.append(trfr2)

perc_diff = []
for i, trfr2 in enumerate(trfr2s):
    if len(trfr2)>1: #check to see if there are two transformers to compare
        name = trfr2_names[i]
        dSn = abs((trfr2_data_df.iloc[trfr2[0]]['Sn [MVA]']-trfr2_data_df.iloc[trfr2[1]]['Sn [MVA]']))/trfr2_data_df.iloc[trfr2[0]]['Sn [MVA]']
        dUn_HV  = abs((trfr2_data_df.iloc[trfr2[0]]['Un_HV [kV]']-trfr2_data_df.iloc[trfr2[1]]['Un_HV [kV]']))/trfr2_data_df.iloc[trfr2[0]]['Un_HV [kV]']
        dUn_LV  = abs((trfr2_data_df.iloc[trfr2[0]]['Un_LV [kV]']-trfr2_data_df.iloc[trfr2[1]]['Un_LV [kV]']))/trfr2_data_df.iloc[trfr2[0]]['Un_LV [kV]']
        dP_Cu  = abs((trfr2_data_df.iloc[trfr2[0]]['P_Cu [pu]']-trfr2_data_df.iloc[trfr2[1]]['P_Cu [pu]']))/trfr2_data_df.iloc[trfr2[0]]['P_Cu [pu]']
        dP_NL  = abs((trfr2_data_df.iloc[trfr2[0]]['P_NL [pu]']-trfr2_data_df.iloc[trfr2[1]]['P_NL [pu]']))/trfr2_data_df.iloc[trfr2[0]]['P_NL [pu]']
        dX1_leak  = abs((trfr2_data_df.iloc[trfr2[0]]['X1_leak [pu]']-trfr2_data_df.iloc[trfr2[1]]['X1_leak [pu]']))/trfr2_data_df.iloc[trfr2[0]]['X1_leak [pu]']
        dI1_mag  = abs((trfr2_data_df.iloc[trfr2[0]]['I1_mag [%]']-trfr2_data_df.iloc[trfr2[1]]['I1_mag [%]']))/trfr2_data_df.iloc[trfr2[0]]['I1_mag [%]']
        perc_diff.append([name, dSn, dUn_HV, dUn_LV, dP_Cu, dP_NL, dX1_leak, dI1_mag])

trfr2_perc_diff_df = pd.DataFrame(perc_diff, columns = ['Transformer name',
                                                        'Sn [MVA]',
                                                        'Un_HV [kV]',
                                                        'Un_LV [kV]',
                                                        'P_Cu [pu]',
                                                        'P_NL [pu]',
                                                        'X1_leak [pu]',
                                                        'I1_mag [%]'])

trfr2_comparative_data_df = pd.concat([powerfactory_trfr2_data_df, pscad_trfr2_data_df, trfr2_perc_diff_df], keys=['PF', 'PSCAD', '% Diff'])

if ELMTR3_EXISTS: 
    #Compare the PowerFactory ElmTr2 data with the PSCAD trfr3 data
    trfr3s = [] 
    for trfr3_name in  trfr3_names:
        trfr3 = trfr3_data_df.index[trfr3_data_df['Transformer name'] == trfr3_name].tolist()
        trfr3s.append(trfr3)
    
    perc_diff = []
    for i, trfr3 in enumerate(trfr3s):
        if len(trfr3)>1: #check to see if there are two transformers to compare
            name = trfr3_names[i]
            dSn_HV = abs((trfr3_data_df.iloc[trfr3[0]]['Sn_HV [MVA]']-trfr3_data_df.iloc[trfr3[1]]['Sn_HV [MVA]']))/trfr3_data_df.iloc[trfr3[0]]['Sn_HV [MVA]']
            dUn_HV  = abs((trfr3_data_df.iloc[trfr3[0]]['Un_HV [kV]']-trfr3_data_df.iloc[trfr3[1]]['Un_HV [kV]']))/trfr3_data_df.iloc[trfr3[0]]['Un_HV [kV]']
            dUn_MV  = abs((trfr3_data_df.iloc[trfr3[0]]['Un_MV [kV]']-trfr3_data_df.iloc[trfr3[1]]['Un_MV [kV]']))/trfr3_data_df.iloc[trfr3[0]]['Un_MV [kV]']
            dUn_LV  = abs((trfr3_data_df.iloc[trfr3[0]]['Un_LV [kV]']-trfr3_data_df.iloc[trfr3[1]]['Un_LV [kV]']))/trfr3_data_df.iloc[trfr3[0]]['Un_LV [kV]']
            dP_Cu_HV_MV  = abs((trfr3_data_df.iloc[trfr3[0]]['P_Cu (HV-MV) [pu]']-trfr3_data_df.iloc[trfr3[1]]['P_Cu (HV-MV) [pu]']))/trfr3_data_df.iloc[trfr3[0]]['P_Cu (HV-MV) [pu]']
            dP_Cu_MV_LV  = abs((trfr3_data_df.iloc[trfr3[0]]['P_Cu (MV-LV) [pu]']-trfr3_data_df.iloc[trfr3[1]]['P_Cu (MV-LV) [pu]']))/trfr3_data_df.iloc[trfr3[0]]['P_Cu (MV-LV) [pu]']
            dP_Cu_LV_HV  = abs((trfr3_data_df.iloc[trfr3[0]]['P_Cu (LV-HV) [pu]']-trfr3_data_df.iloc[trfr3[1]]['P_Cu (LV-HV) [pu]']))/trfr3_data_df.iloc[trfr3[0]]['P_Cu (LV-HV) [pu]']
            dP_NL  = abs((trfr3_data_df.iloc[trfr3[0]]['P_NL [pu]']-trfr3_data_df.iloc[trfr3[1]]['P_NL [pu]']))/trfr3_data_df.iloc[trfr3[0]]['P_NL [pu]']
            dX1_leak_HV_MV  = abs((trfr3_data_df.iloc[trfr3[0]]['X1_leak (HV-MV) [pu]']-trfr3_data_df.iloc[trfr3[1]]['X1_leak (HV-MV) [pu]']))/trfr3_data_df.iloc[trfr3[0]]['X1_leak (HV-MV) [pu]']
            dX1_leak_MV_LV  = abs((trfr3_data_df.iloc[trfr3[0]]['X1_leak (MV-LV) [pu]']-trfr3_data_df.iloc[trfr3[1]]['X1_leak (MV-LV) [pu]']))/trfr3_data_df.iloc[trfr3[0]]['X1_leak (MV-LV) [pu]']
            dX1_leak_LV_HV  = abs((trfr3_data_df.iloc[trfr3[0]]['X1_leak (LV-HV) [pu]']-trfr3_data_df.iloc[trfr3[1]]['X1_leak (LV-HV) [pu]']))/trfr3_data_df.iloc[trfr3[0]]['X1_leak (LV-HV) [pu]']
            dI1_mag  = abs((trfr3_data_df.iloc[trfr3[0]]['I1_mag [%]']-trfr3_data_df.iloc[trfr3[1]]['I1_mag [%]']))/trfr3_data_df.iloc[trfr3[0]]['I1_mag [%]']
            perc_diff.append([name, dSn_HV, dUn_HV, dUn_MV, dUn_LV, dP_Cu_HV_MV, dP_Cu_MV_LV, dP_Cu_LV_HV, dP_NL, dX1_leak_HV_MV, dX1_leak_MV_LV, dX1_leak_LV_HV, dI1_mag])
    
    trfr3_perc_diff_df = pd.DataFrame(perc_diff, columns = ['Transformer name',
                                                            'Sn_HV [MVA]',
                                                            'Un_HV [kV]',
                                                            'Un_MV [kV]',
                                                            'Un_LV [kV]',
                                                            'P_Cu (HV-MV) [pu]',
                                                            'P_Cu (MV-LV) [pu]',
                                                            'P_Cu (LV-HV) [pu]',
                                                            'P_NL [pu]',
                                                            'X1_leak (HV-MV) [pu]',
                                                            'X1_leak (MV-LV) [pu]',
                                                            'X1_leak (LV-HV) [pu]',
                                                            'I1_mag [%]'])
    
    trfr3_comparative_data_df = pd.concat([powerfactory_trfr3_data_df, pscad_trfr3_data_df, trfr3_perc_diff_df], keys=['PF', 'PSCAD', '% Diff'])


#Add these two DataFrames a new Sheets to the PowerFactory Poject Data Excel file
with pd.ExcelWriter(excel_path, mode = "a", if_sheet_exists = 'replace', engine = "openpyxl") as project_data_writer:
  cable_comparative_data_df.to_excel(project_data_writer, sheet_name = 'Comparative Cable Data')
  trfr2_comparative_data_df.to_excel(project_data_writer, sheet_name = 'Comparative ElmTr2 Data')
  if ELMTR3_EXISTS: trfr3_comparative_data_df.to_excel(project_data_writer, sheet_name = 'Comparative ElmTr3 Data')

print(f'Output written to \'{excel_path}\'')






































