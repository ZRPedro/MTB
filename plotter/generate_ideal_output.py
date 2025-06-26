from pylab import *
import numpy as np
import pandas as pd
import os, sys
from datetime import datetime
from process_psout import getSignals

def lfsm(Pref, f, DK=1, FSM=False, s_fsm=10, db=0):
    '''
    With Pref in pu, the function calculates the new value of P given the
    frequency f, for either DK1 or DK2
    Pref in [pu] -- for Power Park Modules, Pref is the actual Active Power 
    output and not the nominal power
    The LFSM droop, s in [%]
    If f > fn (50Hz),
        the output will be the LFSM-O value
    else if f < fn (50Hz),
        it will be the LFSM-U value
    If FSM is True,
        the FSM droop, "s_fsm" has to be specified 
        as well as the deadband, "db" if any
    '''
    Pn = 1.0 #pu
    fn = 50  #Hz
    if DK == 1:
        f1 = 50.2 if f > fn else 49.8
        s = 5
    elif DK == 2:
        f1 = 50.5 if f > fn else 49.5
        s = 4
    else:
        print('"DK" can either be "1" or "2"!')
        return 0
    
    if FSM:
        Pref = fsm(Pref, f, DK, s_fsm, db)
        
    if f > fn and f > f1 or f < fn and f < f1:
        Pnew = Pref-100/s*(f-f1)/fn*Pn
    else:
        Pnew = Pref
        
    if Pnew > 1.0:
        Pnew = 1.0
        
    return Pnew


def fsm(Pref, f, DK=1, s=10, db=0):
    '''
    With Pref in pu, the function calculates the new value of P given the
    frequency f, for the FSM droop, s, given
    Pref in [pu] -- for Power Park Modules, Pref is the actual Active Power
    output and not the nominal power
    The FSM droop, s in [%]
    The FSM deadband, db in [Hz]
    '''
    Pn = 1.0 #pu
    fn = 50  #Hz
    if DK == 1:
        fRU = 49.8
        fRO = 50.2
    elif DK == 2:
        fRU = 49.5
        fRO = 50.5
    else:
        print('"DK" can either be "1" or "2"!')
        return 0

    f = fRU if f < fRU else f
    f = fRO if f > fRO else f
    
    if f<fn:
        Pnew = Pref-100/s*(f-fn+db)/fn*Pn if f<fn-db else Pref
    else:
        Pnew = Pref-100/s*(f-fn-db)/fn*Pn if f>fn+db else Pref        
    
    return Pnew

def ffc(vpos_val, id_val, iq_val, DK=1 , DSO=False):
    '''
    This function calculates the fast fault currunt (FFC) contribution,
    id_ffc and iq_ffc, based on the positive sequence voltage magnitude,
    based on RfG (EU) 2016/631, 20.2 (b) NC 2025 (Version 3) for DK1 and DK2
    '''
    if DK == 1:
        vposFrtLimit = 0.85
    elif DK == 2 or DSO:
        vposFrtLimit = 0.9
    else:
        print(f"DK = {DK} is not a valid option!")
        sys.exit(1)
    
    Imax = 1/vposFrtLimit # Could actually be higher
    
    if vpos_val >= vposFrtLimit: # No FRT
        iq_ffc = iq_val
        id_ffc = id_val
    elif vpos_val < vposFrtLimit and vpos_val > 0.5:
        if DK==2 or DSO:
            iq_ffc = -1/0.4*vpos_val+2.25     
        else: # DK==1
            iq_ffc = -1/0.35*vpos_val+2.42857 
            
        id_ffc = np.sqrt(Imax**2-iq_ffc**2)
    else: # vpos_val <= 0.5
        iq_ffc = 1.0
        id_ffc = np.sqrt(Imax**2-iq_ffc**2)
        
    return (id_ffc, iq_ffc) 

def main():
    testcasesFile = '..\\testcases.xlsx'
    psoutFolder = '..\\export\\MTB_04062025164118'
    exportFolder = '..\\export'

    #Create the date, time stampped results folder
    resultsFolder = f'MTB_{datetime.now().strftime(r"%d%m%Y%H%M%S")}'
    os.makedirs(os.path.join(exportFolder, resultsFolder), exist_ok=True)       

    #Read the 'Settings' sheet (with one header row)
    SettingsDf = pd.read_excel(testcasesFile, sheet_name='Settings', header=0)
    SettingDict = dict(zip(SettingsDf['Name'],SettingsDf['Value']))
    
    #Define the project name and signal list
    projectName = SettingDict['Projectname']
    DK = 1 if SettingDict['Area']=='DK1' else 2    #DK area, either 1 or 2
    DSO = True if SettingDict['Un']<110 else False #DSO, either Energinet (TSO))
    s_fsm = SettingDict['FSM droop']               #FSM droop in [%]
    db = SettingDict['FSM deadband']               #FSM deadband in [Hz]
    
    #Read the 'RfG cases' sheet (with two header rows)
    RfgDf = pd.read_excel(testcasesFile, sheet_name='RfG cases', header=[0, 1])

    #Limit the DataFrame to the first 60 columns
    RfgDf = RfgDf.iloc[:, :60]

    #Get all the LFSM and LFSM+FSM cases   
    FrequencyEventsDf = RfgDf.loc[RfgDf['Case']['EMT'].eq(True) & RfgDf['Event 1']['type'].eq('Frequency')]
    
    #Define the signals to be extracted from the psout files
    signalList = ['MTB\\P_pu_PoC', 'MTB\\pll_f_hz']
        
    for _, row in FrequencyEventsDf.iterrows():
        Rank = row['Case']['Rank']
        Name = row['Case']['Name']
        Pref = row['Initial Settings']['P0']
        FSM = row['Initial Settings']['Pmode'] == 'LFSM+FSM'
        psoutFileName = f"{psoutFolder}\\{projectName}_{Rank}.psout"
        print(f"Rank: {Rank}, Name: {Name}, Pref: {Pref}, FSM: {FSM}", {psoutFileName})
        caseDf = getSignals(psoutFileName, signalList)
        caseDf.rename(columns=dict(zip(signalList, [val.split('\\')[-1] for val in signalList])), inplace=True)
        caseDf.P_pu_PoC = pd.Series(lfsm(Pref=Pref, f=val, DK=DK, FSM=FSM, s_fsm=s_fsm, db=db) for val in caseDf.pll_f_hz)        
        caseDf.to_csv(f"{exportFolder}\\{resultsFolder}\\{projectName}_{Rank}.csv", index=False, sep=';', decimal=',')
        
    #Get all the Fault and Voltage-support cases      
    FaultEventsDf = RfgDf.loc[RfgDf['Case']['EMT'].eq(True) & RfgDf['Case']['Name'].str.contains('LVFRT') | 
                              RfgDf['Case']['EMT'].eq(True) & RfgDf['Event 1']['type'].str.contains('fault') |
                              RfgDf['Case']['EMT'].eq(True) & RfgDf['Case']['Name'].str.contains('support')]
    
    #Define the signals to be extracted from the psout files
    signalList = ['MTB\\fft_pos_Vmag_pu', 'MTB\\fft_pos_Id_pu', 'MTB\\fft_pos_Iq_pu']
    
    for _, row in FaultEventsDf.iterrows():
        Rank = row['Case']['Rank']
        Name = row['Case']['Name']
        psoutFileName = f"{psoutFolder}\\{projectName}_{Rank}.psout"
        print(f"Rank: {Rank}, Name: {Name}", {psoutFileName})
        caseDf = getSignals(psoutFileName, signalList)
        caseDf.rename(columns=dict(zip(signalList, [val.split('\\')[-1] for val in signalList])), inplace=True)
        for _, row_ in caseDf.iterrows():
            row_.fft_pos_Id_pu, row_.fft_pos_Iq_pu = ffc(row_.fft_pos_Vmag_pu, row_.fft_pos_Id_pu, row_.fft_pos_Iq_pu, DK, DSO)        
            
        caseDf.to_csv(f"{exportFolder}\\{resultsFolder}\\{projectName}_{Rank}.csv", index=False, sep=';', decimal=',')
        
        
if __name__ == "__main__":
    main()