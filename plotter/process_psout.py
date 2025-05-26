# -*- coding: utf-8 -*-
'''
A set of functions to process Manitoba Hydro International (MHI) PSOUT files.
This library provides functions to read and process PSOUT files.
It uses the signal names defined in a figure setup CSV file to extract relevant signals from the PSOUT files.
It is designed to be used in conjunction with the Manitoba Hydro International (MHI) PSOUT File Reader Library.
'''

import sys
import numpy as np
import pandas as pd
try:
    import mhi.psout
except ImportError:
    print("Could not import mhi.psout. Make sure Manitoba Hydro International (MHI) PSOUT File Reader Library is installed and available in your Python environment.")
    sys.exit(1)

def getAllSignalnames(figureSetupPath):
    '''
    Get all the EMT signal names required from figureSetup.csv for the figures in the HTML and PNG plotter output
    '''
    
    figureSetupDF = pd.read_csv(figureSetupPath, sep=';')
    
    signamecase=[]
    for index, row in figureSetupDF.iterrows():
        if pd.isnull(row['include_in_case']):
            if not pd.isnull(row['emt_signal_1']): signamecase.append([row['emt_signal_1'], np.nan])
            if not pd.isnull(row['emt_signal_2']): signamecase.append([row['emt_signal_2'], np.nan])
            if not pd.isnull(row['emt_signal_3']): signamecase.append([row['emt_signal_3'], np.nan])
        else:
            if not pd.isnull(row['emt_signal_1']): signamecase.append([row['emt_signal_1'], row['include_in_case']])
            if not pd.isnull(row['emt_signal_2']): signamecase.append([row['emt_signal_2'], row['include_in_case']])
            if not pd.isnull(row['emt_signal_3']): signamecase.append([row['emt_signal_3'], row['include_in_case']])
            
    return pd.DataFrame(signamecase, columns=['Signalname', 'Case'])


def getCaseSignalnames(emtSignalnamesDF, case):
    '''
    Get a list of required signalnames for the specific case
    '''
    signalnames = list()
    for index, row in emtSignalnamesDF.iterrows():
        if pd.isnull(row['Case']):
            signalnames.append(row['Signalname'])
        else:
            cases = [int(val) for val in row['Case'].split(',')]    # The Cases are stored as comma separated string, and needs to be converted to a list
            if case in cases:
                signalnames.append(row['Signalname'])
    return signalnames


def getSignal(psoutFilePath, signalname):
    ''' Get time and trace/signal from psout file '''
    data_path = 'Root/Main/MTB/'+ signalname +'/0'    
    data = psoutFilePath.call(data_path)
    run = psoutFilePath.run(0)
    signal = list()
    for call in data.calls():
        trace = run.trace(call)
        time = trace.domain.data
        signal.append(trace.data)
        
    return time, signal


def getSignals(psoutFilePath, signalnames):
    '''
    Get all signals from the .psout file who's name appear in signalnames
    '''
    signalnames_not_found = list()
    with mhi.psout.File(psoutFilePath) as psoutFile:
        t, _ = getSignal(psoutFile, signalnames[0])           # Get time values to get the length of all the signals in the .psout file
        t = np.array(t)                                           # Convert to a numpy array
        t = t.reshape(1,-1)                                       # Reshape the t from (N,) to (1,N)
        signals = np.array(t)                                     # Use as first row for the signals array
        for signalname in signalnames:
            try:                                                
                _, signal = getSignal(psoutFile, signalname)  # Try to get each signal in the signalnames list
                signal = np.array(signal)                         # Convert to numpy array
                signals = np.append(signals, signal, axis=0)      # And append to the signals array
            except:
                signalnames_not_found.append(signalname)          # Make a list of all the signal names that could not be found
            
    # Remove all the signal names that could not be found
    for signalname_not_found in signalnames_not_found:
        signalnames.remove(signalname_not_found)

    signalnames = ['time']+signalnames                            # Add the lable to be used for the time column
    
    return pd.DataFrame(np.transpose(signals), columns=signalnames)
