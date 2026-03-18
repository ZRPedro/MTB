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


def getPsoutSignal(psoutFilePath, signalPathName):
    '''
    Get time and trace/signal from the opened psout file
    This function should not be called directly
    '''
    data_path = 'Root\\Main\\'+ signalPathName +'\\0'       # figureSetup.csv uses '\\' to be consistend with PowerFactory (and MHI's Enerplot)
    data_path = data_path.replace('\\', '/')                # But mhi.psout want to use '/'
    
    with mhi.psout.File(psoutFilePath) as psoutFile:
        try:
            data = psoutFile.call(data_path)
        except:
            print(f'The signal data path, {data_path} could not be found in {psoutFilePath}!')
            sys.exit(1)
            
        run = psoutFile.run(0)
        signal = list()
        for call in data.calls():
            trace = run.trace(call)
            time = trace.domain.data
            signal.append(trace.data)

    return time, signal


def getPsoutSignals(psoutFilePath, signalPathNames):
    '''
    Get all signals from the .psout file whose names appear in the signalnames list
    '''    
    columnNames = signalPathNames.copy()                                # Column names to be used for the returned DataFrame
    idxOffset = 0                                                       # To keep track of the columnnames size as signals are converter to signal arrays
    
    t, _ = getPsoutSignal(psoutFilePath, signalPathNames[0])            # Get time values to get the length of all the signals in the .psout file
    t = np.array(t)                                                     # Convert to a numpy array
    t = t.reshape(1,-1)                                                 # Reshape the t from (N,) to (1,N)
    psoutSignals = np.array(t)                                          # Use as first row for the signals array
    
    for signalPathName in signalPathNames:
        _, psoutSignal = getPsoutSignal(psoutFilePath, signalPathName)  # Try to get each signal in the signalnames list
        psoutSignal = np.array(psoutSignal)                             # Convert to numpy array

        # Test for signal array and convert signal name to a set of signal names for each signal in the array
        signalRows = psoutSignal.shape[0]
        if  signalRows > 1:
            idx = signalPathName.index(signalPathName)
            columnNames.remove(signalPathName)                          # Remove the signal name
            for i in range(signalRows):                                 # And replace it with signal array names
                columnNames.insert(idxOffset+idx+i,f'{signalPathName}_{i+1}')
                
            idxOffset = idxOffset + signalRows-1                        # Add the offset due to the signal array, to the column name index
            
        psoutSignals = np.append(psoutSignals, psoutSignal, axis=0)     # Append to the signals array containing the time
        
    columnNames = ['time']+columnNames                                  # Add the lable to be used for the time column
    
    return pd.DataFrame(np.transpose(psoutSignals), columns=columnNames)
