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


def _getSignal(psoutFile, signalname):
    '''
    Get time and trace/signal from the opened psout file
    This function should not be called directly
    '''
    data_path = 'Root\\Main\\'+ signalname +'\\0'           # figureSetup.csv uses '\\' to be consistend with PowerFactory (and MHI's Enerplot)
    data_path = data_path.replace('\\', '/')                # But mhi.psout want to use '/'
    data = psoutFile.call(data_path)
    run = psoutFile.run(0)
    signal = list()
    for call in data.calls():
        trace = run.trace(call)
        time = trace.domain.data
        signal.append(trace.data)
        
    return time, signal


def getSignals(psoutFilePath, signalnames):
    '''
    Get all signals from the .psout file whose names appear in the signalnames list
    '''    
    columnnames = signalnames.copy()                            # Column names to be used for the returned DataFrame
    idx_add = 0                                                 # To keep track of the columnnames size as signals are converter to signal arrays
    with mhi.psout.File(psoutFilePath) as psoutFile:
        t, _ = _getSignal(psoutFile, signalnames[0])            # Get time values to get the length of all the signals in the .psout file
        t = np.array(t)                                         # Convert to a numpy array
        t = t.reshape(1,-1)                                     # Reshape the t from (N,) to (1,N)
        signals = np.array(t)                                   # Use as first row for the signals array
        for signalname in signalnames:
            _, signal = _getSignal(psoutFile, signalname)       # Try to get each signal in the signalnames list
            signal = np.array(signal)                           # Convert to numpy array

            # Test for signal array and convert signal name to a set of signal names for each signal in the array
            signal_rows = signal.shape[0]
            if  signal_rows > 1:
                idx = signalnames.index(signalname)
                columnnames.remove(signalname)                  # Remove the signal name
                for i in range(signal_rows):                    # And replace it with signal array names
                    columnnames.insert(idx_add+idx+i,f'{signalname}_{i+1}')
                    
                idx_add = idx_add + signal_rows-1               # Add the offset due to the signal array, to the column name index
                                
            signals = np.append(signals, signal, axis=0)        # Append to the signals array containing the time
            
    columnnames = ['time']+columnnames                          # Add the lable to be used for the time column
    
    return pd.DataFrame(np.transpose(signals), columns=columnnames)
