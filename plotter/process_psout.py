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


def findPsoutSignalPath(psoutFilePath: str, signalPathName: str, sourceType: str ='Module') -> str:
    """
    Recursively find the path relative to Root/Main/ of a signal in a .psout file.
    Returns empty string if the signal could not be found
    
    Parameters:
        psoutFilePath: The path to the .psout file
        signalPathName: Can be in the format 'MTB\\signal_name' or just 'signal_name' as the leading path will be stripped anyway
        sourceType: Should be a valid source type, e.g. 'Module', 'Graphic', 'PGB', etc.
    """
    
    # Strip any leading path prefix (e.g. 'MTB\\' or 'SomePath\\MTB\\') to get just the signal name
    signalName = signalPathName.split('\\')[-1]

    def search_node(node, current_path='', depth=0):
        if depth > 10:
            return None
        try:
            for call in node.calls():
                call_str = str(call)
                name = call_str.split("Name='")[1].split("'")[0]
                source = call_str.split("Source='")[1].split("'")[0]
                new_path = current_path + '\\' + name if current_path else name

                if source == sourceType: 
                    try:
                        child_path = 'Root/Main/' + new_path.replace('\\', '/')
                        child_node = psout_file.call(child_path)
                        for child_call in child_node.calls():
                            child_str = str(child_call)
                            child_name = child_str.split("Name='")[1].split("'")[0]
                            if child_name == signalName:
                                return new_path
                        result = search_node(child_node, new_path, depth + 1)
                        if result:
                            return result
                    except Exception:
                        pass
        except Exception:
            pass
        return None

    with mhi.psout.File(psoutFilePath) as psout_file:
        root = psout_file.call('Root/Main')
        path = search_node(root)
        return path if path else ''

       
def getPsoutSignal(psoutFilePath, signalPathName):
    '''
    Get time and trace/signal from the opened psout file
    '''
    data_path = 'Root\\Main\\'+ signalPathName +'\\0'       # figureSetup.csv uses '\\' to be consistend with PowerFactory (and MHI's Enerplot)
    data_path = data_path.replace('\\', '/')                # But mhi.psout want to use '/'
    
    with mhi.psout.File(psoutFilePath) as psoutFile:
        try:
            data = psoutFile.call(data_path)
        except:
            print(f'The signal data path, {data_path} could not be found in {psoutFilePath}!')
            return None, None
            
        run = psoutFile.run(0)
        signal = list()
        for call in data.calls():
            trace = run.trace(call)
            time = trace.domain.data
            signal.append(trace.data)

    return time, signal


def getPsoutSignals(psoutFilePath, signalPathNames):
    '''
    Get all signals from the .psout file whose names appear in the signalPathNames list
    Exit if the first signal (time refrence) is missing.
    Missing signals are ignored.
    '''    
    
    if not signalPathNames:                                             # Return an empy DataFrame is the signalPathNames list is empy.
        return pd.DataFrame()
    
    t, _ = getPsoutSignal(psoutFilePath, signalPathNames[0])            # Get time values to get the length of all the signals in the .psout file

    if t is None:
            print(f"CRITICAL: Primary signal '{signalPathNames[0]}' (Time) not found. Cannot proceed.")
            sys.exit(1)

    t = np.array(t)                                                     # Convert to a numpy array
    t = t.reshape(1,-1)                                                 # Reshape the t from (N,) to (1,N)
    psoutSignals = np.array(t)                                          # Use as first row for the signals array
    
    columnNames = signalPathNames.copy()                                # Column names to be used for the returned DataFrame
    idxOffset = 0                                                       # To keep track of the columnnames size as signals are converter to signal arrays
    
    for signalPathName in signalPathNames:
        _, psoutSignal = getPsoutSignal(psoutFilePath, signalPathName)  # Try to get each signal in the signalnames list

        if psoutSignal is None:
            print(f"Warning: Signal '{signalPathName}' not found. Skipping...")
            columnNames.remove(signalPathName)                          # Remove the signal name
            continue        
                
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
