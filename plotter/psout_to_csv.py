# -*- coding: utf-8 -*-
'''
This is a script to convert Manitoba Hydro International (MHI) PSOUT files to CSV format.
It reads the PSOUT files from a specified folder, extracts the required signals based on a figure setup CSV file, and writes the data to CSV files. 
It supports optional compression of the output files, i.e. .zip, .gz, .bz2, or .xz formats.
It is designed to be used in conjunction with the process_psout set of function and uses Manitoba Hydro International (MHI) PSOUT File Reader Library.
It is designed to be run from the command line with various options for input and output paths, compression type, and verbosity.
'''
__version__ = 1.0

import os, glob, time
from datetime import datetime
from process_psout import getAllSignalnames, getCaseSignalnames, getSignals
import argparse

#-----------------------------------------------------------------------------#
parser = argparse.ArgumentParser(prog = 'psout_to_csv',
                                 description = 'Convert .psout to .csv with optional compression if required')
parser.add_argument('-p', '--psoutFolder',
                    action = 'store',
                    dest = 'psoutFolder',
                    nargs = '?',
                    metavar = 'PSOUTFOLDER',
                    default = '..\\export\\MTB_16042025101945',
                    help = 'the folder where the .psout files are located')
parser.add_argument('-o', '--outputRootFolder',
                    action = 'store',
                    dest = 'outputRootFolder',
                    nargs = '?',
                    metavar = 'OUTPUTROOTFOLDER',
                    default = '..\\export',
                    help = 'the folder where the .psout files are located')
parser.add_argument('-f', '--figureSetupPath',
                    action = 'store',
                    dest = 'figureSetupPath',
                    nargs = '?',
                    metavar = 'FIGURESETUPPATH',
                    default = 'figureSetup.csv',
                    help = 'the folder where the .psout files are located')
parser.add_argument('-c', '--compressionType',
                    action = 'store',
                    dest = 'compressionType',
                    type = str,           
                    nargs = '?',          
                    metavar = 'COMPRESSIONTYPE',
                    default = '.csv',
                    help = 'the output compression type e.g. .zip, .bx2, .gz or .xz')
parser.add_argument('-q', '--quiet',
                    action = 'store_true',
                    dest = 'QUIET',
                    default = False,
                    help = 'run quietly')
parser.add_argument('-v','--version',
                    action='version',
                    version='Version: %(prog)s '+str(__version__))

args = parser.parse_args()
#-----------------------------------------------------------------------------#


def convertPsouts(psoutFolder, csvFolder, outFileType, dfSignalnames):
    psoutFilesPath = glob.glob(os.path.join(psoutFolder,'*.psout'))
    os.mkdir(csvFolder)

    for psoutFilePath in psoutFilesPath:
        psoutFileNameExt = os.path.basename(psoutFilePath)
        psoutFileName = os.path.splitext(psoutFileNameExt)[0]
        projectname, case = psoutFileName.split('_')
        case = int(case)
        if not args.QUIET: print(f'Processing {psoutFileNameExt}')
        signalnames = getCaseSignalnames(dfSignalnames, case)
        dfSignals = getSignals(psoutFilePath, signalnames)
        dfSignals.set_index('time', inplace=True)
        if not args.QUIET: print(f'Writing {projectname}_{case:02}{outFileType}\n')
        dfSignals.to_csv(os.path.join(csvFolder, f'{projectname}_{case:02}{outFileType}'), sep=';', header=True, compression='infer', decimal=',') #Note: For a Danish computer, decimal=',' else numbers are read in incorrelty in Excel

def main():
    # For .psout post processing
    start = time.time()
    signalnamesDF = getAllSignalnames(args.figureSetupPath)    
    outputFolder = os.path.join(args.outputRootFolder, f'MTB_{datetime.now().strftime(r"%d%m%Y%H%M%S")}')
    convertPsouts(args.psoutFolder, outputFolder, args.compressionType, signalnamesDF)
    end = time.time()
    elapsed = end - start
    print(f'Done! ({elapsed:.2f} s)')

if __name__ == '__main__':
    main()