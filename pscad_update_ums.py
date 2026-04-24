'''
Update unit measurement pgb names
'''
from __future__ import annotations
import os
import sys

if __name__ == '__main__':
    print(sys.version)
    #Ensure right working directory
    executePath = os.path.abspath(__file__)
    executeFolder = os.path.dirname(executePath)
    os.chdir(executeFolder)
    sys.path.append(executeFolder)
    print(executeFolder)

from typing import List

if __name__ == '__main__':
    from execute_pscad import connectPSCAD

import mhi.pscad

def updateUMs(pscad : mhi.pscad.PSCAD, legacy : bool = True, verbose : bool = False) -> None:
    """
    Update all unit measurements' instances signal names
    
    Parameters:
        pscad: PSCAD instance
        legacy: If True, all signal names will be prefixed with "alias" to create unique signal names as was required for .out format.
        verbose: If True, print the signal names being updated
    """
    projectLst = pscad.projects()
    for prjDic in projectLst:
        if prjDic['type'].lower() == 'case':
            project = pscad.project(prjDic['name'])½
            canvas = project.canvas('Main')
            print(f'Updating unit measurements in project: {project}')
            for comp in canvas.components():
                if 'unit_meas' in str(comp.defn_name[1]):
                    print(f'\t{comp}')
                    compCanvas = comp.canvas()
                    compParams = comp.parameters()
                    alias = compParams['alias']
                    pgbs = compCanvas.find_all('master:pgb')
                    for pgb in pgbs:
                        if verbose:
                            print(f'\t\t{pgb}')
                        pgbParams = pgb.parameters()
                        if legacy:
                            pgb.parameters(Name = alias + '_' + pgbParams['Group'])
                        else:
                            pgb.parameters(Name = pgbParams['Group'])
def main():
    pscad = connectPSCAD()  
    updateUMs(pscad, legacy=True, verbose=False)
    print()

if __name__ == '__main__':
    main()





