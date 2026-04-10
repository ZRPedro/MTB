"""
pscad_pgb.py

Library for managing PGB (Output Channel) components in PSCAD projects.

Usage as library:
    import pscad_pgb_manager as pgb_mgr
    
    pscad = mhi.pscad.application()
    
    keep_signals = getSignalsFromFigureSetup(figureSetup)
    missing_total = validateFigureSetupAgainstWorkspace(pscad, keep_signals)
        
    if not missing_total:
        case_names = [p['name'] for p in pscad.projects() if p['type'] == 'Case']
        
        for case_name in case_names:
            proj = pscad.project(case_name)
            disablePGBsInProject(proj, keep_signals, disable=False)

========================================================================================================================
Can also be run as a script from within the PSCAD Python console to list the status of all PGBs across all Case projects
========================================================================================================================

"""

from __future__ import annotations
import csv
import xml.etree.ElementTree as ET
from typing import List, Dict, Tuple, Optional


# ============================================================
# SIGNAL PATH UTILITIES
# ============================================================

def _buildParentMap(proj) -> Dict[str, str]:
    """Build a map of {child_canvas -> parent_canvas} for a project."""
    parent_map = {}
    for defn_name in proj.definitions():
        try:
            canvas = proj.canvas(defn_name)
            for comp in canvas.components():
                if hasattr(comp, 'defn_name') and isinstance(comp.defn_name, tuple):
                    child_project, child_defn = comp.defn_name
                    if child_project == proj.name:
                        if child_defn not in parent_map:
                            parent_map[child_defn] = defn_name
        except Exception:
            pass
    return parent_map


def _getCanvasPath(canvas_name: str, parent_map: Dict[str, str]) -> str:
    """Build full canvas path e.g. Main\\BePPC\\Pcontrol_30"""
    path = [canvas_name]
    current = canvas_name
    while current in parent_map:
        current = parent_map[current]
        path.append(current)
    path.reverse()
    return '\\'.join(path)


def _getSignalPath(canvas_name: str, signal_name: str, parent_map: Dict[str, str]) -> str:
    """
    Build full psout signal path excluding Main.
    e.g. BePPC\\Pcontrol_30\\Pout
    Signals directly on Main are returned as just the signal name.
    """
    canvas_path = _getCanvasPath(canvas_name, parent_map)
    parts = canvas_path.split('\\')
    if parts[0] == 'Main':
        parts = parts[1:]
    if not parts:
        return signal_name
    return '\\'.join(parts) + '\\' + signal_name


def _getDisabledIds(proj) -> set:
    """Get set of component IDs that are disabled in the project XML."""
    disabled_ids = set()
    try:
        tree = ET.parse(proj.filename)
        root = tree.getroot()
        for elem in root.iter('User'):
            if elem.get('defn') == 'master:pgb':
                if elem.get('disable', 'false').lower() == 'true':
                    disabled_ids.add(int(elem.get('id')))
    except Exception as e:
        print('Warning: Could not parse project XML: ' + str(e))
    return disabled_ids


# ============================================================
# FIGURE SETUP CSV READER AND SIGNAL VALIDATION
# ============================================================

def getSignalsFromFigureSetup(figureSetupPath: str) -> List[str]:

    emt_signals = []

    try:
        with open(figureSetupPath, newline='') as f:
            reader = csv.DictReader(f, delimiter=';')
            for row in reader:
                for key in ['emt_signal_1', 'emt_signal_2', 'emt_signal_3']:
                    signal = row.get(key) or ''  # Handle None explicitly
                    signal = signal.strip()
                    if signal and signal not in emt_signals:
                        emt_signals.append(signal)
    except Exception as e:
        print('Warning: Could not read figureSetup.csv: ' + str(e))
    
    emt_signals = [s for s in emt_signals if not s.startswith('MTB\\') ]
    
    return emt_signals


def validateFigureSetupAgainstWorkspace(pscad, keep_signals: List[str]) -> List[str]:
    """
    Checks if the signals requested in the CSV exist anywhere in the workspace.
    """
    case_names = [p['name'] for p in pscad.projects() if p['type'] == 'Case']
    all_existing_paths = set()

    # Collect every single PGB path available in the workspace
    for name in case_names:
        proj = pscad.project(name)
        status_map = getPGBStatus(proj)
        for signals in status_map.values():
            for _, signal_path, _, _ in signals:
                all_existing_paths.add(signal_path)

    # Find signals in CSV that don't exist in ANY project
    missing_globally = [s for s in keep_signals if s not in all_existing_paths]

    if missing_globally:
        print("ERROR: These signals from figureSetup.csv are missing from the ENTIRE workspace:")
        for missing in missing_globally:
            print(f"  - {missing}")
    else:
        print("Success: All signals in figureSetup.csv were located in the workspace.")

    return missing_globally


# ============================================================
# PGB INSPECTION
# ============================================================

def getPGBStatus(proj) -> Dict[str, List[Tuple[str, str, bool, object]]]:
    """
    Get all PGB components in a project grouped by canvas path.

    Returns:
        Dict of {canvas_path: [(signal_name, signal_path, is_disabled, pgb_component), ...]}
    """
    parent_map = _buildParentMap(proj)
    disabled_ids = _getDisabledIds(proj)
    pgb_components = proj.find_all('master:pgb')

    canvas_path_dict = {}
    for pgb in pgb_components:
        canvas_str = str(pgb.parent)
        canvas_name = canvas_str.split(':')[-1].strip('")')
        signal_name = pgb.parameters().get('Name', '')
        is_disabled = pgb.iid in disabled_ids
        canvas_path = _getCanvasPath(canvas_name, parent_map)
        signal_path = _getSignalPath(canvas_name, signal_name, parent_map)

        if canvas_path not in canvas_path_dict:
            canvas_path_dict[canvas_path] = []
        canvas_path_dict[canvas_path].append((signal_name, signal_path, is_disabled, pgb))

    return canvas_path_dict


def printPGBStatus(proj, keep_signals: Optional[List[str]] = None) -> None:
    """
    Print all PGB components in a project grouped by canvas path.

    Parameters:
        proj: PSCAD project object
        keep_signals: Optional list of signal paths to mark as KEEP
    """
    canvas_path_dict = getPGBStatus(proj)
    total_enabled = 0
    total_disabled = 0
    will_keep = 0
    will_disable = 0

    print('\n' + '=' * 60)
    print('Project: ' + proj.name)
    print('=' * 60)

    for canvas_path in sorted(canvas_path_dict.keys()):
        signals = canvas_path_dict[canvas_path]
        n_enabled = sum(1 for _, _, d, _ in signals if not d)
        n_disabled = sum(1 for _, _, d, _ in signals if d)
        total_enabled += n_enabled
        total_disabled += n_disabled

        print('\n' + canvas_path)
        print('=' * len(canvas_path))

        for signal_name, signal_path, is_disabled, pgb in sorted(signals, key=lambda x: x[0]):
            if keep_signals:
                if signal_path in keep_signals:
                    action = 'KEEP'
                    will_keep += 1
                elif not is_disabled:
                    action = 'WILL DISABLE'
                    will_disable += 1
                else:
                    action = 'already disabled'
            else:
                action = 'disabled' if is_disabled else 'enabled'

            print('  ' + signal_name + ' [' + action + ']')

        print('\n  --------------')
        print('  Enabled : ' + str(n_enabled))
        print('  Disabled: ' + str(n_disabled))

    print('\n' + '=' * 60)
    print('Total enabled : ' + str(total_enabled))
    print('Total disabled: ' + str(total_disabled))
    if keep_signals:
        print('Will keep     : ' + str(will_keep))
        print('Will disable  : ' + str(will_disable))


# ============================================================
# PGB DISABLE/ENABLE
# ============================================================

def disablePGBsInProject(proj, keep_signals: List[str],
                          disable: bool = False,
                          verbose: bool = True) -> None:
    """
    Disable PGB components not in the keep_signals list for a SPECIFIC project.
    """
    # Use the existing status gathering tool
    canvas_path_dict = getPGBStatus(proj)

    if verbose:
        printPGBStatus(proj, keep_signals)

    disabled_count = 0
    kept_count = 0

    for canvas_path, signals in canvas_path_dict.items():
        for signal_name, signal_path, is_disabled, pgb in signals:
            in_keep_list = signal_path in keep_signals
            
            if not in_keep_list and not is_disabled:
                if disable:
                    pgb.disable()
                disabled_count += 1
            elif in_keep_list:
                kept_count += 1

    if disable:
        proj.save()
        print(f'\nProject {proj.name} Saved!')
    else:
        print(f'\nProject {proj.name} DRY RUN:')

    print('=' * 60)
    print()


def enableAllPGBs(pscad, verbose: bool = True) -> None:
    """
    Re-enable all PGB components across all Case projects.

    Parameters:
        pscad: PSCAD application object
        verbose: If True, print status.
    """
    case_names = [p['name'] for p in pscad.projects() if p['type'] == 'Case']

    for case_name in case_names:
        proj = pscad.project(case_name)
        pgb_components = proj.find_all('master:pgb')
        disabled_ids = _getDisabledIds(proj)

        enabled_count = 0
        for pgb in pgb_components:
            if pgb.iid in disabled_ids:
                pgb.enable()
                enabled_count += 1

        proj.save()
        if verbose:
            print('Project ' + case_name + ': re-enabled ' + str(enabled_count) + ' pgbs. Saved!')


# ============================================================
# MAIN
# ============================================================

if __name__ == '__main__':
    import mhi.pscad
    
    pscad = mhi.pscad.application()
    
    figureSetup = r'.\plotter\figureSetup.csv'

    if figureSetup:
        # 1. Load the "Keep" list from CSV
        keep_signals = getSignalsFromFigureSetup(figureSetup)
        
        # 2. Validate the list against the whole Workspace
        # This catches typos like 'PCC_p' vs 'PCC_P' regardless of which project they are in.
        missing_total = validateFigureSetupAgainstWorkspace(pscad, keep_signals)
        
        if not missing_total:
            # 3. Only if the CSV is 100% correct, proceed to modify projects
            case_names = [p['name'] for p in pscad.projects() if p['type'] == 'Case']
            
            for case_name in case_names:
                proj = pscad.project(case_name)
                # This function (from previous step) now handles one project at a time
                disablePGBsInProject(proj, keep_signals, disable=False)
        else:
            print("\nAborting: Please fix the signal names in figureSetup.csv before proceeding.")

    else:
        # Just print status of all PGBs
        print('No --figureSetup provided. Printing PGB status for all Case projects...')
        case_names = [p['name'] for p in pscad.projects() if p['type'] == 'Case']
        for case_name in case_names:
            proj = pscad.project(case_name)
            printPGBStatus(proj)


