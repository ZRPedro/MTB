import argparse
import os
import shutil
from glob import glob
from datetime import datetime
import pandas as pd

def readCaseRankTaskID(csv_path):
    df = pd.read_csv(csv_path)
    # Expecting columns: "Case Rank", "Task ID", "Case Name"
    mapping = dict(zip(df["Task ID"].astype(int), df["Case Rank"].astype(int)))
    return mapping

def moveAndRenamePsoutFiles(buildPath, exportPath, mapping, projectName=None):
    timeStamp = datetime.now().strftime("%Y%m%d%H%M%S")
    destFolder = os.path.join(exportPath, f"OOM_{timeStamp}")
    os.makedirs(destFolder, exist_ok=True)
    psoutFilePaths = glob(os.path.join(buildPath,f"{projectName}*.psout")) if projectName else glob(os.path.join(buildPath,"*.psout"))
    for psoutFilePath in psoutFilePaths:
        psoutFileName = os.path.basename(psoutFilePath)
        parts = psoutFileName.split("_")
        taskID = parts[-1].replace(".psout", "")
        destPath = os.path.join(destFolder, psoutFileName.replace(".psout", ".psout_taskid"))
        print(f"Moving {psoutFilePath} to {destPath}")
        shutil.move(psoutFilePath, destPath)
    
    print()
    for psoutTaskIdFilePath in glob(os.path.join(destFolder,"*.psout_taskid")):
        psoutTaskIdFileName = os.path.basename(psoutTaskIdFilePath)
        parts = psoutTaskIdFileName.split("_")
        projectName = parts[0]
        taskID = int(parts[1].replace(".psout", ""))
        caseRank = mapping.get(taskID)
        psoutFileName = f"{projectName}_{caseRank:02d}.psout"
        print(f"Renaming {psoutTaskIdFileName} to {psoutFileName}")
        os.rename(
            os.path.join(destFolder, psoutTaskIdFileName),
            os.path.join(destFolder, psoutFileName)
        )

def main():
    parser = argparse.ArgumentParser(description="The script recovers the .psout files from the build folder after a PSCAD crash, e.g OOM")
    parser.add_argument("-b", "--buildPath", required=True, help="Path to .psout files that needs to be recovered")
    parser.add_argument("-e", "--exportPath", default=r"..\export", help=r'Root folder for export (default = "..\export")')
    parser.add_argument("-n", "--projectName", required=False, help="Project name (optional)")
    parser.add_argument("-c", "--csvPath", default=r"..\caseRankTaskID.csv", help=r'CSV file with Task ID to Case Rank mapping (default = "..\caseRankTaskID.csv)"')
    args = parser.parse_args()

    mapping = readCaseRankTaskID(args.csvPath)
    moveAndRenamePsoutFiles(args.buildPath, args.exportPath, mapping, args.projectName)

if __name__ == "__main__":
    main()
