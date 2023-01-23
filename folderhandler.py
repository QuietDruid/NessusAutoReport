import os
import glob
from pathlib import Path
from datetime import datetime
from datetime import date
from datetime import time


def doesfolderexist():
    dirResults = str(Path.home() / "Downloads")
    reportfolder = dirResults + "\\Reports"
    scansfolder = dirResults + "\\Scans"
    print("Report Folder Exists At: " + reportfolder)
    print("Scans Folder Exists At: " + scansfolder)
    if not os.path.exists(reportfolder):
        print("Report Folder Not Found, Making New Folder")
        os.mkdir(reportfolder)
    if not os.path.exists(scansfolder):
        print("Scans Folder Not Found, Making New Folder")
        os.mkdir(scansfolder)


def movetofolder():
    path = str(Path.home() / "Downloads")
    latest_file = max(glob.glob(os.path.join(path, "*.csv")), key=os.path.getctime)
    latest_report = path + "\\latestreport.pdf"
    dest_file = latest_file.replace("\\Downloads", "\\Downloads\\Scans")

    todays_date = date.today()
    report_timestamp = str(todays_date) + "--" + str(datetime.now().hour) + "-" + str(datetime.now().minute)
    dest_report = path + "\\Reports" + "\\" + report_timestamp + "--report.pdf"

    print("\nMoving Files to Respective Folders:")
    print(dest_report)
    os.rename(latest_report, dest_report)
    os.rename(latest_file, dest_file)
    print(dest_file)


def main():
    doesfolderexist()

    movetofolder()


if __name__ == '__main__':
    doesfolderexist()

    movetofolder()
