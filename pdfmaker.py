import csv
import time

import pandas as pd
import os
import glob
from docx import Document
from pathlib import Path
from docx2pdf import convert


def pdfwriter(pathtofile, crit, high, medium, low):
    docu = Document('rep-template-v0.docx')
    pathtofile = pathtofile + "\\latestreport.docx"
    print("\nPath To Latest Report in Docx: " + pathtofile)
    docu.save(pathtofile)

    docu.tables[0].cell(0, 1).text = '1'
    docu.tables[0].cell(0, 3).text = '1'
    docu.tables[0].cell(1, 1).text = 'Dec 1'
    docu.tables[0].cell(1, 3).text = 'Dec 17'

    docu.tables[1].cell(0, 1).text = str(crit)
    docu.tables[1].cell(0, 1).paragraphs[0].alignment = 1

    docu.tables[1].cell(1, 1).text = str(high)
    docu.tables[1].cell(1, 1).paragraphs[0].alignment = 1

    docu.tables[1].cell(2, 1).text = str(medium)
    docu.tables[1].cell(2, 1).paragraphs[0].alignment = 1

    docu.tables[1].cell(3, 1).text = str(low)
    docu.tables[1].cell(3, 1).paragraphs[0].alignment = 1

    docu.tables[1].cell(4, 1).text = str(crit + high + medium + low)
    docu.tables[1].cell(4, 1).paragraphs[0].alignment = 1
    docu.save(pathtofile)


def vulnwriter(pathtofile, sortedcsv):
    docu = Document(pathtofile + "\\latestreport.docx")
    count = 1
    print("\nVulnerabilities Being Written to Report:")
    for vulnscore in sortedcsv.values:
        if pd.isnull(vulnscore[2]):
            break
        if len(docu.tables[2].rows) == count:
            docu.tables[2].add_row()
            docu.tables[2].style = 'Table Grid'
        docu.tables[2].cell(count, 0).text = vulnscore[4]
        docu.tables[2].cell(count, 0).paragraphs[0].alignment = 1

        docu.tables[2].cell(count, 1).text = vulnscore[8]
        docu.tables[2].cell(count, 1).paragraphs[0].alignment = 1

        docu.tables[2].cell(count, 2).text = str(vulnscore[2])
        docu.tables[2].cell(count, 2).paragraphs[0].alignment = 1

        count += 1
        print(str(vulnscore[2]) + " " + vulnscore[8])
    docu.save(pathtofile + "\\latestreport.docx")
    print()


def callpdfmaker():
    path = str(Path.home() / "Downloads")
    csv_files = glob.glob(os.path.join(path, "*.csv"))
    latest_file = max(glob.glob(os.path.join(path, "*.csv")), key=os.path.getctime)
    output = pd.read_csv(latest_file)
    print("Latest CSV File To Be Read: " + latest_file)
    risks = output['Risk'].values.tolist()

    print("\nCurrent Vulnerabilities")
    print("High Risks: " + str(risks.count('High')))
    print("Medium Risks: " + str(risks.count('Medium')))
    print("Low Risks: " + str(risks.count('Low')))

    pdfwriter(path, risks.count('Critical'), risks.count('High'), risks.count('Medium'), risks.count('Low'))

    sorted_df = output.sort_values(by=["CVSS v2.0 Base Score"], ascending=False)

    vulnwriter(path, sorted_df)

    docxtemp = path + "\\latestreport.docx"
    outpdf = path + "\\latestreport.pdf"
    time.sleep(10)
    print("Temp Files For Editing:")
    print(docxtemp)
    print(outpdf + "\nConverting Docx to PDF")
    convert(docxtemp)


if __name__ == '__main__':
    callpdfmaker()
