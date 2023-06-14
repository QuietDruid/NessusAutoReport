# Nessus AutoReport 
## NOTICE CURRENT VERSION PROBABLY DOESN'T WORK ON WINDOWS DUE TO DIRECTORY SLASHES USE ON LINUX OR WSL

## Idea

Short "portable" script designed to use vulnerablility results from Tenable's Nessus. Taking a docx template and display the number of vulnerabilities, sorted by severity. Then listing which losts are impacted by IP and by which potential vulnerability. While Nessus does provide its own reports, simplfying down to certain vulnerabilities assigning names and numbers to the problem helps getting action taken to resolve the issue. Only requiring technician intervention when initially downloading the program onto work device and modifying the template logo, as well as, before presenting if any vulnerability is a non-issue.

## Design
Decided with python to use take full advantage of the libraries at hand. Heavy usage of selenium for interfacing with the local web interface, pandas to handle the csv created, docx to deal with the template, docx2pdf for the final presentable form, and then pyinstaller to create and exe that can potential be deployed to windows 10/11 without having to install additional packages other than the exe ad the docx template.

## Usage
Needed: Windows 10/11, Tenable's Nessus (with atleast 1 Scan and 1 scan done), Word

 1. Insert username and password into script in `main.py` then compile using pyinstaller
 2. Download the exe and modifying the image logo in the template
 3. Make sure template and exe are in the same folder, and that you have read/write permissions within the system.
 4. Double-check that Tenable has a scan result to download.
 5. Run the exe.
 6. After execution check the `Reports` folder which is located within the local `Downloads` folder for the latest report.

### Version 0.6
It works tentatively. 

**TO DO:**
* Need to test and implement timing mechanism to run on intervals or use the clock to launch reports
* Potentially implement a secure file uploading system to a centralized server for manual review
* Change/Receive feedback on file location
* Test potential issues with permissions with creating files and folders

