# Specialisterne---PDF Downloader

A python script to download a bunch of pdf files based on a .xlsx sheet.

## Description

The script ("downloadPDF.py"), when run correctly, consist of 4 parts:
    -Step 1: Extracting "BRnum", "link 1" and "link 2" from every row in a excel sheet "GRI_2017_2020 (1).xlsx" and appeneding them as a list into pdf_files[].
    -Step 2: Creating a new uniquely named folder "Downloaded pdfsX".
    -Step 3: For each item in pdf_files[], try to download the pdf from "link 1", and then from "link 2" if it failed, into the new folder. Note in item[0] whether it was downloaded (yes) or wasn't (no).
    -Step 4: Make a .cvs report with all the info from pdf_files[].

This script is very fucking slow though, so that should be the number 1 thing to optimize.
I tried to run it where it should have gone through the entire thing, it ended up downloading 105 pdfs (from BR50042 to BR50382) before the computer paused itself.
So it technically hasn't been tested if it can download all the files.

The project was given to me by "Specialisterne - Academy".
In the project we were given an old python script that didn't complete the task optimally, so we were told to either improve it, or make our own solution.
I didn't know enough about python to optimize it, so i made my own from scratch.

The files in the "\other files\" folder are as follows:
    -"downloadPDF copy.py": A slightly older version of the script (but with the same full functionality) that has the authors old print() lines for testing.
    -"Kravspecifikation.pdf": A requirement specification for the project.
    -"Metadata2006_2016.xlsx": A uneeded file given as part of the project files.
    -"OLDdownload_files.py": The old code that was given as part of the project files.
    -"PDF Downloader (Opgavebeskrivelse).pdf": The full project description.

OBS!:
The script is so far only made to handle excel sheets, and only download pdf files.
Currently you have to manually alter 2 lines in the code if you want to make the script download files from another excel sheet.
The script has only been tried on a windows 11 computer, and only through VS Code
I didn't have time to make a UML Diagram, but hopefully this readme file and the comments in the python file is descriptive enough.

## Getting Started

### Dependencies

* You need python and the following packages installed: pandas and requests

### Installing

* Download the "downloadPDF.py" and the "GRI_2017_2020 (1).xlsx" file into the same folder on your computer.

### Executing program

* Make sure that "downloadPDF.py" and "GRI_2017_2020 (1).xlsx" are located in the same folder.
    * (If you want the script to handle a different excel sheet, then you must manually alter line 24 and 37 in the code)
* Run "downloadPDF.py" (ideally by opening it in VS Code and pressing F5).
* Wait for the program to finish (can take up to 30 minutes), see the terminal for progress messages.
* When finished, open the new \Downloaded pdfsX\ folder, check it includes PDF files.
* Open the "Report.cvs" file (in excel, sheet, notepad, etc.) and see that the script has noted which files has and hasn't been downloaded.

That's it!

## Help

To make sure the script only takes 10 or so downloads, uncomment line 35, 38, 39 and 40.

To make sure the script doesn't run it's main function when run, and instead is ready for testing, change the str parameter from "__main__" on line 120 to something else. (Make sure you change it back whne you need to run it properly.)

The authors test print() lines has been removed from the current script so as not to clutter the script.
If need inspiration on where to put your own print() lines to get relevent test information, then see the a slightly older version (but with the same full functionality) of the script here: "\other files\downloadPDF copy.py".

In the bottom of the script there are some test code and functions, as well as a note of possible future improvements to the code.

I tried the script where "r = requests.get(link,timeout=3)" was set to timeout=0.5, which was faster, but it also skipped some actual pdf files such as BR50055.

## Authors

Thomas Dyrholm Siemsen