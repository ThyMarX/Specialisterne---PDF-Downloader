import os
import pandas
import requests
import csv

#   ---
#
#   This version of the script has all the old lines with the test prints
#  
#   ---


inputFiles = pandas.read_excel("GRI_2017_2020 (1).xlsx")        #   Get the excel file
pdf_files = []          #   The list where we make a new list for every file [["was_downloaded_status", "BRnum", "link1", "link2"],...]
folderNr = 1            #   A int to make sure we can create a uniquely named folder
downloadedNr = 0        #   Total number of succesfully downloaded files
notDownloadedNr = 0     #   Total number of files that wasn't downloaded
rowsProcesedNr = 0      #   Total number of rows that has been processed to keep the user updated while downloading


# ---Converts the excel sheet into a python list[]
def listHandling():
    global pdf_files
    max = 0     #   when testing with max nr of files
    for index, row in inputFiles.iterrows():    #   For loop for each row of data in the .xslx file. I don't fully understand how this works, but it does ¯\_(ツ)_/¯ and the seemingly unused 'index' variable is important to declare because of iterrows() https://stackoverflow.com/questions/16476924/how-can-i-iterate-over-rows-in-a-pandas-dataframe
        pdf_files.append(["", row['BRnum'], row['Pdf_URL'], row['Report Html Address']])    #   Add the importat parts of the row as a new list to pdf_files[] 
        max += 1     #  when testing with max nr of files
        if max == 12:        #  when testing with max nr of files, set nr here
           break


# ---Makes a folder to put the files
def createDir():        
    global folderNr
    try:
        path = 'Downloaded pdfs'+str(folderNr)
        os.makedirs(path)
        #print(f"Successfully created directory: {path}")       #   test status message
    except FileExistsError:     #   if the directory already exists, try again
        #print(f"Directory already exists: {path}")     #   test status message
        folderNr += 1
        createDir()
        #       Might be a good idea to make an exception for if folderNr is above 10 or something?


# ---Downloads the files into the folder and tracks progress
def downloadAllPDFS():     
    global pdf_files
    global downloadedNr
    global notDownloadedNr
    for item in pdf_files:
        downloadOnePdf(item, 1)     #   Splits the download functions as to not repeat code and so you can more easily test one row at a time
    print("\nDone with downloading.\nFiles succusfully downloaded: "+str(downloadedNr)+" | Files failed to be downloaded: "+str(notDownloadedNr))

# -Download a file based on a specific row from the excel sheet
def downloadOnePdf(item, attempt):
    global downloadedNr
    global notDownloadedNr
    global folderNr
    global rowsProcesedNr
    error = 0       #   Error checker

    if attempt == 1:            #   First or second attempt = first or second link
        link = item[2]
    elif attempt == 2:
        link = item[3]
    
    _, file_extension = os.path.splitext(str(link))        #    link = PDF?, get the file extension (aka ".pdf")

    try:                                        #   If the program can download the file
        r = requests.get(link,timeout=3)
        r.raise_for_status()
        
        file_Path = 'Downloaded pdfs'+str(folderNr)+'/'+item[1]+'.pdf'
        
        if file_extension == ".pdf":                      #    link = PDF?, might not be good enough (see BR50051)?
            if r.status_code == 200:
                with open(file_Path, 'wb') as file:
                    file.write(r.content)
                #print("File "+item[1]+' File downloaded successfully')     #   test
                item[0] = "yes"
                downloadedNr += 1
                rowsProcesedNr += 1
        else:       #   If it's not a pdf, but the request still goes through ()
            error = 1

    except requests.exceptions.RequestException as err:     #   if the program can't download the file
        #print ("File "+item[1]+" error: ",err)     #   test error message, not needed for final output
        error = 1

    if error == 1:      #   If something went wrong, one way or another
        if attempt == 1:
            #print("Attempting second link")        #   test checking that it does try the second link
            downloadOnePdf(item, 2)
        if attempt == 2:
            item[0] = "no"
            notDownloadedNr += 1
            rowsProcesedNr += 1

    if rowsProcesedNr % 1000 == 0:      #   Progress message
        print('\n Status: The program has processed a total of '+str(rowsProcesedNr)+' rows in the excel sheet so far...')


# ---Creates final report of every file, and a summary, into the new folder
def createReport():     
    global pdf_files
    global folderNr

    with open('Downloaded pdfs'+str(folderNr)+'/Report.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(pdf_files)


# ---What the script runs:
if __name__ == "__main__":      #   Change  the str variable from "__main__" to not run the following code for test purposes
    print('\nConverting the excel sheet into a list for easier handling...')
    listHandling()
    print('\nCreating a folder to place the downloaded files and report...')
    createDir()
    print('\nDownloading files...')
    downloadAllPDFS()
    print('\nCreating the report...')
    createReport()
    print('\nThe program is done! :)\nFind your files and the report.cvs in the "Downloaded pdfs'+str(folderNr)+'" folder.')






# ---Space for test functions (remember to change "__main__" to something else)

# -When testing downloadOnePdf(), make sure you don't already have a "Downloaded pdfs1" folder.
#testRow = ["",	"BR50050", "https://www.akb.ch/documents/30573/92310/nachhaltigkeitsbericht-2016.pdf", "https://www.akb.ch/die-akb/kommunikation-medien/geschaeftsberichte"]
#downloadOnePdf(testRow, 1)
#print(testRow)

# -Usuable exceptions for the try part of downloadOnePdf() 
''' 
    #   The following code is from https://stackoverflow.com/questions/16511337/correct-way-to-try-except-using-python-requests-module

    except requests.exceptions.HTTPError as errh:       
        print ("File"+item[1]+" Http Error:",errh)
        error = 1
    except requests.exceptions.ConnectionError as errc:
        print ("File"+item[1]+" Error Connecting:",errc)
        error = 1
    except requests.exceptions.Timeout as errt:
        print ("File"+item[1]+" Timeout Error:",errt)
        error = 1
    except requests.exceptions.TooManyRedirects as errr:
        print ("File"+item[1]+" Too many redirects Error:",errr)
        error = 1
'''


# Future nice to have
#   -I could optimize the speed by not putting the excel data into a list, and just download right away in listHandling() (also by somehow download multiple files at once but i don't know how to do that?)
#   -Mention in the readme and as a print how long the program might take pr 1000 rows or 1000 downloads (figure out how long)
#   -Make the script tell the user how many rows there are in the excel sheet
#   -Make a better file type checker (see BR50051)