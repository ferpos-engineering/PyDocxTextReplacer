# ------------------------------------------------------------------------------
# This script sets the following parameters:
# 1. the name of a ".docx" file having text and embedded placeholders,
# 2. the name of a text file key,values formatted one per row. The value is 
#    the string replacing the associated placeholder (key). This file is called
#    database.
# 3. the name of the ".docx" file which consist in the output of this script.
# 4. the name of the temporary folder saving the unzipped archive for processing
#    only.
#
# This script builds a new ".docx" file replacing the placeholders found in the
# input ".docx" file with the corresponding values gathered in the database.
#
# (C) 2022 Marco Ferrero Poschetto, Reutlingen, Germany
# This is a software product released under GNU Public License (GPL)
# email marco.fer.pos@gmail.com
# ------------------------------------------------------------------------------

inputDocx         = "doc_sample.docx"
databaseName      = "database_sample.txt"
outputDocx        = "output.docx"
unzipTempFolder   = "TempFolder"

# ------------------------------------------------------------------------------
# readDatabase: This function extracts the content of the database file and
#               it returns to the caller it as a key, value dictionary.
# ------------------------------------------------------------------------------
def readDatabase(fileName):
    output={}
    file = open(fileName, "r")

    while True:
        line = file.readline()
 
        if not line:
            # EOF has been reached, therefore I quit the loop.
            break
        words=line.split(",")
        key=words[0]
        value=words[1]
        valueLen = len(value)
        value = value[0:valueLen-1]
        output[key]=value

    file.close()
    return output

# ------------------------------------------------------------------------------
# analyzeText:  This function reads the given text file (fileName) and replaces
#               its placeholders with the values of the specified database.
#               If placeholders found in the text file do not hit keys
#               in the database, they are replaced with some underscores.
# ------------------------------------------------------------------------------
def analyzeText(fileName, database):
    import os
    import shutil

    # Name of the temporary file, necessary to shortly park then replaced text. 
    tempFile = "./temp.xml"

    file = open(fileName, "r")
    fileOut = open(tempFile, "w")

    while True:
        line = file.readline()
 
        if not line:
            # EOF has been reached, therefore I quit the loop.
            break
        words=findPlaceholders(line)

        risultato=line
        for word in words:
            if word in database:
                value = database[word]
                risultato = risultato.replace(word, value)
            else:
                risultato = risultato.replace(word, "_______")
        fileOut.write(risultato)

    fileOut.close()
    file.close()

    # I overwrite the source file with the temporary one.
    shutil.copy(tempFile, fileName);

    # I finally delete the temporary file.
    os.remove(tempFile)
    return

# ------------------------------------------------------------------------------
# findPlaceholders:  This function finds all of the placeholders in the input
#                    text file and it returns them in a set to the caller.
#                    No placeholders found means empty set.
# ------------------------------------------------------------------------------
def findPlaceholders(text):
    placeHolders = set() 
    pos=0
    start=0
    while pos != -1:
        key, pos = findPlaceholder(text, start)
        if pos !=-1:   
            placeHolders.add(key)
            start=pos

    return placeHolders

# ------------------------------------------------------------------------------
# findPlaceholder:   This function finds the first placeholder in the input text
#                    starting from the position given as second argument.
#                    It returns a tuple (placeholder, pos2). Placeholders
#                    is the string containing the placeholder and pos2 is the
#                    position of the following character after the last % of the
#                    placeholder.
#                    For example:  this is an %%EXAMPLE%% of a %%TEXT%%. 
#                              If start is 0, the result is the tuple:
#                              %%EXAMPLE%%, 22
#                              If start is 22, the result is the tuple:
#                              %%TEXT%%, 36
#                              
#                    It the execution of this function does not hit any
#                    placeholder in the input text, the function returns the 
#                    tuple "", -1.
# ------------------------------------------------------------------------------
def findPlaceholder(text, start):
    placeholder = ""
    pos1 = text.find('%%', start) #first %%
    if(pos1==-1):
        return placeholder, pos1
    else:
        pos2 = text.find("%%", pos1 + 2) #second %%
        placeholder = text[pos1:pos2+2]
    return placeholder, pos2+2

# ------------------------------------------------------------------------------
# unzipFile:  This function unzip the sourceFile (it must be a zip archive)
#             in the destFolder folder.
# ------------------------------------------------------------------------------
def unzipFile(sourceFile, destFolder):
    import zipfile
    with zipfile.ZipFile(sourceFile,"r") as zip_ref:
        zip_ref.extractall(destFolder)
    return

# ------------------------------------------------------------------------------
# zipFolder:  This function creates a zip archived appending the content of the
#             folder (sourceFolder). The name of the zip archive is the
#             second parameter destFile.
# ------------------------------------------------------------------------------
def zipFolder(sourceFolder, destFile):
    import os
    import zipfile
  
    relroot = os.path.abspath(os.path.join(sourceFolder, os.pardir))
    with zipfile.ZipFile(destFile, "w", zipfile.ZIP_DEFLATED) as zip:
        for root, dirs, files in os.walk(sourceFolder):
            # add directory (needed for empty dirs)
            arcname = os.path.relpath(root, relroot)
            arcname = arcname.replace(sourceFolder, ".")
            #if arcname != ".":
            zip.write(root, arcname)
            for file in files:
                filename = os.path.join(root, file)
                if os.path.isfile(filename): # regular files only
                    arcname = os.path.join(os.path.relpath(root, relroot), file)
                    arcname = arcname.replace(sourceFolder, ".")
                    zip.write(filename, arcname)
    return

# ------------------------------------------------------------------------------
# deleteFolder:  This function delete the folder folderName as well as all of
#                its content.
# ------------------------------------------------------------------------------
def deleteFolder(folderName):
    import os
    import shutil

    dirpath = os.path.join('./', folderName)
    if os.path.exists(dirpath) and os.path.isdir(dirpath):
        shutil.rmtree(dirpath)
    return

# I read the content of the database file in memory.
db = readDatabase(databaseName)

# I unzip the ".docx" file in a temporary folder.
unzipFile(inputDocx, unzipTempFolder)

# I substitute the text with the database values.
analyzeText("./" + unzipTempFolder + "/word/document.xml", db)               

# I zip the content of the temporary folder in a new ".docx" file.
zipFolder(unzipTempFolder, outputDocx)

# I delete the temporary folder, because at this stage it is not useful anymore.
deleteFolder(unzipTempFolder)

#Sup