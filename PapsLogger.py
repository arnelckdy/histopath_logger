#!/usr/bin/env python
# coding: utf-8

# In[1]:


#PapsLogger script will (1) read doc or docx files, (2) extract Cytology No., name, age, date received, date reported, consultant, and residents, and (3) output to a csv file

#Initialize
#!/home/asus/anaconda3/bin/python

#Version 0.1
#Last modified 09-19-21

#Import OS module to handle multiple Word files
import os

#Import RegEx module to specify search parameters
import re

#Import textract module to extract data from doc and docx Word documents by turning them into plain text
#If not installed, go to this URL https://textract.readthedocs.io/en/stable/installation.html
#Or use pip install textract
import textract

#Import json module to export extracted data to Excel readable file
import csv


# In[2]:


#Create new csv file named output.csv
outputFile = open('output.csv', 'w', newline='')

#Write headers
outputDictWriter = csv.DictWriter(outputFile, ['Acc. No.', 'Date requested', 'Date reported', 'Name', 'Age', 'Specimen type', 'Specimen adequacy', 'General categorization', 'Interpretation', 'Non-neoplastic findings', 'Organisms', 'Maturation index', 'Remarks', 'Consultant', 'Resident'])
outputDictWriter.writeheader()


# In[3]:


#Use os module to list doc and docx files in folder
for file in os.listdir(os.getcwd()):
    if file.endswith(".doc") or file.endswith(".docx"):
    
        #For each doc and docx file listed, print file name
        print(file)
    
        #Use textract to read/convert doc or docx into UTF-8 plain text
        text = textract.process(file)
        text = text.decode("utf-8")
        
        #Find accession number (SP NO.:) by matching XXX-0000-00 pattern
        #\d*\d allows for flexibility if written as one, two, three or four digit code, eg. S-1-21. S-11-21, S-011-21, S-0011-21
        SPNoRegex = re.compile(r'PS-\d*\d-\d\d')
        SPNo = SPNoRegex.search(text)
        if SPNo is None:
            outputSPNo = "NA"
        else:
            outputSPNo = SPNo.group()
        print(outputSPNo)
    
        #Find patient name after phrase "Name:" and stop before "Age:"
        #Look ahead (?<=) of Name: and copy everything (.*) until look behind (?=) Age:
        #re.I to ignore case
        nameRegex = re.compile(r'(?<=Name:)(.*)(?=Age:)', re.I)
        name = nameRegex.search(text)
        if name is None:
            outputName = "NA"
        else:
            outputName = name.group()
        print(outputName)
    
        #Find age after phrase "Age:" and matching 000 pattern
        #\d*\d allows for flexibility in age from single to triple digits, eg. 1, 10, 100
        #re.I to ignore case
        ageRegex = re.compile(r'Age:\s*(\d*\d)', re.I)
        age = ageRegex.search(text)
        if age is None:
            outputAge = "NA"
        else:
            outputAge = age.group(1)
        print(outputAge)
         
        #Find date submitted after phrase "Date processed:" and matching 00-00-00 pattern
        #\d?\d allows flexibility for 0 or 00 date, eg. September written as 9 or 09
        #\d*\d allows flexibility for 00 or 0000 year, eg. 2020 written as 21 or 2021
        #re.I to ignore case
        dateReceivedRegex = re.compile(r'Date Received: (\d?\d-\d?\d-\d*\d)', re.I)
        dateReceived = dateReceivedRegex.search(text)
        if dateReceived is None:
            outputDateReceived = "NA"
        else:
            outputDateReceived = dateReceived.group(1)
        print(outputDateReceived)
      
        #Find date reported after phrase "Date reported:" and "Final report:" and matching 00-00-00 pattern
        #\d?\d allows flexibility for 0 or 00 date, eg. September written as 9 or 09
        #\d*\d allows flexibility for 00 or 0000 year, eg. 2020 written as 21 or 2021
        #re.I to ignore case
        dateReportedRegex = re.compile(r'(Date reported:|Date final.*:) (\d?\d-\d?\d-\d*\d)', re.I)
        dateReported = dateReportedRegex.search(text)
        if dateReported is None:
            outputDateReported = "NA"
        else:
            outputDateReported = dateReported.group(2)
        print(outputDateReported)

        #Look ahead (?<=) of SPECIMEN TYPE: and copy everything (.*) until look behind (?=):
        #re.I to ignore case
        spTypeRegex = re.compile(r'(?<=SPECIMEN TYPE:)(.*)', re.I)
        spType = spTypeRegex.search(text)
        if spType is None:
            outputSpType = "NA"
        else:
            outputSpType = spType.group()
        print(outputSpType)      

        #Look ahead (?<=) of ADEQUACY: and copy everything (.*) until look behind (?=):
        #re.I to ignore case
        spAdeqRegex = re.compile(r'(?<=ADEQUACY:)(.*)', re.I)
        spAdeq = spAdeqRegex.search(text)
        if spAdeq is None:
            outputSpAdeq = "NA"
        else:
            outputSpAdeq = spAdeq.group()
        print(outputSpAdeq)      

        #Look ahead (?<=) of GENERAL CATEGORIZATION: and copy everything (.*) until look behind (?=):
        #re.I to ignore case
        genCatRegex = re.compile(r'(?<=GENERAL CATEGORIZATION:)(.*)', re.I)
        genCat = genCatRegex.search(text)
        if genCat is None:
            outputGenCat = "NA"
        else:
            outputGenCat = genCat.group()
        print(outputGenCat)      

        #Look ahead (?<=) of INTERPRETATION/RESULT: and copy everything (.*) until look behind (?=):
        #re.I to ignore case
        resultRegex = re.compile(r'(?<=INTERPRETATION/RESULT:)(.*)', re.I)
        result = resultRegex.search(text)
        if result is None:
            outputResult = "NA"
        else:
            outputResult = result.group()
        print(outputResult)      
        
        #Look ahead (?<=) of NON-NEOPLASTIC FINDINGS: and copy everything (.*) until look behind (?=):
        #re.I to ignore case
        nonNeoRegex = re.compile(r'(?<=NON-NEOPLASTIC FINDINGS:)(.*)', re.I)
        nonNeo = nonNeoRegex.search(text)
        if nonNeo is None:
            outputNonNeo = "NA"
        else:
            outputNonNeo = nonNeo.group()
        print(outputNonNeo)      

        #Look ahead (?<=) of ORGANISMS: and copy everything (.*) until look behind (?=):
        #re.I to ignore case
        orgRegex = re.compile(r'(?<=ORGANISMS:)(.*)', re.I)
        org = orgRegex.search(text)
        if org is None:
            outputOrg = "NA"
        else:
            outputOrg = org.group()
        print(outputOrg)      

        #MATURATION INDEX
        #re.I to ignore case
        matIndexRegex = re.compile(r'(?<=MATURATION INDEX:)(.*\d-.*\d-.*\d)', re.I)
        matIndex = matIndexRegex.search(text)
        if matIndex is None:
            outputMatIndex = "NA"
        else:
            outputMatIndex = matIndex.group()
        print(outputMatIndex)      

        #Look ahead (?<=) of REMARKS: and copy everything (.*) until look behind (?=):
        #re.I to ignore case
        remRegex = re.compile(r'(?<=REMARKS:)(.*)(?=Pathologist)', re.DOTALL)
        rem = remRegex.search(text)
        if rem is None:
            outputRem = "NA"
        else:
            outputRem = rem.group()
        print(outputRem)      
        
        #Find first consultant by matching known list of names
        #For those with similar last names, add wildcard after distinguishing last letter
        #Ex. If Katherine Hepburn is a consultant and Audrey Hepburn is a resident:
        #       A*Hepburn and K*Hepburn so the script will distinguish Aubrey Hepburn from Katherine Hepburn
        #       This also works if both Katherine and Audrey Hepburn are both consultants or both residents.
        #re.I to ignore case
        consultantRegex = re.compile(r'(K*Hepburn|Monroe|Flynn|Wayne|Hayworth|Garland|MD)',re.I)
        consultant = consultantRegex.search(text)
        if consultant is None:
            outputConsultant = "NA"
        else:
            outputConsultant = consultant.group()
        print(outputConsultant)
        
        #Find first resident by matching known list of names
        #Last names allows for maximum flexibility
        #findall to list all residents involve
        #re.I to ignore case
        residentRegex = re.compile(r'(A*HepburnTracy|Bergman|Crawford|Fontaine|Dunne|Robinson)',re.I)
        resident = residentRegex.search(text)
        if resident is None:
            outputResident = "NA"
        else:
            outputResident = resident.group()
        print(outputResident)

        #Compile extracted elements into list named dataDump
        dataDump = [outputSPNo, outputDateReceived, outputDateReported, outputName, outputAge, outputSpType, outputSpAdeq, outputGenCat, outputResult, outputNonNeo, outputOrg, outputMatIndex, outputRem, outputConsultant, outputResident]
        print(dataDump)
    
        #Append dataDump to the end of the csv file
        outputWriter = csv.writer(outputFile)
        outputWriter.writerow(dataDump)


# In[4]:


#Close output csv file
outputFile.close()

