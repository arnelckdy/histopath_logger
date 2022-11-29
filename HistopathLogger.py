#!/usr/bin/env python
# coding: utf-8

# In[1]:


#HistopathLogger script will (1) read doc or docx files, (2) extract SP No., name, age, sex, date submitted, diagnosis, consultant, and residents, and (3) output to a csv file

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


#export dependencies
#%pip freeze > requirements.txt


# In[3]:


#Create new csv file named output.csv
outputFile = open('output.csv', 'w', newline='')

#Write headers
outputDictWriter = csv.DictWriter(outputFile, ['Acc. No.', 'Date requested', 'Date submitted', 'Date processed', 'Date reported', 'Name', 'Age', 'Sex', 'Diagnosis', 'Consultant', 'Resident'])
outputDictWriter.writeheader()


# In[4]:


#Use os module to list doc and docx files in folder
for file in os.listdir(os.getcwd()):
    if file.endswith(".doc") or file.endswith(".docx"):
    
        #For each doc and docx file listed, print file name
        print(file)
    
        #Use textract to read/convert doc or docx into UTF-8 plain text
        text = textract.process(file)
        text = text.decode("utf-8")
        
        #Find accession number (SP NO.:) for Surgical, Cytology, Pap Smear, IHC, and Re-Reading by matching XXX-0000-00 pattern
        #Accession number starts with S for Surgical, C for Cyto, PS for Pap Smear, IHC for Immunohistochemistry Stain, RR for Re-Reading
        #\d*\d allows for flexibility if written as one, two, three or four digit code, eg. S-1-21. S-11-21, S-011-21, S-0011-21
        SPNoRegex = re.compile(r'(S|C|PS|IHC|RR)-\d*\d-\d\d')
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
        ageRegex = re.compile(r'Age: (\d*\d)', re.I)
        age = ageRegex.search(text)
        if age is None:
            outputAge = "NA"
        else:
            outputAge = age.group(1)
        print(outputAge)
         
        #Find sex after phrase "Sex:" and matching M or F
        #re.I to ignore case
        sexRegex = re.compile(r'Sex: (M|F)',re.I)
        sex = sexRegex.search(text)
        if sex is None:
            outputSex = "NA"
        else:
            outputSex = sex.group(1)
        print(outputSex)

        #Find date submitted after phrase "Date processed:" and matching 00-00-00 pattern
        #\d?\d allows flexibility for 0 or 00 date, eg. September written as 9 or 09
        #\d*\d allows flexibility for 00 or 0000 year, eg. 2020 written as 21 or 2021
        #re.I to ignore case
        dateProcessedRegex = re.compile(r'Date processed: (\d?\d-\d?\d-\d*\d)', re.I)
        dateProcessed = dateProcessedRegex.search(text)
        if dateProcessed is None:
            outputDateProcessed = "NA"
        else:
            outputDateProcessed = dateProcessed.group(1)
        print(outputDateProcessed)

        #Find date submitted after phrase "Date requested:" and matching 00-00-00 pattern
        #\d?\d allows flexibility for 0 or 00 date, eg. September written as 9 or 09
        #\d*\d allows flexibility for 00 or 0000 year, eg. 2020 written as 21 or 2021
        #re.I to ignore case
        dateRequestedRegex = re.compile(r'Date requested: (\d?\d-\d?\d-\d*\d)', re.I)
        dateRequested = dateRequestedRegex.search(text)
        if dateRequested is None:
            outputDateRequested = "NA"
        else:
            outputDateRequested = dateRequested.group(1)
        print(outputDateRequested)
        
        #Find date submitted after phrase "Date submitted:" and matching 00-00-00 pattern
        #\d?\d allows flexibility for 0 or 00 date, eg. September written as 9 or 09
        #\d*\d allows flexibility for 00 or 0000 year, eg. 2020 written as 21 or 2021
        #re.I to ignore case
        dateSubmittedRegex = re.compile(r'Date submitted: (\d?\d-\d?\d-\d*\d)', re.I)
        dateSubmitted = dateSubmittedRegex.search(text)
        if dateSubmitted is None:
            outputDateSubmitted = "NA"
        else:
            outputDateSubmitted = dateSubmitted.group(1)
        print(outputDateSubmitted)
      
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
        
        #Extract diagnosis
        #Check first if ER PR HER2 receptor report
        #If not receptor report, then generic diagnosis extraction
        receptorRegex = re.compile(r'RECEPTOR STATUS REPORT', re.I)
        receptor = receptorRegex.search(text)

        if receptor is None:
            
            #Find diagnosis after phrase "INTERPRETATION" and stop before first consultant
            #Look ahead (?<=) of INTERPRETATION and copy everything (.*) until look behind (?=) list of known consultants first names and most common abbreviations
            #re.DOTALL to read multiple lines
            #cleanText remove all \n by replacing them with blank space

            cleanText = text.replace("\n", " ")
            dxRegex = re.compile(r'(?<=INTERPRETATION)(.*)(?=SOCORRO|EVETTE|MARIA|JO|ARLENE|IMELDA|JANE|JR|JANET|ARACELI)',re.DOTALL)
            dx = dxRegex.search(cleanText)
            if dx is None:
                outputDx = "NA"
            else:
                outputDx = dx.group()
            print(outputDx)
        
        else:
            cleanText = text.replace("\n", " ")
            dxRegex = re.compile(r'(?<=RECEPTOR STATUS REPORT:)(.*)(?=ASCO)',re.DOTALL)
            dx = dxRegex.search(cleanText)
            if dx is None:
                outputDx = "NA"
            else:
                outputDx = dx.group()
            print(outputDx)
                
        #Find first consultant by matching known list of names
        #Last names allows for maximum flexibility eg. Socorro C Yanez, SC Yanez, Socorro Yanez, all will be recognized as Yanez
        #Exceptions are Cu and Dy since multiple doctors with same last name
        #re.I to ignore case
        consultantRegex = re.compile(r'(Ya.ez|Demaisip|Santos|Rivera|Mesina|Tilbe|Janet.*Dy|Ledesma|Jacoba|J.*Cu|MD)',re.I)
        consultant = consultantRegex.search(text)
        if consultant is None:
            outputConsultant = "NA"
        else:
            outputConsultant = consultant.group()
        print(outputConsultant)
        
        #Find first resident by matching known list of names
        #Last names allows for maximum flexibility
        #Exceptions are Cu and Dy since multiple doctors with same last name
        #findall to list all residents involve
        #re.I to ignore case
        #Not compatible with Janelyn as resident
        residentRegex = re.compile(r'(Mark.*Chua|Pastores|Ricka.*Cu|RVR.*Cu|Jill.*Perez|De.*Reyes|VE.*Cruz|ACK Dy|Ar.*Dy|Valera|Dinopol)',re.I)
        resident = residentRegex.search(text)
        if resident is None:
            outputResident = "NA"
        else:
            outputResident = resident.group()
        print(outputResident)

        #Compile extracted elements into list named dataDump
        dataDump = [outputSPNo, outputDateRequested, outputDateSubmitted, outputDateProcessed, outputDateReported, outputName, outputAge, outputSex, outputDx, outputConsultant, outputResident]
        print(dataDump)
    
        #Append dataDump to the end of the csv file
        outputWriter = csv.writer(outputFile)
        outputWriter.writerow(dataDump)


# In[5]:


#Close output csv file
outputFile.close()

