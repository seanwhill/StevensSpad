'''
Created on Jun 15, 2018

@author: Sean Hill
'''
from tabula import read_pdf
import os
import re
import pandas as pd
import xlsxwriter

''' key is the class name while the value is going to be a list in the format index = outcome. then string in the format of
 "numStudPerform, performResult, numStudSurv, survResult" '''
dct = {}


PDFDIRECTORY = '2017-18' 

'''
loads pdfs into an array
'''
def loadPdfs():
    files = os.listdir()
    pdfs = []
    rgx = re.compile('.+[.]pdf')
    for x in files:
        if rgx.match(x):
            pdfs.append(x)
    return pdfs


def isFloat(element):
    try:
        float(element)
        return True
    except ValueError:
        return False
    

'''
Exctracts the data from the pdf and stores it in a dictionary
key = pdfNmae
value = [numStudPerform, performResult, numStudSurv, survResult]
'''
def extractPdfData(df, pdfName):
    index = 0
    tmp = ''
    while (index < df['Program'].size):
        if type(df['Program'][index]) == str and isFloat(df['Program'][index]):
            tmp= df['Student'][index]
            numStudPerform = tmp[:-4]
            performResult = tmp[-4:]
            
            if (len(numStudPerform) == 0 or len(performResult) < 4):
                numStudPerform = performResult = 'ERROR'
                
            
            tmp = df['ACE Survey'][index]
            
            numStudSurv = tmp[:-4]
            survResult = tmp[-4:]
            
            if (len(numStudSurv) == 0 or len(survResult) < 4):
                numStudPerform = performResult = 'ERROR'
            
            dataArry = [numStudPerform, performResult, numStudSurv, survResult]
            dct[pdfName].append(dataArry) 
            
        index += 1

'''
Opens the excel sheet that we want to write to and goes through each sheet and fills them in with data from the dictionary
'''        
def FillExcel():
    excl = pd.ExcelFile("NewExcelSheet.xlsx")
    writer = pd.ExcelWriter('NewExcelSheet.xlsx', engine='xlsxwriter')
    sheetNames = excl.sheet_names
    print(sheetNames)
    
    df = excl.parse('Summary')
    df.to_excel(writer, 'Summary')
    
    for sheet in sheetNames[1:]:
        df = excl.parse(sheet)
        fillSheet(df, sheet, writer)
    writer.save()

'''
fills in each sheet with the correct data for each outcome number and each course(year and session)
'''
def fillSheet(df, sheetName, writer):
    survey = False
    for x in range(0, df['Course'].size):
        if df['Course'][x] == 'Direct Assessment Average:':
            survey = True
        if type(df['Course'][x]) == str and df['Course'][x][-1] == ')':
            #TODO all of this splitting with REGEX
            courseNumber = df['Course'][x].split(',')[0]
            courseYear = df['Course'][x].split('(')[1][:-1]
            outcomeNumber = df['Outcome Number'][x] - 1
            fullCourse = '_'.join([courseNumber, courseYear])
            #TODO: Deal with cases where pdf does not match spreadsheet.
            print(fullCourse)
            print(dct[fullCourse])
            print(outcomeNumber)
            if fullCourse in dct.keys() and outcomeNumber != 'ERROR' and 'ERROR' not in dct[fullCourse][outcomeNumber]:
                
                courseData = dct[fullCourse][outcomeNumber]
                numStud = float(courseData[0])
                studPerf = float(courseData[1])
                numSurv = float(courseData[2])
                survRes = float(courseData[3])
                
                if studPerf < 4.0 and survRes < 4.0:
                
                    if survey == True:
                        df['Number of Students'][x] = numSurv
                        df['Outcome Score'][x] = survRes
                        df['Weighted Direct'][x] = numSurv * survRes 
                    else:
                        df['Number of Students'][x] = numStud
                        df['Outcome Score'][x] = studPerf
                        df['Weighted Direct'][x] = numStud * studPerf
                else:
                    if survey == True:
                        df['Number of Students'][x] = numSurv
                        df['Outcome Score'][x] = 'ERROR'
                        df['Weighted Direct'][x] = 'ERROR' 
                    else:
                        df['Number of Students'][x] = numStud
                        df['Outcome Score'][x] = 'ERROR'
                        df['Weighted Direct'][x] = 'ERROR'
                    
            else:
                #print("course Pdf Cannot Be Found")
                pass
    df.to_excel(writer, sheetName)

        
        
    

if __name__ == '__main__':
    pdfs = loadPdfs()
    for x in pdfs:
        print(x)
        
        try:
            pdfData = read_pdf(x)
            #TODO use REGEX for this
            courseName = '_'.join(x.split('_', 2)[:2])
            dct[courseName] = []
            print(pdfData['Student'])
            extractPdfData(pdfData, courseName)
        except:
            print('Error reading pdf')
        

        
    FillExcel()
