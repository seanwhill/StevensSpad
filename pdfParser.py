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
def loadTxts():
    files = os.listdir()
    txts = []
    rgx = re.compile('.+_SPAD[.]txt')
    for x in files:
        if rgx.match(x):
            txts.append(x)
    return txts


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
def extractData(txt, courseName):
    file = open(txt, 'r')
    rgx = re.compile('\s*[0-9]+[.][0-9].*\s+([0-9]+)\s+([0-9]+[.]*[0-9]*)\s+([0-9]+)\s+([0-9]+[.]*[0-9]*)')
    for line in file:
        m = rgx.match(line)
        if m:
            numStudPerform = float(m.group(1))
            performResult = float(m.group(2))
            numStudSurv = float(m.group(3))
            survResult = float(m.group(4))
                
            print(m.group(1) + ' ' + m.group(2) + ' ' + m.group(3) + ' ' + m.group(4))                
            dataArry = [numStudPerform, performResult, numStudSurv, survResult]
            dct[courseName].append(dataArry)
        

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
    totalNumStudents = 0
    totalWeightedDirect = 0
    survey = False
    for x in range(0, df['Course'].size):
        if df['Course'][x] == 'Total Students':
            df['Number of Students'][x] = totalNumStudents
            df['Weighted Direct'][x] = totalWeightedDirect
            survey = True
            
        elif df['Course'][x] == 'Direct Assessment Average:':
            df['Number of Students'][x] = totalWeightedDirect / totalNumStudents
            totalNumStudents = 0
            totalWeightedDirect = 0
        
        elif df['Course'][x] == 'Indirect Assessment Average:':
            df['Number of Students'][x] = totalWeightedDirect / totalNumStudents
            

        elif type(df['Course'][x]) == str and df['Course'][x][-1] == ')':
            #TODO all of this splitting with REGEX
            courseNumber = df['Course'][x].split(',')[0]
            courseYear = df['Course'][x].split('(')[1][:-1]
            outcomeNumber = df['Outcome Number'][x] - 1
            fullCourse = '_'.join([courseNumber, courseYear])
            #TODO: Deal with cases where pdf does not match spreadsheet.
            print('AAAAAAA')
            print(sheetName)
            print(fullCourse)
            print(dct.keys())
            print(len(dct[fullCourse]))
            print(outcomeNumber)
            if len(dct[fullCourse]) < outcomeNumber:
                df['Outcome Score'][x] = 'ERROR Outcome Number is not in the pdf'
                continue
            
            if fullCourse in dct.keys() and outcomeNumber != 'ERROR' and 'ERROR' not in dct[fullCourse][outcomeNumber]:
                
                courseData = dct[fullCourse][outcomeNumber]
                
                numStud = courseData[0]
                studPerf = courseData[1]
                numSurv = courseData[2]
                survRes = courseData[3]
                
                
                if survey == True:
                    df['Number of Students'][x] = numSurv
                    
                    if survRes <= 4.0:
                        df['Outcome Score'][x] = survRes
                        df['Weighted Direct'][x] = numSurv * survRes
                    else:
                        survRes = round(survRes / 25, 2)
                        df['Outcome Score'][x] = str(survRes) + ', ERROR data not recorded on 4.0 scale'
                        
                    df['Weighted Direct'][x] = numSurv * survRes
                    
                    totalNumStudents += numSurv
                    totalWeightedDirect += numSurv * survRes
                    
                else:
                    df['Number of Students'][x] = numStud
                    
                    if studPerf <= 4.0:
                        df['Outcome Score'][x] = studPerf
                    else:
                        studPerf = round(studPerf / 25, 2)
                        df['Outcome Score'][x] = str(studPerf) + ', ERROR data not recorded on 4.0 scale'
                        
                    df['Weighted Direct'][x] = numStud * studPerf
                    
                    totalNumStudents += numStud
                    totalWeightedDirect += numStud * studPerf

            else:
                #print("course Pdf Cannot Be Found")
                pass
    df.to_excel(writer, sheetName)

        
        
    

if __name__ == '__main__':
    txts = loadTxts()
    for x in txts:
        print(x)
        
        try:
            #TODO use REGEX for this
            courseName = '_'.join(x.split('_', 2)[:2])
            dct[courseName] = []
            extractData(x, courseName)
        except:
            print('Error reading pdf')
        

        
    FillExcel()
