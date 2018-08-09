'''
Created on Jun 15, 2018

@author: Sean Hill
'''
import os
import re
import pandas as pd

''' key is the class name while the value is going to be a list in the format index = outcome. then string in the format of
 "numStudPerform, performResult, numStudSurv, survResult" '''
dct = {}

'''{SheetName, [DirectAssesmentAvg, IndirectAssesAvg]}'''
summaryDct = {}

TXTDIRECTORY = '2017-18' 

'''
loads text file names into an array
'''
def loadTxts():
    files = os.listdir(TXTDIRECTORY)
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
key = CourseName ie. (CS115_17F)
value = [numStudPerform, performResult, numStudSurv, survResult]
'''
   
def extractData(txt, courseName):
    file = open(TXTDIRECTORY +  '/' + txt, 'r')
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
    
    df = excl.parse('Summary')
    df.to_excel(writer, 'Summary')
    
    for sheet in sheetNames[1:]: 
        df = excl.parse(sheet)
        fillSheet(df, sheet, writer)
    
    writer.save()
    
    excl2 = pd.ExcelFile("NewExcelSheet.xlsx")
    writer2 = pd.ExcelWriter('NewExcelSheet.xlsx', engine='xlsxwriter')
    
    '''Go through the sheets again to fill in data for Summary'''
    for sheet in sheetNames:
        df = excl2.parse(sheet)
        if sheet == 'Summary':
            for x in range(0, df['Outcome'].size):
                directAssesmentAverage = summaryDct[df['Outcome'][x]][0]
                indirectAssesmentAverage = summaryDct[df['Outcome'][x]][1]

                df['Direct Assessment Average'][x] = directAssesmentAverage
                df['Indirect Assessment Average'][x] = indirectAssesmentAverage
                df['|Difference|'][x] = abs(directAssesmentAverage - indirectAssesmentAverage)
            df.to_excel(writer2, sheet, index=False)
            workbook  = writer2.book
            worksheet = writer2.sheets[sheet]
            
            format1 = workbook.add_format({'num_format': '#,##0.00'})

            worksheet.set_column('B:D', None, format1)            
        else:
            df.to_excel(writer2, sheet, index=False)
            workbook  = writer2.book
            worksheet = writer2.sheets[sheet]
            format1 = workbook.add_format({'num_format': '#,##0.00'})
            worksheet.set_column('D:E', None, format1)
    
    writer2.save()

'''
fills in each sheet with the correct data for each outcome number and each course(year and session)
'''
def fillSheet(df, sheetName, writer):
    totalNumStudents = 0
    totalWeightedDirect = 0
    survey = False
    summaryDct[sheetName] = []
    courseRgx = re.compile('(.+), .+ .+\((.+)\)')
        
    for x in range(0, df['Course'].size):
        m = courseRgx.match(df['Course'][x])
        
        if df['Course'][x] == 'Total Students':
            df['Number of Students'][x] = totalNumStudents
            df['Weighted Direct'][x] = totalWeightedDirect
            survey = True
            
        elif df['Course'][x] == 'Direct Assessment Average:':
            df['Number of Students'][x] = totalWeightedDirect / totalNumStudents
            summaryDct[sheetName].append(totalWeightedDirect / totalNumStudents)
            totalNumStudents = 0
            totalWeightedDirect = 0
        
        elif df['Course'][x] == 'Indirect Assessment Average:':
            df['Number of Students'][x] = totalWeightedDirect / totalNumStudents
            summaryDct[sheetName].append(totalWeightedDirect / totalNumStudents)

            
        elif  m:
            courseNumber = m.group(1)
            courseYear = m.group(2)
            outcomeNumber = df['Outcome Number'][x] - 1
            fullCourse = courseNumber + '_' + courseYear
                        
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
                
            
    df.to_excel(writer, sheetName, index=False)

        
        
    

if __name__ == '__main__':
    txts = loadTxts()
    txtRgx = re.compile('(.+_.+)_SPAD\.txt')
    for x in txts:  
        try:
            course = txtRgx.search(x).group(1)
            dct[course] = []
            extractData(x, course)
        except:
            print('Error reading txt')
            
    FillExcel()
