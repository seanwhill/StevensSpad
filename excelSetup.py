'''
Created on Jun 22, 2018

@author: seasaltsean
'''
import os
import re
import pandas as pd


pdfs = []
PDFDIRECTORY = '2017-18' 
EXCELSHEET_R = 'MasterSheet.xlsx'
EXCELSHEET_W = 'NewExcelsheet.xlsx'

'''
Creates an array of pdfs
'''
def loadPdfs():
    files = os.listdir(PDFDIRECTORY)
    rgx = re.compile('(.+)_(.+)_SPAD[.]pdf')
    for x in files:
        m = rgx.match(x)
        if m:
            courseNum = m.group(1)
            courseSem = m.group(2)
            pdf = courseNum + '_' + courseSem
            print(pdf)
            pdfs.append(pdf)
            
'''
Goes through each sheet and fills them with the correct courses(year and session) for this years pdfs
'''
def FillExcel():
    excl = pd.ExcelFile(EXCELSHEET_R)
    writer = pd.ExcelWriter(EXCELSHEET_W, engine='xlsxwriter')
    sheetNames = excl.sheet_names
    
    
    for sheet in sheetNames[:-1]:
        df = excl.parse(sheet)
        dct = extractExcelData(df)
        df = fillSheet(df, sheet, writer, dct)
        df.to_excel(writer, sheet, index=False)
    
    df = excl.parse('Summary')
    df.to_excel(writer, 'Summary', index=False)
    
    writer.save()

'''
Creates a dictionary of each course name ex. CS135 and the outcome number associated with it
'''
def extractExcelData(df):
    dct = {}
    for x in range(0, df['Course'].size):
        courseNumber = df['Course'][x]
        courseName = df['Name'][x]
        outcomeNumber = df['Outcome Number'][x]
        dct[courseNumber] = [courseName, outcomeNumber]
    return dct
                

'''
creates a dataframe in the format that the excel sheet typically is in with the course(year and session) and outcome number
. if there is no outcome number for that course on that sheet put NA
'''
def fillSheet(df, sheetName, writer, dct):
    courses = []
    courseNumbers = []
    outcomeNumbers = []
    emptylist = []
    rgx = re.compile('([A-Z]+[0-9]+)([A-Z]*)_([0-9]+[A-z]).*')
    
    '''
    Puts Course names in this format: CS 135, Discrete Structures (16F)
    also loads the outcomeNumbers array with the outcome Number associated with a course
    '''
    for x in pdfs:
        g = rgx.search(x)
        courseNumber = g.group(1)
        courseSection = g.group(2)
        courseSession = g.group(3)
        
        if courseNumber in dct:
            outcomeNumbers.append(dct[courseNumber][1])
            courseFormatted =  courseNumber + courseSection + ', ' + dct[courseNumber][0] + '(' + courseSession + ')'
            courseNumbers.append(courseNumber)
            courses.append(courseFormatted)
    
    '''Check if a professor did not submit a pdf'''
    for x in dct.keys():
        if x not in courseNumbers:
            courses.append(x + ' Did not have a pdf!')
            outcomeNumbers.append('ERROR')
    
    for x in range(0, len(courses) + 3):
        emptylist.append(' ')
    
    print (len(courses) * 2 + 8)
    print(len(emptylist) * 2 + 1)
    d = {'Course': courses + ['Total Students', ' ', 'Direct Assessment Average:', ' ', 'Course'] + courses + ['Total Students', ' ', 'Indirect Assessment Average:'],
          'Number of Students': emptylist + [' ', 'Number of Students'] + emptylist, 'Outcome Number': outcomeNumbers + [' ', ' ', ' ', ' ', 'Outcome Number'] + outcomeNumbers + [' ', ' ', ' '],
          'Outcome Score': emptylist + [' ', 'Outcome Score'] + emptylist, 'Weighted Direct': emptylist + [' ', 'Weighted Direct'] + emptylist}    
    newdf = pd.DataFrame(data=d)
    return newdf    

if __name__ == '__main__':
    loadPdfs()
    FillExcel()
    