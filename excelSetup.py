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
    rgx = re.compile('.+[.]pdf')
    for x in files:
        if rgx.match(x):
            #TODO use regex
            pdf = '_'.join(x.split('_', 2)[:2])
            pdfs.append(pdf)
            
'''
Goes through each sheet and fills them with the correct courses(year and session) for this years pdfs
'''
def FillExcel():
    excl = pd.ExcelFile(EXCELSHEET_R)
    writer = pd.ExcelWriter(EXCELSHEET_W, engine='xlsxwriter')
    sheetNames = excl.sheet_names
    
    df = excl.parse('Summary')
    df.to_excel(writer, 'Summary')
    
    for sheet in sheetNames[1:]:
        df = excl.parse(sheet)
        dct = extractExcelData(df)
        df = fillSheet(df, sheet, writer, dct)
        df = cleanUpSheet(df)
        df.to_excel(writer, sheet)
        
    writer.save()

'''
Deletes the rows that do not apply to the outcome number on this page
NOTICE:
If we do not have the pdf and a sheet would typically have that course with an outcome number we delete it.

Solution:
Have a master sheet with every course and outcome number.
'''
def cleanUpSheet(df):
    for i in range(0, df['Outcome Number'].size):
        if df['Outcome Number'][i] == 'NA':
            df = df.drop([i])
    return df
    
'''
Creates a dictionary of each course name ex. CS135 and the outcoeme number associated with it
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
    
    for x in range(0, 2 * len(courses) + 8):
        emptylist.append(' ')
    
    d = {'Course': courses + ['Total Students', ' ', 'Direct Assessment Average:', ' ', 'Course'] + courses + ['Total Students', ' ', 'Indirect Assessment Average:'],
          'Number of Students': emptylist, 'Outcome Number': outcomeNumbers + [' ', ' ', ' ', ' ', 'Outcome Number'] + outcomeNumbers + [' ', ' ', ' '],
          'Outcome Score': emptylist, 'Weighted Direct': emptylist}    
    newdf = pd.DataFrame(data=d)
    return newdf    

if __name__ == '__main__':
    loadPdfs()
    FillExcel()
    