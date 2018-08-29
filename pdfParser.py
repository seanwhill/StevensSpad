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

'''{SheetName, [DirectAssesmentAvg row, IndirectAssesAvg row]}'''
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
                
            dataArry = [numStudPerform, performResult, numStudSurv, survResult]
            dct[courseName].append(dataArry)
        

'''
Opens the excel sheet that we want to write to and goes through each sheet and fills them in with data from the dictionary
'''        
def FillExcel():
    excl = pd.ExcelFile("NewExcelSheet.xlsx")
    writer = pd.ExcelWriter('NewExcelSheet.xlsx', engine='xlsxwriter')
    sheetNames = excl.sheet_names
    
    for sheet in sheetNames[:-1]: 
        df = excl.parse(sheet)
        fillSheet(df, sheet, writer)
        
    df = excl.parse('Summary')
    for x in range(0, df['Outcome'].size):
        directAssesmentAverage = summaryDct[df['Outcome'][x]][0]
        indirectAssesmentAverage = summaryDct[df['Outcome'][x]][1]

        df['Direct Assessment Average'][x] = directAssesmentAverage
        df['Indirect Assessment Average'][x] = indirectAssesmentAverage
        df['|Difference|'][x] = abs(directAssesmentAverage - indirectAssesmentAverage)
        
    df.to_excel(writer, 'Summary', index=False)
    
    workbook  = writer.book
    worksheet = writer.sheets['Summary']
    format1 = workbook.add_format({'num_format': '#,##0.00'})
    format2 = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#FFE699'})
    format3 = workbook.add_format({'bold': True, 'bg_color': '#C0C0C0'})
    border_bottom = workbook.add_format({'bottom': 1})
    
    border_bottom_right = workbook.add_format({'bottom': 1, 'right':1})

    border_right = workbook.add_format({'right': 1})

    worksheet.conditional_format(0,3,0,3, {'type': 'no_errors', 'format': border_bottom_right})
    worksheet.conditional_format(11,0,11,0, {'type': 'no_errors', 'format': border_bottom_right})
    worksheet.conditional_format(11,3,11,3, {'type': 'no_errors', 'format': border_bottom_right})


    worksheet.conditional_format(0,0,11,0, {'type': 'no_errors', 'format': border_right})
    worksheet.conditional_format(11,0,11,3, {'type': 'no_errors', 'format': border_bottom})
    worksheet.conditional_format(0,3,11,3, {'type': 'no_errors', 'format': border_right})

    
    worksheet.conditional_format(0,0,0,3, {'type': 'no_errors', 'format': format2})
    worksheet.conditional_format(1,0,11,0, {'type': 'no_errors', 'format': format3})
    worksheet.set_column('A:A', 18, format1)
    worksheet.set_column('B:C', 25, format1)
    worksheet.set_column('D:D', 12, format1)
    writer.save()

'''
fills in each sheet with the correct data for each outcome number and each course(year and session)
'''
def fillSheet(df, sheetName, writer):
    totalNumStudents = 0
    totalWeightedDirect = 0
    survey = False
    summaryDct[sheetName] = []
    courseRgx = re.compile('(.+), .+\((.+)\)')
    
    courseRow = 0
    indirectAssesmentRow = 0
    directAssesmentRow = 0
    
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
            directAssesmentRow = x + 1

        
        elif df['Course'][x] == 'Indirect Assessment Average:':
            df['Number of Students'][x] = totalWeightedDirect / totalNumStudents
            summaryDct[sheetName].append(totalWeightedDirect / totalNumStudents)
            indirectAssesmentRow = x + 1
        
        elif df['Course'][x] == 'Course':
            courseRow = x + 1

            
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
    
    '''Formatting the worksheet'''
    
    workbook  = writer.book
    worksheet = writer.sheets[sheetName]
    
    format1 = workbook.add_format({'num_format': '#,##0.00'})
    format2 = workbook.add_format({'bold': True,'bg_color': '#FFE699', 'border': 1})
    format3 = workbook.add_format({'align': 'center'})
    right_border = workbook.add_format({'right': 1})
    
    top_border = workbook.add_format({'top': 1})
    border_bottom_right = workbook.add_format({'bottom': 1, 'right': 1})

    align_right = workbook.add_format({'align': 'right'})

    
    worksheet.conditional_format(0,0,0,4, {'type': 'no_errors', 'format': format2})
    worksheet.conditional_format(courseRow,0,courseRow,4, {'type': 'no_errors', 'format': format2})
    worksheet.conditional_format(indirectAssesmentRow,0,indirectAssesmentRow,4, {'type': 'no_errors', 'format': format2})
    worksheet.conditional_format(directAssesmentRow,0,directAssesmentRow,4, {'type': 'no_errors', 'format': format2})
    
    worksheet.conditional_format(directAssesmentRow-3,0,directAssesmentRow-3,0, {'type': 'no_errors', 'format': border_bottom_right})
    worksheet.conditional_format(indirectAssesmentRow-3,0,indirectAssesmentRow-3,0, {'type': 'no_errors', 'format': border_bottom_right})


    
    worksheet.conditional_format(0,0,directAssesmentRow,0, {'type': 'no_errors', 'format': right_border})
    worksheet.conditional_format(courseRow,0,indirectAssesmentRow,0, {'type': 'no_errors', 'format': right_border})

    
    worksheet.conditional_format(directAssesmentRow - 2,0,directAssesmentRow - 2,4, {'type': 'no_errors', 'format': top_border})
    worksheet.conditional_format(indirectAssesmentRow - 2,0,indirectAssesmentRow - 2,4, {'type': 'no_errors', 'format': top_border})



    
    worksheet.set_column('A:A', 40, None)
    worksheet.set_column('B:C', 18, None)
    worksheet.set_column('D:E', 14, format1)
    
    worksheet.set_row(directAssesmentRow - 2, None, align_right)
    worksheet.set_row(indirectAssesmentRow - 2, None, align_right)

    worksheet.set_row(0, None, format3)
    worksheet.set_row(courseRow, None, format3)
    worksheet.set_row(indirectAssesmentRow, None, format3)
    worksheet.set_row(directAssesmentRow, None, format3)



    
        
        
    

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
