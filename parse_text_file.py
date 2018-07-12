'''
Created on Jul 12, 2018

@author: seasaltsean
'''
import re
    

if __name__ == '__main__':
    file = open('output.txt', 'r')
    rgx = re.compile('\s*[0-9]+[.][0-9].*\s+([0-9]+)\s+([0-9]+[.]*[0-9]*)\s+([0-9]+)\s+([0-9]+[.]*[0-9]*)')
    for line in file:
        m = rgx.match(line)
        if m:
            print(m.group(1) + ' ' + m.group(2) + ' ' + m.group(3) + ' ' + m.group(4))    