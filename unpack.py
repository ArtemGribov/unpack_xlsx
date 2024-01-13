#Developed by Artem Gribov 2023
#gribov.ag8@gmail.com

import openpyxl
import sys

from sys import argv
if len(argv) > 1:
    namescript = argv[0]
    pathsource = str(argv[1])
    pathdestiny = str(argv[2])

#pathsource = str("C:\\Users\\User\\Desktop\\file.xlsx")
#pathdestiny = str("C:\\Users\\User\\Desktop\\Temp\\")

book = openpyxl.open(pathsource, read_only = True)

def convert_tuple(t):
    s = ''
    for row in t:
        count = 0
        for c in row:
            if count == 0:
                s = s + c
                count += 1
            else:
                s = s + '\t'
                s = s + str(c)
        s = s + '\n'
    return s

#Список имен всех листов
sheetNames = book.sheetnames
with open(pathdestiny+"_$pipeline.tmp", 'r+', encoding='utf-8') as outfile2:
    if outfile2.readline() == "_%start_process":
        outfile2.write('\n')
        for name in sheetNames:
            outfile2.write(name)
            outfile2.write('\n')
            # Распаковка листа в отдельный файл
            with open(pathdestiny+name, 'w', encoding='utf-8') as outfile:
                sheet = book[name]
                t = tuple(sheet.values)
                outfile.write(str(t))
            outfile.close()
    else:
        outfile2.write("error")
        sys.exit()
    outfile2.write("_%end_process")
outfile2.close()
