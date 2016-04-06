The MIT License (MIT)

Copyright (c) 2014 Allen Plummer, https://www.linkedin.com/in/aplummer

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
from comtypes.client import CreateObject
import shutil
import os
import argparse

parser = argparse.ArgumentParser(description='Mounts Access via COM wrapper, and extracts meta objects.')
parser.add_argument('-file','--file',help='File',required=True)
parser.add_argument('-exportpath','--exportpath',help='Export Path',required=False)
args = vars(parser.parse_args())

filename = args['file']
basefilename = os.path.basename(filename)
workingfilename = os.path.dirname(filename) + os.path.sep + basefilename + '-WORKING' + os.path.splitext(basefilename)[1]
exportPath = os.path.dirname(filename) + "\\export"
if args['exportpath'] != None:
    exportPath = args['exportpath']
else:
    print "Since none provided, will use default export path: " + exportPath

if not os.path.exists(exportPath):
    os.makedirs(exportPath)
#copy file to working version.
shutil.copyfile(filename, workingfilename)
accessApplication = CreateObject("Access.Application")
accessApplication.OpenAccessProject(workingfilename)

#constants
acForm = 2
acModule= 5
acMacro = 4
acReport = 3
dictionaryDelete = CreateObject("Scripting.Dictionary")

#FORMS
for f in accessApplication.CurrentProject.AllForms:
    accessApplication.SaveAsText(acForm, f.FullName, exportPath + '\\' + f.FullName + '.form' )
    accessApplication.DoCmd.Close(acForm, f.FullName)
    dictionaryDelete.Add("FO" + f.FullName, acForm)

#MODULES
for m in accessApplication.CurrentProject.AllModules:
    accessApplication.SaveAsText(acModule, m.FullName, exportPath + '\\' + m.FullName + '.bas' )
    accessApplication.DoCmd.Close(acModule, m.FullName)
    dictionaryDelete.Add("MO" + m.FullName, acModule)

#MACROS
for m in accessApplication.CurrentProject.AllMacros:
    accessApplication.SaveAsText(acMacro, m.FullName, exportPath + '\\' + m.FullName + '.mac' )
    accessApplication.DoCmd.Close(acMacro, m.FullName)
    dictionaryDelete.Add("MA" + m.FullName, acMacro)

#REPORTS
for r in accessApplication.CurrentProject.AllReports:
    accessApplication.SaveAsText(acReport, r.FullName, exportPath + '\\' + r.FullName + '.report' )
    accessApplication.DoCmd.Close(acReport, r.FullName)
    dictionaryDelete.Add("RE" + r.FullName, acReport)

accessApplication.CloseCurrentDatabase()
os.remove(workingfilename)
accessApplication.Quit()
