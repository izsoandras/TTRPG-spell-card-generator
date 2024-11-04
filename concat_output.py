
import win32com.client
import os
from os import listdir
from os.path import isfile, join

mypath = './output'
onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]

powerPoint = win32com.client.Dispatch("PowerPoint.Application")
outputPresentation = powerPoint.Presentations.Open(onlyfiles[0])
windowRef = outputPresentation.Application.Windows(1)
outputPresentation.SaveAs("output.pptx")

for name in onlyfiles[1:]:
    currentPresentation = powerPoint.Presentations.Open(name)
    currentPresentation.Slides.Range(range(1, currentPresentation.Slides.Count + 1)).copy()
    windowRef.Activate()
    outputPresentation.Application.CommandBars.ExecuteMso("PasteSourceFormatting")
    currentPresentation.Close()

outputPresentation.save()
outputPresentation.saveAs("output.pdf", 32)
outputPresentation.close()
powerPoint.Quit()