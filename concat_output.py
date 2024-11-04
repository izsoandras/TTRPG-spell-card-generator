
import win32com.client
import os

# gather files to merge
mypath = os.path.join(os.getcwd(), 'output')
onlyfiles = [f for f in os.listdir(mypath) if os.path.isfile(os.path.join(mypath, f))]

# create base output presentation
powerPoint = win32com.client.Dispatch("PowerPoint.Application")
outputPresentation = powerPoint.Presentations.Open(os.path.join(mypath, onlyfiles[0]))
windowRef = outputPresentation.Application.Windows(1)
outputPresentation.SaveAs(mypath+"/merged.pptx")

# insert every presentation into the output one
for name in onlyfiles[1:]:
    currentPresentation = powerPoint.Presentations.Open(os.path.join(mypath, name))
    currentPresentation.Slides.Range(range(1, currentPresentation.Slides.Count + 1)).copy()
    windowRef.Activate()
    outputPresentation.Application.CommandBars.ExecuteMso("PasteSourceFormatting")
    currentPresentation.Close()

# save the output presentation and export to pdf
outputPresentation.save()
outputPresentation.saveAs(os.path.join(mypath, "merged.pdf"), 32)
outputPresentation.close()
powerPoint.Quit()