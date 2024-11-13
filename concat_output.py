
import win32com.client
import os

# gather files to merge
savepath = os.path.join(os.getcwd(), 'output')
readpath = os.path.join(savepath, 'cards')
onlyfiles = [f for f in os.listdir(readpath) if os.path.isfile(os.path.join(readpath, f))]

# create base output presentation
powerPoint = win32com.client.Dispatch("PowerPoint.Application")
outputPresentation = powerPoint.Presentations.Open(os.path.join(readpath, onlyfiles[0]))
windowRef = outputPresentation.Application.Windows(1)
outputPresentation.SaveAs(savepath+"/merged.pptx")

# insert every presentation into the output one
for name in onlyfiles[1:]:
    currentPresentation = powerPoint.Presentations.Open(os.path.join(readpath, name))
    currentPresentation.Slides.Range(range(1, currentPresentation.Slides.Count + 1)).copy()
    windowRef.Activate()
    outputPresentation.Application.CommandBars.ExecuteMso("PasteSourceFormatting")
    currentPresentation.Close()

# save the output presentation and export to pdf
outputPresentation.save()
outputPresentation.saveAs(os.path.join(savepath, "merged.pdf"), 32)
outputPresentation.close()
powerPoint.Quit()