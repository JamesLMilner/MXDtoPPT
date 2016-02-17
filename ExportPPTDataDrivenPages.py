import sys
import os
import arcpy
import win32com.client

mxd = arcpy.mapping.MapDocument(sys.argv[1])

def exportppt_datadriven(mxd, filepath, filename, pageNum ):
    '''Export MXD toh Powerpoint Presentation'''
    
    #Get map as vector (.emf)
    arcpy.mapping.ExportToEMF(mxd, filepath + "\\Map.emf")

    #Mapimage
    mapimg = filepath + "\\Map.emf"

    #Input SVG into Powerpoint
    Application = win32com.client.Dispatch("PowerPoint.Application")
    Presentation = Application.Presentations.Add()
    Mapslide = Presentation.Slides.Add(1, 12)
    picture = Mapslide.Shapes.AddPicture(FileName=mapimg, LinkToFile=False, SaveWithDocument=True,
                                       Left=1, Top=1)

    #Finish Up
    Presentation.SaveAs(filepath + "\\" + filename + str(pageNum) +".ppt")
    Presentation.Close()
    Application.Quit()
    del mxd
    os.remove(mapimg)


for pageNum in range(1, mxd.dataDrivenPages.pageCount + 1):
    mxd.dataDrivenPages.currentPageID = pageNum
    filepath = sys.argv[2]
    filename = sys.argv[3]    
    exportppt_datadriven(mxd, filepath, filename, pageNum)


