import os
from styleframe import StyleFrame
import pandas as pd
import win32com.client
import glob
import random

cwd = os.getcwd()
mockups = {1: {"loc": "C:\\Users\\asus\\Desktop\\mockup1.psd",
               "smartObjLoc": "C:\\Users\\asus\\AppData\\Local\\Temp\\Capa 1412111.psb",
               "smartObjSize": (714, 824)},
           2: {"loc": "C:\\Users\\asus\\Desktop\\mockup2.psd",
               "smartObjLoc": "C:\\Users\\asus\\AppData\\Local\\Temp\\Capa 1102.psb",
               "smartObjSize": (948, 1240)},
           3: {"loc": "C:\\Users\\asus\\Desktop\\mockup3.psd",
               "smartObjLoc": "C:\\Users\\asus\\AppData\\Local\\Temp\\Capa 1412112.psb",
               "smartObjSize": (714, 824)}}

# open photoshop and smartobject
psApp = win32com.client.Dispatch("Photoshop.Application")
print("Opened photoshop")

# load and paste image
allPngs = glob.glob("C:\\Users\\asus\\Desktop\\Back to School" +
                    "/**/*.png", recursive=True)
print(f"Found {len(allPngs)} png to paste on mockup")
i = 1

# loop for all pngs
for png in allPngs:
    pngName = png.split("\\")[-1]
    rand = random.randint(1, len(mockups))
    tshirt = psApp.Open(mockups[rand]["loc"])
    smartObject = psApp.Open(mockups[rand]["smartObjLoc"])
    psApp.Load(png)
    psApp.ActiveDocument.Selection.SelectAll()
    psApp.ActiveDocument.Selection.Copy()
    psApp.ActiveDocument.Close()
    psApp.ActiveDocument.Paste()
    smartObject.ArtLayers[1].Delete()

    # resize image
    layer = smartObject.ArtLayers[1]
    layerSize = layer.bounds[2:]
    docSize = mockups[rand]["smartObjSize"]
    if docSize[0] <= docSize[1]:
        resizeFactor = min(docSize)/layerSize[0]*100
        layer.Resize(Horizontal=resizeFactor, Vertical=resizeFactor, Anchor=1)
        layer.Translate(0, (docSize[1]-layer.bounds[-1])//2)

    # save and close image and export jpg
    psApp.ActiveDocument.Save()
    options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
    options.Format = 6
    options.Quality = 100

    jpgFile = f"C:\\Users\\asus\\Desktop\\{pngName}-Mockup.jpg"
    psApp.ActiveDocument = tshirt
    psApp.ActiveDocument.Export(ExportIn=jpgFile, ExportAs=2, Options=options)
    psApp.ActiveDocument.Save()
    print(
        f"exported tshirt {i} of {len(allPngs)} ({int(i/len(allPngs)*100)}%)")
    i += 1


allPngs = [i.split("\\")[-1] for i in allPngs]

columns = ['Folder', 'File', 'Posted']
df = pd.DataFrame({'Folder': {0: "Back to School"},
                   'File': {0: allPngs[1]},
                   'Posted': {0: "No"}}, columns=columns)

for i in allPngs:
    df = df.append({'Folder': "Back to School",
                    'File': i,
                    'Posted': "No"}, ignore_index=True)

excel_writer = StyleFrame.ExcelWriter('ExcelFileRecords.xlsx')
sf = StyleFrame(df)
sf.to_excel(
    excel_writer=excel_writer,
    best_fit=columns,
    columns_and_rows_to_freeze='B2',
    row_to_add_filters=0,
)
excel_writer.save()
print("Exported excel file too")
