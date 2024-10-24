
import os
import fitz
import win32com.client
from pyxll import xl_app
import easyocr

class myclass(object):
    @staticmethod
    def get_all_file_path_include_Subfolder(folderFullPath):
        result = []
        for maindir, subdir, file_name_list in os.walk(folderFullPath):
            for filename in file_name_list:
                apath = os.path.join(maindir, filename)
                result.append(apath)
        return result   
    @staticmethod
    def getExcelApp():
        firstopen =  win32com.client.dynamic.Dispatch("Excel.Application")
        excel = xl_app()
        return excel
    @staticmethod
    def pdf_extract_text_by_pt_area(pdfFilePath,pt1,pt2, pageIndex=0, zoom_scale=10):
        '''
        pt1, left top point
        pt2, right bottom point
        '''
        imageExportFullPath = r"C:\export_temp\tempImage.png"
        if os.path.exists(r"C:\export_temp")==False: 
            os.mkdir(r"C:\export_temp")
        pdfDoc = fitz.open(pdfFilePath)
        page = pdfDoc[pageIndex]
        rotate = int(0)
        zoom_x = zoom_scale
        zoom_y = zoom_scale
        mat = fitz.Matrix(zoom_x, zoom_y).prerotate(rotate)
        clip = fitz.Rect(pt1, pt2)  
        pix = page.get_pixmap(matrix=mat, alpha=False, clip=clip)  
        if os.path.exists(imageExportFullPath):
            os.remove(imageExportFullPath)
        pix.save(imageExportFullPath)
        extractText = \
            myclass.extract_text_from_image(imageExportFullPath).\
                replace("\n\n", "\n").\
                replace("\n","")
        # print(extractText)
        if os.path.exists(imageExportFullPath):
            os.remove(imageExportFullPath)
        return extractText
    @staticmethod
    def extract_text_from_image(imagePath):
        reader = easyocr.Reader(['en'])
        results = reader.readtext(imagePath)
        texts=[]
        for (bbox, text, prob) in results:
            texts.append(text)
        result_text=str.join(" ",texts)
        return result_text



inputtext=input("please input pdf folder path: \n\r")


folderPath = inputtext
if os.path.exists(folderPath)==False:
    exit()
allFileList=myclass.get_all_file_path_include_Subfolder(folderPath)

excel=myclass.getExcelApp()
excel.Visible = True
myworkbook=excel.Workbooks.Add()
worksheet = myworkbook.Sheets[1]

worksheet.Cells(1, 1).Value = "File Name"
worksheet.Cells(1, 2).Value = "Title"
worksheet.Cells(1, 3).Value = "Drawing No"
worksheet.Cells(1, 4).Value = "Revision"
worksheet.Cells(1, 5).Value = "filePath"


rowNumber=2
for f in allFileList:
    if str(f).lower().endswith(".pdf")==False:
        continue
    filePath=f
    filename=os.path.basename(filePath)
    worksheet.Cells(rowNumber, 1).Value = filename
    
    title_pt1=(2820,2180)
    title_pt2=(3334,2278)
    
    drawingNo_pt1=(2820,2294)
    drawingNo_pt2=(3176,2330)
    
    revision_pt1=(3283,2294)
    revision_pt2=(3334,2330)
    try:
        title=myclass.pdf_extract_text_by_pt_area(filePath,title_pt1,title_pt2)
    except:
        title="error"
    try:
        drawingNo=myclass.pdf_extract_text_by_pt_area(filePath,drawingNo_pt1,drawingNo_pt2)
    except:
        drawingNo="error"
    try:
        revision=myclass.pdf_extract_text_by_pt_area(filePath,revision_pt1,revision_pt2)
    except:
        revision="error"
    worksheet.Cells(rowNumber, 2).Value = title
    worksheet.Cells(rowNumber, 3).Value = drawingNo
    worksheet.Cells(rowNumber, 4).Value = revision
    worksheet.Cells(rowNumber, 5).Value = filePath

    
    print(filename)
    rowNumber+=1

print("finish")
    
    
  









    
