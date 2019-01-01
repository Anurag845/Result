import PyPDF2
import xlsxwriter
 
workbook = xlsxwriter.Workbook('tecomp12.xlsx')
worksheet = workbook.add_worksheet()

pdfFileObj = open('fe.pdf','rb')

pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

pageObj = pdfReader.getPage(0)
content = pageObj.extractText()
contents = content.split()
newcontent = contents[28:]

i = 0
while(i<len(contents)):
    print(str(i) + " " + contents[i])
    i+=1
    
workbook.close()
pdfFileObj.close()

