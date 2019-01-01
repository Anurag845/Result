import PyPDF2
import xlsxwriter

workbook = xlsxwriter.Workbook('feres.xlsx')
worksheet = workbook.add_worksheet()

pdfFileObj = open('fe.pdf','rb')

pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

def main():
    pgno = 0
    i = 1
    num = pdfReader.getNumPages()
    print(num)
    pageObj = pdfReader.getPage(0)
    content = pageObj.extractText()
    contents = content.split()
    bindex1 = 26
    newcontent = contents[28:]
    n1,n2 = initialize(newcontent) 
    while(pgno < num):
        
        pageObj = pdfReader.getPage(pgno)
        content = pageObj.extractText()
        contents = content.split()
        newcontent = contents[bindex1:]
        noofsub1,noofsub2 = NoOfSub(newcontent)
        m,flag = writer(i,newcontent,noofsub1,noofsub2,n1)					#writing marks of 1st student
        i+=9
        if(flag):
            newcontent1 = newcontent[m+27:]
        else:
            newcontent1 = newcontent[m+17:]
        noofsub1,noofsub2 = NoOfSub(newcontent1)
        m,flag = writer(i,newcontent1,noofsub1,noofsub2,n1)					#writing marks of 2nd student
        pgno+=1
        i+=9


def initialize(array):
    k = 0
    j = 1
    #m = 0
    worksheet.write(0,0,'PRN')
    
    checkprn = array[0]
    while(checkprn != 'PICT' and k < len(array) and not checkprn.endswith('PICT')):
        checkprn = array[k]
        k += 1
        
    if k == len(array):
        return 0,False
    if(checkprn.endswith('PICT') and len(checkprn)>4):
        checkprn = checkprn[:-4]
        array = array[k-2:]
        array[0] = checkprn
    else:
        array = array[k-2:]
        check = array[0]
    
        k = 0
        while(check[k] != '7' and k < len(array)):
            k+=1
        
        check = check[k:]
        array[0] = check
    
    array = array[27:] #1st subject code
    j = 3
    
    m = 0 
    cnt = 0
    while(array[m] != 'SEM.:2' and array[m] != 'SGPA1'):
        str = array[m]
        if(not str[len(str)-1].isalpha()):
            worksheet.write(0,j,array[m])
        if(array[m+1] == '*'):
            m += 13
        else:
            m += 12
        cnt+=1
        j += 1
        
    noofsub1 = cnt
    
    array = array[m+1:]
    j = cnt + 3
    m = 0
    cnt = 0
    while(array[m] != 'FIRST' and array[m] != 'SECOND' and array[m] != 'THIRD' and array[m] != 'FOURTH' ):
        str = array[m]
        if(not str[len(str)-1].isalpha() and str!="107010"):
            worksheet.write(0,j,array[m])
        if(array[m+1] == '*'):
            m += 13
        else:
            m += 12
        cnt+=1
        j += 1
            
    noofsub2 = cnt
    return noofsub1,noofsub2
    
def NoOfSub(array):
    k = 0
    #j = 1
    Flag = False
    checkprn = array[0]
    while(checkprn != 'PICT' and k < len(array) and not checkprn.endswith('PICT')):
        checkprn = array[k]
        k += 1
    if k == len(array):
        return 0,False
    if(checkprn.endswith('PICT') and len(checkprn)>4):
        checkprn = checkprn[:-4]
        array = array[k-1:]
        Flag = True
        array[0] = checkprn
    else:   
        array = array[k-2:]
    checkprn = array[0]
    
    k = 0
    while(checkprn[k] != '7' and k < len(checkprn)):
        k+=1
        
    checkprn = checkprn[k:]
    array[0] = checkprn
    if(Flag):
        array = array[26:]
    else:
        array = array[27:]
    #j = 1
    
    m = 0 
    cnt = 0
    while(array[m] != 'SEM.:2' and array[m] != 'SGPA1'):
        if(array[m+1] == '*'):
            m += 13
        else:
            m += 12
        cnt+=1
        #j += 1
        
    noofsub1 = cnt
    if(array[m] == 'SGPA1'):
        return noofsub1,0
    array = array[m+1:]
    #j = cnt + 1
    m = 0
    cnt = 0
    while(array[m] != 'FIRST' and array[m] != 'SECOND' and array[m] != 'THIRD' and array[m] != 'FOURTH' and array[m] != 'SGPA1'):
        if(array[m+1] == '*'):
            m += 13
        else:
            m += 12
        cnt+=1
        #j += 1
    noofsub2 = cnt
    
    return noofsub1,noofsub2
        
def writer(i,array,noofsub1,noofsub2,n1):
    k = 0
    j = 1
    #m = 0
    Flag = False
    checkprn = array[0]
    while(checkprn != 'PICT' and k < len(array) and not checkprn.endswith('PICT')):
        checkprn = array[k]
        k += 1
    if k == len(array):
        return 0,False
    if(checkprn.endswith('PICT') and len(checkprn)>4):
        checkprn = checkprn[:-4]
        array = array[k-1:]
        Flag = True
        array[0] = checkprn
    else:   
        array = array[k-2:]
    checkprn = array[0]
    
    k = 0
    while(checkprn[k] != '7' and k < len(checkprn)):
        k+=1
        
    checkprn = checkprn[k:]
    array[0] = checkprn

    worksheet.write(i,0,array[0])
    
    i=i+1
    
    if(Flag):
        array = array[3:]
    else:
        array = array[4:]
    j = 3
    #correct uptill here
    
    insertmarks1(i,j,array,noofsub1)
  
    m = noofsub1*12
        
    while(array[m] != 'SEM.:2' and array[m] != 'SGPA1'):
        m += 1
    flag = False
    if(array[m]=='SEM.:2'):
        array = array[m+1:]
        j = n1 + 3
        m,flag = insertmarks2(i,j,array,noofsub2)
    return m,flag
    
def insertmarks1(i,j,array,cnt1):
    m = 0 
    cnt = 0
    k = 0
    rowstart = i
    while(k < 6):
    	worksheet.write(i,1,array[k])
    	i += 1
    	k += 1
    array = array[23:]
    flag = False
    while(m < 134 and cnt!=cnt1):
    	i = rowstart
    	while(i - rowstart != 6):
    	    str = array[m]
    	    if(str[len(str)-1].isalpha() or str=="210250"):
    	        worksheet.write(rowstart+6,1,str)
    	        flag = True
            if(array[m+1] == '*'):
                m += 2
            else:
                m += 1
            if(flag):
                worksheet.write(rowstart+6,2,array[m+3])
                break
            else:
                worksheet.write(i,j,array[m])
            i += 1
        
        cnt+=1
        m += 6
        j += 1
        
def insertmarks2(i,j,array,cnt2):
    m = 0 
    cnt = 0
    rowstart = i
    flag = False
    while(m < 139 and cnt!=cnt2):
    	i = rowstart
    	while(i - rowstart != 6):
    	    str = array[m]
    	    if(str[len(str)-1].isalpha() or str=="214458" or str=="210258" or str=="107010"):
    	        worksheet.write(rowstart+7,1,str)
    	        flag = True
            if(array[m+1] == '*'):
                m += 2
            else:
                m += 1
            if(flag):
                worksheet.write(rowstart+7,2,array[m+3])
                break
            else:
                worksheet.write(i,j,array[m])
            i += 1
            
        cnt+=1
        m += 6
        j += 1
    return m,True

main()

workbook.close()
pdfFileObj.close()

