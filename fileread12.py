import PyPDF2
import xlsxwriter

workbook = xlsxwriter.Workbook('beit.xlsx')
worksheet = workbook.add_worksheet()

pdfFileObj = open('BEit.pdf','rb')

pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

def main2():
    pgno = 0
    i = 0
    num = pdfReader.getNumPages()
    print(num)
    pageObj = pdfReader.getPage(0)
    content = pageObj.extractText()
    contents = content.split()
    dept = 'none'
    code = ''
    if('B.E.' in contents[8]):
        if('COMPUTER' in contents[10]):
            dept = 'COMP'
        elif('ELECTRONICS' in contents[10]):
            dept = 'ETC'
        else:
            dept = 'IT'
    code = contents[8] + contents[9] + contents[10]
    bindex1 = 68
    newcontent = contents[68:]
    initialize2(newcontent) 
    while(pgno < num):        
        pageObj = pdfReader.getPage(pgno)
        content = pageObj.extractText()
        contents = content.split()
        newcontent = contents[bindex1:]
        
        i,m = writer2(i,newcontent,dept,code)
        newcontent = newcontent[m+6:]
        print(str(m) + 'value')
        i,m = writer2(i,newcontent,dept,code)
        newcontent = newcontent[m+6:]
        print(str(m) + 'value')
        if(m > 20):
            i,m = writer2(i,newcontent,dept,code)	
        
        pgno+=1
        
def initialize2(array):
    k = 0
    j = 1
    #m = 0
    worksheet.write(0,0,'PRN')
    worksheet.write(0,1,'SEMESTER')
    worksheet.write(0,2,'SUBJECT')
    #worksheet.write(0,3,'CATEGORY')
    worksheet.write(0,3,'MAXIMUM MARKS')
    worksheet.write(0,4,'MARKS OBTAINED')
    worksheet.write(0,5,'GRADE')
    worksheet.write(0,6,'CARRY OVER')
    worksheet.write(0,7,'SYLLABUS')
    
    '''checkprn = array[0]
    while(checkprn != ',PICT' and k < len(array) and not checkprn.endswith('PICT')):
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
    while(m < len(array) and array[m] != 'SEM.:2' and array[m] != 'SGPA1'):
        str = array[m]
        #if(not str[len(str)-1].isalpha()):
            #worksheet.write(0,j,array[m])
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
    while(m < len(array) and array[m] != 'FIRST' and array[m] != 'SECOND' and array[m] != 'THIRD' and array[m] != 'FOURTH' ):
        str = array[m]
        #if(not str[len(str)-1].isalpha() and str!="107010"):
            #worksheet.write(0,j,array[m])
        if(array[m+1] == '*'):
            m += 13
        else:
            m += 12
        cnt+=1
        j += 1
            
    noofsub2 = cnt
    return noofsub1,noofsub2'''
    

def writer2(i,array,dept,code):
    k = 0
    j = 1
    #m = 0
    Flag = False
    
    checkprn = array[0]
    while(checkprn != ',PICT' and k < len(array) and not checkprn.endswith(',PICT')):
        checkprn = array[k]
        k += 1
    if k == len(array):
        #print('returning midway')
        return i,k
    if(checkprn.endswith(',PICT') and len(checkprn)>5):
        checkprn = checkprn[:-4]
        array = array[k-1:]
        Flag = True
        array[0] = checkprn
    else:   
        array = array[k-2:]
    checkprn = array[0]
    
    k = 0
    while(k < len(checkprn) and checkprn[k] != '7'):
        k+=1
        
    checkprn = checkprn[k:]
    array[0] = checkprn
    print(checkprn)

    i += 1
    m = 2
    
    array = array[3:]
    count = 0
    sem = array[0]
    while(count < 10):
        check = array[m]
        while(len(check) != 6 and (not '310' in check) and (not '410' in check) and (not '404' in check) and (not '414' in check)):
            m += 1
            check = array[m]
        
        worksheet.write(i,0,checkprn)
    
        worksheet.write(i,1,sem)
        
        #worksheet.write(i,2,array[m])
        print(array[m] + ' ' + str(m))
        while(array[m] != 'PP' and array[m] != 'PR' and array[m] != 'OR' and array[m] != 'TW'):
            m += 1
        worksheet.write(i,2,check + '$' + array[m])
        #worksheet.write(i,3,array[m])
        m += 1
        worksheet.write(i,3,array[m])
        if(count >= 5):
            m += 2
        else:
            m += 4
        #if(array[m] == '--'):
        worksheet.write(i,4,array[m])
        m += 1
        worksheet.write(i,5,array[m])
        if(array[m+1] == 'C'):
            m += 1
            worksheet.write(i,6,array[m])  
        else:
            worksheet.write(i,6,'NA')
            m += 2
        worksheet.write(i,7,code)
        i += 1
        count += 1
        m += 7
    
    count = 0
    sem = array[1]
    m = 12
    if(dept == 'IT'):
        sub2 = 12
    else:
        sub2 = 10
    while(count < sub2):
        check = array[m]
        while(m < len(array)-1 and len(check) != 6 and (not '310' in check) and (not '410' in check) and (not '404' in check) and (not '414' in check)):
            m += 1
            check = array[m]
        
        worksheet.write(i,0,checkprn)
    
        worksheet.write(i,1,sem)
        
        #worksheet.write(i,2,array[m])
        #print(m)
        while(m < len(array) and array[m] != 'PP' and array[m] != 'PR' and array[m] != 'OR' and array[m] != 'TW'):
            m += 1
        if(m == len(array)):
            break
        worksheet.write(i,2,check + '$' + array[m])
        m += 1
        worksheet.write(i,3,array[m])
        if(dept == 'COMP' or dept=='ETC'):
            if(count >= 4):
                m += 2
            else:
                m += 4
        elif(dept == 'IT'):
            if(count == 0 or count == 1 or count == 2 or count == 5):
                m += 4
            else:
                m += 2
        else:
            if(count >= 5):
                m += 2
            else:
                m += 4
        worksheet.write(i,4,array[m])
        m += 1
        worksheet.write(i,5,array[m])
        if(array[m+1] == 'C'):
            m += 1
            worksheet.write(i,6,array[m])
        else:
            worksheet.write(i,6,'NA')
            m += 2
        worksheet.write(i,7,code)
        i += 1
        count += 1
        if(count < 10):
            m += 7
        else:
            m -= 1
    i -= 1
    
    return i,m
    

main2()

workbook.close()
pdfFileObj.close()

