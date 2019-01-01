import PyPDF2
import xlsxwriter

workbook = xlsxwriter.Workbook('teetc2.xlsx')
worksheet = workbook.add_worksheet()

pdfFileObj = open('teetc.pdf','rb')

pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

syllabus = " "

def main():
    pgno = 0
    i = 0
    num = pdfReader.getNumPages()
    print(num)
    pageObj = pdfReader.getPage(0)
    content = pageObj.extractText()
    contents = content.split()
    bindex1 = 26
    syllabus = contents[17] + contents[18] + contents[19]
    print(contents[19])
    newcontent = contents[28:]
    n1,n2 = initialize(newcontent) 
    while(pgno < num):        
        pageObj = pdfReader.getPage(pgno)
        content = pageObj.extractText()
        contents = content.split()
        newcontent = contents[bindex1:]
        noofsub1,noofsub2 = NoOfSub(newcontent)
        m,i,flag = writer(i,newcontent,noofsub1,noofsub2,n1,syllabus)					#writing marks of 1st student
        
        if(flag):
            newcontent1 = newcontent[m+27:]
        else:
            newcontent1 = newcontent[m+17:]
        noofsub1,noofsub2 = NoOfSub(newcontent1)
        m,i,flag = writer(i,newcontent1,noofsub1,noofsub2,n1,syllabus)					#writing marks of 2nd student
        pgno+=1
        

def initialize(array):
    k = 0
    j = 1
    #m = 0
    worksheet.write(0,0,'PRN')
    worksheet.write(0,1,'SEMESTER')
    worksheet.write(0,2,'SUBJECT')
    worksheet.write(0,3,'CREDITS')
    worksheet.write(0,4,'MAXIMUM MARKS')
    worksheet.write(0,5,'MARKS OBTAINED')
    worksheet.write(0,6,'APPEARED')
    worksheet.write(0,7,'SYLLABUS')
    
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
    while(array[m] != 'FIRST' and array[m] != 'SECOND' and array[m] != 'THIRD' and array[m] != 'FOURTH' ):
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
        
def writer(i,array,noofsub1,noofsub2,n1,syllabus):
    k = 0
    j = 1
    #m = 0
    Flag = False
    checkprn = array[0]
    while(checkprn != 'PICT' and k < len(array) and not checkprn.endswith('PICT')):
        checkprn = array[k]
        k += 1
    if k == len(array):
        print('returning midway')
        return 0,i,False
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

    i=i+1
    
    if(Flag):
        array = array[3:]
    else:
        array = array[4:]
        
    #correct uptill here
    
    newarray = array[23:]
    
    k = 0
    category = []
    while(k < 7):
        category.append(array[k])
    	k += 1
    
    count = 0
    m = 0
    subcode1 = ""
    credit = '5'
    while(count < noofsub1):
        countpersub = 0
        subcode1 = newarray[m]
        if(newarray[m+1] == '*'):
            m += 2
            appeared = 'A'
        else:
            m += 1
            appeared = 'P'
        marker = True
        
        while(countpersub < 7):
       
            while(countpersub < 7 and (category[countpersub] == 'OE' or category[countpersub] == 'IN' or category[countpersub] == 'TH' or category[countpersub] == 'Tot%' or newarray[m]=='-------' or ((subcode1[len(subcode1)-1].isalpha() or subcode1=="214458" or subcode1=="210258" or subcode1=="107010" or subcode1=="210250") and (category[countpersub]=='Crd' or category[countpersub]=='Grd Pts' or category[countpersub]=='Crd Pts')))):
                m += 1
                countpersub += 1
            if(countpersub >= 7):
                m += 4
                break

            maxmark = []
            maxmark1 = []
            cat = False
            if('/' in newarray[m]):
                maxmark = newarray[m].split('/')
                if(category[countpersub] == 'TW' and ('/' in newarray[m+1] and (not 'AB' in newarray[m+1]))):
                    #print(checkprn)
                    maxmark1 = newarray[m+1].split('/')
                    if('#' in maxmark1[0] or '$' in maxmark1[0] or '!' in maxmark1[0]):
                        mark = maxmark1[0]
                        mark = mark[:-1]
                        maxmark1[0] = mark
                    max0 = int(maxmark[0]) + int(maxmark1[0])
                    max1 = int(maxmark[1]) + int(maxmark1[1])
                    maxmark[0] = str(max0)
                    maxmark[1] = str(max1)
                    cat = True 
               
            elif((subcode1[len(subcode1)-1].isalpha() or subcode1=="214458" or subcode1=="210258" or subcode1=="107010" or subcode1=="210250")):
                maxmark.append(newarray[m])
                maxmark.append('PP')
            '''elif(category[countpersub] == 'Tot%' or category[countpersub] == 'TW'):
                if(newarray[m]=='FF'):
                    maxmark.append('00')
                else:
                    maxmark.append(newarray[m])
                if((subcode1[len(subcode1)-1].isalpha() or subcode1=="214458" or subcode1=="210258" or subcode1=="107010" or subcode1=="210250") and marker):
                    maxmark.append('PP')
                else:
                    maxmark.append("100")
            elif(category[countpersub] == 'Grd'):
                maxmark.append(newarray[m])
                if((subcode1[len(subcode1)-1].isalpha() or subcode1=="214458" or subcode1=="210258" or subcode1=="107010" or subcode1=="210250") and marker):
                    maxmark.append('P')
                else:
                    maxmark.append('O')
            elif(category[countpersub] == 'Crd Pts'):
                maxmark.append(newarray[m])
                maxmark.append(str(10*int(newarray[m-3])))   #10*int()
            elif(category[countpersub] == 'Grd Pts'):
                maxmark.append(newarray[m])
                maxmark.append("10")'''
           
            
            worksheet.write(i,0,checkprn)
            sem = array[22]
    
            worksheet.write(i,1,sem)
    
            if(cat == True):
                worksheet.write(i,2,subcode1 + '$[TW+PR]')
            else:
                worksheet.write(i,2,subcode1 + '$' + category[countpersub])
            
            cnt = 0
            
            while(newarray[m] != '00' and newarray[m] != '01' and newarray[m] != '02' and newarray[m] != '03' and newarray[m] != '04' and newarray[m] != '05'):
               m += 1
               cnt += 1
    
            worksheet.write(i,3,newarray[m])
            m -= cnt
                
            worksheet.write(i,4,maxmark[1])
    
            worksheet.write(i,5,maxmark[0])
            
            worksheet.write(i,6,appeared)
            
            worksheet.write(i,7,syllabus)
            
            i += 1
            countpersub += 1
            if(cat == True):
                countpersub += 1
                m += 1
            if(countpersub == 7):
                m += 5
            else:
                m += 1
        count += 1
    
    m += 1
    count = 0
    
    subcode2 = ""
    sem = newarray[m-1]
    while(count < noofsub2):
        countpersub = 0
        subcode2 = newarray[m]
        if(newarray[m+1] == '*'):
            m += 2
            appeared = 'A'
        else:
            m += 1
            appeared = 'P'
        
        while(countpersub < 7):
       
            while(countpersub < 7 and (category[countpersub] == 'OE' or category[countpersub] == 'IN' or category[countpersub] == 'TH' or category[countpersub] == 'Tot%' or newarray[m]=='-------' or ((subcode2[len(subcode2)-1].isalpha() or subcode2=="214458" or subcode2=="210258" or subcode2=="107010" or subcode2=="210250") and (category[countpersub]=='Crd' or category[countpersub]=='Grd Pts' or category[countpersub]=='Crd Pts')))):
                m += 1
                countpersub += 1
            if(countpersub >= 7):
                m += 4
                break

            maxmark = []
            maxmark1 = []
            cat = False
            if('/' in newarray[m]):
                maxmark = newarray[m].split('/')
                if(category[countpersub] == 'TW' and ('/' in newarray[m+1] and (not 'AB' in newarray[m+1]))):
                    #print(checkprn)
                    maxmark1 = newarray[m+1].split('/')
                    if('$' in maxmark1[0] or '!' in maxmark1[0] or '#' in maxmark1[0]):
                        mark = maxmark1[0]
                        mark = mark[:-1]
                        maxmark1[0] = mark
                    max0 = int(maxmark[0]) + int(maxmark1[0])
                    max1 = int(maxmark[1]) + int(maxmark1[1])
                    maxmark[0] = str(max0)
                    maxmark[1] = str(max1)
                    cat = True
            elif((subcode2[len(subcode2)-1].isalpha() or subcode2=="214458" or subcode2=="210258" or subcode2=="107010" or subcode2=="210250")):
                maxmark.append(newarray[m])
                maxmark.append('PP')
            '''elif(category[countpersub] == 'Tot%' or category[countpersub] == 'TW'):
                if(newarray[m]=='FF'):
                    maxmark.append('00')
                else:
                    maxmark.append(newarray[m])
                if((subcode1[len(subcode1)-1].isalpha() or subcode1=="214458" or subcode1=="210258" or subcode1=="107010" or subcode1=="210250") and marker):
                    maxmark.append('PP')
                else:
                    maxmark.append("100")
            elif(category[countpersub] == 'Grd'):
                maxmark.append(newarray[m])
                if((subcode1[len(subcode1)-1].isalpha() or subcode1=="214458" or subcode1=="210258" or subcode1=="107010" or subcode1=="210250") and marker):
                    maxmark.append('P')
                else:
                    maxmark.append('O')
            elif(category[countpersub] == 'Crd Pts'):
                maxmark.append(newarray[m])
                maxmark.append(str(10*int(newarray[m-3])))   #10*int()
            elif(category[countpersub] == 'Grd Pts'):
                maxmark.append(newarray[m])
                maxmark.append("10")'''

            worksheet.write(i,0,checkprn)
    
            worksheet.write(i,1,sem)
    
            if(cat == True):
                worksheet.write(i,2,subcode2 + '$[TW+PR]')
            else:
                worksheet.write(i,2,subcode2 + '$' + category[countpersub])
            
            cnt = 0
            
            while(newarray[m] != '00' and newarray[m] != '01' and newarray[m] != '02' and newarray[m] != '03' and newarray[m] != '04' and newarray[m] != '05'):
               m += 1
               cnt += 1
    
            worksheet.write(i,3,newarray[m])
            m -= cnt
                
            worksheet.write(i,4,maxmark[1])
    
            worksheet.write(i,5,maxmark[0])
            
            worksheet.write(i,6,appeared)
            
            worksheet.write(i,7,syllabus)
            
            i += 1
            countpersub += 1
            if(cat == True):
                countpersub += 1
                m += 1
            if(countpersub == 7):
                m += 5
            else:
                m += 1
        count += 1  
        
    '''if(newarray[m] == 'FIRST' or newarray[m] == 'SECOND' or newarray[m] == 'THIRD'):
    	worksheet.write(i,0,checkprn)
    
        worksheet.write(i,1,newarray[m]+ " " +newarray[m+1])
        m+=2
        worksheet.write(i,2,newarray[m])
        m+=1
        worksheet.write(i,3," ")
                
        worksheet.write(i,4,"10")
        if(newarray[m+1] == '--,'):
            sgpa = '0'
        else:
            sgpa = newarray[m+1]
            sgpa = sgpa[:-1]
        worksheet.write(i,5,sgpa)
    else:
        m -= 1
        worksheet.write(i,0,checkprn)
        
        worksheet.write(i,1,'SEM1')
        
        worksheet.write(i,2,newarray[m])
        m+=2
        
        worksheet.write(i,3," ")
                
        worksheet.write(i,4,"10")
        if(newarray[m] == '--,'):
            sgpa = '0'
        else:
            sgpa = newarray[m]
            sgpa = sgpa[:-1]
        worksheet.write(i,5,sgpa)
        m -= 1
    
    i += 1
    worksheet.write(i,0,checkprn)
    worksheet.write(i,1,'TOTAL CREDITS EARNED')
    credits = newarray[m+6]
    k = 0
    while(k < len(credits) and (not credits[k].isalpha())):
        k += 1
    credits = credits[:k]
    worksheet.write(i,5,credits)
    i += 1'''
    
    #insertmarks1(i,j,array,noofsub1)
    i -= 1
    m = noofsub1*12
        
    while(array[m] != 'SEM.:2' and array[m] != 'SGPA1'):
        m += 1
    flag = False
    if(array[m]=='SEM.:2'):
        array = array[m+1:]
        j = n1 + 3
        #m,flag = insertmarks2(i,j,array,noofsub2)
    return m,i,flag
    
'''def insertmarks1(i,j,array,cnt1):
    m = 0 
    cnt = 0
    k = 0
    rowstart = i
    category = []
    while(k < 6):
        category.append(array[k])
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
    return m,True'''

main()

workbook.close()
pdfFileObj.close()

