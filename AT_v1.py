import xlwt,xlrd
from xlutils.copy import copy
print('WELCOME ROHAN JALI')
res=input('Are you Entering Here First time? Y/N:-')
if res=='y':
    workbook=xlwt.Workbook(encoding='utf-8')
    sheet1=workbook.add_sheet('Attendence sheet')
    sheet1.write(0,0,'COURSE')
    nest=[x for x in input('Enter the courses you study:').split()]
    n=len(nest)
    for i in range(1,n+1):
      # st=input('Enter the %d course name:'%(i))
        sheet1.write(i,0,nest[i-1])
        workbook.save('ROHAN JALI.xls')
    s=int(input('Enter the number of courses you registered:'))
    book=xlrd.open_workbook("ROHAN JALI.xls")
    wb=copy(book)
    w_sheet=wb.get_sheet(0)
    w_sheet.write(10,0,s)
    wb.save('ROHAN JALI.xls')


book=xlrd.open_workbook("ROHAN JALI.xls")
fsheet=book.sheet_by_index(0)
jali=(fsheet.cell(10,0).value)
s=int(jali)





d=input('Enter the date DD/MM:')
loc="ROHAN JALI.xls"
book=xlrd.open_workbook(loc)
fsheet=book.sheet_by_index(0)
lst=fsheet.row_values(0)
c=len(lst)
wb=copy(book)
w_sheet=wb.get_sheet(0)
w_sheet.write(0,c,d)
wb.save('ROHAN JALI.xls')


nest=[]
book=xlrd.open_workbook("ROHAN JALI.xls")
fsheet=book.sheet_by_index(0)
for i in range(1,s+1):
    nest.append((fsheet.cell(i,0).value))



    
for i in range(1,s+1):
    at=input('Enter the %s sub status:'%(nest[i-1]))
    book=xlrd.open_workbook(loc)
    wb=copy(book)
    w_sheet=wb.get_sheet(0)
    w_sheet.write(i,c,at)
    wb.save('ROHAN JALI.xls')

print('Updated Successfully :) ')
ch=input('Do you want to check your attendance status [Y/N]:')
print('---------------------------------------------------------')

if ch=='Y'or ch=='y':
    k=b=f=w=0
    for i in range(1,s+1):
        book=xlrd.open_workbook(loc)
        fsheet=book.sheet_by_index(0)
        lst=fsheet.row_values(i)
        z=(len(lst))-1
        for j in range(z+1):
            soh=(fsheet.cell(i,j).value)
            if soh=='P'or soh=='p':
                k+=1
            if soh=='-':
                f+=1
            
            if soh=='E'or soh=='e':
                b+=2
        if b>0:
            w=b-1
        else:
            w=0
                
        
        print('For %s course num of attendance is %f%%:'%(nest[i-1],((k+b)/(z-f+w))*100))
        if ((k+b)/(z-f+w))*100<85:
            lov=int(input('Enter the total number of classes for %s course:'%(nest[i-1])))
            x=((z-f+w)*85-(k+b)*100)/15
            if x<lov-(z-f+w):
                 print('You need to attend %f more class to overcome NSAR'%(x))
            else:
                print('You Will get NSAR')
                pm=lov-(z-f+w)
                k=(k+b)+pm
                z=(z-f+w)+pm
                print('Your final attendence status is %f %%'%((k/z)*100))
        else:
            print('Yor are in safe zone :D')
        
        k=b=f=w=0
        print('----------------------------------------------------------')
        

        # HI Rohan

    
