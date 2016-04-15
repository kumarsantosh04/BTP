import urllib
import urllib.request
import re
import sys
import xlrd


def get_grade(rollno,subject):
        url = 'https://erp.iitkgp.ernet.in/StudentPerformance/view_performance.jsp?rollno='+rollno
        htmlfile = urllib.request.urlopen(url)
        htmlstring =str(htmlfile.read())
        #get name
        regex='Name</b></td><td>(.+?)</td>'
        pattern = re.compile(regex)
        name = re.search(pattern,htmlstring)
        name=name.group(1)
        #get CGPA
        regex = 'CGPA</b></td><td>(.+?)</td>'
        pattern = re.compile(regex)
        cgpa = re.findall(pattern,htmlstring)
        #format htmlstring
        regex= '<tr bgcolor="#FFF3FF">(.+?)</html>'
        pattern = re.compile(regex)
        htmlstring=re.findall(pattern,htmlstring)
        #find the subject
        regex ='(<td>'+subject+'.+?)</tr>'
        pattern = re.compile(regex)
        result =re.findall(pattern,str(htmlstring))
        if not result:
                print(rollno+"'s subject not found : "+subject)
                return
        
        #pattern='<td align="center">(.+?)</td>'
        pattern='<td(.+?)</td>'
        #print(result)
        print(rollno+' '+name+' CGPA: '+cgpa[2]+' : ')
        for i in range(len(result)):
                grade = re.findall(pattern, result[i])
                print('         '+grade[0]+' : credit : '+grade[-3][-1]+' '+grade[-2][-2:])
        return

def read_excel(keyword,subject):
        #change the file location to where you put the file
        file_loc ='./btp_evaluator_list_1st_september-1.xlsx'
        workbook = xlrd.open_workbook(file_loc)
        for i in range(4):
                sheet= workbook.sheet_by_index(i)
                for row in range(sheet.nrows):
                        
                        if sheet.cell_value(row,2)==keyword :
                                get_grade(sheet.cell_value(row,0),subject)
        return
        
def main():
        if len(sys.argv) != 4:
                print ('usage: BTP.py get_grade|btp rollno|keyword  subject')
                sys.exit(1)
        if sys.argv[1]=='get_grade':
                get_grade(sys.argv[2],sys.argv[3].upper())
        if sys.argv[1]=='btp':
                read_excel(sys.argv[2],sys.argv[3].upper())

if __name__ == '__main__':
        main()
