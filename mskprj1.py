import xlrd
import re
from urllib.request import *
from bs4 import BeautifulSoup
import xlsxwriter
import sqlite3
from xlsxwriter.workbook import Workbook

data1=xlrd.open_workbook("F:\\niit\msk.xlsx")
vel=data1.sheet_by_index(0)

tbl=[]
tbl=vel.col_slice(colx=0,start_rowx=1,end_rowx=8)
reg=re.compile('text:')
sw=[]
for i1 in tbl:
    s1=i1.value
    sw.append(re.sub(reg,'',s1))
print(sw)
velLink=(vel.cell(0,4))
Link1=velLink.value
velurl=re.sub(reg,'',Link1)
print(velurl)
class KeywordSearch:
    "Class for web scrapping"
    def __init__(self, url=""):
        self.url=url
    def PageSearch(self):
        "Function for getting content of url"
        self.req=Request(self.url,data=None,headers={'User-Agent':'Mozilla/5.0 (Macintosh;Intel Mac OS X 10_9_3)AppleWebKit/537.36 (KHTML, like Gecko) Chrome/35.0.1916.47 Safari/537.36'})
        self.f=urlopen(self.req)
        self.soup =BeautifulSoup(self.f,'html.parser')
        self.words=self.soup.findAll(text=True)
        return self.words
    def exclude(self,element):
        "Function for getting text from the content"
        if element.parent.name in ['style','script','[document]','head','title']:
            return False
        elif re.match('<!--.*-->',str(element)):
            return False
        return True
    def SearchCount(self,word):
        "Function for counting word occured in the text"
        tags=self.PageSearch()
        ExcludeTags=list(filter(self.exclude,tags))
        wordcount=[]
        for j in ExcludeTags:
                if(j.count(word)!=0):
                    wordcount.append(j.count(word))
        total=0
        for k in wordcount:
            total+=k
        return word,total
K1=KeywordSearch(velurl)
FK=dict(map(K1.SearchCount,sw))
print("word count",FK)
totalFreq1=0
for l1 in FK.values():
    totalFreq1+=l1
velout={ll1:{'count':FK.get(ll1),'frequency': FK.get(ll1)/totalFreq1*100}for ll1 in FK.keys()}
flag=0
while(flag==0):
    try:
        flag=1
        conn=sqlite3.connect('TEST1.db')     
        conn.execute("CREATE TABLE velammal (WORD TEXT NOT NULL, COUNT INT NOT NULL,FREQUENCY FLOAT NOT NULL);")
        print("output table created")
    except sqlite3.OperationalError:
        conn.execute("DROP TABLE IF EXISTS VELAMMAL")
        flag=0
for m1 in velout.keys():
    conn.execute("""INSERT INTO velammal(WORD,COUNT,FREQUENCY) VALUES(?,?,?)""",(m1,velout[m1]['count'],velout[m1]['frequency']))
velammal=conn.execute("SELECT * FROM velammal")
for n1 in velammal:
    print("Data input ",n1)
workbook = Workbook('F:\\niit\msk.xlsx')
worksheet1= workbook.add_worksheet("vel")
def Format(worksheet):
    worksheet.write("A1","Words")
    worksheet.write("B1","Count")
    worksheet.write("C1","Frequency")
    worksheet.write("E1",velurl)
Format(worksheet1)
c=conn.cursor()
c.execute("select * from velammal")
Flip=c.execute("select * from velammal")
for i,row in enumerate(Flip):
    print (row)
    i=i+1
    worksheet1.write_row(i,0,row)
chart1=workbook.add_chart({'type':'pie'})
chart1.add_series({'name':'FlipKart Search','categories': '=vel!$A2:$A5','values':'=vel!$C2:$C5','data_labels':{'value':True}})
chart1.set_title({'name':'Velammal Search',})
chart1.set_style(10)
worksheet1.insert_chart("H6",chart1, {'x_offset': 25, 'y_offset': 10})


workbook.close()
conn.commit()
