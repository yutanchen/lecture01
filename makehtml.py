from pandas import read_excel                                                                                                                             
from datetime import date
from sys import argv

# fill
doctors=['授課教師','王名儒']
titles=['A','B','C','D','E','F','G','H','I','J','K']
mywidth={'A':50,'B':100,'C':30,'D':30,'E':30,'F':30,'G':30,'H':30,'I':30,'J':30,'K':50}
df = read_excel(argv[1], names=titles, usecols="A:K")

f=open('out.html',encoding='utf-8',mode ='w')
f.write('<!DOCTYPE html><html><body style="margin 0px 0px;">\n')
f.write('<h1 style="text-align:center;">Patient List ('+str(date.today())+')</h1>\n')
f.write('<table style="font-size:8px; line-height:50px;">\n')
for i, row in df.iterrows():
    if row['K'] in doctors:
        f.write('<tr>')
        for j in ['A','B','K']:
            if (isinstance(row[j], datetime)): row[j]=row[j].date()
            if (str(row[j])=='nan'):f.write('<td width=' + str(mywidth[j]) + '> </td>')
            else:f.write('<td width=' + str(mywidth[j]) + '>' + str(row[j]) +'</td>')
        f.write('</tr>\n')
f.write('</table></body></html>')
