from pandas import read_excel                                                                                                                             
from datetime import date, datetime
from sys import argv
from fpdf import FPDF

# fill
doctors=['授課教師','王名儒', '郭光宇']
titles=['A','B','C','D','E','F','G','H','I','J','K']
mywidth={'A':15,'B':15,'C':15,'D':15,'E':15,'F':15,'G':15,'H':15,'I':15,'J':15,'K':15}
df = read_excel(argv[1], names=titles, usecols="A:K")


def print_line(y,row):
    x = 10
    for j, title in enumerate(titles):
        f.set_xy(x,y)
        if (isinstance(row[title], datetime)): row[title]=row[title].date()
        if (str(row[title])=='nan'): row[title]= ''
        f.cell(200, 20, str(row[title]), align = 'L')
        x += mywidth[title]


f = FPDF('P', 'mm', (210, 297))
#f.add_font('myF', '', 'DejaVuSansCondensed.ttf', uni=True)
f.add_page()
f.set_font('Times', 'B', 20)
f.cell(190, 15, 'Patient List ( '+str(date.today())+' )', align = 'C')

y, N, heads = 15, 0, []
f.add_font('myF', '', '/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc')
f.set_font('myF', '', 8)
for i, row in df.iterrows():
    if row['K'] in doctors:
        if (N%10==0 and N!=0):
            [f.set_xy(10,260), f.cell(190, 10, 'Page '+str(N//10), align = 'C')]
            f.add_page()
            y = 35
            print_line(y, heads)
        N += 1
        y += 20
        if (row['K']=='授課教師'): heads, N= row, N-1
        print_line(y,row)
        print(N,y)

[f.set_xy(10,260), f.cell(190, 10, 'Page '+str(N//10+1), align = 'C')]


f.output('tuto1.pdf')


