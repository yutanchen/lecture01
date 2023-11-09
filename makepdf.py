from pandas import read_excel                                                                                                                             
from datetime import date, datetime
from sys import argv
from fpdf import FPDF

doctors=['王名儒', '郭光宇'] 
titles=['A','B','C','D','E','F','G','H','I','J','K']
mywidth={'A':15,'B':15,'C':15,'D':15,'E':15,'F':15,'G':15,'H':15,'I':15,'J':15,'K':15}
df = read_excel(argv[1], names=titles, usecols="A:K")

class PDF(FPDF):
    def header(self):
        self.set_font_size(8)
        [self.set_xy(-30,+5), self.cell(30, 10, self.author, align = 'C')]
        [self.set_xy(-30,+10), self.cell(30, 10, str(date.today()), align = 'C')]
        self.set_font_size(20)
        [self.set_xy(+0,+0), self.cell(0, 20, self.title, align = 'C')]
    def footer(self):
        self.set_y(-15) 
        self.set_font('Times', '', 8) 
        self.cell(0, 10, 'Page %s' % self.page_no(), align = 'C')
    def print_line(self, y, row):
        x = 10 
        for j, title in enumerate(titles):
            if (isinstance(row[title], datetime)): row[title]=row[title].date()
            if (str(row[title])=='nan'): row[title]= ''
            self.set_xy(x,y)
            self.cell(200, 20, str(row[title]), align = 'L')
            x += mywidth[title]

f = PDF()
f.set_title('Patient List') # title
f.set_author('Jules Verne') # author
f.add_font('myF', '', '/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc') 
f.set_font('myF', '', 8)

y0, N = 40, 0 # y0: y of first content (special titles excluded)
for i, row in df.iterrows():
    if row['K']=='授課教師':SP=row # change 'K', '授課教師'
    elif row['K'] in doctors: # change 'K'
        if N%15==0: # 15: 15 lines per page
            f.add_page()
            f.print_line(25, SP) # 25: y of special titles
            y = y0
        f.print_line(y, row)
        y += 15 # 15: height of each line
        N += 1
f.output('out.pdf')


