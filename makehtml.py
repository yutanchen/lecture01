from pandas import read_excel                                                                                                                
from datetime import date
from sys import argv 
doctors=['主治醫師',]
titles=['','']
mywidth=[]
#df = read_excel("temp.xlsx", names=NL, usecols="A:K")
df = read_excel(argv[1], names=titles, usecols="")

for i, row in df.iterrows():
  if row['doctor'] in doctors:
    print(row[''], row[''])
