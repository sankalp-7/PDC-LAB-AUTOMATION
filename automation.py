import os, sys  
from docxtpl import DocxTemplate ,InlineImage
import pandas as pd  
from docx.shared import Cm,Inches,Mm,Emu
doc=DocxTemplate('template.docx')
df = pd.read_excel('excel_template.xlsx')
keys=df['key'].values
values = df['value'].values
context1 = {keys[i]: values[i] for i in range(len(keys))}
p1=InlineImage(doc,'placeholder1.png',Cm(16))
p2=InlineImage(doc,'placeholder2.png',Cm(16))
context1['pic1']=p1 
context1['pic2']=p2
s1=f"C:/Users/sodag/OneDrive/Pictures/Screenshots/{context1['code_pic']}"
s2=f"C:/Users/sodag/OneDrive/Pictures/Screenshots/{context1['code_output']}"
doc.render(context1)
doc.replace_pic("placeholder1.png",s1)
doc.replace_pic("placeholder2.png",s2)
doc.save('output/pdc_lab.docx')


