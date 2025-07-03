# importing required modules
from pypdf import PdfReader
import pandas as pd
import re

pdf = 'UoN _ Blue Castle _ My Marks _ My Transcript.pdf'

def wa2cgpa(wa):
    if wa >= 80:
        return 4
    elif wa >= 70:
        return 3
    elif wa >= 60:
        return 2
    elif wa >= 50:
        return 1

pattern = re.compile(r'''
    ^([A-Z]{4}\s?\d{4})        # course code  
    ([A-Za-z0-9&/:\-',.\s]+?)\s# title  
    (\d+)\s                    # credits
    ([A-Za-z]+(?:\s[A-Za-z]+)*)# semester
    \s([A-Za-z]+)              # country
    \s(\d+)?                   # mark
    $''', re.MULTILINE | re.VERBOSE)
year = []
marks = []
wa = []
cgpa = []

# creating a pdf reader object
reader = PdfReader(pdf)

for page in reader.pages:
    # extracting text from page
    text = page.extract_text().split('\n')
    # separate data by year
    for t in text:
        if len([s for s in t.split('/') if s.isdigit()]) == 2:
            year.append(t)  

# separate marks by academic year
text = "\n".join([page.extract_text() for page in reader.pages]).split(year[0])[1]
for t in year[1:]:
    marks.append(text.split(t)[0])
    text = text.split(t)[1]
marks.append(text)

# match the 
with pd.ExcelWriter('output.xlsx') as writer:  
    for i, m in enumerate(marks):
        d = {'Code':[], 'Title':[], 'Credit':[], 'Semester':[], 
             'Mark':[], 'Weighted Mark':[], 'CGPA':[], 'Weighted CGPA':[]}
        for mm in pattern.finditer(m.strip()):
            code, title, credit, sem, _, mark = mm.groups()
            code = code.replace(' ', '')   # collapse any internal space
            d['Code'].append(code)
            d['Title'].append(title)
            d['Credit'].append(int(credit))
            d['Semester'].append(sem)
            d['Mark'].append(int(mark))
            d['Weighted Mark'].append(int(credit)*int(mark))
            d['CGPA'].append(wa2cgpa(int(mark)))
            d['Weighted CGPA'].append(int(credit)*wa2cgpa(int(mark)))

        t_credit = sum(d['Credit'])
        t_wmarks = sum(d['Weighted Mark'])
        t_wcgpa = sum(d['Weighted CGPA'])
        wa.append(t_wmarks/t_credit) 
        cgpa.append(t_wcgpa/t_credit) 

        d['Code'].append("")
        d['Title'].append("")
        d['Credit'].append(t_credit)
        d['Semester'].append("")
        d['Mark'].append(t_wmarks/t_credit)
        d['Weighted Mark'].append(t_wmarks)
        d['CGPA'].append(t_wcgpa/t_credit)
        d['Weighted CGPA'].append(t_wcgpa)

        df = pd.DataFrame.from_dict(d)
        df.to_excel(writer, sheet_name=year[i].replace('/','-')) 


    if len(year) == 4:
        d = {}
        idx = []
        t_wa = 0 
        t_cgpa = 0 
        d['WA'] = wa
        d['CGPA'] = cgpa
        d['Percentage'] = [0.4, 0.4, 0.2, 0, None]
        for i in range(4):
            t_wa = t_wa + d['WA'][i]*d['Percentage'][i]
            t_cgpa = t_cgpa + d['CGPA'][i]*d['Percentage'][i]
            idx.append(f'Year {4-i}')
            
        d['WA'].append(t_wa)
        d['CGPA'].append(t_cgpa)

        df = pd.DataFrame.from_dict(d)
        idx.append("Overall")
        df.index = idx
        df.to_excel(writer, sheet_name="Overall") 

    print(df)

