import PyPDF2 as pdf
import pandas as pd
import os
import xlsxwriter.utility as xl_utility

def gradeFinder(str):
    grade = ''
    if 'A+' in str:
        grade = 'A+'
    elif 'A' in  str:
        grade = 'A'
    elif 'B' in  str:
        grade = 'B'
    elif 'C' in  str:
        grade = 'C'
    elif 'D' in  str:
        grade = 'D'
    elif 'E' in  str:
        grade = 'E'
    elif 'F' in  str:
        grade = 'F'
    elif 'X' in str:
        grade = 'X'
    return grade

def backlogCounter(li):
    backlog = 0
    for i in li:
        if i == 'F' or i == 'X':
            backlog = backlog + 1
    return backlog if backlog > 0 else '-'

def extract(path):
    file = open(path, 'rb')
    pdf_file = pdf.PdfReader(file).pages[0].extract_text()
    content = pdf_file.split()
    NAME = ' '.join(content[content.index('Name') + 2: content.index('Reg')])
    REG = content[content.index('Reg') + 2]
    OS = gradeFinder(content[content.index('System') + 1])
    EJP = gradeFinder(content[content.index('Java') + 2])
    PYT = gradeFinder(content[content.index('Python') + 2])
    WEB = gradeFinder(content[content.index('Technology') + 1])
    INF = gradeFinder(content[content.index('Security') + 1])
    GENERICNAME = ' '.join(pdf_file.split('\n')[32].split()[1:-2])
    GENERIC = gradeFinder(''.join(pdf_file.split('\n')[32].split()[-2]))
    SGPA = content[content.index('(%)') + 3]
    PERCENTAGE = '-' if SGPA == '-' else float(SGPA) * 10
    CGPA = content[content.index('brought') + 17]
    GRADE = gradeFinder(content[content.index('(%)') + 5][:-6])
    GRADElist = [OS, EJP, PYT, WEB, INF]
    BACKLOGS = backlogCounter(GRADElist)

    data = {
        'Reg': [REG],
        'Name': [NAME],
        'Operating Systems': [OS],
        'Enterprise Java Programming': [EJP],
        'Python Programming': [PYT],
        'Web Technology': [WEB],
        'Information Security': [INF],
        'Exercise is Medicine': [GENERIC if GENERICNAME == 'Exercise is Medicine' else None],
        'Basic Accounting': [GENERIC if GENERICNAME == 'Basic Accounting' else None],
        "India's Struggle for Freedom": [GENERIC if GENERICNAME == 'Indias Struggle for Freedom' else None],
        'Percentage': [PERCENTAGE],
        'SGPA': [float(SGPA) if SGPA != '-' else '-'],
        'CGPA': [float(CGPA) if CGPA != '-' else '-'],
        'Grade': [GRADE if GRADE else '-'],
        'S5 Backlog': [BACKLOGS]
    }
    return pd.DataFrame(data)

directory = './reservoir'
dfs = []
for filename in os.listdir(directory):
    if filename.endswith('.pdf'):
        pdfpath = os.path.join(directory, filename)
        df = extract(pdfpath)
        dfs.append(df)
finaldf = pd.concat(dfs, ignore_index=True)
finaldf = finaldf.drop_duplicates(subset='Reg', keep='first')
finaldf = finaldf.sort_values(by='Reg')
center_align = ['Operating Systems', 'Enterprise Java Programming', 'Python Programming', 'Web Technology', 'Information Security', 'SGPA', 'CGPA', 'Grade', 'S5 Backlog', 'Percentage', 'Exercise is Medicine','Basic Accounting',"India's Struggle for Freedom"]
column_widths = {'Reg': 12, 'Name': 24, 'Operating Systems': 10, 'Enterprise Java Programming': 10, 'Python Programming': 10, 'Web Technology': 10, 'Information Security': 10, 'Percentage': 10, 'SGPA': 10, 'CGPA': 10, 'Grade': 10, 'S5 Backlog': 10}
with pd.ExcelWriter('Analysis.xlsx', engine='xlsxwriter') as writer:
    finaldf.to_excel(writer, index=False, sheet_name='Sheet1', startrow=1, header=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    for col, width in column_widths.items():
        col_letter = xl_utility.xl_col_to_name(finaldf.columns.get_loc(col))
        worksheet.set_column(col_letter + ':' + col_letter, width)
    for col_num, value in enumerate(finaldf.columns.values):
        worksheet.write(0, col_num, value)
    for col_num, col_name in enumerate(finaldf.columns.values):
        if col_name in center_align:
            col_index = finaldf.columns.get_loc(col_name)
            worksheet.set_column(col_index, col_index, width, workbook.add_format({'align': 'center'}))
    for row_num, row in enumerate(finaldf.values, start=2):
        for col_num, value in enumerate(row):
            if value in ['F', 'X']:
                worksheet.write(row_num - 1, col_num, value, workbook.add_format({'bg_color': 'red', 'font_color': 'white', 'align': 'center'}))
