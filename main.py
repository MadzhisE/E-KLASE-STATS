from playwright.sync_api import sync_playwright
from io import StringIO
import xlsxwriter
import pandas as pd
import re

PERSONAS_KODS = input("Ievadat savu personas kodu: ")
PAROLE = input("Ievadat savu paroli: ")

NERADIT_CHROME = True 
ATZIMES_LINK = "https://my.e-klase.lv/Family/ReportPupilMarks/Get"

def setwidthforcolumns(df):
        for idx, col in enumerate(df):  # loop through all columns
            series = df[col]
            max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
                len(str(series.name))  # len of column name/header
                )) + 1  # adding a little extra space
            worksheet.set_column(idx, idx, max_len)  # set column width

with sync_playwright() as p:
    browser = p.chromium.launch(headless=NERADIT_CHROME)
    page = browser.new_page()
    page.goto("https://www.e-klase.lv/")
    persontextbox = page.locator("(//input[@placeholder='Lietotājvārds'])[2]")
    passtextbox = page.locator("(//input[@placeholder='Parole'])[2]")
    persontextbox.fill(PERSONAS_KODS)
    passtextbox.fill(PAROLE)
    page.locator("(//button[@type='submit'])[3]").click()
    page.goto(ATZIMES_LINK)

    workbook = xlsxwriter.Workbook("Atzimes.xlsx")
    workbook.close()
    
    dfs = pd.read_html(StringIO(page.content()))
    df = pd.concat(dfs)

    writer = pd.ExcelWriter("Atzimes.xlsx", engine='xlsxwriter')
    worksheet = writer.book.add_worksheet('Sheet1')
    
    CountGrade = 0
    SumGrade = 0
    MidGradesPerClass = []

    for row_idx, row in df.iterrows():
        ClassSum = 0
        ClassCount = 0
        ClassMidGrade = 0

        for col_idx_month, value in row.items():
            if pd.notna(value):
                if isinstance(value, str):
                    matches = re.findall(r'(\d+)\(p\.d\.\)', value) # extracto visus value ar (p.d.) blakus
                    for match in matches:
                        #print(row, match)
                        ClassCount += 1
                        ClassSum += int(match) #klases vid atzime
                        CountGrade += 1
                        SumGrade += int(match) # vispareja vid atzime

        if ClassCount == 0:
            ClassMidGrade = "No grades given"
        else:
            ClassMidGrade = str(ClassSum/ClassCount)
        
        MidGradesPerClass.insert(int(row_idx), ClassMidGrade)
        
    df["Vidējās atzīmes"] = MidGradesPerClass

    df.loc[len(df), "Vidējās atzīmes"] = str("Vispārējā vidējā atzīme: "+str(round(SumGrade/CountGrade, 2))) #pievieno visparejo gada atzimi

    setwidthforcolumns(df)

    df.to_excel(writer, index=False)

    writer.close()

    browser.close()
