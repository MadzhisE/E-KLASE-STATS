from playwright.sync_api import sync_playwright
from io import StringIO
import xlsxwriter
import pandas as pd

PERSONAS_KODS = input("Ievadat savu personas kodu: ")
PAROLE = input("Ievadat savu paroli: ")

NERADIT_CHROME = True
ATZIMES_LINK = "https://my.e-klase.lv/Family/ReportPupilMarks/Get"

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

    for idx, col in enumerate(df):
        series = df[col]
        max_len = max((
        series.astype(str).map(len).max(),
            len(str(series.name))
            )) + 1
        worksheet.set_column(idx, idx, max_len)

    df.to_excel(writer, index=False)

    writer.close()

    browser.close()
