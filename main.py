import openpyxl
from selenium import webdriver
from bs4 import BeautifulSoup
import time

# Taking inputs
old = input('Enter the name of the excel file to be modified: ')

# Loading the file
wb = openpyxl.load_workbook(old)
ws = wb.worksheets[0]

# Deleting first 3 rows
ws.unmerge_cells('A1:C1')
ws.unmerge_cells('A3:F3')
ws.delete_rows(1, 3)

# Inserting 23 columns at the beginning
ws.insert_cols(1, 23)

# Calculating no of rows
rows = ws.max_row

# Shifting the columns
ws.move_range(cell_range='AR1:AR' + str(rows), cols=-43)
ws.cell(1, 1).value = 'Store'                             # Column A
ws.move_range(cell_range='AC1:AC' + str(rows), cols=-27)  # Column B
ws.move_range(cell_range='AK1:AK' + str(rows), cols=-34)  # Column C
ws.move_range(cell_range='AS1:AS' + str(rows), cols=-41)
ws.cell(1, 4).value = 'Customer Paid'                     # Column D
ws.move_range(cell_range='AJ1:AJ' + str(rows), cols=-31)
ws.cell(1, 5).value = 'Postage Paid'                      # Column E
ws.move_range(cell_range='AT1:AT' + str(rows), cols=-40)
ws.cell(1, 6).value = 'Profit/Unit'                       # Column F
ws.move_range(cell_range='AQ1:AQ' + str(rows), cols=-36)  # Column G
ws.move_range(cell_range='AB1:AB' + str(rows), cols=-20)  # Column H
ws.move_range(cell_range='AA1:AA' + str(rows), cols=-18)  # Column I
ws.move_range(cell_range='AG1:AG' + str(rows), cols=-23)  # Column J
ws.move_range(cell_range='AD1:AD' + str(rows), cols=-19)  # Column K
ws.move_range(cell_range='Z1:Z' + str(rows), cols=-14)    # Column L
ws.move_range(cell_range='AE1:AE' + str(rows), cols=-18)  # Column M
ws.move_range(cell_range='AF1:AF' + str(rows), cols=-18)  # Column N
ws.move_range(cell_range='X1:X' + str(rows), cols=-9)     # Column O
ws.move_range(cell_range='AH1:AH' + str(rows), cols=-18)  # Column P
ws.move_range(cell_range='Y1:Y' + str(rows), cols=-8)     # Column Q
ws.move_range(cell_range='AI1:AI' + str(rows), cols=-17)  # Column R
ws.move_range(cell_range='AL1:AL' + str(rows), cols=-19)  # Column S
ws.move_range(cell_range='AM1:AM' + str(rows), cols=-19)  # Column T
ws.move_range(cell_range='AN1:AN' + str(rows), cols=-19)  # Column U
ws.move_range(cell_range='AO1:AO' + str(rows), cols=-19)  # Column V
ws.move_range(cell_range='AP1:AP' + str(rows), cols=-19)  # Column W

# Deleting remaining columns
ws.delete_cols(47, 2)

# Loading the Webdriver
driver = webdriver.Chrome('chromedriver.exe')

# Log in
driver.get('https://ss.shipstation.com/')
input('Hit "ENTER" after you have successfully logged in: ')
driver.get('https://ship6.shipstation.com/orders/awaiting-shipment')
time.sleep(30)
driver.find_elements_by_class_name('advanced-search-text-6ODW4Fd')[0].click()

# Filling the columns
# Column D
i = 2
while i <= rows:
    driver.find_elements_by_class_name('flex-input-3IbU2o7')[0].clear()
    driver.find_elements_by_class_name('flex-input-3IbU2o7')[0].send_keys(str(ws.cell(i, 2).value))
    date_old = str(ws.cell(i, 12).value)[0:10]
    date_new = date_old[5:7] + '/' + date_old[8:10] + '/' + date_old[0:4]
    driver.find_elements_by_class_name('full-width-input-3F2-Knx')[0].clear()
    driver.find_elements_by_class_name('full-width-input-3F2-Knx')[0].send_keys(date_new)
    driver.find_elements_by_id('advanced-search-submit-button')[0].click()
    time.sleep(1)
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    try:
        store = soup.select('.message-Wl5p7I-')[0].getText()
        ws.cell(i, 1).value = str(store)  # Column A
        total = soup.select('.currency-column-value-1VfNyxR')[0].getText()
        ws.cell(i, 4).value = str(total)  # Column D
        ws.cell(i, 4).value = '$' + str(float(ws.cell(i, 4).value[1:]) - ws.cell(i, 5).value)  # Column F
        driver.find_elements_by_class_name('checkbox-CUegX-s')[0].click()
        driver.find_element_by_xpath("//div[contains(text(), 'Edit Tags')]").click()
        driver.find_element_by_xpath("//span[contains(text(), 'AUDITED')]").click()
        time.sleep(1)
    except:
        i += 1
        continue
    i += 1

# Save the file
wb.save('New.xlsx')
