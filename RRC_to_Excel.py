import xlrd
import xlsxwriter
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import os


x = []
Oil_Gas = []
Well_Name = []
Lease_Number = []
District = []
From_Month = []
From_Year = []
To_Month = []
To_Year = []
row_new_excel = []
header = []
current_lease_operator = []
field = []
lease = []
prod_month_range = []
sch_start = []
sch_remove = []
on_sch_flag = []
district = []
rrc_id = []
gas_well_number = []
sev_seal = []
lease_name = []
oil = ['Oil Lease', 'oil', 'Oil', 'O', 'o']
gas = ['Gas Well', 'Gas', 'G', 'g', 'gas']
counter = []
next = []
header2 = []


def site_login():
    global workbook
    workbook = xlsxwriter.Workbook(r'02-Python AppNew RRC File.xlsx')
    global k
    k = []
    global driver
    # options = webdriver.ChromeOptions()

    # options.add_argument('headless')
    # # set the window size
    # options.add_argument('window-size=1200x600')
    #
    # # initialize the driver

    driver = webdriver.Chrome()

    # driver = webdriver.Chrome(options=options)
    driver.get('https://webapps.rrc.texas.gov/security/login.do;jsessionid'
               '=bDcjYYAw9cXpON9IK2I2GiGBEuTCF35IX02l6AtB99ZQT-UJxKNY!-1603648408')
    driver.find_element_by_name('login').send_keys('Pptech')
    driver.find_element_by_name('password').send_keys('$Proline19')
    driver.find_element_by_xpath("//input[@name='submit' and @value='Submit']").click()
    driver.get("https://webapps.rrc.texas.gov/PR/prHomeAction.do?action=home")
    ###############################################################################################
    global loc
    # loc = r'C:\Users\mark.stanford\Desktop\RRC Production History App_V6 - Mark.xlsx'
    loc = r"C:\Users\mark.stanford\PycharmProjects\ProLine_Work\RRC\Excel_to_RRC_Login\RRC Production History App_V6 - Mark.xlsm"
    global wb
    wb = xlrd.open_workbook(loc, "Sheet 1")
    global sheet
    sheet = wb.sheet_by_index(0)

    global rowcount
    rowcount = sheet.nrows  # Get number of rows with data in excel sheet
    rowcount -= 2
    open_excel_and_extract_data()
    return


def open_excel_and_extract_data():
    for i in range(rowcount):
        Oil_Gas.append(sheet.cell_value(2 + i, 1))
        Well_Name.append(sheet.cell_value(2 + i, 2))
        Lease_Number.append(str(int(sheet.cell_value(2 + i, 3))))
        District.append((sheet.cell_value(2 + i, 4)))
        From_Month.append((sheet.cell_value(2 + i, 6)))
        From_Year.append(int(sheet.cell_value(2 + i, 7)))
        To_Month.append((sheet.cell_value(2 + i, 9)))
        To_Year.append(int(sheet.cell_value(2 + i, 10)))
        i += 1
    writing_to_website()
    return


def writing_to_website():
    driver.get('https://webapps.rrc.texas.gov/PR/queriesMainAction.do')

    select = Select(driver.find_element_by_name('district'))
    row_new_excel.clear()
    while (len(k)) <= (len(Oil_Gas) - 1):
        if Oil_Gas[(len(k))] in oil:
            driver.find_element_by_xpath("//input[@name='leaseType' and @value='Oil']").click()
            if len(Lease_Number[len(k)]) < 5:
                while len((Lease_Number[(len(k))])) < 5:
                    add_zeros = '0' + Lease_Number[len(k)]
                    Lease_Number[len(k)] = add_zeros
                driver.find_element_by_name('leaseNumber').send_keys((Lease_Number[(len(k))]))
            else:
                driver.find_element_by_name('leaseNumber').send_keys((Lease_Number[(len(k))]))
        elif Oil_Gas[(len(k))] in gas:
            driver.find_element_by_xpath("//input[@name='leaseType' and @value='Gas']").click()
            if len(Lease_Number[(len(k))]) < 6:
                while len(Lease_Number[len(k)]) < 6:
                    add_zeros = '0' + Lease_Number[len(k)]
                    Lease_Number[len(k)] = add_zeros
                driver.find_element_by_name('leaseNumber').send_keys((Lease_Number[(len(k))]))
            else:
                driver.find_element_by_name('leaseNumber').send_keys((Lease_Number[(len(k))]))
        else:
            driver.find_element_by_xpath("//input[@name='leaseType' and @value='Pending']").click()
        if District[(len(k))] == 1:
            select.select_by_visible_text('01')
            select.select_by_value('01')
        elif District[(len(k))] == 2:
            select.select_by_visible_text('02')
            select.select_by_value('02')
        elif District[(len(k))] == 3:
            select.select_by_visible_text('03')
            select.select_by_value('03')
        elif District[(len(k))] == 4:
            select.select_by_visible_text('04')
            select.select_by_value('04')
        elif District[(len(k))] == 5:
            select.select_by_visible_text('05')
            select.select_by_value('05')
        elif District[(len(k))] == 6:
            select.select_by_visible_text('06')
            select.select_by_value('06')
        elif District[(len(k))] == "6E":
            select.select_by_visible_text('6E')
            select.select_by_value('6E')
        elif District[(len(k))] == "7B":
            select.select_by_visible_text('7B')
            select.select_by_value('7B')
        elif District[(len(k))] == "7C":
            select.select_by_visible_text('7C')
            select.select_by_value('7C')
        elif District[(len(k))] == 8:
            select.select_by_visible_text('08')
            select.select_by_value('08')
        elif District[(len(k))] == '8A':
            select.select_by_visible_text('8A')
            select.select_by_value('8A')
        elif District[(len(k))] == 9:
            select.select_by_visible_text('09')
            select.select_by_value('09')
        elif District[(len(k))] == 10:
            select.select_by_visible_text('10')
            select.select_by_value('10')
        else:
            select.select_by_visible_text('None Selected')
            select.select_by_value('None Selected')

        driver.find_element_by_name('startMonth').send_keys(From_Month[(len(k))])
        driver.find_element_by_name('startYear').send_keys(From_Year[(len(k))])
        driver.find_element_by_name('endMonth').send_keys(To_Month[(len(k))])
        driver.find_element_by_name('endYear').send_keys(To_Year[(len(k))])
        driver.find_element_by_xpath("//input[@name='submit' and @value='Lease Query']").click()
        collect_data()
    return


def collect_data():
    html = driver.page_source
    soup = BeautifulSoup(html, features='lxml')
    table_body = soup.find('table', {'class': 'DataGrid'})
    rows = table_body.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        cols = [x.text.strip() for x in cols]
        row_new_excel.append(cols)
    while len(k) <= (len(Well_Name) - 1):
        html = driver.page_source
        soup = BeautifulSoup(html, features='lxml')
        words = soup.find(text='[ Next > ]')
        if words == '[ Next > ]':
            driver.find_element_by_link_text('[ Next > ]').click()
            collect_data()
        else:
            html = driver.page_source
            soup = BeautifulSoup(html, features='lxml')
            table_body = soup.find('table', {'cellpadding': '2', 'align': 'center', "width": '100%', 'border': '0',
                                             'cellspacing': '0'})
            rows = table_body.find_all('tr')
            for row in rows:
                cols = row.find_all('td')
                cols = [y.text.strip() for y in cols]
                header.append(cols)

            html = driver.page_source
            soup = BeautifulSoup(html, features='lxml')
            table_body = soup.find('table', {'class': 'TabBox2'})
            rows = table_body.find_all('tr')
            for row in rows:
                cols = row.find_all('td')
                cols = [y.text.strip() for y in cols]
                header2.append(cols)

            for i in range(1):
                current_lease_operator.append(header[i][1].strip("Current Lease Operator:\n"))
                sch_start.append(header[i][2].strip('Sch Start:\n'))
                district.append(header[i][3].strip('District:\n \n'))
                sev_seal.append(header[i][4].strip('Sev/Seal:\n      \n        \n'))
                field.append(header[i][5].strip('Field:\n '))
                if header[i][6].strip('Sch. Remove:') =="":
                    sch_remove.append("-")
                else:
                    sch_remove.append(header[i][6].strip('Sch. Remove:'))
                rrc_id.append(header[i][7].strip('RRC ID:\n'))
                lease_name.append(header[i][9].strip('Lease Name:\n'))
                on_sch_flag.append(header[i][10].strip('On Sch. Flag:\n'))
                gas_well_number.append(header[i][11].strip('Gas Well #:\n '))
                prod_month_range.append(header2[i][14].strip("Prod Month Range:\xa0\n "))
            header.clear()
            header2.clear()
            create_an_excel_file()
    return


def create_an_excel_file():
    row = 8
    col = 0
    df = [x for x in row_new_excel if x != []]

    merge_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '06C3FC', 'text_wrap': 1})
    header_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': 1,"fg_color":'C0C0C0'})
    format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': 1, "fg_color":'C0C0C0'})
    boarderless_format = workbook.add_format({
        'border': 0,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': 1,"fg_color":'000000'})
    cell_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': 'FFFFFF', 'text_wrap': 1})
    counter_for_well = 0

    worksheet = workbook.add_worksheet(Well_Name[len(k)])
    worksheet.freeze_panes(0,'A:K')
    worksheet.freeze_panes(1, 0)
    worksheet.freeze_panes(2, 0)
    worksheet.freeze_panes(3, 0)
    worksheet.freeze_panes(4, 0)
    worksheet.freeze_panes(5, 0)
    worksheet.freeze_panes(6, 0)
    worksheet.freeze_panes(7, 0)
    worksheet.freeze_panes(8, 0)
    worksheet.set_row(1, 22)
    worksheet.set_row(2, 22)
    worksheet.set_column("A:K", 10, cell_format=cell_format)
    worksheet.merge_range(5, 0, 5, 7, "Oil/Condensate (Whole Barrels)", merge_format)
    worksheet.merge_range(5, 8, 5, 10, "Gas/Casinghead Gas - MCF", merge_format)
    worksheet.merge_range(6, 0, 7, 0, "Multiple Reports", merge_format)
    worksheet.merge_range(6, 1, 7, 1, "Prod Month", merge_format)
    worksheet.merge_range(6, 2, 7, 2, "Commingle Permit No.", merge_format)
    worksheet.merge_range(6, 3, 7, 3, "On Hand Beginning of Month", merge_format)
    worksheet.merge_range(6, 4, 7, 4, "Production", merge_format)
    worksheet.merge_range(6, 5, 6, 6, "Disposition", merge_format)
    worksheet.write(7, 5, "Volume", merge_format)
    worksheet.write(7, 6, "Code", merge_format)
    worksheet.merge_range(6, 7, 7, 7, "On Hand End of Month", merge_format)
    worksheet.merge_range(6, 8, 7, 8, "Formation Production", merge_format)
    worksheet.merge_range(6, 9, 6, 10, "Disposition", merge_format)
    worksheet.write(7, 9, "Volume", merge_format)
    worksheet.write(7, 10, "Code", merge_format)
    counter_for_well += 1

    for j in range(len(df)):
        for r in range(len(df[j])):
            worksheet.write(row, col, df[j][r], cell_format)
            row += 0
            col += 1
        row += 1
        col = 0
    worksheet.merge_range(0, 10, 1, 10, "", boarderless_format)
    worksheet.merge_range(3, 10, 4, 10, "", boarderless_format)
    worksheet.merge_range(0, 0, 0, 1, 'Current Lease Operator: ', header_format)
    worksheet.merge_range(1, 0, 1, 1, "Field:", header_format)
    worksheet.merge_range(2, 0, 2, 1, "Lease Name:", header_format)
    worksheet.merge_range(3, 0, 3, 1, "Prod Month Range:", header_format)
    worksheet.merge_range(4, 0, 4, 1, "Sch Start:", header_format)
    worksheet.merge_range(0, 5, 0, 6, "Sch. Remove:", header_format)
    worksheet.merge_range(1, 5, 1, 6, "On Sch. Flag:", header_format)
    worksheet.merge_range(2, 5, 2, 6, "District:", header_format)
    worksheet.merge_range(3, 5, 3, 6, "RRC ID:", header_format)
    worksheet.merge_range(4, 5, 4, 6, "Gas Well #:", header_format)
    worksheet.write(2, 9, "Sev/Seal:", header_format)
    worksheet.merge_range(0,4,4,4,"",boarderless_format)
    worksheet.merge_range(0,9,1,9,"",boarderless_format)
    worksheet.merge_range(3,9,4,9,"",boarderless_format)

    worksheet.merge_range(0, 2, 0, 3, current_lease_operator[len(k)], format)
    worksheet.merge_range(1, 2, 1, 3, field[len(k)], format)
    worksheet.merge_range(2, 2, 2, 3, Well_Name[len(k)], format)
    worksheet.merge_range(3, 2, 3, 3, prod_month_range[len(k)], format)
    worksheet.merge_range(4, 2, 4, 3, sch_start[len(k)], format)
    worksheet.merge_range(0, 7, 0, 8, sch_remove[len(k)], format)
    worksheet.merge_range(1, 7, 1, 8, on_sch_flag[len(k)], format)
    worksheet.merge_range(2, 7, 2, 8, district[len(k)], format)
    worksheet.merge_range(3, 7, 3, 8, rrc_id[len(k)], format)
    worksheet.merge_range(4, 7, 4, 8, gas_well_number[len(k)], format)
    worksheet.write(2, 10, sev_seal[len(k)], format)





    k.append(1)
    writing_to_website()


site_login()
workbook.close()
os.startfile('02-Python AppNew RRC File.xlsx')