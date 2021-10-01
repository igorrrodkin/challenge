"""Template robot with Python."""
from RPA.Browser.Selenium import Selenium
import time
import os
from RPA.Excel.Files import Files


browser = Selenium()
excel = Files()

url = "https://itdashboard.gov/"
dive_in_button = "css=a.btn.btn-default.btn-lg-2x.trend_sans_oneregular"
rows_agencies = '//*[@id="agency-tiles-widget"]/div/div'
agencies_in_one_row = '//*[@id="agency-tiles-widget"]/div/div[agency_index]/div'
agency_link = "css=span.h4.w200"
table_elements_quantity = "css=select.c-select"
table_rows = '//*[@id="investments-table-object"]/tbody/tr'
agency_website = "https://itdashboard.gov/drupal/summary/005/"
pdf_download = "css=div#business-case-pdf a"
agency_name = '//*[@id="agency-tiles-widget"]/div/div[agency_row]/div[index_in_row]/div/div/div/div[1]/a/span[1]'
agency_price = '//*[@id="agency-tiles-widget"]/div/div[agency_row]/div[index_in_row]/div/div/div/div[1]/a/span[2]'
index_table_row = '//*[@id="investments-table-object"]/tbody/tr[index]/td'
table_cell_value = '//*[@id="investments-table-object"]/tbody/tr[index]/td[ind2]'


def open_browser():
    browser.open_available_browser(url)
    browser.click_element(dive_in_button)
    time.sleep(3)


def creating_workbook_and_worksheets():
    excel.create_workbook()
    excel.create_worksheet("agency")
    excel.create_worksheet("big_table")


def finding_agencies():
    rows_quantity = browser.get_webelements(rows_agencies)
    for agency_row in range(len(rows_quantity)):
        agency_row = agency_row + 1
        row_elements = browser.get_webelements(agencies_in_one_row.replace("agency_index", str(agency_row)))
        for index in range(len(row_elements)):
            index = index + 1
            agency = browser.get_text(
                agency_name.replace("agency_row", str(agency_row)).replace("index_in_row", str(index))
            )
            spending = browser.get_text(
                agency_price.replace("agency_row", str(agency_row)).replace("index_in_row", str(index))
            )
            dictionary = {"agency": agency, "amount": spending}
            excel.append_rows_to_worksheet(dictionary, "agency")
    # excel.save_workbook("output/Table.xlsx")


def working_with_one_agency():
    browser.execute_javascript("window.scrollTo(0,700)")
    browser.click_element(agency_link)
    time.sleep(10)
    browser.select_from_list_by_value(table_elements_quantity, "10")
    time.sleep(20)
    tramount = browser.get_webelements(table_rows)
    for row in range(len(tramount)):
        row = row + 1
        row_elements = browser.get_webelements(index_table_row.replace("index", str(row)))
        tableInvestment = []
        for row_index in range(len(row_elements)):
            row_index = row_index + 1
            value = browser.get_text(table_cell_value.replace("index", str(row)).replace("ind2", str(row_index)))
            tableInvestment.append(value)
        # print(tableInvestment)
        dictionary2 = {
            "2": tableInvestment[0],
            "3": tableInvestment[1],
            "4": tableInvestment[2],
            "5": tableInvestment[3],
            "6": tableInvestment[4],
            "7": tableInvestment[5],
            "8": tableInvestment[6],
        }
        excel.append_rows_to_worksheet(dictionary2, "big_table")
    excel.save_workbook("output/Tables.xlsx")


def reading_worksheet():
    excel.open_workbook("output/Tables.xlsx")
    links = excel.read_worksheet_as_table()
    excel.close_workbook()
    return links


def downloading_pdf(links):
    for link in links:
        browser.set_download_directory(os.path.abspath(os.curdir) + "/output")
        try:
            browser.open_available_browser(agency_website + str(link["A"]))
            browser.click_element(pdf_download)
            time.sleep(6)
        except:
            browser.close_browser()


if __name__ == "__main__":
    creating_workbook_and_worksheets()
    open_browser()
    finding_agencies()
    working_with_one_agency()
    downloading_pdf(reading_worksheet())
