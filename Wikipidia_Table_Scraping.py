from selenium import webdriver
import csv

driver = webdriver.Chrome("C:\\Users\\click\\PycharmProjects\\Scraping\\chromedriver.exe")
driver.get("https://en.wikipedia.org/wiki/Demographics_of_India")
Table_Name = driver.find_element_by_class_name('wikitable')
rows = Table_Name.find_elements_by_tag_name('tr')
print(type(rows))
print(rows)

output = []
rowspan = -1
for row in rows:
    output_row = []
    headings = row.find_elements_by_tag_name('th')
    cells = row.find_elements_by_tag_name('td')
    if headings:
        for num, cell in enumerate(headings):
            output_row.append(cell.text)
            if cell.get_attribute("colspan") == "2":
                output_row.append("")
            if cell.get_attribute("rowspan") == "2":
                rowspan = num
    if rowspan > -1:
        output_row.insert(num, "")
        rowspan = -1
    if cells:
        for cell in cells:
            output_row.append(cell.text)
    output.append(output_row)
print(output)

with open('output.csv','w', newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerows(output)

# row = Table_Name.find_element_by_xpath(".//tr[1]")
# header_1 = ([th.text for th in row.find_elements_by_xpath(".//th")])
# print(header_1)

