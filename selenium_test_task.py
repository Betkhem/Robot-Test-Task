# Selenium
# https://sites.google.com/a/chromium.org/chromedriver/downloads
# chromedriver -v == 95...
#chrome -v == 95...
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import os

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Writing to an excel 
# sheet using Python
import xlwt
from xlwt import Workbook



PATH = "C:\Program Files (x86)\chromedriver.exe"

#changing downloads path to output folder
chrome_options = webdriver.ChromeOptions()
goal_dir = os.path.join(os.getcwd(), "output")
prefs = {'download.default_directory' : os.path.normpath(goal_dir)}
chrome_options.add_experimental_option('prefs', prefs)
driver = webdriver.Chrome(PATH, chrome_options=chrome_options)


driver.get("https://itdashboard.gov")

print(driver.title)

time.sleep(1)
wb = Workbook() #excel file config
DIVE_IN_LINK_TEXT = driver.find_element_by_xpath('//*[@id="node-23"]/div/div/div/div/div/div/div/a')
DIVE_IN_LINK_TEXT.send_keys(Keys.RETURN)
agencies = driver.find_element_by_xpath('//*[@id="agency-tiles-widget"]/div')

li_es1 = agencies.find_elements_by_tag_name("div")

ls = []
for item in li_es1:
	span = item.find_elements_by_tag_name("span")
	for text in span:
		ls.append(text.get_attribute("innerHTML"))


agencies1 = wb.add_sheet('agencies 1')
for i in range(len(ls)):
	if i % 2==0:
		agencies1.write(i, 0, ls[i])
	else:
		agencies1.write(i, 1, ls[i])
#write down to excel all agencies and spending

agencies_table = wb.add_sheet('agencies table')

#choose agency
Agency = WebDriverWait(driver, 10).until(
	EC.presence_of_element_located((By.XPATH, '//*[@id="agency-tiles-2-widget"]/div/div[8]/div[3]/div/div/div/div[2]/a'))
)
Agency.send_keys(Keys.RETURN)

content = WebDriverWait(driver, 10).until(
	EC.presence_of_element_located((By.NAME, 'investments-table-object_length'))
)

content.send_keys(Keys.RETURN)
options = content.find_elements_by_tag_name('option')
options[-1].click()
#content variable stands for table length, which is set to "All"

time.sleep(10)
Table = WebDriverWait(driver, 10).until(
	EC.presence_of_element_located((By.XPATH, '//*[@id="investments-table-object"]/tbody'))
)
tr = driver.find_elements_by_xpath ('//*[@id="investments-table-object"]/tbody/tr')
for i in range(len(tr)):
	k = i
	agencies_table.write(i, 0, tr[i].text)

dest_filename = 'Agencies.xls'
wb.save(os.path.join('output', dest_filename))
#write down to new excell sheet all table data

links = [td.find_elements_by_tag_name('a') for td in tr if td.find_elements_by_tag_name('a') != []]

for i in range(len(links)):
	time.sleep(3)
	driver.execute_script("window.open('https://itdashboard.gov/drupal/summary/{a}/{b}/','new_window')".format(a = links[0][0].get_attribute("innerHTML")[:3], b = links[i][0].get_attribute("innerHTML")))
	time.sleep(2)
	driver.switch_to.window(driver.window_handles[1])
	link = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.LINK_TEXT, 'Download Business Case PDF'))
	)
	link.send_keys(Keys.RETURN)
	driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL+'w')
	driver.switch_to.window(driver.window_handles[0])
#download all pdf files
#all files are stored in "output" folder

time.sleep(5)

driver.close()

#driver.quit() to close entire browser
