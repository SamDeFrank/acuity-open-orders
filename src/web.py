from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def fetchTSV(username, password):
  driver = webdriver.Chrome()
  driver.get("https://isupplier12.acuitybrandslighting.net/OA_HTML/AppsLocalLogin.jsp")

  #login
  print('Logging in\n')
  user = driver.find_element_by_name("usernameField")
  pswrd = driver.find_element_by_name("passwordField")
  user.clear()
  pswrd.clear()
  user.send_keys(username)
  pswrd.send_keys(password)
  submit_button = driver.find_element_by_xpath("//*[@id=\"ButtonBox\"]/button[1]")
  submit_button.click()

  #Wait for 'Delivery Schedule' link to exist, then click it.
  print('Navigating website\n')
  delivery_schedule = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PosHpgGenericUrl3")))
  delivery_schedule.click()
  
  #Wait for 'GO' button to exist, then click it.
  search = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="DelivSchedSrchRN"]/tbody/tr[4]/td/table/tbody/tr/td/div/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/button[1]')))
  search.click()

  #wait for data to populate on page.
  table_row = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="DelivSchedTblRN:Content"]/tbody/tr[3]')))
  
  #Click 'export' button.
  print('Requesting new tsv file\n')
  driver.find_element_by_id("ExportBtn").click()

  return True