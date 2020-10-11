from update import update
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import SessionNotCreatedException

def fetchTSV(username, password, download_path, webDriver_path):
  try:
    driver = webdriver.Chrome()
  except (SessionNotCreatedException) as ex:
    print("Chrome webDriver out of date. Updating now.")
    update.update_webDriver(download_path, webDriver_path)
    driver = webdriver.Chrome()


  driver.get("https://isupplier12.acuitybrandslighting.net/OA_HTML/AppsLocalLogin.jsp")

  #login
  print('Logging in')
  user = driver.find_element_by_name("usernameField")
  pswrd = driver.find_element_by_name("passwordField")
  user.clear()
  pswrd.clear()
  user.send_keys(username)
  pswrd.send_keys(password)
  submit_button = driver.find_element_by_xpath("//*[@id=\"ButtonBox\"]/button[1]")
  submit_button.click()

  #Wait for 'Delivery Schedule' link to exist, then click it.
  print('Navigating website')
  delivery_schedule = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "PosHpgGenericUrl3")))
  delivery_schedule.click()
  
  #Wait for 'GO' button to exist, then click it.
  search = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="DelivSchedSrchRN"]/tbody/tr[4]/td/table/tbody/tr/td/div/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[2]/button[1]')))
  search.click()

  #wait for data to populate on page.
  table_row = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="DelivSchedTblRN:Content"]/tbody/tr[3]')))
  
  #Click 'export' button.
  print('Requesting new tsv file')
  driver.find_element_by_id("ExportBtn").click()

  return True