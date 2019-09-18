from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def fetchTSV(username, password):
  driver = webdriver.Chrome()
  driver.get("https://isupplier12.acuitybrandslighting.net/OA_HTML/AppsLocalLogin.jsp")

  #login
  user = driver.find_element_by_name("usernameField")
  pswrd = driver.find_element_by_name("passwordField")
  user.clear()
  pswrd.clear()
  user.send_keys(username)
  pswrd.send_keys(password)
  submit_button = driver.find_element_by_xpath("//*[@id=\"ButtonBox\"]/button[1]")
  submit_button.click()
  



  #click on the "Delivery Schedule" link
  try:
    delivery_schedule = WebDriverWait(driver, 10).Until(
      EC.element_to_be_clickable((By.ID, "PosHpgGenericUrl3"))
    )
  finally:
    pass

  delivery_schedule.click()


fetchTSV('frank@belairmfg.com', 'BelAir3525!2')