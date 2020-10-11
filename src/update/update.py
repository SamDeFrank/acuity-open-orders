import os
import bs4
import asyncio
import zipfile
import subprocess
import urllib.request

URL_FRONT = "https://chromedriver.storage.googleapis.com/"
URL_BACK = "chromedriver_win32.zip"
FILE_NAME = "webDriver.zip"
ZIP_EXTRACT_PATH =  r"C:\Users\Sam\Documents\code\work\WebDrivers"

#grabs current version of chrome
async def get_chrome_version():
    print("Checking local Chrome Version")
    local_machine = {}
    current_user = {}

    output = subprocess.Popen((os.path.dirname(__file__) + "\get_chrome_version.bat"),stdout=subprocess.PIPE).stdout
    output = output.read().decode().split("\r\n"*5)
    output = [x.replace("\r\n", "") for x in output]
    output = [x.split(" "*4) for x in output]
    
    for x in output:
      if "HKEY_LOCAL_MACHINE" in x[0]:
        local_machine.update({x[1]: x[3]})
      else:
        chrome_info['current_user'][x[1]] = x[3]

    if 'pv' in current_user:
      return current_user['pv']
    else:
      return local_machine['pv']


#use web scraping to find most recent chrome driver that matches local chrome version
#this only works because the first link on the page that contains the correct major version is the one we want
#if the web page were to be re-designed, this will most likely break
async def get_latest_webdriver_version():

    task1 = asyncio.create_task(get_chrome_version())

    page = urllib.request.urlopen("https://chromedriver.chromium.org/downloads")
    soup = bs4.BeautifulSoup(page, features="html.parser")
    links = [x.get('href') for x in soup.find_all('a') if isinstance(x.get('href'), str)]

    await task1
    
    chrome_version = task1.result().split(".")[0]
    links = [link for link in links if chrome_version in link]
    webDriver_version = links[0].split("=")[1]
    
    return webDriver_version

#download webdriver and extract it to the correct directory
async def fetch_webDriver(download_path, webDriver_path):

    task2 = asyncio.create_task(get_latest_webdriver_version())
    zip_download_path = download_path + "/" + FILE_NAME
    if os.path.isfile(zip_download_path):
        os.remove(zip_download_path)

    await task2
    print("Downloading correct webDriver")
    url = URL_FRONT + task2.result() + URL_BACK

    urllib.request.urlretrieve(url, zip_download_path)

    with zipfile.ZipFile(zip_download_path, 'r') as z:
        print("Extracting file")
        z.extractall(path=webDriver_path)
    
    os.remove(zip_download_path)
    print("Succesfully updated webDriver")

def update_webDriver(download_path, webDriver_path):
  asyncio.run(fetch_webDriver(download_path, webDriver_path))