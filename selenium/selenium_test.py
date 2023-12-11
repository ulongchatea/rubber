from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

driver.get("https://naver.com")
naver_searchbox = driver.find_element(By.NAME, "query")
naver_searchbox.send_keys("Automation")
naver_searchbox.send_keys(Keys.RETURN)
#driver.find_element(By.ID, "search-btn").click()



time.sleep(10)
