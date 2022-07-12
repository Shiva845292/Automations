from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://na12.replicon.com/DXCTechnology/home/")
driver.maximize_window()
driver.implicitly_wait(100)
driver.find_element(By.XPATH,"//*[@id='LoginNameTextBox']").send_keys("s.mallapalli@dxc.com")
driver.find_element(By.XPATH,"//*[@id='PasswordTextBox']").send_keys("Ju@22845292")
driver.implicitly_wait(10)
driver.find_element(By.XPATH,"//*[@id='LoginButton']").click()
driver.find_element(By.XPATH,"//*[@id='okta-signin-username']").send_keys("s.mallapalli@dxc.com")
driver.find_element(By.XPATH,"//*[@id='okta-signin-password']").send_keys("Ju@22845292")
driver.implicitly_wait(10)
driver.find_element(By.XPATH,"//*[@id='okta-signin-submit']").click()
driver.find_element(By.XPATH,"//*[@id='form75']/div[2]/input").click()
time.sleep(60)
driver.find_element(By.XPATH,"//*[@id='timesheet-card']/timesheet-card/div/article/current-timesheet-card-item/div/ul/li/span").click()
print(driver.title)
time.sleep(100)
driver.quit()
