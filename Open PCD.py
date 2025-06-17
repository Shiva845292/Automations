from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import glob

USERNAME = "s.mallapalli"
PASSWORD = "o9y[uihuew0Dgsb@"
DOWNLOAD_DIR = r"C:\Users\smallapalli\Downloads"
URL = "https://ctx-ldc.cag.zurich.com/Citrix/OKTAXAWeb/"
RETRY_LIMIT = 3

def setup_driver():
    options = Options()
    options.use_chromium = True
    options.add_argument("start-maximized")
    driver = webdriver.Edge(options=options)
    return driver

def get_latest_ica_file(folder_path):
    list_of_files = glob.glob(os.path.join(folder_path, "*.ica"))
    if not list_of_files:
        raise FileNotFoundError("No ICA file found in Downloads.")
    latest_file = max(list_of_files, key=os.path.getctime)
    return latest_file

def login(driver):
    wait = WebDriverWait(driver, 30)
    driver.get(URL)

    username_input = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="input28"]')))
    username_input.clear()
    username_input.send_keys(USERNAME)
    driver.find_element(By.XPATH, '//*[@id="form20"]/div[2]/input').click()

    password_input = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="input64"]')))
    password_input.clear()
    password_input.send_keys(PASSWORD)
    driver.find_element(By.XPATH, '//*[@id="form56"]/div[2]/input').click()

    push_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form82"]/div[2]/div/div[2]/div[2]/div[2]/a')))
    push_btn.click()

    print("Waiting for 2FA confirmation (30 seconds)...")
    time.sleep(30)

    desktops_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="desktopsBtn"]/div')))
    desktops_btn.click()

    time.sleep(3)
    app_icon = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="home-screen"]/div[2]/section[5]/div[5]/div/ul/li[3]/a[2]/img')))
    app_icon.click()

    print("Waiting for ICA file download...")
    time.sleep(5)
    latest_ica = get_latest_ica_file(DOWNLOAD_DIR)
    print(f"Launching ICA file: {latest_ica}")
    os.startfile(latest_ica)

def main():
    for attempt in range(1, RETRY_LIMIT + 1):
        try:
            print(f"Attempt {attempt} of {RETRY_LIMIT}...")
            driver = setup_driver()
            login(driver)
            print("Login and launch successful!")
            break
        except Exception as e:
            print(f"Error on attempt {attempt}: {e}")
            driver.quit()
            if attempt == RETRY_LIMIT:
                print("Max retries reached. Exiting.")
        time.sleep(3)

if __name__ == "__main__":
    main()
