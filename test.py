from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# === CONFIG ===
USERNAME = "s.mallapalli"
PASSWORD = "o9y[uihuew0Dgsb@"  # Keep secure in real use

# URLs to check
urls_to_check = [
    "https://sn.itsmservice.com/nav_to.do?uri=%2Fsn_vul_vulnerable_item_list.do%3Fsysparm_query%3Dactive%3Dtrue%5Ecmdb_ci.sys_class_nameINSTANCEOFcmdb_ci_server%5Ecmdb_ci.companyLIKEdxc%5Ecmdb_ci.nameLIKEpvs%5EORcmdb_ci.nameLIKEtsr%5Evulnerability.summaryLIKEAdobe%5Eassignment_group%3Dac5004376fe7a100a53d15ef1e3ee489%5Eu_detection.proofLIKEC%3A%5CAPPS%5CAdobe%5CAcrobat%20DC%26sysparm_view%3D",
    "https://sn.itsmservice.com/nav_to.do?uri=%2Fsn_vul_vulnerable_item_list.do%3Fsysparm_query%3Dactive%3Dtrue%5Ecmdb_ci.sys_class_nameINSTANCEOFcmdb_ci_server%5Ecmdb_ci.companyLIKEdxc%5Ecmdb_ci.nameLIKEpvs%5EORcmdb_ci.nameLIKEtsr%5Evulnerability.summarySTARTSWITHAzul%5Eassignment_group%3Dac5004376fe7a100a53d15ef1e3ee489%26sysparm_clear_stack%3Dtrue",
    "https://sn.itsmservice.com/nav_to.do?uri=%2Fsn_vul_vulnerable_item_list.do%3Fsysparm_query%3Dactive%3Dtrue%5Ecmdb_ci.sys_class_nameINSTANCEOFcmdb_ci_server%5Ecmdb_ci.companyLIKEdxc%5Ecmdb_ci.nameLIKEpvs%5EORcmdb_ci.nameLIKEtsr%5Evulnerability.summarySTARTSWITHNotepad%2B%2B%5Eassignment_group%3Dac5004376fe7a100a53d15ef1e3ee489",
    "https://sn.itsmservice.com/now/nav/ui/classic/params/target/sn_vul_vulnerable_item_list.do%3Fsysparm_query%3Dactive%253Dtrue%255Ecmdb_ci.sys_class_nameINSTANCEOFcmdb_ci_server%255Ecmdb_ci.companyLIKEdxc%255Ecmdb_ci.nameLIKEtsr%255EORcmdb_ci.nameLIKEpvs%255Ecmdb_ci.nameNOT%2520LIKEZURNE%255Evulnerability.summarySTARTSWITHMicrosoft%2520Visual%2520Studio%255Ecmdb_ci.u_business_owner.u_business_unit%253DIreland%255Eassignment_group%253Dac5004376fe7a100a53d15ef1e3ee489%26sysparm_first_row%3D1%26sysparm_view%3D",
    "https://sn.itsmservice.com/sn_vul_vulnerable_item_list.do?sysparm_query=active%3Dtrue%5Ecmdb_ci.sys_class_nameINSTANCEOFcmdb_ci_server%5Ecmdb_ci.companyLIKEdxc%5Ecmdb_ci.nameINDETSR0006%2CGEPVS0200%2CGEPVS0215%2CGEPVS0287%2CGEPVS0288%2CGEPVS0399%2CGEPVS0507%2CGEPVS0527%2C%5Evulnerability.summarySTARTSWITHMicrosoft%20SQL%20Server%2C%20ODBC%20and%20OLE%20DB%20Driver%20for%20SQL%20Server%20Remote%20Code%20Execution%20(RCE)%20Vulnerabilities%20for%20June%202023%5Eassignment_group%3Dac5004376fe7a100a53d15ef1e3ee489%5EORassignment_group%3D10934f69dba9ff009a9c1252399619a3%5Eu_detection.proofLIKE18.4.0.0&sysparm_first_row=1&sysparm_view=",
    "https://sn.itsmservice.com/sn_vul_vulnerable_item_list.do?sysparm_query=active%3Dtrue%5Ecmdb_ci.sys_class_nameINSTANCEOFcmdb_ci_server%5Ecmdb_ci.companyLIKEdxc%5Ecmdb_ci.nameINDETSR0010%2CGEPVS0511%2CGEPVS0531%2CGEPVS0418%2CGEPVS0419%2CGEPVS0296%2CGEPVS0297%5Evulnerability.summarySTARTSWITHMicrosoft%20SQL%20Server%2C%20ODBC%20and%20OLE%20DB%20Driver%20for%20SQL%20Server%20Remote%20Code%20Execution%20(RCE)%20Vulnerabilities%20for%20June%202023%5Eassignment_group%3Dac5004376fe7a100a53d15ef1e3ee489%5EORassignment_group%3D10934f69dba9ff009a9c1252399619a3%5Eu_detection.proofLIKE17.6.1.1&sysparm_first_row=1&sysparm_view=",
    "https://sn.itsmservice.com/sn_vul_vulnerable_item_list.do?sysparm_query=active%3Dtrue%5Ecmdb_ci.sys_class_nameINSTANCEOFcmdb_ci_server%5Ecmdb_ci.companyLIKEdxc%5Ecmdb_ci.nameINDETSR0001%2CDETSR0015%5Evulnerability.summarySTARTSWITHMicrosoft%20SQL%20Server%2C%20ODBC%20and%20OLE%20DB%20Driver%20for%20SQL%20Server%20Remote%20Code%20Execution%20(RCE)%20Vulnerabilities%20for%20June%202023%5Eassignment_group%3Dac5004376fe7a100a53d15ef1e3ee489%5EORassignment_group%3D10934f69dba9ff009a9c1252399619a3%5EORassignment_group%3D385004376fe7a100a53d15ef1e3ee493&sysparm_first_row=1&sysparm_view=",
    "https://sn.itsmservice.com/sn_vul_vulnerable_item_list.do?sysparm_tiny=AN4ZstkPaB76qG7u2CtWcuAUTC45WBn8",
    "https://sn.itsmservice.com/sn_vul_vulnerable_item_list.do?sysparm_query=active%3Dtrue%5Ecmdb_ci.sys_class_nameINSTANCEOFcmdb_ci_server%5Ecmdb_ci.companyLIKEdxc%5Ecmdb_ci.nameINDETSR0013%2CGEPVS0236%2CGEPVS0338%2CGEPVS0339%2CGEPVS0340%2CGEPVS0346%2CGEPVS0347%2CGEPVS0398%2CGEPVS0407%2CGEPVS0439%2CGEPVS0510%2CGEPVS0530%5Evulnerability.summarySTARTSWITHMicrosoft%20SQL%20Server%2C%20ODBC%20and%20OLE%20DB%20Driver%20for%20SQL%20Server%20Remote%20Code%20Execution%20(RCE)%20Vulnerabilities%20for%20June%202023%5Eassignment_group%3Dac5004376fe7a100a53d15ef1e3ee489%5EORassignment_group%3D10934f69dba9ff009a9c1252399619a3%5EORassignment_group%3D385004376fe7a100a53d15ef1e3ee493&sysparm_first_row=1&sysparm_view=",
    "https://sn.itsmservice.com/sn_vul_vulnerable_item_list.do?sysparm_query=active%3Dtrue%5Ecmdb_ci.sys_class_nameINSTANCEOFcmdb_ci_server%5Ecmdb_ci.companyLIKEdxc%5Ecmdb_ci.name%3DDETSR0012%5Evulnerability.summarySTARTSWITHMicrosoft%20SQL%20Server%2C%20ODBC%20and%20OLE%20DB%20Driver%20for%20SQL%20Server%20Remote%20Code%20Execution%20(RCE)%20Vulnerabilities%20for%20June%202023%5Eassignment_group%3Dac5004376fe7a100a53d15ef1e3ee489%5EORassignment_group%3D10934f69dba9ff009a9c1252399619a3&sysparm_first_row=1&sysparm_view=",
    "https://sn.itsmservice.com/now/nav/ui/classic/params/target/sn_vul_vulnerable_item_list.do%3Fsysparm_query%3Dactive%253Dtrue%255Ecmdb_ci.sys_class_nameINSTANCEOFcmdb_ci_server%255Ecmdb_ci.companyLIKEdxc%255Ecmdb_ci.nameLIKEpvs%255EORcmdb_ci.nameLIKEtsr%255Evulnerability.summarySTARTSWITHGhostscript%255Eassignment_group%253D%26sysparm_first_row%3D1%26sysparm_view%3D"
]

def setup_driver():
    options = Options()
    options.use_chromium = True
    driver = webdriver.Edge(options=options)
    driver.maximize_window()
    return driver

def login(driver, wait):
    driver.get(urls_to_check[0])

    username_input = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="input44"]')))
    username_input.clear()
    username_input.send_keys(USERNAME)

    driver.find_element(By.XPATH, '//*[@id="form36"]/div[2]/input').click()

    password_input = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="input80"]')))
    password_input.clear()
    password_input.send_keys(PASSWORD)

    driver.find_element(By.XPATH, '//*[@id="form72"]/div[2]/input').click()

    push_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form98"]/div[2]/div/div[2]/div[2]/div[2]/a')))
    push_btn.click()

    print("Waiting for 2FA confirmation (30 seconds)...")
    time.sleep(30)

def check_data_on_page(driver):
    try:
        # Check for any <tr> with id starting with 'row_sn_vul_vulnerable_item_'
        rows = driver.find_elements(By.XPATH, "//tr[starts-with(@id, 'row_sn_vul_vulnerable_item_')]")
        # Check if any row has visible text content
        if rows and any(row.text.strip() for row in rows):
            return True
        return False
    except Exception:
        return False

def main():
    driver = setup_driver()
    wait = WebDriverWait(driver, 15)

    try:
        login(driver, wait)

        for idx, url in enumerate(urls_to_check, start=1):
            if idx == 1:
                driver.get(url)
            else:
                driver.execute_script("window.open('');")
                driver.switch_to.window(driver.window_handles[-1])
                driver.get(url)

            time.sleep(5)  # Wait for page to load fully

            if check_data_on_page(driver):
                print(f"✅ URL {idx}: Data found")
            else:
                print(f"❌ URL {idx}: No data found")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
