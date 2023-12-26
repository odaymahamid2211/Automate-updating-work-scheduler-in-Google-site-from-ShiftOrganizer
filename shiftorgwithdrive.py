from datetime import datetime
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.http import MediaFileUpload
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import time
import os
from google.auth.transport.requests import Request
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def ShiftOrganizer():
    SCOPES = ['https://www.googleapis.com/auth/drive.file']
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json')

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    drive_service = build('drive', 'v3', credentials=creds)

    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": os.getcwd(),  # Set the download directory to the current working directory
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    chrome_driver_path = ChromeDriverManager().install()
    service = ChromeService(chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)

    driver.get("https://app.shiftorganizer.com/app/rota")

    company_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "company"))
    )

    username_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "username"))
    )

    password_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "password"))
    )

    company_field.send_keys("1853")
    username_field.send_keys("odaym")
    password_field.send_keys("AbuShibli2211!")

    login_button = driver.find_element(By.ID, "log-in")
    login_button.click()

    time.sleep(1)

    excel_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[2]/div[2]/div[2]/div[1]/div/div[2]/div[5]/div[1]/button"))
    )
    time.sleep(5)
    excel_button.click()

    time.sleep(10)

    files = os.listdir(os.getcwd())

    excel_file = [file for file in files if file.endswith('.xlsx')][0]
    current_date = datetime.now().strftime("%Y-%m-%d")

    new_excel_file_name = f"schedule_{current_date}.xlsx"

    os.rename(excel_file, new_excel_file_name)

    file_metadata = {'name': new_excel_file_name}
    media = MediaFileUpload(new_excel_file_name, resumable=True)
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()

    file_id = file.get('id')
    file_link = f'https://drive.google.com/file/d/{file_id}/edit'
    print(file_link)
    chrome_driver_path = ChromeDriverManager().install()
    service = ChromeService(chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)

    google_sites_url = "https://sites.google.com/d/17bwVJndsskCUBeatExeMw6E6XoPU1RnF/p/1MF5zsbrXkvO12Id5jJZCRf_nxvzRML1L/edit"
    driver.get(google_sites_url)

    email_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "identifierId"))
    )
    email_field.send_keys("oday.mahamid@wideops.com")
    email_field.send_keys(Keys.RETURN)

    time.sleep(7)

    password_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "Passwd"))
    )
    password_field.send_keys("AbuShibli2211")
    password_field.send_keys(Keys.RETURN)

    embed_label = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@aria-label='Embed']"))
    )
    embed_label.click()

    input_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "poFWNe"))
    )

    input_field.send_keys(file_link)

    time.sleep(10)

    insert_button = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'NPEfkd') and contains(text(), 'Insert')]"))
    )
    driver.execute_script("arguments[0].click();", insert_button)

    time.sleep(2)

    publish_button = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'NPEfkd') and contains(text(), 'Publish')]"))
    )
    driver.execute_script("arguments[0].click();", publish_button)

    time.sleep(5)
    driver.quit()
    return


if __name__ == '__main__':
    ShiftOrganizer()
