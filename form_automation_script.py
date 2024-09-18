import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Initialize Tkinter and hide the main window
Tk().withdraw()

# Open file dialog to select the Excel file
filetypes = [("Excel files", "*.xlsx *.xls")]
excel_path = askopenfilename(title="Select the Excel file", filetypes=filetypes)

if not excel_path:
    print("No file selected.")
    exit()

# Load the Excel file
data = pd.read_excel(excel_path)

# Set up the web driver with WebDriver Manager
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# Open the website
driver.get('your_form_link')

# Handle login form
username = 'username_or_email'
password = 'password'

try:
    # Wait until the username field is present and interactable
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'email')))
    driver.find_element(By.NAME, 'email').send_keys(username)
    driver.find_element(By.NAME, 'password').send_keys(password)
    
    # Wait until the Login button is clickable and then click it
    login_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "Login")]'))
    )
    login_button.click()

    # Wait for redirection and ensure the form page has loaded
    WebDriverWait(driver, 20).until(EC.url_contains('/emigrant-registrations/create'))
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'field_name')))
    
    # Loop through each row in the Excel file
    for index, row in data.iterrows():
        #for dropdown if it is in your form
        ptn_dropdown = Select(driver.find_element(By.NAME, 'dropdown_field_name'))
        ptn_dropdown.select_by_value(str(row['drop_down_field_name']))  # Ensure this is a string

        # Wait for the job_id dropdown to become enabled and populate based on the PTN selection
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'field_name')))
    
        # Select the Job ID option
        job_id_dropdown = Select(driver.find_element(By.NAME, 'job_id'))
        job_id_dropdown.select_by_value(str(row['job_id']))  # Ensure this is a string

        # Convert the date format
        date_of_birth = row['date_of_birth'].strftime('%d/%m/%Y')
        passport_issue_date = row['passport_issue_date'].strftime('%d/%m/%Y')
        passport_expiry_date = row['passport_expiry_date'].strftime('%d/%m/%Y')
        deposite_date = row['deposite_date'].strftime('%d/%m/%Y')

        # Fill in the form fields with data from the Excel row
        driver.find_element(By.NAME, 'full_name').send_keys(row['full_name'])
        driver.find_element(By.NAME, 'father_name').send_keys(row['father_name'])
        driver.find_element(By.NAME, 'cnic').send_keys(row['cnic'])
        
        gender_dropdown = Select(driver.find_element(By.NAME, 'gender'))
        gender_dropdown.select_by_value(str(row['gender']))
        
        driver.find_element(By.NAME, 'date_of_birth').send_keys(date_of_birth)
        driver.find_element(By.NAME, 'bank_account_title').send_keys(row['bank_account_title'])
        driver.find_element(By.NAME, 'bank_account_no').send_keys(row['bank_account_no'])
        
        bank_id_dropdown = Select(driver.find_element(By.NAME, 'bank_id'))
        bank_id_dropdown.select_by_value(str(row['bank_id']))
        
        driver.find_element(By.NAME, 'place_of_birth').send_keys(row['place_of_birth'])
        driver.find_element(By.NAME, 'phone_no').send_keys(str(row['phone_no']).zfill(11))
        driver.find_element(By.NAME, 'email').send_keys(row['email'])
        driver.find_element(By.NAME, 'address').send_keys(row['address'])
        
        province_dropdown = Select(driver.find_element(By.NAME, 'province_id'))
        province_dropdown.select_by_value(str(row['province_id']))
        
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'district_id')))
        district_id_dropdown = Select(driver.find_element(By.NAME, 'district_id'))
        district_id_dropdown.select_by_value(str(row['district_id']))
        
        domicile_dropdown = Select(driver.find_element(By.NAME, 'domicile_district_id'))
        domicile_dropdown.select_by_value(str(row['domicile_district_id']))
        
        education_dropdown = Select(driver.find_element(By.NAME, 'education_id'))
        education_dropdown.select_by_value(str(row['education_id']))
        
        driver.find_element(By.NAME, 'passport_no').send_keys(row['passport_no'])
        driver.find_element(By.NAME, 'passport_place_of_issue').send_keys(row['passport_place_of_issue'])
        driver.find_element(By.NAME, 'passport_issue_date').send_keys(passport_issue_date)
        driver.find_element(By.NAME, 'passport_expiry_date').send_keys(passport_expiry_date)
        
        driver.find_element(By.NAME, 'nominee_full_name').send_keys(row['nominee_full_name'])
        driver.find_element(By.NAME, 'nominee_age').send_keys(row['nominee_age'])
        driver.find_element(By.NAME, 'nominee_relation').send_keys(row['nominee_relation'])
        driver.find_element(By.NAME, 'nominee_cnic').send_keys(row['nominee_cnic'])
        driver.find_element(By.NAME, 'nominee_phone_no').send_keys(str(row['nominee_phone_no']).zfill(11))
        driver.find_element(By.NAME, 'nominee_address').send_keys(row['nominee_address'])
        
        driver.find_element(By.NAME, 'voucher_no').send_keys(row['voucher_no'])
        driver.find_element(By.NAME, 'deposite_date').send_keys(deposite_date)
        driver.find_element(By.NAME, 'payment_recieved_from_the_emigrant').send_keys(row['payment_recieved_from_the_emigrant'])
        
        picture_url_input = driver.find_element(By.NAME, 'picture_url')
        picture_url_input.send_keys(row['picture_url'])  # File path from Excel
        
        driver.find_element(By.NAME, 'salary').send_keys(str(row['salary']))
        # Submit the form
        submit_button = driver.find_element(By.XPATH, '//button[contains(text(), "Register New Emigrant")]')
        submit_button.click()

        # Optionally wait for the form to submit and any post-submission processing
        time.sleep(20)
        
        # Wait for the page to reload and be ready for the next input
        driver.get('your_web_from_link')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'full_name')))
        
except Exception as e:
    print(f"An error occurred: {e}")
