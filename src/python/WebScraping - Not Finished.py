# pip install selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# Set up the WebDriver
driver = webdriver.Chrome()

# Open the login page
driver.get("http://104.41.27.207:8030/Login")
time.sleep(2)

# Fill in the login form
username_field = driver.find_element(By.NAME, "txtLogin")
username_field.send_keys("dboliveira")
username_field.send_keys(Keys.TAB)

password_field = driver.find_element(By.NAME, "txtSenha")
password_field.send_keys("Abcd1234,.;")
password_field.send_keys(Keys.RETURN)
time.sleep(2)

# Check if login was successful
if "Logout" in driver.page_source:
    print("Login successful!")
    
    # Directly navigate to the Lista de Documentos page
    driver.get("http://104.41.27.207:8030/ModuloGED/ListadeDocumentos")
    time.sleep(3)  # Wait for the page to load
    
    # Locate the label or checkbox using its class or attributes
    Click_00 = driver.find_element(By.CLASS_NAME, "multiselect-all")
    Click_00.click()  # Perform the click action
    
    click_01 = driver.find_element(By.ID, "ContentPlaceHolder1_btnVerProjetos")
    driver.execute_script("arguments[0].click();", button)
    
    Click_02 = driver.find_element(By.CLASS_NAME, "multiselect-all")
    Click_02.click()  # Perform the click action
    
    # Example of extracting table data or any specific elements
    # Assuming there's a table with data
    table_rows = driver.find_elements(By.TAG_NAME, "tr")
    for row in table_rows:
        cells = row.find_elements(By.TAG_NAME, "td")
        cell_data = [cell.text for cell in cells]
        print(cell_data)  # Print or store each row's data
    
    # If there's an Export button, locate and click it
    try:
        export_button = driver.find_element(By.XPATH, "//button[text()='Exportar']")
        export_button.click()
        time.sleep(5)  # Wait for the file to download if necessary
    except:
        print("Export button not found.")

else:
    print("Login failed.")

# Close the driver
#driver.quit()
