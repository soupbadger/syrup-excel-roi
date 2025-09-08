from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# Replace your local chromedriver path here
chrome_driver_path = r"/path/to/chromedriver"

# Generic placeholder search URL (no real property or personal info)
search_url = "https://example.com/property-search?query=sample+street"

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(chrome_driver_path), options=options)
wait = WebDriverWait(driver, 20)

driver.get(search_url)

all_data = []

while True:
    try:
        rows = wait.until(EC.visibility_of_all_elements_located(
            (By.XPATH, "//table[@id='searchResultsTable']/tbody/tr")
        ))
    except:
        print("Results table did not load.")
        break

    for row in rows:
        try:
            view_link = row.find_element(By.CLASS_NAME, "viewDetails").get_attribute("href")
            owner_name = row.find_elements(By.TAG_NAME, "td")[1].text  # 2nd <td> is Owner
        except:
            continue

        driver.execute_script("window.open(arguments[0]);", view_link)
        driver.switch_to.window(driver.window_handles[1])

        try:
            sale_rows = wait.until(EC.visibility_of_all_elements_located(
                (By.XPATH, "//table[contains(@class,'table')]/tbody/tr")
            ))

            for sale in sale_rows:
                try:
                    sale_date = sale.find_element(By.XPATH, "./td[1]").text
                except:
                    sale_date = ""
                try:
                    price = sale.find_element(By.XPATH, "./td[2]").text
                except:
                    price = ""
                all_data.append({"Owner": owner_name, "Sale Date": sale_date, "Price": price})

        except:
            print(f"No sale data found for {view_link}")

        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(0.5)

    # Pagination
    try:
        next_btn = driver.find_element(By.ID, "searchResultsTable_next")
        if "disabled" in next_btn.get_attribute("class"):
            break
        else:
            next_btn.click()
            time.sleep(2)
    except:
        break

df = pd.DataFrame(all_data)
df.to_csv("sample_sales.csv", index=False)
print("Done! Data saved to sample_sales.csv")

driver.quit()
