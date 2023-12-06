from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
from time import sleep
import random
import pandas as pd

# ChromeDriver path
chrome_driver_path = "chromedriver"

# Setting up ChromeDriver
service = Service(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service)

def scrape_product_urls(url):
    driver.get(url)
    sleep(random.uniform(1, 3))  # Random sleep to mimic human behavior

    product_urls = set()
    page_number = 1

    while True:
        try:
            product_links = driver.find_elements(By.CSS_SELECTOR, "a.product-item-photo")
            for link in product_links:
                product_urls.add(link.get_attribute('href'))

            # Save to Excel every 5 pages
            if page_number % 5 == 0:
                save_to_excel(product_urls, 'product_urls_partial_save.xlsx')

            next_page_link = driver.find_element(By.CSS_SELECTOR, "a.action.next:not(.disabled)")
            next_page_url = next_page_link.get_attribute('href')
            driver.get(next_page_url)
            page_number += 1
            sleep(random.uniform(1, 3))
        except NoSuchElementException:
            break
        except Exception as e:
            print(f"Error occurred: {e}")
            break

    return product_urls

def save_to_excel(data, filename):
    df = pd.DataFrame(list(data), columns=["Product URLs"])
    df.to_excel(filename, index=False)
    print(f"Data saved to {filename}")

# URL to scrape
url_to_scrape = "https://www.matahari.com/anak/bayi.html"

# Scrape the product URLs
product_urls = scrape_product_urls(url_to_scrape)

# Close the driver
driver.quit()

# Save the final results to Excel
final_excel_filename = "product_urls_final.xlsx"
save_to_excel(product_urls, final_excel_filename)

print(f"\nScraping complete. The URLs have been saved to '{final_excel_filename}'.")
