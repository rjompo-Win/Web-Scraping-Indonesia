import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from time import sleep

# Fungsi untuk menyimpan data ke Excel
def save_data_to_excel(data, filename):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"Data telah disimpan ke {filename}")

# Setup ChromeDriver
chrome_driver_path = "chromedriver"
service = Service(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service)

# Membaca URL dari file Excel
excel_path = "lovebonito_url.xlsx"
urls_df = pd.read_excel(excel_path)
urls = urls_df['url'].tolist()

all_data = []

try:
    for url in urls:
    # Membuka URL
        driver.get(url)
        sleep(2)  # Tunggu agar halaman sepenuhnya dimuat

        # Mencari elemen untuk rata-rata bintang dan jumlah rating
        try:
            average_rating_element = driver.find_element(By.CLASS_NAME, "average-rating")
            average_rating = average_rating_element.text.strip()
        except NoSuchElementException:
            average_rating = "Tidak tersedia"
            print("Rata-rata bintang tidak ditemukan.")

        try:
            rating_count_element = driver.find_element(By.CLASS_NAME, "rating-count")
            rating_count = rating_count_element.text.strip()
        except NoSuchElementException:
            rating_count = "Tidak tersedia"
            print("Jumlah rating tidak ditemukan.")

        # Menemukan semua tombol warna
        color_buttons = driver.find_elements(By.CSS_SELECTOR, ".sf-button.sf-button--pure.sf-color")
        color_data = []

        for button in color_buttons:
            try:
                if "sf-color--active" not in button.get_attribute("class"):
                    button.click()
                    sleep(2)  # Tunggu data dimuat
            except WebDriverException as e:
                print(f"Tidak dapat mengklik tombol warna: {e}")
                color_data.append({
                    "URL": url,
                    "Error": "Tidak dapat mengklik tombol warna"
                })
                continue  # Lanjutkan ke tombol warna berikutnya

            color_name_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "product__color-name"))
            )
            color_name = color_name_element.text.strip()

            title_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "sf-heading__title"))
            )
            title = title_element.text.strip()

            price_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "sf-price__value"))
            )
            price = price_element.text.strip()

            sizes_element = driver.find_elements(By.CLASS_NAME, "product__select-size-sizes")
            sizes = [size.text.strip() for size in sizes_element if size.text.strip() != ""]

            color_data.append({
                "URL": url,
                "Color Style": button.get_attribute("style"),
                "Color Name": color_name,
                "Title": title,
                "Price": price,
                "Sizes": sizes,
                "Average Rating": average_rating,
                "Rating Count": rating_count
            })

        all_data.extend(color_data)

except Exception as e:
    print(f"Terjadi kesalahan: {e}")

finally:
    # Menutup driver
    driver.quit()

# Menyimpan data ke file Excel
save_data_to_excel(all_data, "love_bonito_product.xlsx")
