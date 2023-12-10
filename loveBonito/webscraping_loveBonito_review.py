import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
import re

def scrape_reviews(driver, main_url):
    reviews_data = []

    try:
        driver.get(main_url)
        time.sleep(2)

        while True:
            current_url = driver.current_url
            reviews = driver.find_elements(By.CSS_SELECTOR, "div.review-item")
            for review in reviews:
                rating_style = review.find_element(By.CSS_SELECTOR, "div.m-star-ratings").get_attribute('style')
                rating = re.search(r'--rating: (\d+);', rating_style).group(1) if rating_style else "No rating"

                fit_size_text = review.find_element(By.CSS_SELECTOR, "div.fit-additional").text.split("\n")
                fit = fit_size_text[0].split(":")[1].strip() if len(fit_size_text) > 0 else "No fit info"
                size_color = fit_size_text[1] if len(fit_size_text) > 1 else "No size/color info"

                fit_labels_elements = review.find_elements(By.CSS_SELECTOR, "div.fit-labels")
                fit_labels = [element.text for element in fit_labels_elements]

                review_text = review.find_element(By.CSS_SELECTOR, "div.review-body").text

                reply_elements = review.find_elements(By.CSS_SELECTOR, "div.reply-body")
                reply_text = reply_elements[0].text if reply_elements else "No reply"

                review_date_element = review.find_element(By.CSS_SELECTOR, "div.review-date i")
                review_date = review_date_element.text if review_date_element else "No date"

                reviews_data.append({
                    "Main URL": main_url,
                    "Page URL": current_url,
                    "Rating": rating,
                    "Fit": fit,
                    "Size and Color": size_color,
                    "Fit Labels": " | ".join(fit_labels),
                    "Review": review_text,
                    "Reply": reply_text,
                    "Date": review_date
                })

            next_button = driver.find_elements(By.CSS_SELECTOR, "div.right_arrows > div.sf-pagination__item--next:not(.disabled) > button")
            disabled_next = driver.find_elements(By.CSS_SELECTOR, "div.sf-pagination__item--next.disabled_last")

            if next_button and not disabled_next:
                next_button[0].click()
                time.sleep(2)
            else:
                break

    except Exception as e:
        print(f"Terjadi kesalahan: {e}")
        return pd.DataFrame(reviews_data), e

    return pd.DataFrame(reviews_data), None

# Inisialisasi WebDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# Membaca daftar URL dari file Excel
excel_path = 'lovebonito_products_review_urls.xlsx'  # Ganti dengan path file Excel Anda
df_urls = pd.read_excel(excel_path)
all_reviews = pd.DataFrame()

# Iterasi melalui setiap URL dan mengumpulkan ulasan
for index, main_url in enumerate(df_urls['urls']):
    reviews_df, error = scrape_reviews(driver, main_url)
    all_reviews = pd.concat([all_reviews, reviews_df], ignore_index=True)
    
    if error:
        all_reviews.to_csv(f"reviews_error_{index+1}.csv", index=False)
        break

# Menutup WebDriver setelah proses selesai
driver.quit()

# Menampilkan DataFrame hasil gabungan
print(all_reviews.head())

# Menyimpan hasil gabungan ke file CSV
all_reviews.to_csv("product_reviews.csv", index=False)
