import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import random
from bs4 import BeautifulSoup
import re
import pandas as pd
import os
import tkinter as tk
from tkinter import messagebox

def initializing_driver():

    options = Options()

    # Adding argument to disable the AutomationControlled flag
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--incognito")

    # Exclude the collection of enable-automation switches
    options.add_experimental_option("excludeSwitches", ["enable-automation"])

    # Turn-off userAutomationExtension
    options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(
        service=ChromeService(ChromeDriverManager().install()),
        options=options)
    
    # Changing the property of the navigator value for webdriver to undefined
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    useragentarray = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/535.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/532.36",
        "Mozilla/4.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.0.0 Safari/533.36",
        "Mozilla/2.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/532.36",
        "Mozilla/3.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/532.36" ]

    driver.execute_cdp_cmd("Network.setUserAgentOverride", {"userAgent": random.choice(useragentarray)})
    update_status("driver initialized")

    return driver

def startScrape(product):
    driver = initializing_driver()
    driver.get(f'https://www.amazon.in/s?k={product}')
    driver.implicitly_wait(4)
    element =  driver.page_source
    driver.quit()
    update_status("driver closed and page sent for scrapping information")
    return element

def scrapeAmazon(page_source):
    driver = initializing_driver()
    soup = BeautifulSoup(page_source, 'html.parser')
    elements = soup.find_all('div', {'data-asin': True, 'data-component-type': 's-search-result'})
    update_status(f"Number of prodcuts found= {len(elements)}")
    
    data = []
    update_status("Collecting data...")
    
    for element in elements:
        member = dict()
        text = element.text
        try:
            member['Product Title'] = element.find('h2').text.strip()
            member['Discounted Price'] = element.find(class_='a-price-whole').text.strip()
        except:
            if re.search(r'₹(\d+(?:,\d+)*)', text):
                member['Discounted Price'] = re.search(r'₹(\d+(?:,\d+)*)', text).group(1).replace(',', '')
            else:
                continue
        url = "https://www.amazon.in"+element.find('a')['href']
        if re.search(r'Sponsored', text):
            member['Sponsored'] = True
        else:
            member['Sponsored'] = False
        price = re.search(r'M\.R\.P\:\s*₹(\d+(?:,\d+)*)', text)
        if price:
            member['Original Price'] = price.group(1).replace(',', '')
        else:
            member['Original Price'] = member['Discounted Price']
        rating = re.search(r'(\d+(\.\d+)?)\s+out\s+of\s+5\s+stars', text)
        if rating:
            member['Ratings'] = rating.group(1)
        else:
            member['Ratings'] = 0
        reviews = re.search(r'out\s+of\s+5\s+stars\s+(\d+)\s+', text)
        if reviews:
            member['Reviews'] = reviews.group(1)
        else:
            member['Reviews'] = 0
        try:
            driver.get(url)
            subsoup = BeautifulSoup(driver.page_source, 'html.parser')
            content = subsoup.find('div', id='productDetailsWithModules_feature_div').text.strip('\n')
            search = re.search(r'Net Quantity\s+‎(\d+(?:\.\d+)?)\s+([a-zA-Z]+)', content)
            if search:
                member['Net Quantity'] = f'{search.group(1)} {search.group(2)}'
            else:
                member['Net Quantity'] = 'N/A'
        except:
            member['Net Quantity'] = 'N/A'
        data.append(member)
    driver.quit()
    update_status("Data Collected making excel file!")
    return data

def data_filtering(data, product):
    #colomns = ['Product Title', 'Original Price', 'Discounted Price', 'Ratings', 'Reviews', 'Quantity', 'Product Dimensions', 'Sponsored']
    
    df = pd.DataFrame(data)

    # Sort DataFrame based on NA count, Reviews, and Ratings
    #df = df.sort_values(by=['Reviews'], ascending=[False])

    # Save DataFrame to Excel
    output_path = f'output/{product}.xlsx'
    output_directory = os.path.dirname(output_path)
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    df.to_excel(output_path, index=False)
    update_status(f"Data saved to excel file! at {output_path}")

def main(product):
    data = scrapeAmazon(startScrape(product))

    # Perform data analysis and save to Excel
    data_filtering(data, product)

def start_scraping():
    product = product_entry.get()
    if not product:
        messagebox.showerror("Error", "Please enter a product")
        return
    main(product)

def update_status(message):
    status_label.config(text=message)
    status_label.update()

# Tkinter setup
root = tk.Tk()
root.title("Report Generator for Village Compamany")
root.geometry("400x300")

# Product entry
product_label = tk.Label(root, text="Enter Product Name:")
product_label.pack()
product_entry = tk.Entry(root, width=90)  # Set width to make it larger
product_entry.pack(padx=40, pady=20)  # Add padx for padding around the entry widget

# Start scraping button
scrape_button = tk.Button(root, text="Start Scraping", command=start_scraping, width=44)  # Set width to make it larger
scrape_button.pack(pady=22)

# Status label
status_label = tk.Label(root, text="")
status_label.pack()

author_info_label = tk.Label(root, text="Author Information: Shruti Dubey\n (https://github.com/Dubeyshruti)", font=("Helvetica", 12), fg="blue")
author_info_label.pack()

warning_label = tk.Label(root, text="prerequisite: You must have Chrome browser Installed in your system", font=("Helvetica", 9), fg="red")
warning_label.pack(padx=9, pady=9)

root.mainloop()