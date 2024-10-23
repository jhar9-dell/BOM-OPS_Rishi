from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import tkinter as tk
from tkinter import messagebox

def show_popup(message):
    root = tk.Tk()
    root.withdraw() 
    messagebox.showinfo("Notification", message)
    root.destroy()

def automate_browser(excel_file_path):
    driver = webdriver.Edge()
    driver.get("https://agile.us.dell.com/Agile/")

    time.sleep(10)
    driver.refresh()
    driver.maximize_window()

    time.sleep(10)

    data = pd.read_excel(excel_file_path, sheet_name='ECO-MOD')

    for index, row in data.iterrows():
        searchbox_value = row['PES ID']
        mod_value = row['MOD']
        eco_value = row['ECO']

        # Check for blank values
        if pd.isna(searchbox_value) or pd.isna(mod_value) or pd.isna(eco_value):
            show_popup(f"Sheet has an empty row at index: {index + 1}")  
            driver.quit() 
            return 

        searchbox = driver.find_element(By.ID, 'QUICKSEARCH_STRING')
        searchbox.clear()
        searchbox.send_keys(searchbox_value)
        searchbox.send_keys(Keys.RETURN)
        time.sleep(5)

        editco = driver.find_element(By.ID, 'MSG_Edit')
        editco.click()
        time.sleep(3)

        modvalue = driver.find_element(By.ID, 'R1_1585_6')
        modvalue.clear()
        modvalue.send_keys(mod_value)
        modvalue.send_keys(Keys.RETURN)
        time.sleep(3)

        ecovalue = driver.find_element(By.ID, 'R1_1586_6')
        ecovalue.clear()
        ecovalue.send_keys(eco_value)
        ecovalue.send_keys(Keys.RETURN)
        time.sleep(3)

        dell_save = driver.find_element(By.ID, 'MSG_Save')
        dell_save.click()
        time.sleep(2)

    show_popup("Process Completed")
    driver.quit() 

if __name__ == "__main__":
    excel_file_path = 'C:/Users/jhar9/Downloads/origi2.xlsx'
    automate_browser(excel_file_path)