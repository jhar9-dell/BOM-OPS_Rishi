from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import tkinter as tk
from tkinter import messagebox
import customtkinter

def show_popup(message):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    messagebox.showinfo("Notification", message)
    root.destroy()


def automate_browser(excel_file_path):
    driver = webdriver.Edge()
    driver.get("https://agile.us.dell.com/Agile/PLMServlet")

    time.sleep(10)
    driver.maximize_window()

    time.sleep(10)

    
    data = pd.read_excel(excel_file_path, sheet_name='Analyst')

    for index, row in data.iterrows():
        searchbox_value = row['PES ID']
        dellana_value = row['Dell Analyst'] 

        if pd.isna(searchbox_value) or pd.isna(dellana_value): 
            continue

        
        searchbox = driver.find_element(By.ID, 'QUICKSEARCH_STRING')
        searchbox.clear()
        searchbox.send_keys(searchbox_value)
        searchbox.send_keys(Keys.RETURN)
        time.sleep(7)

        editco = driver.find_element(By.ID, 'MSG_Edit')
        editco.click()
        time.sleep(3)

        dellana = driver.find_element(By.ID, 'search_query_R1_1099_0_display')
        dellana.clear()
        dellana.send_keys(dellana_value)
        dellana.send_keys(Keys.RETURN)
        time.sleep(3)

        dell_save = driver.find_element(By.ID, 'MSG_Save')
        dell_save.click()
        time.sleep(2)
    
    show_popup("Process Completed")
    
if __name__ == "__main__":
    excel_file_path = 'C:/Users/jhar9/Downloads/origi2.xlsx'
    automate_browser(excel_file_path)