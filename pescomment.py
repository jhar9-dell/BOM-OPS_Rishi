from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import tkinter as tk
from tkinter import messagebox

def show_popup(message):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    messagebox.showinfo("Notification", message)
    root.destroy()

def automate_browser(excel_file_path):
    driver = webdriver.Edge()
    driver.get("https://agile.us.dell.com/Agile/")

    time.sleep(10)
    driver.refresh()
    driver.maximize_window()

    time.sleep(5)

    data = pd.read_excel(excel_file_path, sheet_name='Comment')

    for index, row in data.iterrows():
        searchbox_value = row['PES ID']
        commentto_value = row['Comment To']
        commentas_value = row['Comment As']

        # Check for blank values
        if pd.isna(searchbox_value) or pd.isna(commentto_value) or pd.isna(commentas_value):
            show_popup(f"Sheet has an empty row at index: {index + 1}")  
            driver.quit() 
            return 

        searchbox = driver.find_element(By.ID, 'QUICKSEARCH_STRING')
        searchbox.clear()
        searchbox.send_keys(searchbox_value)
        searchbox.send_keys(Keys.RETURN)
        time.sleep(5)

        Comment = driver.find_element(By.ID, 'MSG_Comment')
        Comment.send_keys(Keys.RETURN)
        time.sleep(3)

        initial_window = driver.current_window_handle
        windows = driver.window_handles

        for window in windows:
            if window != initial_window:  
                driver.switch_to.window(window)
                if driver.title == " Comment": 
                    driver.close()  
                    break

        time.sleep(3)

        commentto = driver.find_element(By.ID, 'search_query_notify_display')
        commentto.clear() 
        commentto.send_keys(commentto_value)
        time.sleep(3)
        commentto.send_keys(Keys.RETURN)
        time.sleep(5)

        commentas = driver.find_element(By.ID, 'Comments')
        commentas.clear() 
        commentas.send_keys(commentas_value)
        time.sleep(2)

        sendas = driver.find_element(By.ID, 'save')
        sendas.send_keys(Keys.RETURN)
        time.sleep(2)

        driver.switch_to.window(initial_window)
        time.sleep(3)

    show_popup("Process Completed")

    driver.quit()


if __name__ == "__main__":
    excel_file_path = 'C:/Users/jhar9/Downloads/origi2.xlsx'
    automate_browser(excel_file_path)