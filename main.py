import tkinter as tk
from tkinter import ttk, filedialog
import openpyxl
import openexcel


#remove data from treeview
def clear_all():
    for item in treeview.get_children():
          treeview.delete(item)

#upload file
def upload_file():
        path = filedialog.askopenfilename(
            title="Select A File", filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        clear_all()
        load_data(path)

def load_data(path):
        workbook = openpyxl.load_workbook(path)
        #sheet = Select_sheet.get()
        sheet = workbook[Select_sheet.get()]
        #print(sheet)

        list_values = list(sheet.values)
        for col_name in list_values[0]:
            treeview.heading(col_name, text=col_name)

        for value_tuple in list_values[1:]:
            treeview.insert('', tk.END, values=value_tuple)

#main function to sub function
def run_all():
        sheet = Select_sheet.get()
        if sheet == "Dell_Analyst":
              dell_analyst()
        if sheet == "Comment":
              Comment()
        if sheet == "ECO-MOD":
              ecomod()
      
def dell_analyst():
        print("dellanalyst")

def Comment():
        print("comment")
    
def ecomod():
        print("eco")

root = tk.Tk()

style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text = "Upload and Run")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)


#Dropdown for sheets
combo_list = ["Dell_Analyst", "Comment", "ECO-MOD"]

Select_sheet = ttk.Combobox(widgets_frame, values= combo_list)
Select_sheet.current(0)
Select_sheet.grid(row=1, column=0, sticky="ew", padx=10, pady=10)

#upload Button
Upload = ttk.Button(widgets_frame, text="Upload file", command=upload_file)
Upload.grid(row=2, column=0, sticky="nsew", padx=10, pady=10)

#run button
run = ttk.Button(widgets_frame, text="Run", command=run_all)
run.grid(row=3, column=0, sticky="nsew", padx=10, pady=10)

#setting up the frame
treeframe = ttk.Frame(frame)
treeframe.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeframe)
treeScroll.pack(side="right", fill="y")

#table view
cols = ("PES ID", "MOD", "ECO")
treeview = ttk.Treeview(treeframe, show="headings", yscrollcommand=treeScroll.set, columns=cols, height=10)
treeview.column("PES ID", width=100)
treeview.column("MOD", width=50)
treeview.column("ECO",width=50)
treeview.pack()
treeScroll.config(command=treeview.yview)

root.mainloop()