#CODE STRUCTRE
# 1. Modules/Libraries to Import
# 2. Variables and labels initialization (or) declaration
# 3. Button Definitions
# 4. Functions( Main: Update Page)







from time import strftime
import os
import tkinter as tk                  #import GUI module
from tkinter import filedialog        # to bring the file-opening possibilities
import excelreader

#                                     inititalising the root-window(GUI)
root=tk.Tk()                          #initialising window object
root.title("nCURES-Excel to Word")           #window title
root.geometry("1200x600")             # page length-width
root.config(bg="white")               # page background

#Intialising some basic text, to be able to change later on
label=tk.Label(root,text="nCURES",font=("Times New Roman",25),bg="white",fg="#000000")
label.pack(pady=80,padx=20)
filename=tk.Label(text="")
timelabel=tk.Label(root,text="Time",font=("Nirmala UI",10))
details=tk.Label(text="Here is some text-details",font=("Times New Roman",15),bg="white",fg="#000000")
#main variables
page=1                               #current page (variable), resposible for page-display



button_convert=tk.Button(            #Convert Button
    root,
    text="",
    padx=20,
    pady=12,
    font=("Century Gothic",10),
    fg="#000000",
    bg="#90EFF6",
)
button_prv=tk.Button(            #previous page(page1)
    root,
    text="PREVIOUS",
    padx=20,
    pady=12,
    font=("Century Gothic",10),
    fg="#FFFFFF",
    bg="#000000",
)
button_more=tk.Button(            #Web Button
    root,
    text="DETAILS",
    padx=20,
    pady=12,
    font=("Century Gothic",10),
    fg="#000000",
    bg="#EEC477",
)
button_echo=tk.Button(               #next page(page2)
    root,
    text="Echocardiography Report", #Echocardiography Report
    padx=20,
    pady=2,
    font=("Gadugi",15),
    fg="#FFFFFF",
    bg="#000000",
)
button_obs=tk.Button(               #next page(page-6)
    root,
    text="Obstetric Ultrasound Report", #Obstetric Ultrasound Report
    padx=12,
    pady=2,
    font=("Gadugi",15),
    fg="#FFFFFF",
    bg="#1C1C1C",
)
button_ni=tk.Button(               #next page(page-7)
    root,
    text="NeuroImaging Report", #NeuroImaging Report
    padx=27,
    pady=6,
    font=("Gadugi",15),
    fg="#FFFFFF",
    bg="#1C1C1C",
)
button_select=tk.Button(               #button to convert to word
    root,
    text="SELECT FILE ",
    padx=5,
    pady=1,
    font=("Nirmala UI",12),
    fg="#000000",
    bg="#DDFE00",
)
button_proceed=tk.Button(               #button to convert to word
    root,
    text="CONVERT ",
    padx=5,
    pady=2,
    font=("Nirmala UI",12),
    fg="#000000",
    bg="#02E63B",
)
button_home=tk.Button(            #home button
    root,
    text="RETURN HOME",
    padx=15,
    pady=5,
    font=("Gadugi",8),
    fg="#000000",
    bg="#17F8A5",
)
button_open=tk.Button(               #button to convert to word
    root,
    text=" OPEN DOCUMENT ",
    padx=10,
    pady=10,
    font=("Nirmala UI",15),
    fg="#000000",
    bg="#D5FF86",
)




def update_page():                                            #updates pages
    global page                                               #page variable

    if page==0:  #details page
        button_echo.place_forget()
        button_more.place_forget()
        button_ni.place_forget()
        button_obs.place_forget()
        label.config(text="Info:")
        
        
        details.config(text="""Utility Details

Utility Name: nCURES- Excel to Word Converter

Purpose:
This utility is designed to automate the conversion of Excel files (.xlsx)
 into structured Word documents (.docx). It provides a simple
efficient interface to extract, format, and export data, 
minimizing manual effort and ensuring consistency in reports or documentation.

Developers:

Team Lead:  Chandrashekhar Rao Kuthyar

Python Backend (Data Processing): Dhyan Kulal

Python Backend (File Operations): Ranjan Shetty

Python GUI (Tkinter Interface): Aditya Nayak
                       

                       For any doubts, contact ncures@googlegroups.com""")
        details.place(relx=0.8, rely=0.90, anchor="se")
        button_prv.place(relx=0.05,rely=0.90,anchor="sw")

    if page==1:                               #page1, default page
        button_prv.place_forget()
        details.place_forget()
        button_select.place_forget() 
        filename.place_forget()
        button_home.place_forget()
        button_proceed.place_forget()
        button_open.place_forget()
        timelabel.place(relx=0.99,rely=0.01,anchor="ne")
        label.config(text="""
        Hello World
                This is page1
        """)
        label.place(relx=0.20,rely=0.10)
        button_more.place(relx=0.05,rely=0.90,anchor="sw")
        button_echo.place(relx=0.95,rely=0.90,anchor="se")
        button_ni.place(relx=0.95,rely=0.80,anchor="se")
        button_obs.place(relx=0.95,rely=0.70,anchor="se")
        
    
    if page==2:                                 #page2, conversion page
        label.config(text="""Excel to Docs Converter
    ~Echocardiography~""")
        label.pack()
        filename.config(text="""Please upload the Excel file you wish to convert.
(Tip: Ensure the Excel file is saved and closed before starting the
                conversion to avoid file access errors.)""",font=("Times New Roman",20),bg="white")
        filename.place(rely=0.50,relx=0.20)
        button_select.place(relx=0.95,rely=0.90,anchor="se")
        button_prv.place(relx=0.05,rely=0.90,anchor="sw")
        button_echo.place_forget()
        button_more.place_forget()
        button_ni.place_forget()
        button_obs.place_forget()
        
    if page==3:                                   # final page, page3
        button_proceed.place_forget()
        button_select.place_forget()
        filename.place_forget()
        label.config(text="This is Page3")
        label.pack()
        button_open.place(relx=0.90,rely=0.90,anchor="se")
        button_open.place(relx=0.90,rely=0.90,anchor="se")
        
        button_home.place(relx=0.05,rely=0.95,anchor="sw")
    if page==5:
        button_echo.place_forget()
        button_more.place_forget()
        button_ni.place_forget()
        button_obs.place_forget()
        label.config(text="""This section is currently under development.
It will be used for NeuroImaging Report and Obstetric Ultrasound Report features, 
                which are planned for future updates.
These options are being designed for upcoming iterations of this product.

Please click “Previous” to return to the home screen.""",font=("Times New Roman",20))
        button_prv.place(relx=0.05,rely=0.90,anchor="sw")

def open_doc(file_path):
    print("Opening Doc File")
    os.startfile(file_path)
def morepage():
    global page
    page=0
    update_page()
def obs_page():
    global page
    page=5
    update_page()
def next_page():                        #next page, page2
    global page
    page=2
    update_page()
def prv_page():                         # previous page, page1
    global page
    page=1
    update_page()
def home():                           # return to home page (page1)
    global page
    page=1
    update_page()
def final_page():                      # page 3 , home
    global page
    page=3
    update_page()
def select():                          #Pending
    file_path=filedialog.askopenfilename(
        title="Select an Excel File : ",
        filetypes=[("Excel Files","*.xlsx")])
    
    file_name=str(os.path.basename(str(file_path)))
    if file_path:
        button_prv.place_forget()
        button_select.place_forget()
        filename.config(text=f"File Selected:{file_name} ✅")
        button_proceed.place(relx=0.90,rely=0.90,anchor="se")
        doc_name=excelreader.reader(file_path)
        print(doc_name)
        update_time()
        button_open.config(command=lambda:open_doc(doc_name))
        
        
def update_time():
    current_time = strftime("%H:%M:%S %p")  # Format like 14:35:22 PM
    timelabel.config(text=current_time)
    timelabel.after(1000, update_time)  # Update every 1 second



def web_more():
    print("Coming Soon")             # pending

#button functions-command bridge
button_echo.config(command=next_page)        # command to move to the next page
button_prv.config(command=prv_page)          # command to move to previous page
button_select.config(command=select)         # asking them to select the file(only excel files)
button_proceed.config(command=final_page)    # proceed button after selecting all data
button_home.config(command=home)             # return home 
button_obs.config(command=obs_page)           # to-do obs page
button_ni.config(command=obs_page)
button_more.config(command=morepage)

#final running
update_time()                           # Start the clock
update_page()                           # update page
root.mainloop()                         # running the window