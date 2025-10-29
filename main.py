#CODE STRUCTRE
# 1. Modules/Libraries to Import
# 2. Variables and labels initialization (or) declaration
# 3. Button Definitions
# 4. Functions( Main: Update Page)








import os
import tkinter as tk                  #import GUI module
from tkinter import filedialog        # to bring the file-opening possibilities


#                                     inititalising the root-window(GUI)
root=tk.Tk()                          #initialising window object
root.title("Excel to Word")           #window title
root.geometry("1200x600")             # page length-width
root.config(bg="white")               # page background

#Intialising some basic text, to be able to change later on
label=tk.Label(root,text="Hello World",font=("Times New Roman",25),bg="white",fg="#000000")
label.pack(pady=100)
filename=tk.Label(text="")


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
button_next=tk.Button(               #next page(page2)
    root,
    text="NEXT",
    padx=20,
    pady=12,
    font=("Gadugi",10),
    fg="#FFFFFF",
    bg="#000000",
)
button_select=tk.Button(               #button to convert to word
    root,
    text="Select",
    padx=20,
    pady=12,
    font=("Nirmala UI",10),
    fg="#000000",
    bg="#DDFE00",
)
button_proceed=tk.Button(               #button to convert to word
    root,
    text=" Proceed ",
    padx=20,
    pady=12,
    font=("Nirmala UI",10),
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



def update_page():                                            #updates pages
    global page                                               #page variable


    if page==1:                               #page1, default page
        button_prv.place_forget()
        button_select.place_forget() 
        filename.place_forget()
        button_home.place_forget()
        button_proceed.place_forget()

        label.config(text="This is Page1")
        label.pack()
        button_more.place(relx=0.05,rely=0.90,anchor="sw")
        button_next.place(relx=0.95,rely=0.90,anchor="se")
        
    
    if page==2:                                 #page2, conversion page
        label.config(text="This is Page2")
        label.pack()
        filename.config(text="No File Selected",font=("Times New Roman",20),bg="white")
        filename.place(rely=0.50,relx=0.40)
        button_select.place(relx=0.95,rely=0.90,anchor="se")
        button_prv.place(relx=0.05,rely=0.90,anchor="sw")
        button_next.place_forget()
        button_more.place_forget()
        
    if page==3:                                   # final page, page3
        button_proceed.place_forget()
        button_select.place_forget()
        filename.place_forget()
        label.config(text="This is Page3")
        label.pack()
        button_home.place(relx=0.05,rely=0.95,anchor="sw")


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
        
    

def web_more():
    print("Coming Soon")             # pending

#button functions-command bridge
button_next.config(command=next_page)        # command to move to the next page
button_prv.config(command=prv_page)          # command to move to previous page
button_select.config(command=select)         # asking them to select the file(only excel files)
button_proceed.config(command=final_page)    # proceed button after selecting all data
button_home.config(command=home)             # return home 

#final running
update_page()                           # update page
root.mainloop()                         # running the window