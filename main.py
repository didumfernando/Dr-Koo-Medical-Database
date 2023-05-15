import tkinter as tk
from tkinter import * 
from tkinter import ttk, messagebox, simpledialog, filedialog
from ttkwidgets import LinkLabel
from ttkwidgets.autocomplete import AutocompleteEntry,AutocompleteCombobox
from PIL import Image, ImageTk
import matplotlib.pyplot as plt
import csv
from csv import writer
import pandas as pd
from tkinterdnd2 import *
import warnings
warnings.filterwarnings("ignore", message="normalize=None does not normalize if the sum is less than 1 but this behavior is deprecated since 3.3 until two minor releases later. After the deprecation period the default value will be normalize=True. To prevent normalization pass normalize=False")

# This is the initialisation of the window
window = TkinterDnD.Tk()
window.title("Medical Database Cardiology")
window.iconphoto(False, tk.PhotoImage(file='cardiologist.png'))
window.geometry("1050x760")
window.resizable(0,0) # this is to disallow users from resizing their window
window.tk.call("source", "sun-valley.tcl")
window.tk.call("set_theme", "light")

style = ttk.Style()

# the initialisation of the menubar at the top
menubar = tk.Menu(window)

menuEdit = tk.Menu(menubar,tearoff=0)
menubar.add_cascade(menu=menuEdit, label="Edit")
menuEdit.add_command(label="Undo", command=lambda: window.focus_get().event_generate("<<Undo>>"))
menuEdit.add_command(label="Redo", command=lambda: window.focus_get().event_generate("<<Redo>>"))
menuEdit.add_separator()
menuEdit.add_command(label="Cut", command=lambda: window.focus_get().event_generate("<<Cut>>"))
menuEdit.add_command(label="Copy", command=lambda: window.focus_get().event_generate("<<Copy>>"))
menuEdit.add_command(label="Paste", command=lambda: window.focus_get().event_generate("<<Paste>>"))
menuEdit.add_command(label="Clear", command=lambda: window.focus_get().event_generate("<<Clear>>"))
menuEdit.add_command(label="Go To Start", command=lambda: window.focus_get().event_generate("<LineStart>>"))
menuEdit.add_command(label="Go To End", command=lambda: window.focus_get().event_generate("<<LineEnd>>"))
menuEdit.add_command(label="Select All", command=lambda: window.focus_get().event_generate("<<SelectAll>>"))
menuEdit.add_command(label="Select Line Start", command=lambda: window.focus_get().event_generate("<<SelectLineStart>>"))
menuEdit.add_command(label="Select Line End", command=lambda: window.focus_get().event_generate("<<SelectLineEnd>>"))
menuEdit.add_command(label="Select None", command=lambda: window.focus_get().event_generate("<<SelectNone>>"))

menuEdit.entryconfigure('Undo', accelerator='Command+Z')
menuEdit.entryconfigure('Redo', accelerator='Command+Shift+Z')
menuEdit.entryconfigure('Cut', accelerator='Command+X')
menuEdit.entryconfigure('Copy', accelerator='Command+C')
menuEdit.entryconfigure('Paste', accelerator='Command+V')
menuEdit.entryconfigure('Select All', accelerator='Command+A')

database = tk.Menu(menubar,tearoff=0)
menubar.add_cascade(label="Database",menu=database)
database.add_command(label="Back",command=lambda: back())
database.add_command(label="Home",command=lambda: goToMainPage())
database.add_command(label='Quit',command=window.destroy)
database.entryconfigure('Quit', accelerator='Command+W')

help_ = tk.Menu(menubar, tearoff = 0)
menubar.add_cascade(label='Help', menu = help_)
help_.add_command(label='Instructions for Usage', command=lambda: help_())
help_.add_separator()
help_.add_command(label ='About the Program', command= lambda: goToAbout())

mainOn = 0
settingsOn = 0

# help button
def help_():
    global mainOn
    global settingsOn
    if mainFrame.winfo_ismapped():
        mainFrame.grid_forget()
        mainFrameHelp.grid(row = 0,column = 0)
        mainOn = 1
    elif settingsFrame.winfo_ismapped():
        settingsFrame.grid_forget()
        settingsFrameHelp.grid(row = 0,column = 0)
        settingsOn = 1

def back():
    if mainFrame.winfo_ismapped():
        mainFrame.grid_forget()
        settingsFrame.grid(row=0,column=0)
    elif settingsFrame.winfo_ismapped():
        settingsFrame.grid_forget()
        mainFrame.grid(row=0,column=0)
    elif mainFrameHelp.winfo_ismapped() or settingsFrameHelp.winfo_ismapped():
        mainFrameHelp.grid_forget()
        settingsFrameHelp.grid_forget()
        mainFrame.grid(row=0,column=0)

#editOn = False
addOn = False

#editOpen = False
addOpen = False

noInvalidAddOptions = False
#noInvalidEditOptions = False

invalidAddEntry = False
#invalidEditEntry = False

validGenderCombobox = False
validageSpinbox = False
validsmokingStatusCombobox = False
validbloodPressureSpinbox = False

counter = 0

detached = []
detachedIndexes = []
allRows = []
valuesList = []
filtersList = []

# special characters to be not allowed
specialCharacters = ['`','~','!','@','#','$','%','^','&','*','_','+','=','{','}',':',';','.','<','>','/','\\','|','√','∫','≈','Ω','ß','∂','ƒ','©','˙','∆','˚','¬','œ','∑','´','®','†','¥','¨','ˆ','ø','π','ç','µ']

# initialisation of all frames
mainFrame = tk.Frame()
mainFrame.configure(bg="#fafafa")

mainFrameHelp = tk.Frame()

settingsFrame = tk.Frame()
settingsFrame.configure(bg="#fafafa")

settingsFrameHelp = tk.Frame() 

addNewFrame = tk.Frame(master=mainFrame,bg="#fafafa")

#editFrame = tk.Frame(master=mainFrame,bg="#fafafa")

aboutFrame = tk.Frame(bg="#fafafa")

mainFrame.grid(row=0,column=0)

tvFrame = tk.Frame(master=mainFrame,bg="#fafafa")
tvFrame.grid(row=3,column=2,columnspan=4,rowspan=10)

searchFrame = tk.Frame(master=mainFrame,bg="#fafafa")
searchFrame.grid(row=0,column=2,columnspan=4)

statusFrame = tk.Frame(master=mainFrame,bg="#1FE41C")

statusFramePlaceholder = ttk.Label(master=statusFrame,text="")
statusFramePlaceholder.grid(row=100,column=100)

# function to go to the settings page
def goToSettingsPage():
    global mainOn
    global settingsOn
    mainFrame.grid_forget()
    settingsFrame.grid(row=0,column=0)
    mainOn = 0
    settingsOn = 1

# function to return to the main page
def goToMainPage():
    global mainOn
    global settingsOn
    settingsFrame.grid_forget()
    aboutFrame.grid_forget()
    mainFrame.grid(row=0,column=0)
    settingsOn = 0
    mainOn = 1

# function to go to the about page
def goToAbout():
    global settingsOn
    global mainOn
    if settingsFrame.winfo_ismapped() == 1:
        settingsOn = 1
    elif mainFrame.winfo_ismapped() == 1:
        mainOn = 1
    settingsFrame.grid_forget()
    mainFrame.grid_forget()
    aboutFrame.grid(row=0,column=0)

def done():
    global mainOn
    global settingsOn
    if mainOn == 1:
        aboutFrame.grid_forget()
        mainFrame.grid(row=0,column=0)
        mainFrameHelp.grid_forget()
    elif settingsOn == 1:
        aboutFrame.grid_forget()
        settingsFrame.grid(row=0,column=0)
        settingsFrameHelp.grid_forget()

# about the program
about = '''This program is mainly a a program made for showing the risk score of a patient developing a stroke. The code
converts an excel file into a csv file and then imports the data automatically for display for easier viewing. The
program then calculates the risk score for each individual patient automatically and displays it on the display. In
addition, the program allows to import the data off the excel file automatically to be manipulated.


Developers: Timothy, Didum, Pranav, Dharshan 
    

Credits:'''
aboutText = ttk.Label(master=aboutFrame,text=about,font=("Arial", 21))
aboutText.grid(row=0,column=0,padx=5)

# Credits + links
logoLabel = LinkLabel(aboutFrame,text="Logo for database",link="https://www.student-crm.co.uk/student-database",normal_color='royal blue',hover_color='blue',clicked_color='purple')
backLabel = LinkLabel(aboutFrame,text="Back Icon",link="https://www.visualpharm.com/free-icons/back-595b40b75ba036ed117d877b",normal_color='royal blue',hover_color='blue',clicked_color='purple')
searchLabel = LinkLabel(aboutFrame,text="Search Icon",link="https://upload.wikimedia.org/wikipedia/commons/thumb/7/7e/Vector_search_icon.svg/945px-Vector_search_icon.svg.png",normal_color='royal blue',hover_color='blue',clicked_color='purple')
settingsIconLabel = LinkLabel(aboutFrame,text="Settings Icon",link="https://www.pinclipart.com/picdir/big/65-654632_setting-icon-clipart-png-download.png",normal_color='royal blue',hover_color='blue',clicked_color='purple')
w11Label = LinkLabel(aboutFrame,text="Windows 11 Theme",link="https://www.reddit.com/r/Python/comments/ot4zzx/make_tkinter_look_like_windows_11/",normal_color='royal blue',hover_color='blue',clicked_color='purple')

logoLabel.grid(row=1,column=0,sticky="W",padx=5)
backLabel.grid(row=2,column=0,sticky="W",padx=5)
searchLabel.grid(row=3,column=0,sticky="W",padx=5)
settingsIconLabel.grid(row=4,column=0,sticky="W",padx=5)
w11Label.grid(row=5,column=0,sticky="W",padx=5)

doneButton = ttk.Button(master=aboutFrame,text="Done",command=done,style="Accent.TButton")
doneButton.grid(row=20,column=0,sticky="W")

# Picture for help with main frame
imM = Image.open("main help.png")
imM = imM.resize((1050,700),Image.ANTIALIAS)
mainHelpPic = ImageTk.PhotoImage(imM)
mainHelpLabel = tk.Label(master=mainFrameHelp,image=mainHelpPic,compound="c")
mainHelpLabel.grid(row=0,column=0)

rowHelpLabel = ttk.Label(master=mainFrameHelp,text="How to Select rows in the treeview: Select Single Row: Left-click\nSelect Multiple rows at once / Deselect row: Cmd/Ctrl + left-click each individual row")
rowHelpLabel.grid(row=1,column=0)

# picture for help with settings frame
imH = Image.open("settings help.png")
imH = imH.resize((1050,700),Image.ANTIALIAS)
settingsHelpPic= ImageTk.PhotoImage(imH)
settingsHelpLabel = tk.Label(master=settingsFrameHelp,image=settingsHelpPic,compound="c")
settingsHelpLabel.grid(row=0,column=0)

doneSettings = ttk.Button(master=settingsFrameHelp,text="Done",command=done,style="Accent.TButton")
doneSettings.grid(row=1,column=0)

doneMain = ttk.Button(master=mainFrameHelp,text="Done",command=done,style="Accent.TButton")
doneMain.grid(row=2,column=0)

# settings icon for settings button
im = Image.open("settings icon.png")
im = im.resize((50,50),Image.ANTIALIAS)
settingsIcon = ImageTk.PhotoImage(im)
settingsButton = tk.Button(master=mainFrame,image=settingsIcon,compound="c",command=goToSettingsPage)
settingsButton.grid(row=0,column=0,sticky="nw")
im3 = Image.open("search icon.png")
im3 = im3.resize((20,20),Image.ANTIALIAS)
searchIcon = ImageTk.PhotoImage(im3)
searchIconLabel = tk.Label(master=searchFrame,image=searchIcon,width=17,height=17,compound="c")
searchIconLabel.grid(row=0,column=0,sticky="E")

# Search box placeholder
def focusOutEntryBox(widget,widgetText):
    if len((widget.get()).replace(" ","")) == 0 or (widget.get()).replace(" ","") == placeholder:
        widget.delete(0,tk.END)
        widget.insert(0,widgetText)

def focusInEntryBox(widget):
    if (widget.get()).replace(" ","") == placeholder:
        widget.delete(0,tk.END)
        
# Search function
def search(event):
    global detachedIndexes
    global allRows
    global filtersList
    counter = 0
    i = 0
    criteria = searchBox.get()
    filtersList = []

    for child in allRows:
        counter = 0
        for item in databaseTree.item(child,option='values'):

            if (criteria.lower().strip()) in item.lower().strip():
                
                if filtersList == []:
                    databaseTree.move(child,databaseTree.parent(child),databaseTree.index(child))
                    break

                else:
                    if item not in filtersList:
                        counter += 1
            else:

                counter += 1

            if filtersList == []:
                pass
            else:
                if item in filtersList:
                    break

            if counter >= 5 or i >= 5:
                databaseTree.detach(child)

    if len(criteria.replace(" ","")) == 0 or criteria == "":
        for row in allRows[::-1]:
            databaseTree.detach(row)    
            databaseTree.move(row,databaseTree.parent(row),databaseTree.index(row))

searchText = tk.StringVar()

# Search box initialisation
searchBox = ttk.Entry(master=searchFrame,textvariable=searchText,width=70)

placeholder = "Search"
searchBox.insert(0,placeholder)
searchBox.bind("<FocusIn>", lambda args: focusInEntryBox(searchBox))
searchBox.bind("<FocusOut>", lambda args: focusOutEntryBox(searchBox,placeholder))
searchBox.bind("<KeyRelease>", search)
searchBox.grid(row=0,column=1,columnspan=8,pady=10)

# Function to add new entries
def addNew():
    global addOpen

    addOpen = True
    addNewButton.grid_forget()
    collapseAddButton.grid(row=3,column=0,sticky="NW")
    addNewFrame.grid(row=4,column=0,columnspan=6)
    databaseTree.grid_forget()
    tvFrame.grid_forget()
    searchFrame.grid_forget()
    searchFrame.grid(row=0,column=7,columnspan=10)
    tvFrame.grid(row=3,column=7,rowspan=10,columnspan=10)
    databaseTree.grid(row=1,column=0,columnspan=15)

# Function to close the adding window
def collapseAddFrame():
    global addOpen

    addOpen = False
    collapseAddButton.grid_forget()
    addNewButton.grid(row=3,column=0,sticky="NW")
    addNewFrame.grid_forget()
    databaseTree.grid_forget()
    tvFrame.grid_forget()
    searchFrame.grid_forget()
    searchFrame.grid(row=0,column=4)
    tvFrame.grid(row=3,column=2,columnspan=5,rowspan=10)
    databaseTree.grid(row=1,column=0,ipadx=150,columnspan=15)
        
warningAddLabel = tk.Label(master=addNewFrame,text="No numbers and special characters allowed",fg="#FF0000",font=("Arial",9))

# Validation: Format Check
def invalidEntry(event):
    global validGenderCombobox
    global validageSpinbox
    global validsmokingStatusCombobox
    global validbloodPressureSpinbox
    global invalidAddEntry
    global noInvalidAddOptions
    global specialCharacters
    if addNameEntry.get().strip() == "" or addNameEntry.get().isalnum() == True or " " in addNameEntry.get():
        addNameEntry.state(["!invalid"])
        warningAddLabel.grid_forget()
        h = 1

        if addNameEntry.get().strip() == "":
            add["state"] = "disabled"

        else:
            for specialCharacter in specialCharacters:
                if specialCharacter in addNameEntry.get():
                    addNameEntry.state(["invalid"])
                    warningAddLabel.grid(row=8,column=0)
                    h = 0
                    add["state"] = "disabled"
                    invalidAddEntry = False
                    break

            for character in addNameEntry.get():
                if character.isdigit() == True:
                    addNameEntry.state(["invalid"])
                    warningAddLabel.grid(row=8,column=0)
                    h = 0
                    add["state"] = "disabled"
                    invalidAddEntry = False
                    break
            if noInvalidAddOptions == True and invalidAddEntry == False:
                add["state"] = "enabled"
    else:
        addNameEntry.state(["invalid"])
        warningAddLabel.grid(row=8,column=0)
        h = 0
        add["state"] = "disabled"
        invalidAddEntry = False
    if validGenderCombobox == True and validageSpinbox == True and validsmokingStatusCombobox == True and validbloodPressureSpinbox == True and invalidAddEntry == False and h == 1 and addNameEntry.get().strip() != "" and genderCombobox.get() != "" and ageSpinbox.get().strip() != "" and smokingStatusCombobox.get().strip() != "" and bloodPressureSpinbox.get().strip() != "":
        add["state"] = "enabled"

    else:
        add["state"] = "disabled"

warningInvalidAddOption = tk.Label(master=addNewFrame,text="Invalid Option Chosen",fg="#FF0000",font=("Arial",9))
noTextAllowedAdd = tk.Label(master=addNewFrame,text="No text allowed",fg="#FF0000",font=("Arial",9))
outOfRangeAdd = tk.Label(master=addNewFrame,text="Option chosen is out of range",fg="#FF0000",font=("Arial",9))
noNumbersAllowedAdd = tk.Label(master=addNewFrame,text="No numbers allowed",fg="#FF0000",font=("Arial",9))

genderList = ["M","F"]
statusList = ["N","Y"]

# Validation: Range and Format check
def invalidCombo(event):
    global validGenderCombobox
    global validageSpinbox
    global validsmokingStatusCombobox
    global validbloodPressureSpinbox
    global validcholestrolSpinbox
    global invalidAddEntry
    global noInvalidAddOptions
    noTextAllowedAdd.grid_forget()
    noNumbersAllowedAdd.grid_forget()
    outOfRangeAdd.grid_forget()
    #noTextAllowedEdit.grid_forget()
    #outOfRangeEdit.grid_forget()
    #noNumbersAllowedEdit.grid_forget()
    if genderCombobox.get().strip() == "" or genderCombobox.get() in genderList:
        genderCombobox.state(["!invalid"])
        
        if genderCombobox.get().strip() == "":
            add["state"] = "disabled"
            validGenderCombobox = False

        else:
            validGenderCombobox = True

    else:
        genderCombobox.state(["invalid"])
        add["state"] = "disabled"
        validGenderCombobox = False
        if genderCombobox.get().strip().isdigit() == True:
            noNumbersAllowedAdd.grid(row=7,column=0)

        elif genderCombobox.get() not in genderList:
            outOfRangeAdd.grid(row=7,column=0)

    if ageSpinbox.get().strip() == "" or ageSpinbox.get().strip().isdigit():
        ageSpinbox.state(["!invalid"])

        if ageSpinbox.get().strip() == "":
            add["state"] = "disabled"
            validageSpinbox = False
            
        elif int(ageSpinbox.get().strip()) > 140:
            ageSpinbox.state(["invalid"])
            add["state"] = "disabled"
            outOfRangeAdd.grid(row=7,column=0)
            validageSpinbox = True
        else:
            validageSpinbox = True
            pass
            
    else:
        ageSpinbox.state(["invalid"])
        add["state"] = "disabled"
        validageSpinbox = False
        if not ageSpinbox.get().strip().isdigit():
            noTextAllowedAdd.grid(row=7,column=0)

    if smokingStatusCombobox.get().strip() == "" or smokingStatusCombobox.get() in statusList:
        smokingStatusCombobox.state(["!invalid"])
        if smokingStatusCombobox.get().strip() == "":
            add["state"] = "disabled"
            validsmokingStatusCombobox = False
        else:
            validsmokingStatusCombobox = True

    else:
        smokingStatusCombobox.state(["invalid"])
        add["state"] = "disabled"
        validsmokingStatusCombobox = False
        if smokingStatusCombobox.get().isalpha():
            outOfRangeAdd.grid(row=7,column=0)
        elif not smokingStatusCombobox.get().isalpha():
            noNumbersAllowedAdd.grid(row=7,column=0)

    if bloodPressureSpinbox.get().strip() == "" or bloodPressureSpinbox.get().isdigit():
        bloodPressureSpinbox.state(["!invalid"])
        if bloodPressureSpinbox.get().strip() == "":
            add["state"] = "disabled"
            validbloodPressureSpinbox = False
        elif float(bloodPressureSpinbox.get().strip()) > 180 or float(bloodPressureSpinbox.get().strip()) < 140:
            bloodPressureSpinbox.state(["invalid"])
            add["state"] = "disabled"
            outOfRangeAdd.grid(row=7,column=0)
            validbloodPressureSpinbox = False
        else:
            validbloodPressureSpinbox = True
        
    else:
        try:
            float(bloodPressureSpinbox.get().strip())
            validbloodPressureSpinbox = True
        except:
            bloodPressureSpinbox.state(["invalid"])
            add["state"] = "disabled"
            validbloodPressureSpinbox = False
            noTextAllowedAdd.grid(row=7,column=0)


    if cholestrolSpinbox.get().strip() == "" or cholestrolSpinbox.get().isdigit():
        cholestrolSpinbox.state(["!invalid"])
        if cholestrolSpinbox.get().strip() == "":
            add["state"] = "disabled"
            validcholestrolSpinbox = False
        elif float(cholestrolSpinbox.get().strip()) > 8 or float(cholestrolSpinbox.get().strip()) < 4:
            validcholestrolSpinbox = False
            add["state"] = "disabled"
            cholestrolSpinbox.state(["invalid"])
            outOfRangeAdd.grid(row=7,column=0)
        else:
            validcholestrolSpinbox = True
    else:
        try:
            float(cholestrolSpinbox.get().strip())
            validcholestrolSpinbox = True
        except:
            cholestrolSpinbox.state(["invalid"])
            add["state"] = "disabled"
            validcholestrolSpinbox = False
            noTextAllowedAdd.grid(row=7,column=0)

    if validGenderCombobox == True and validageSpinbox == True and validsmokingStatusCombobox == True and validbloodPressureSpinbox == True and validcholestrolSpinbox == True and invalidAddEntry == False and addNameEntry.get().strip() != "" and genderCombobox.get() != "" and ageSpinbox.get().strip() != "" and smokingStatusCombobox.get().strip() != "" and bloodPressureSpinbox.get().strip() != "":
        add["state"] = "enabled"

allAddFieldsRequired = tk.Label(master=addNewFrame,text="*All fields required",fg="#FF0000",font=("Arial",9))
allAddFieldsRequired.grid(row=0,column=0,sticky="W")

addNewButton = ttk.Button(master=mainFrame,text="Add New +",command=addNew)
addNewButton.grid(row=3,column=0,sticky="NW")

collapseAddButton = ttk.Button(master=mainFrame,text="Collapse -",command=collapseAddFrame)

genderLabel = ttk.Label(master=addNewFrame,text="Gender")
genderLabel.grid(row=1,column=1,sticky="W",padx=15)

genderCombobox = ttk.Combobox(master=addNewFrame,width=5)
genderCombobox["values"] = ("M","F")
genderCombobox.bind("<FocusOut>",invalidCombo)
genderCombobox.bind("<FocusIn>",invalidCombo)
genderCombobox.bind("<KeyRelease>",invalidCombo)
genderCombobox.bind("<<ComboboxSelected>>",invalidCombo)
genderCombobox.grid(row=2,column=1,padx=15,sticky="W")

addNameLabel = ttk.Label(master=addNewFrame,text="Name")
addNameLabel.grid(row=1,column=0,sticky="W",pady=10,padx=10)

addNameEntry = ttk.Entry(master=addNewFrame)
addNameEntry.grid(row=2,column=0,padx=10,sticky="W")
addNameEntry.bind("<FocusOut>",invalidEntry)
addNameEntry.bind("<FocusIn>",invalidEntry)
addNameEntry.bind("<KeyRelease>",invalidEntry)

ageLabel = ttk.Label(master=addNewFrame,text="Age")
ageLabel.grid(row=3,column=0,sticky="W",padx=10,pady=10)

ageSpinbox = ttk.Spinbox(master=addNewFrame,from_=40,to=69,width=5)
#ageSpinbox["values"] = ("1C","1H","1A","1M","1I","1O","1N","2C","2H","2A","2M","2I","2O","2N","3C","3H","3A","3M","3I","3O","3N","3S","4C","4H","4A","4M","4P","4I","4O","4N","4S","5O")
ageSpinbox.bind("<FocusOut>",invalidCombo)
ageSpinbox.bind("<FocusIn>",invalidCombo)
ageSpinbox.bind("<KeyRelease>",invalidCombo)
ageSpinbox.bind("<<SpinboxSelected>>",invalidCombo)
ageSpinbox.grid(row=4,column=0,padx=10,sticky="W")

smokingStatusLabel = ttk.Label(master=addNewFrame,text="Smoking Status")
smokingStatusLabel.grid(row=3,column=1,sticky="W",padx=15)

#levelList = ["Sec 1","Sec 2","Sec 3","Sec 4","Sec 5"]
smokingStatusCombobox = ttk.Combobox(master=addNewFrame,width=5)
smokingStatusCombobox["values"] = ("N","Y")
smokingStatusCombobox.bind("<FocusOut>",invalidCombo)
smokingStatusCombobox.bind("<FocusIn>",invalidCombo)
smokingStatusCombobox.bind("<KeyRelease>",invalidCombo)
smokingStatusCombobox.bind("<<ComboboxSelected>>",invalidCombo)
smokingStatusCombobox.grid(row=4,column=1,sticky="W",padx=15)

bloodPressureLabel = ttk.Label(master=addNewFrame,text="Blood Pressure")
bloodPressureLabel.grid(row=5,column=0,sticky="W",padx=10,pady=10)

bloodPressureSpinbox = ttk.Spinbox(master=addNewFrame,from_=120.00,to=180.00,increment=0.01,width=7)
bloodPressureSpinbox.bind("<FocusOut>",invalidCombo)
bloodPressureSpinbox.bind("<FocusIn>",invalidCombo)
bloodPressureSpinbox.bind("<KeyRelease>",invalidCombo)
bloodPressureSpinbox.bind("<<Increment>>",invalidCombo)
bloodPressureSpinbox.bind("<<Decrement>>",invalidCombo)
bloodPressureSpinbox.bind("<<ComboboxSelected>>",invalidCombo)
bloodPressureSpinbox.grid(row=6,column=0,padx=10,sticky="W")

cholestrolLabel = ttk.Label(master=addNewFrame,text="Cholestrol")
cholestrolLabel.grid(row=5,column=1,sticky="W",padx=15)

cholestrolSpinbox = ttk.Spinbox(master=addNewFrame,from_=4.00,to=8.00,increment=0.01,width=5)
cholestrolSpinbox.bind("<FocusOut>",invalidCombo)
cholestrolSpinbox.bind("<FocusIn>",invalidCombo)
cholestrolSpinbox.bind("<KeyRelease>",invalidCombo)
cholestrolSpinbox.bind("<<Increment>>",invalidCombo)
cholestrolSpinbox.bind("<<Decrement>>",invalidCombo)
cholestrolSpinbox.bind("<<SpinboxSelected>>",invalidCombo)
cholestrolSpinbox.grid(row=6,column=1,padx=15,sticky="W")

# Add data to treeview
def addData():
    global counter
    if invalidAddEntry == False and addNameEntry.get().strip() != "" and genderCombobox.get() != "" and ageSpinbox.get().strip() != "" and smokingStatusCombobox.get().strip() != "" and bloodPressureSpinbox.get().strip() != "":
        add["state"] = "enabled"
        global idNumber
        global allRows
        idNumber += 1
        ID = idNumber
        name = addNameEntry.get()
        gender = genderCombobox.get()
        age = ageSpinbox.get()
        smokingStatus = smokingStatusCombobox.get()
        bloodPressure = bloodPressureSpinbox.get()
        cholestrol = cholestrolSpinbox.get()
        risk = 0
        
        '''#this is getting the data from the boxes inputted by the user
        dataList = [ID,name, gender, age, smokingStatus, bloodPressure,cholestrol,risk] # the data collected is put into a list
        
        with open("patients' data csv.csv", 'a', newline = '') as f_object:
            writer_object = writer(f_object)
            writer_object.writerow(dataList) # and then it is written
            f_object.close()'''

        row = [name,gender,int(age),smokingStatus,float(bloodPressure),float(cholestrol),risk]
        
        tag = ("",)
        
        if row[1] == "M":
            if row[3] == "N":
                if row[2] >= 65:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 14
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 12
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 10
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 9
                            tag = ("high risk",)
                        elif row[5] >= 4:
                            row[6] = 8
                            tag = ("high risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 10
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 5
                            tag = ("medium risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 4
                            tag = ("medium risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                elif row[2] >= 60 and row[2] < 65:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 9
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 5
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                elif row[2] >= 55 and row[2] < 60:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                elif row[2] >= 50 and row[2] < 55:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                elif row[2] >= 40 and row[2] < 50:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 0
                            tag = ("low risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 0
                            tag = ("low risk",)
                    elif row[4] >= 140 or row[4] >= 120:
                        row[6] = 0
                        tag = ("low risk",)
            if row[3] == "Y":
                if row[2] >= 65:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 26
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 23
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 20
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 17
                            tag = ("high risk",)
                        elif row[5] >= 4:
                            row[6] = 15
                            tag = ("high risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 19
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 16
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 14
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 12
                            tag = ("high risk",)
                        elif row[5] >= 4:
                            row[6] = 10
                            tag = ("high risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 13
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 11
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 9
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 4:
                            row[6] = 7
                            tag = ("medium risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 9
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 5
                            tag = ("medium risk",)
                elif row[2] >= 60 and row[2] < 65:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 18
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 15
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 13
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 11
                            tag = ("high risk",)
                        elif row[5] >= 4:
                            row[6] = 10
                            tag = ("high risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 13
                        elif row[5] >= 7:
                            row[6] = 11
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 9
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 4:
                            row[6] = 7
                            tag = ("medium risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 9
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 5
                            tag = ("medium risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                elif row[2] >= 55 and row[2] < 60:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 12
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 10
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 6
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 4
                            tag = ("medium risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif choles3rol == 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif cholesrol == 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                elif row[2] >= 50 and row[2] < 55:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 12
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 10
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 6
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 4
                            tag = ("medium risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                elif row[2] >= 40 and row[2] < 50:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[4] >= 160:
                        row[6] = 1
                        tag = ("low risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 0
                            tag = ("low risk",)                
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 0
                            tag = ("low risk",)
        elif row[1] == "F":
            if row[3] == "N":
                if row[2] >= 65:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 4
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                elif row[2] >= 60 and row[2] < 65:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[4] >= 120:
                        row[6] = 1
                        tag = ("low risk",)
                elif row[2] >= 55 and row[2] < 60:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[4] >= 160:
                        row[6] = 1
                        tag = ("low risk",)
                    elif row[4] >= 140:
                        row[6] = 1
                        tag = ("low risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 0
                            tag = ("low risk",)
                elif row[2] >= 50 and row[2] < 55:
                    if row[4] >= 180:
                        row[6] = 1
                        tag = ("low risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 0
                            tag = ("low risk",)
                    elif row[4] >= 140:
                        row[6] = 0
                        tag = ("low risk",)
                    elif row[4] >= 120:
                        row[6] = 0
                        tag = ("low risk",)
                elif row[2] >= 40 and row[2] < 50:
                    row[6] = 0
                    tag = ("low risk",)
            if row[3] == "Y":
                if row[2] >= 65:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 14
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 12
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 11
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 9
                            tag = ("high risk",)
                        elif row[5] >= 4:
                            row[6] = 9
                            tag = ("high risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 10
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 6
                            tag = ("medium risk",)
                    elif row[4] >= 140 or row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                elif row[2] >= 60 and row[2] < 65:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 5
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                elif row[2] >= 55 and row[2] < 60:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[4] >= 120:
                        row[6] = 1
                        tag = ("low risk",)
                elif row[2] >= 50 and row[2] < 55:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[4] >= 160 or row[4] >= 140:
                        row[6] = 1
                        tag = ("low risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 0
                            tag = ("low risk",)
                elif row[2] >= 40 and row[2] < 50:
                    row[6] = 0
                    tag = ("low risk",)

        risk = row[6]
        
        #this is getting the data from the boxes inputted by the user
        dataList = [ID,name, gender, age, smokingStatus, bloodPressure,cholestrol,risk] # the data collected is put into a list
        
        with open("patients' data csv.csv", 'a', newline = '') as f_object:
            writer_object = writer(f_object)
            writer_object.writerow(dataList) # and then it is written
            f_object.close()

        databaseTree.tag_configure("high risk",background='red',foreground="white")
        databaseTree.tag_configure("medium risk",background="orange")
        databaseTree.tag_configure("low risk",background="#00FF17")
        
        databaseTree.insert(parent='',index='end',iid=idNumber,text=idNumber,values=tuple(row),tags=tag)
        allRows.append(databaseTree.get_children()[counter])
        counter += 1
        valuesList.append(databaseTree.item(idNumber,option='values'))
        addNameEntry.delete(0,'end')
        genderCombobox.set('')
        ageSpinbox.set('')
        smokingStatusCombobox.set('')
        bloodPressureSpinbox.set('')
        cholestrolSpinbox.set('')
        add["state"] = "disabled"

    else:
        add["state"] = "disabled"

def openFile():
    global databaseTree
    #global treeviewScrollbar
    global counter
    global idNumber
    global allRows
    global valuesList
    filename = filedialog.askopenfilename(
initialdir="C:/gui/",
title = "Open A File",filetypes=(("xlsx files", "*.xlsx"), ("All Files", "*.*"))
)

    if filename:
        try:
            filename = r"{}".format(filename)
            df = pd.read_excel(filename)
        except ValueError:
            my_label.config(text="File must be xlsx format.")
        except FileNotFoundError:
            my_label.config(text="File Couldn't Be Found...please choose a file againtry again!")
    else:
        return
    excel_name = filename[filename.rfind("/")+1:]
    try:
        read_file = pd.read_excel(excel_name)
        read_file.to_csv("patients' data csv.csv",
                         index = None,
                         header = True)
        df = pd.DataFrame(pd.read_csv("patients' data csv.csv"))

    except FileNotFoundError:
        messagebox.showwarning(message="File Not Found",detail="Excel file needs to be placed in same folder as the program!")
        return

    # Clear old treeview
    databaseTree.destroy()
    valuesList = []

    
    file_name = "patients' data csv.csv"
    file = open(file_name)
    file_reader = csv.reader(file)
    data = list(file_reader)
    print(data)

    '''if data[0][0] != 'IID':
        a = 'IID'
        data[0].insert(0,a)
        for i in range(1,len(data)):
            num = str(i)
            data[i].insert(0,num)
    
        with open("patients' data csv.csv",'w', newline = '') as writeFile:
            writer_object = writer(writeFile)
            writer_object.writerows(data)
            writeFile.close()'''

    # initialise treeview
    databaseTree = ttk.Treeview(master=tvFrame,height=25,selectmode="extended",show="headings")
    #treeviewScrollbar.destroy()
    # initialise and connect scrollbar to treeview
    #treeviewScrollbar = ttk.Scrollbar(orient='vertical',command=databaseTree.yview)
    #treeviewScrollbar.grid(row=1,column=16,sticky=(N,S))
    databaseTree.configure(yscrollcommand = treeviewScrollbar.set)

    # Define the columns
    databaseTree["columns"] = ("Name","Sex","Age","Smoking Status","Blood Pressure","Cholestrol","Risk")

    # Format the columns
    databaseTree.column("#0",anchor=CENTER,width=50,minwidth=50)
    databaseTree.column("Sex",anchor=CENTER,width=30,minwidth=30)
    databaseTree.column("Age",anchor=CENTER,width=30,minwidth=30)
    databaseTree.column("Name",anchor=W,width=230,minwidth=60)
    databaseTree.column("Smoking Status",anchor=CENTER,width=80,minwidth=90)
    databaseTree.column("Blood Pressure",anchor=CENTER,width=50,minwidth=90)
    databaseTree.column("Cholestrol",anchor=CENTER,width=60,minwidth=60)
    databaseTree.column("Risk",anchor=CENTER,width=30,minwidth=30)

    # Create Headings
    databaseTree.heading("#0",text="#",anchor=CENTER)
    databaseTree.heading("Sex",text="Sex",anchor=CENTER)
    databaseTree.heading("Age",text="Age",anchor=CENTER)
    databaseTree.heading("Name",text="Name",anchor=W)
    databaseTree.heading("Smoking Status",text="Smoking Status",anchor=CENTER)
    databaseTree.heading("Blood Pressure",text="Blood Pressure",anchor=CENTER)
    databaseTree.heading("Cholestrol",text="Cholestrol",anchor=CENTER)
    databaseTree.heading("Risk",text="Risk",anchor=CENTER)
    df_rows = df.to_numpy().tolist()

    idNumber = 0
    counter += len(df_rows)
    print(df_rows)
    yuiop = 1
    for row in df_rows:

        tag = ("",)
        
        if row[1] == "M":
            if row[3] == "N":
                if row[2] >= 65:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 14
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 12
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 10
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 9
                            tag = ("high risk",)
                        elif row[5] >= 4:
                            row[6] = 8
                            tag = ("high risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 10
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 5
                            tag = ("medium risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 4
                            tag = ("medium risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                elif row[2] >= 60 and row[2] < 65:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 9
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 5
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                elif row[2] >= 55 and row[2] < 60:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                elif row[2] >= 50 and row[2] < 55:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                elif row[2] >= 40 and row[2] < 50:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 0
                            tag = ("low risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 0
                            tag = ("low risk",)
                    elif row[4] >= 140 or row[4] >= 120:
                        row[6] = 0
                        tag = ("low risk",)
            if row[3] == "Y":
                if row[2] >= 65:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 26
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 23
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 20
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 17
                            tag = ("high risk",)
                        elif row[5] >= 4:
                            row[6] = 15
                            tag = ("high risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 19
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 16
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 14
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 12
                            tag = ("high risk",)
                        elif row[5] >= 4:
                            row[6] = 10
                            tag = ("high risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 13
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 11
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 9
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 4:
                            row[6] = 7
                            tag = ("medium risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 9
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 5
                            tag = ("medium risk",)
                elif row[2] >= 60 and row[2] < 65:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 18
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 15
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 13
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 11
                            tag = ("high risk",)
                        elif row[5] >= 4:
                            row[6] = 10
                            tag = ("high risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 13
                        elif row[5] >= 7:
                            row[6] = 11
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 9
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 4:
                            row[6] = 7
                            tag = ("medium risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 9
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 5
                            tag = ("medium risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                elif row[2] >= 55 and row[2] < 60:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 12
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 10
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 6
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 4
                            tag = ("medium risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif choles3rol == 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif cholesrol == 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                elif row[2] >= 50 and row[2] < 55:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 12
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 10
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 6
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 4
                            tag = ("medium risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                elif row[2] >= 40 and row[2] < 50:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[4] >= 160:
                        row[6] = 1
                        tag = ("low risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 0
                            tag = ("low risk",)                
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 0
                            tag = ("low risk",)
            data[yuiop][6] = row[6]
            yuiop += 1
        elif row[1] == "F":
            if row[3] == "N":
                if row[2] >= 65:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 4
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                elif row[2] >= 60 and row[2] < 65:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[4] >= 120:
                        row[6] = 1
                        tag = ("low risk",)
                elif row[2] >= 55 and row[2] < 60:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[4] >= 160:
                        row[6] = 1
                        tag = ("low risk",)
                    elif row[4] >= 140:
                        row[6] = 1
                        tag = ("low risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 0
                            tag = ("low risk",)
                elif row[2] >= 50 and row[2] < 55:
                    if row[4] >= 180:
                        row[6] = 1
                        tag = ("low risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 0
                            tag = ("low risk",)
                    elif row[4] >= 140:
                        row[6] = 0
                        tag = ("low risk",)
                    elif row[4] >= 120:
                        row[6] = 0
                        tag = ("low risk",)
                elif row[2] >= 40 and row[2] < 50:
                    row[6] = 0
                    tag = ("low risk",)
            if row[3] == "Y":
                if row[2] >= 65:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 14
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 12
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 11
                            tag = ("high risk",)
                        elif row[5] >= 5:
                            row[6] = 9
                            tag = ("high risk",)
                        elif row[5] >= 4:
                            row[6] = 9
                            tag = ("high risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 10
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 6:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 6
                            tag = ("medium risk",)
                    elif row[4] >= 140 or row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                elif row[2] >= 60 and row[2] < 65:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 8
                            tag = ("high risk",)
                        elif row[5] >= 7:
                            row[6] = 7
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 6
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 5
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 5
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                elif row[2] >= 55 and row[2] < 60:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 4
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 5:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 4:
                            row[6] = 3
                            tag = ("medium risk",)
                    elif row[4] >= 160:
                        if row[5] >= 8:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 7:
                            row[6] = 3
                            tag = ("medium risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 2
                            tag = ("low risk",)
                    elif row[4] >= 140:
                        if row[5] >= 8:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[4] >= 120:
                        row[6] = 1
                        tag = ("low risk",)
                elif row[2] >= 50 and row[2] < 55:
                    if row[4] >= 180:
                        if row[5] >= 8:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 2
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[4] >= 160 or row[4] >= 140:
                        row[6] = 1
                        tag = ("low risk",)
                    elif row[4] >= 120:
                        if row[5] >= 8:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 7:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[5] >= 6:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[5] >= 5:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[5] >= 4:
                            row[6] = 0
                            tag = ("low risk",)
                elif row[2] >= 40 and row[2] < 50:
                    row[6] = 0
                    tag = ("low risk",)
            data[yuiop][6] = row[6]
            yuiop += 1

        databaseTree.insert("", "end",idNumber,text=idNumber,values=row,tags=tag)
        valuesList.append(databaseTree.item(idNumber,option='values'))
        allRows.append(databaseTree.get_children()[idNumber])
        idNumber += 1

        '''risk = row[6]
        name = row[0]
        gender = row[1]
        age = row[2]
        smokingStatus = row[3]
        bloodPressure = row[4]
        cholestrol = row[5]
        
        #this is getting the data from the boxes inputted by the user
        dataList = [name, gender, age, smokingStatus, bloodPressure,cholestrol,risk] # the data collected is put into a list
        
        with open("patients' data csv.csv", 'a', newline = '') as f_object:
            writer_object = writer(f_object)
            writer_object.writerow(dataList) # and then it is written
            f_object.close()'''
        
    if data[0][0] != 'IID':
        a = 'IID'
        data[0].insert(0,a)
        for i in range(1,len(data)):
            num = str(i)
            data[i].insert(0,num)
    
        with open("patients' data csv.csv",'w', newline = '') as writeFile:
            writer_object = writer(writeFile)
            writer_object.writerows(data)
            writeFile.close()        
    databaseTree.tag_configure("high risk",background='red',foreground="white")
    databaseTree.tag_configure("medium risk",background="orange")
    databaseTree.tag_configure("low risk",background="#00FF17")
    #databaseTree.bind('<ButtonRelease-1>',collectDataForEdit)

    style.map("Treeview",foreground=[("selected","white")],background=[("selected","#0953FF")])
    databaseTree.grid(row=1,column=0,columnspan=15)


def dnd(event):
    global databaseTree
    global counter
    global idNumber
    global valuesList
    global allRows
    if event.data.startswith("{") and event.data.endswith("}"):
        event.data = event.data.replace("{","",1)
        event.data = event.data.replace("}","",len(event.data)-1)
    if event.data.endswith(".xlsx") or event.data.endswith(".csv") or event.data.endswith(".xls") or event.data.endswith(".ods"):
        read_file = pd.read_excel(event.data)
        read_file.to_csv("patients' data csv.csv",
                         index = None,
                         header = True)
        df = pd.DataFrame(pd.read_csv("patients' data csv.csv"))
        valuesList = []
        allRows = []
        databaseTree.destroy()
        data = list(csv.reader(open("patients' data csv.csv")))

        if data[0][0] != 'IID':
            a = 'IID'
            data[0].insert(0,a)
            for i in range(1,len(data)):
                num = str(i)
                data[i].insert(0,num)
        
            with open("patients' data csv.csv",'w', newline = '') as writeFile:
                writer_object = writer(writeFile)
                writer_object.writerows(data)
                writeFile.close()

        # initialise treeview
        databaseTree = ttk.Treeview(tvFrame,height=25,selectmode="extended",show="headings")

        # initialise and connect scrollbar to treeview
        #treeviewScrollbar = ttk.Scrollbar(tvFrame,orient='vertical',command=databaseTree.yview)
        #treeviewScrollbar.grid(row=1,column=16,sticky=(N,S))
        databaseTree.configure(yscrollcommand = treeviewScrollbar.set)

        # Define the columns
        databaseTree["columns"] = ("Name","Sex","Age","Smoking Status","Blood Pressure","Cholestrol","Risk")

        # Format the columns
        databaseTree.column("#0",anchor=CENTER,width=50,minwidth=50)
        databaseTree.column("Sex",anchor=CENTER,width=30,minwidth=30)
        databaseTree.column("Age",anchor=CENTER,width=30,minwidth=30)
        databaseTree.column("Name",anchor=W,width=230,minwidth=60)
        databaseTree.column("Smoking Status",anchor=CENTER,width=80,minwidth=90)
        databaseTree.column("Blood Pressure",anchor=CENTER,width=50,minwidth=90)
        databaseTree.column("Cholestrol",anchor=CENTER,width=60,minwidth=60)
        databaseTree.column("Risk",anchor=CENTER,width=30,minwidth=30)

        # Create Headings
        databaseTree.heading("#0",text="#",anchor=CENTER)
        databaseTree.heading("Sex",text="Sex",anchor=CENTER)
        databaseTree.heading("Age",text="Age",anchor=CENTER)
        databaseTree.heading("Name",text="Name",anchor=W)
        databaseTree.heading("Smoking Status",text="Smoking Status",anchor=CENTER)
        databaseTree.heading("Blood Pressure",text="Blood Pressure",anchor=CENTER)
        databaseTree.heading("Cholestrol",text="Cholestrol",anchor=CENTER)
        databaseTree.heading("Risk",text="Risk",anchor=CENTER)

        idNumber = 0
        counter += len(data)
        for row in data:
            row.pop(0)
            try:
                row = list(row)
                row[2] = int(row[2])
                row[4] = float(row[4])
                row[5] = float(row[5])
                tag = ("",)
            except:
                continue
            
            if row[1] == "M":
                if row[3] == "N":
                    if row[2] >= 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 14
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 12
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 10
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 8
                                tag = ("high risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 10
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 5
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 4
                                tag = ("medium risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                    elif row[2] >= 60 and row[2] < 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 5
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                    elif row[2] >= 55 and row[2] < 60:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                    elif row[2] >= 50 and row[2] < 55:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                    elif row[2] >= 40 and row[2] < 50:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
                        elif row[4] >= 140 or row[4] >= 120:
                            row[6] = 0
                            tag = ("low risk",)
                if row[3] == "Y":
                    if row[2] >= 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 26
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 23
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 20
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 17
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 15
                                tag = ("high risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 19
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 16
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 14
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 12
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 10
                                tag = ("high risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 13
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 11
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 7
                                tag = ("medium risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 5
                                tag = ("medium risk",)
                    elif row[2] >= 60 and row[2] < 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 18
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 15
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 13
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 11
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 10
                                tag = ("high risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 13
                            elif row[5] >= 7:
                                row[6] = 11
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 7
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 5
                                tag = ("medium risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                    elif row[2] >= 55 and row[2] < 60:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 12
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 10
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 6
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 4
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif choles3rol == 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif cholesrol == 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                    elif row[2] >= 50 and row[2] < 55:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 12
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 10
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 6
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 4
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                    elif row[2] >= 40 and row[2] < 50:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 160:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)                
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
            elif row[1] == "F":
                if row[3] == "N":
                    if row[2] >= 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 4
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                    elif row[2] >= 60 and row[2] < 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[2] >= 55 and row[2] < 60:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 160:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[4] >= 140:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
                    elif row[2] >= 50 and row[2] < 55:
                        if row[4] >= 180:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
                        elif row[4] >= 140:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[4] >= 120:
                            row[6] = 0
                            tag = ("low risk",)
                    elif row[2] >= 40 and row[2] < 50:
                        row[6] = 0
                        tag = ("low risk",)
                if row[3] == "Y":
                    if row[2] >= 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 14
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 12
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 11
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 9
                                tag = ("high risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 10
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 6
                                tag = ("medium risk",)
                        elif row[4] >= 140 or row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                    elif row[2] >= 60 and row[2] < 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 5
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                    elif row[2] >= 55 and row[2] < 60:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[2] >= 50 and row[2] < 55:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 160 or row[4] >= 140:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
                    elif row[2] >= 40 and row[2] < 50:
                        row[6] = 0
                        tag = ("low risk",)
            databaseTree.insert("", "end",idNumber,text=idNumber, values=row,tags=tag)
            valuesList.append(databaseTree.item(idNumber,option='values'))
            allRows.append(databaseTree.get_children()[idNumber])
            idNumber += 1

        databaseTree.tag_configure("high risk",background='red',foreground="white")
        databaseTree.tag_configure("medium risk",background="orange")
        databaseTree.tag_configure("low risk",background="#00FF17")


        style.map("Treeview",foreground=[("selected","white")],background=[("selected","#0953FF")])
        databaseTree.grid(row=1,column=0,columnspan=15)
        
    else:
        messagebox.showerror(message='Only xlsx and csv files can be dropped')
    
add = ttk.Button(master=addNewFrame,text="Add",command=addData,style="Accent.TButton")
add["state"] = "disabled"
add.grid(row=10,column=1,sticky="E",padx=10)

dragLabel = tk.Label(master=addNewFrame,text="Drag Excel file to open here",width=20,height=10,fg="grey",relief="sunken")
dragLabel.grid(row=9,column=0,sticky="W",padx=10,pady=40)
dragLabel.drop_target_register(DND_FILES)
dragLabel.dnd_bind('<<Drop>>', dnd)

importFile = ttk.Button(master=addNewFrame,text="Choose Excel file to import",command=openFile)
importFile.grid(row=10,column=0,sticky="W",padx=10)

im2 = Image.open("back icon.png")
im2= im2.resize((20,20),Image.ANTIALIAS)
backIcon = ImageTk.PhotoImage(im2)
backToMainButton = tk.Button(master=settingsFrame,text="Database",image=backIcon,compound=LEFT,command=goToMainPage)
backToMainButton.grid(row=0,column=0,sticky="nw")

# Switch between light and dark mode
def switchThemes():
    global databaseTree
    global treeviewScrollbar
    global editOn
    global addOn
    global idNumber
    if modeVar.get() == "Dark":
        window.tk.call("set_theme", "dark")
        settingsFrame.configure(bg="#1c1c1c")
        mainFrame.configure(bg="#1c1c1c")
        addNewFrame.configure(bg="#1c1c1c")

        tvFrame.configure(bg="#1c1c1c")
        searchFrame.configure(bg="#1c1c1c")
        treeviewScrollbar.destroy()
        databaseTree.destroy()
        if addOn == True:
            # initialise treeview
            databaseTree = ttk.Treeview(master=tvFrame,height=25,selectmode="extended",show="headings")

            # initialise and connect scrollbar to treeview
            treeviewScrollbar = ttk.Scrollbar(tvFrame,orient='vertical',command=databaseTree.yview)
            treeviewScrollbar.grid(row=1,column=16,sticky=(N,S))
            databaseTree.configure(yscrollcommand = treeviewScrollbar.set)

            # Define the columns
            databaseTree["columns"] = ("Sex","Age","Name","Smoking Status","Blood Pressure","Cholestrol Level")
            databaseTree["columns"] = ("Name","Sex","Age","Smoking Status","Blood Pressure","Cholestrol","Risk")

            # Format the columns
            databaseTree.column("#0",anchor=CENTER,width=50,minwidth=50)
            databaseTree.column("Sex",anchor=CENTER,width=30,minwidth=30)
            databaseTree.column("Age",anchor=CENTER,width=30,minwidth=30)
            databaseTree.column("Name",anchor=W,width=230,minwidth=60)
            databaseTree.column("Smoking Status",anchor=CENTER,width=80,minwidth=90)
            databaseTree.column("Blood Pressure",anchor=CENTER,width=50,minwidth=90)
            databaseTree.column("Cholestrol",anchor=CENTER,width=60,minwidth=60)
            databaseTree.column("Risk",anchor=CENTER,width=30,minwidth=30)

            # Create Headings
            databaseTree.heading("#0",text="#",anchor=CENTER)
            databaseTree.heading("Sex",text="Sex",anchor=CENTER)
            databaseTree.heading("Age",text="Age",anchor=CENTER)
            databaseTree.heading("Name",text="Name",anchor=W)
            databaseTree.heading("Smoking Status",text="Smoking Status",anchor=CENTER)
            databaseTree.heading("Blood Pressure",text="Blood Pressure",anchor=CENTER)
            databaseTree.heading("Cholestrol",text="Cholestrol",anchor=CENTER)
            databaseTree.heading("Risk",text="Risk",anchor=CENTER)

            #databaseTree.bind('<ButtonRelease-1>',collectDataForEdit)

            style.map("Treeview",foreground=[("selected","white")],background=[("selected","#0953FF")])

            databaseTree.grid(row=1,column=0,ipadx=150,columnspan=15)
            
        else:
            # initialise treeview
            databaseTree = ttk.Treeview(master=tvFrame,height=25,selectmode="extended",show="headings")

            # initialise and connect scrollbar to treeview
            treeviewScrollbar = ttk.Scrollbar(tvFrame,orient='vertical',command=databaseTree.yview)
            treeviewScrollbar.grid(row=1,column=16,sticky=(N,S))
            databaseTree.configure(yscrollcommand = treeviewScrollbar.set)

            # Define the columns
            databaseTree["columns"] = ("Sex","Age","Name","Smoking Status","Blood Pressure","Cholestrol Level")
            databaseTree["columns"] = ("Name","Sex","Age","Smoking Status","Blood Pressure","Cholestrol","Risk")

            # Format the columns
            databaseTree.column("#0",anchor=CENTER,width=50,minwidth=50)
            databaseTree.column("Sex",anchor=CENTER,width=30,minwidth=30)
            databaseTree.column("Age",anchor=CENTER,width=30,minwidth=30)
            databaseTree.column("Name",anchor=W,width=230,minwidth=60)
            databaseTree.column("Smoking Status",anchor=CENTER,width=80,minwidth=90)
            databaseTree.column("Blood Pressure",anchor=CENTER,width=50,minwidth=90)
            databaseTree.column("Cholestrol",anchor=CENTER,width=60,minwidth=60)
            databaseTree.column("Risk",anchor=CENTER,width=30,minwidth=30)

            # Create Headings
            databaseTree.heading("#0",text="#",anchor=CENTER)
            databaseTree.heading("Sex",text="Sex",anchor=CENTER)
            databaseTree.heading("Age",text="Age",anchor=CENTER)
            databaseTree.heading("Name",text="Name",anchor=W)
            databaseTree.heading("Smoking Status",text="Smoking Status",anchor=CENTER)
            databaseTree.heading("Blood Pressure",text="Blood Pressure",anchor=CENTER)
            databaseTree.heading("Cholestrol",text="Cholestrol",anchor=CENTER)
            databaseTree.heading("Risk",text="Risk",anchor=CENTER)

            style.map("Treeview",foreground=[("selected","white")],background=[("selected","#0953FF")])

            databaseTree.grid(row=1,column=0,ipadx=150,columnspan=15)

        idNumber = len(databaseTree.get_children())

        databaseTree.tag_configure("high risk",background='red',foreground="white")
        databaseTree.tag_configure("medium risk",background="orange")
        databaseTree.tag_configure("low risk",background="#00FF17")
        
        for row in valuesList:
            row = list(row)
            row[2] = int(row[2])
            row[4] = float(row[4])
            row[5] = float(row[5])
            tag = ("",)
        
            if row[1] == "M":
                if row[3] == "N":
                    if row[2] >= 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 14
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 12
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 10
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 8
                                tag = ("high risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 10
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 5
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 4
                                tag = ("medium risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                    elif row[2] >= 60 and row[2] < 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 5
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                    elif row[2] >= 55 and row[2] < 60:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                    elif row[2] >= 50 and row[2] < 55:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                    elif row[2] >= 40 and row[2] < 50:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
                        elif row[4] >= 140 or row[4] >= 120:
                            row[6] = 0
                            tag = ("low risk",)
                if row[3] == "Y":
                    if row[2] >= 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 26
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 23
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 20
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 17
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 15
                                tag = ("high risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 19
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 16
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 14
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 12
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 10
                                tag = ("high risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 13
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 11
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 7
                                tag = ("medium risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 5
                                tag = ("medium risk",)
                    elif row[2] >= 60 and row[2] < 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 18
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 15
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 13
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 11
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 10
                                tag = ("high risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 13
                            elif row[5] >= 7:
                                row[6] = 11
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 7
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 5
                                tag = ("medium risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                    elif row[2] >= 55 and row[2] < 60:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 12
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 10
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 6
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 4
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif choles3rol == 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif cholesrol == 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                    elif row[2] >= 50 and row[2] < 55:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 12
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 10
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 6
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 4
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                    elif row[2] >= 40 and row[2] < 50:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 160:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)                
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
            elif row[1] == "F":
                if row[3] == "N":
                    if row[2] >= 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 4
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                    elif row[2] >= 60 and row[2] < 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[2] >= 55 and row[2] < 60:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 160:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[4] >= 140:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
                    elif row[2] >= 50 and row[2] < 55:
                        if row[4] >= 180:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
                        elif row[4] >= 140:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[4] >= 120:
                            row[6] = 0
                            tag = ("low risk",)
                    elif row[2] >= 40 and row[2] < 50:
                        row[6] = 0
                        tag = ("low risk",)
                if row[3] == "Y":
                    if row[2] >= 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 14
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 12
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 11
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 9
                                tag = ("high risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 10
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 6
                                tag = ("medium risk",)
                        elif row[4] >= 140 or row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                    elif row[2] >= 60 and row[2] < 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 5
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                    elif row[2] >= 55 and row[2] < 60:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[2] >= 50 and row[2] < 55:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 160 or row[4] >= 140:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
                    elif row[2] >= 40 and row[2] < 50:
                        row[6] = 0
                        tag = ("low risk",)

            idNumber += 1
            databaseTree.insert(parent='',index='end',iid=idNumber,text=idNumber,values=tuple(row),tags=tag)

    elif modeVar.get() == "Light":
        window.tk.call("set_theme", "light")
        mainFrame.configure(bg="#fafafa")
        settingsFrame.configure(bg="#fafafa")
        addNewFrame.configure(bg='#fafafa')
        tvFrame.configure(bg="#fafafa")
        searchFrame.configure(bg="#fafafa")
        treeviewScrollbar.destroy()
        databaseTree.destroy()
        if addOn == True:
            # initialise treeview
            databaseTree = ttk.Treeview(master=tvFrame,height=25,selectmode="extended",show="headings")

            # initialise and connect scrollbar to treeview
            treeviewScrollbar = ttk.Scrollbar(tvFrame,orient='vertical',command=databaseTree.yview)
            treeviewScrollbar.grid(row=1,column=16,sticky=(N,S))
            databaseTree.configure(yscrollcommand = treeviewScrollbar.set)

            # Define the columns
            databaseTree["columns"] = ("Sex","Age","Name","Smoking Status","Blood Pressure","Cholestrol Level")
            databaseTree["columns"] = ("Name","Sex","Age","Smoking Status","Blood Pressure","Cholestrol","Risk")

            # Format the columns
            databaseTree.column("#0",anchor=CENTER,width=50,minwidth=50)
            databaseTree.column("Sex",anchor=CENTER,width=30,minwidth=30)
            databaseTree.column("Age",anchor=CENTER,width=30,minwidth=30)
            databaseTree.column("Name",anchor=W,width=230,minwidth=60)
            databaseTree.column("Smoking Status",anchor=CENTER,width=80,minwidth=90)
            databaseTree.column("Blood Pressure",anchor=CENTER,width=50,minwidth=90)
            databaseTree.column("Cholestrol",anchor=CENTER,width=60,minwidth=60)
            databaseTree.column("Risk",anchor=CENTER,width=30,minwidth=30)

            # Create Headings
            databaseTree.heading("#0",text="#",anchor=CENTER)
            databaseTree.heading("Sex",text="Sex",anchor=CENTER)
            databaseTree.heading("Age",text="Age",anchor=CENTER)
            databaseTree.heading("Name",text="Name",anchor=W)
            databaseTree.heading("Smoking Status",text="Smoking Status",anchor=CENTER)
            databaseTree.heading("Blood Pressure",text="Blood Pressure",anchor=CENTER)
            databaseTree.heading("Cholestrol",text="Cholestrol",anchor=CENTER)
            databaseTree.heading("Risk",text="Risk",anchor=CENTER)

            style.map("Treeview",foreground=[("selected","white")],background=[("selected","#0953FF")])

            databaseTree.grid(row=1,column=0,ipadx=150,columnspan=15)
            
        else:
            # initialise treeview
            databaseTree = ttk.Treeview(master=tvFrame,height=25,selectmode="extended",show="headings")

            # initialise and connect scrollbar to treeview
            treeviewScrollbar = ttk.Scrollbar(tvFrame,orient='vertical',command=databaseTree.yview)
            treeviewScrollbar.grid(row=1,column=16,sticky=(N,S))
            databaseTree.configure(yscrollcommand = treeviewScrollbar.set)

            # Define the columns
            databaseTree["columns"] = ("Sex","Age","Name","Smoking Status","Blood Pressure","Cholestrol Level")
            databaseTree["columns"] = ("Name","Sex","Age","Smoking Status","Blood Pressure","Cholestrol","Risk")

            # Format the columns
            databaseTree.column("#0",anchor=CENTER,width=50,minwidth=50)
            databaseTree.column("Sex",anchor=CENTER,width=30,minwidth=30)
            databaseTree.column("Age",anchor=CENTER,width=30,minwidth=30)
            databaseTree.column("Name",anchor=W,width=230,minwidth=60)
            databaseTree.column("Smoking Status",anchor=CENTER,width=80,minwidth=90)
            databaseTree.column("Blood Pressure",anchor=CENTER,width=50,minwidth=90)
            databaseTree.column("Cholestrol",anchor=CENTER,width=60,minwidth=60)
            databaseTree.column("Risk",anchor=CENTER,width=30,minwidth=30)

            # Create Headings
            databaseTree.heading("#0",text="#",anchor=CENTER)
            databaseTree.heading("Sex",text="Sex",anchor=CENTER)
            databaseTree.heading("Age",text="Age",anchor=CENTER)
            databaseTree.heading("Name",text="Name",anchor=W)
            databaseTree.heading("Smoking Status",text="Smoking Status",anchor=CENTER)
            databaseTree.heading("Blood Pressure",text="Blood Pressure",anchor=CENTER)
            databaseTree.heading("Cholestrol",text="Cholestrol",anchor=CENTER)
            databaseTree.heading("Risk",text="Risk",anchor=CENTER)

            style.map("Treeview",foreground=[("selected","white")],background=[("selected","#0953FF")])

            databaseTree.grid(row=1,column=0,ipadx=150,columnspan=15)

        idNumber = len(databaseTree.get_children())

        databaseTree.tag_configure("high risk",background='red',foreground="white")
        databaseTree.tag_configure("medium risk",background="orange")
        databaseTree.tag_configure("low risk",background="#00FF17")
        
        for row in valuesList:
            row = list(row)
            row[2] = int(row[2])
            row[4] = float(row[4])
            row[5] = float(row[5])
            if row[1] == "M":
                if row[3] == "N":
                    if row[2] >= 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 14
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 12
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 10
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 8
                                tag = ("high risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 10
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 5
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 4
                                tag = ("medium risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                    elif row[2] >= 60 and row[2] < 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 5
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                    elif row[2] >= 55 and row[2] < 60:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                    elif row[2] >= 50 and row[2] < 55:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                    elif row[2] >= 40 and row[2] < 50:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
                        elif row[4] >= 140 or row[4] >= 120:
                            row[6] = 0
                            tag = ("low risk",)
                if row[3] == "Y":
                    if row[2] >= 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 26
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 23
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 20
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 17
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 15
                                tag = ("high risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 19
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 16
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 14
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 12
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 10
                                tag = ("high risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 13
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 11
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 7
                                tag = ("medium risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 5
                                tag = ("medium risk",)
                    elif row[2] >= 60 and row[2] < 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 18
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 15
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 13
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 11
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 10
                                tag = ("high risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 13
                            elif row[5] >= 7:
                                row[6] = 11
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 7
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 5
                                tag = ("medium risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                    elif row[2] >= 55 and row[2] < 60:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 12
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 10
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 6
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 4
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif choles3rol == 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif cholesrol == 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                    elif row[2] >= 50 and row[2] < 55:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 12
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 10
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 6
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 4
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                    elif row[2] >= 40 and row[2] < 50:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 160:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)                
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
            elif row[1] == "F":
                if row[3] == "N":
                    if row[2] >= 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 4
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                    elif row[2] >= 60 and row[2] < 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[2] >= 55 and row[2] < 60:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 160:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[4] >= 140:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
                    elif row[2] >= 50 and row[2] < 55:
                        if row[4] >= 180:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
                        elif row[4] >= 140:
                            row[6] = 0
                            tag = ("low risk",)
                        elif row[4] >= 120:
                            row[6] = 0
                            tag = ("low risk",)
                    elif row[2] >= 40 and row[2] < 50:
                        row[6] = 0
                        tag = ("low risk",)
                if row[3] == "Y":
                    if row[2] >= 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 14
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 12
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 11
                                tag = ("high risk",)
                            elif row[5] >= 5:
                                row[6] = 9
                                tag = ("high risk",)
                            elif row[5] >= 4:
                                row[6] = 9
                                tag = ("high risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 10
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 6:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 6
                                tag = ("medium risk",)
                        elif row[4] >= 140 or row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                    elif row[2] >= 60 and row[2] < 65:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 8
                                tag = ("high risk",)
                            elif row[5] >= 7:
                                row[6] = 7
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 6
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 5
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 5
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                    elif row[2] >= 55 and row[2] < 60:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 4
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 5:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 4:
                                row[6] = 3
                                tag = ("medium risk",)
                        elif row[4] >= 160:
                            if row[5] >= 8:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 7:
                                row[6] = 3
                                tag = ("medium risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 2
                                tag = ("low risk",)
                        elif row[4] >= 140:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 120:
                            row[6] = 1
                            tag = ("low risk",)
                    elif row[2] >= 50 and row[2] < 55:
                        if row[4] >= 180:
                            if row[5] >= 8:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 2
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 1
                                tag = ("low risk",)
                        elif row[4] >= 160 or row[4] >= 140:
                            row[6] = 1
                            tag = ("low risk",)
                        elif row[4] >= 120:
                            if row[5] >= 8:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 7:
                                row[6] = 1
                                tag = ("low risk",)
                            elif row[5] >= 6:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 5:
                                row[6] = 0
                                tag = ("low risk",)
                            elif row[5] >= 4:
                                row[6] = 0
                                tag = ("low risk",)
                    elif row[2] >= 40 and row[2] < 50:
                        row[6] = 0
                        tag = ("low risk",)

            idNumber += 1
            databaseTree.insert(parent='',index='end',iid=idNumber,text=idNumber,values=tuple(row),tags=tag)

modeVar = tk.StringVar()

themeSwitch = ttk.Checkbutton(settingsFrame,text="Dark Mode",variable=modeVar,onvalue="Dark",offvalue="Light",command=switchThemes,style="Switch.TCheckbutton")
themeSwitch.grid(row=15,column=20)
modeVar.set("Light")

settingsLabel = ttk.Label(master=settingsFrame,text="Settings",font=("Arial",30))
settingsLabel.grid(row=0,column=30)

themeLabel = ttk.Label(master=settingsFrame,text="Theme",font=("Arial",20))
themeLabel.grid(row=10,column=20)

othersLabel = ttk.Label(master=settingsFrame,text="Others",font=("Arial",20))
othersLabel.grid(row=10,column=40)

versionLabel = ttk.Label(master=settingsFrame,text="Version 1.4")
versionLabel.grid(row=65,column=50)

ENCVar = tk.StringVar()

autoExpandNCollapseCheckbutton = ttk.Checkbutton(master=settingsFrame,text="Auto Expand and Collapse",variable=ENCVar,onvalue=True,offvalue=False,style="Switch.TCheckbutton")
ENCVar.set(True)
autoExpandNCollapseCheckbutton.grid(row=15,column=40)

selectAllVar = tk.StringVar()

def selectAll():
    if selectAllVar.get() == "On":
        databaseTree.selection_set(databaseTree.get_children())

    elif selectAllVar.get() == "Off":
        databaseTree.selection_set(())

selectAllCheckbutton = ttk.Checkbutton(master=tvFrame,text="Select All",onvalue="On",offvalue="Off",variable=selectAllVar,command=selectAll)
selectAllCheckbutton.grid(row=0,column=0,sticky="W")
selectAllVar.set("Off")

# initialise treeview
databaseTree = ttk.Treeview(master=tvFrame,height=25,selectmode="extended",show="headings")

# initialise and connect scrollbar to treeview
treeviewScrollbar = ttk.Scrollbar(tvFrame,orient='vertical',command=databaseTree.yview)
treeviewScrollbar.grid(row=1,column=16,sticky=(N,S))
databaseTree.configure(yscrollcommand = treeviewScrollbar.set)

# Define the columns
databaseTree["columns"] = ("Name","Sex","Age","Smoking Status","Blood Pressure","Cholestrol","Risk")

# Format the columns
databaseTree.column("#0",anchor=CENTER,width=50,minwidth=50)
databaseTree.column("Sex",anchor=CENTER,width=30,minwidth=30)
databaseTree.column("Age",anchor=CENTER,width=30,minwidth=30)
databaseTree.column("Name",anchor=W,width=230,minwidth=60)
databaseTree.column("Smoking Status",anchor=CENTER,width=80,minwidth=90)
databaseTree.column("Blood Pressure",anchor=CENTER,width=50,minwidth=90)
databaseTree.column("Cholestrol",anchor=CENTER,width=60,minwidth=60)
databaseTree.column("Risk",anchor=CENTER,width=30,minwidth=30)

# Create Headings
databaseTree.heading("#0",text="#",anchor=CENTER)
databaseTree.heading("Sex",text="Sex",anchor=CENTER)
databaseTree.heading("Age",text="Age",anchor=CENTER)
databaseTree.heading("Name",text="Name",anchor=W)
databaseTree.heading("Smoking Status",text="Smoking Status",anchor=CENTER)
databaseTree.heading("Blood Pressure",text="Blood Pressure",anchor=CENTER)
databaseTree.heading("Cholestrol",text="Cholestrol",anchor=CENTER)
databaseTree.heading("Risk",text="Risk",anchor=CENTER)

style.map("Treeview",foreground=[("selected","white")],background=[("selected","#0953FF")])

databaseTree.grid(row=1,column=0,ipadx=150,columnspan=15)

idNumber = len(databaseTree.get_children())

# delete rows
def deleteRow():
    global counter
    rowsToDelete = databaseTree.selection()
    if len(rowsToDelete) == 0:
        messagebox.showerror(title='Error',message="Please select something to delete!")
        return
    
    numberOfRows = len(rowsToDelete)
    deleteConfirmation = messagebox.askokcancel(title=f"Delete {numberOfRows} row(s)?",message=f"Are you sure you want to delete {numberOfRows} row(s)?",detail="This action cannot be undone.",default="cancel")

    curitem = databaseTree.selection()

    for itm in curitem:
        intervar = databaseTree.item(itm)
        delnames = intervar["values"]
    
    if deleteConfirmation == False:
        return

    counter -= numberOfRows
    
    for rowToDelete in rowsToDelete:
        valuesList.remove(databaseTree.item(rowToDelete,option='values'))
        allRows.remove(rowToDelete)
        databaseTree.delete(rowToDelete)

    lines = []

    with open("patients' data csv.csv", 'r') as readFile:
        reader = csv.reader(readFile)
        for row in reader:
            lines.append(row)
            for field in row:
                for ID in curitem:
                    if field == ID:
                        lines.remove(row)
                    
    with open("patients' data csv.csv", 'w', newline = '') as writeFile:
        writer = csv.writer(writeFile)
        writer.writerows(lines)

    if addOpen == True:
        statusFrame.grid_forget()
        statusFrame.grid(row=18,column=0,columnspan=28)

    else:
        statusFrame.grid_forget()
        statusFrame.grid(row=18,column=0,columnspan=29)
    deleteStatusLabel = tk.Label(master=statusFrame,text= str(numberOfRows) + " row(s) deleted.",bg="#1FE41C",fg="#FFFFFF")
    deleteStatusLabel.grid(row=1,column=0,sticky="W",padx=20)
    window.after(3000, statusFrame.grid_forget)
    window.after(3000, deleteStatusLabel.destroy)

imd = Image.open("delete (2).png")
imd = imd.resize((20,20),Image.ANTIALIAS)
deleteIcon = ImageTk.PhotoImage(imd)
deleteButton = ttk.Button(master=tvFrame,image=deleteIcon,command=deleteRow)
deleteButton.grid(row=0,column=1,sticky="W")

# The pie chart
def pieChart():
    file_name = "patients' data csv.csv"
    file = open(file_name)
    file_reader = csv.reader(file)
    filtered = list(file_reader)
    print(filtered)
    
    #Initialisation for values, under 50, above 50 and total for each
    low_risk_u50 = 0
    medium_risk_u50 = 0
    high_risk_u50 = 0

    low_risk_a50 = 0
    medium_risk_a50 = 0
    high_risk_a50 = 0

    low_risk_total = 0
    medium_risk_total = 0
    high_risk_total = 0

    num_patients_u50 = 0
    num_patients_a50 = 0
    num_patients_total = 0
    
    #Calculating the variables above using for loop

    for i in range(1,len(filtered)):
        if int(filtered[i][3]) < 50 and float(filtered[i][7]) < 2.5:
            low_risk_u50 += 1
        if int(filtered[i][3]) < 50 and 2.5 <= float(filtered[i][7]) < 7.5 :
            medium_risk_u50 += 1
        if int(filtered[i][3]) < 50 and float(filtered[i][7]) >= 7.5:
            high_risk_u50 += 1
        if int(filtered[i][3]) < 50:
            num_patients_u50 += 1

        if 50 <= int(filtered[i][3]) <= 69 and float(filtered[i][7]) < 5:
            low_risk_a50 += 1
        if 50 <= int(filtered[i][3]) <= 69 and 5 <= float(filtered[i][7]) < 8:
            medium_risk_a50 += 1
        if 50 <= int(filtered[i][3]) <= 69 and float(filtered[i][7]) >= 8:
            high_risk_a50 += 1
        if 50 <= int(filtered[i][3]) <= 69:
            num_patients_a50 += 1
            
    #Finding totals that will be used for generation of chart
    low_risk_total = low_risk_u50 + low_risk_a50
    medium_risk_total = medium_risk_u50 + medium_risk_a50
    high_risk_total = high_risk_u50 + high_risk_a50
    num_patients_total = num_patients_u50 + num_patients_a50

    
    portion = [0,0,0]
    portion[0] += low_risk_total
    portion[1] += medium_risk_total
    portion[2] += high_risk_total
    labels = 'Low Risk','Medium Risk','High Risk'
    colors = ['green', 'yellow', 'red']
    colors = ['#71E960','#F7E841','#FF5034']
    explode = (0, 0, 0)
    chart_title = "Patients based on their Risk Score (Total no. of Patients = " + str(num_patients_total) + ")"
    plt.title(chart_title,bbox={'facecolor':'0.8', 'pad':5})
    plt.pie(portion, explode = explode, labels = labels, colors = colors, autopct = '%1.1f%%', shadow = True, startangle = 140)
    plt.show()
    
pieChartButton = ttk.Button(master=mainFrame,text="Pie Chart",command=pieChart)
pieChartButton.grid(row=8,column=0,sticky="W")

def copyToXL():   #excel_name has to be the name of the xlsx file user put in , with .xlsx extension in string format
    xlToCopyBack = filedialog.askopenfilename(title="Choose a file to copy to",filetypes=(("xlsx files", "*.xlsx"), ("All Files", "*.*"))
)   
    if xlToCopyBack == "":
        return
    excel_name = xlToCopyBack[xlToCopyBack.rfind("/")+1:]
    file_name = "patients' data csv.csv"
    file = open(file_name)
    file_reader = csv.reader(file)
    filtered = list(file_reader)

    if filtered[0][0] == 'IID':
        for i in range(len(filtered)):
            d = filtered[i][0]
            filtered[i].remove(d)
    with open("patients' data csv.csv",'w',newline = '') as writeFile:
        writer_object = writer(writeFile)
        writer_object.writerows(filtered)
        writeFile.close()

    read_file = pd.read_csv ("patients' data csv.csv")   #^like "patients' data xlsx.xlsx"
    read_file.to_excel (excel_name, index = None, header = True)

copyToXLButton = ttk.Button(master=mainFrame,text="Copy Data to Excel File",command=copyToXL)
copyToXLButton.grid(row=10,column=0,sticky="W")

'''def addIID():
    file_name = "patients' data csv.csv"
    file = open(file_name)
    file_reader = csv.reader(file)
    filtered = list(file_reader)

    if filtered[0][0] != 'IID':
        a = 'IID'
        filtered[0].insert(0,a)
        for i in range(1,len(filtered)):
            num = str(i)
            filtered[i].insert(0,num)

        with open("patients' data csv.csv",'w', newline = '') as writeFile:
            writer_object = writer(writeFile)
            writer_object.writerows(filtered)
            writeFile.close()

addIIDButton = ttk.Button(master=mainFrame,text="Add IID column to Excel File",command=addIID)
addIIDButton.grid(row=11,column=0,sticky="W")'''

mainColumnCount, mainRowCount = mainFrame.grid_size()
for column in range(mainColumnCount):
    
    mainFrame.columnconfigure(column, minsize=10)

for row in range(mainRowCount):
    mainFrame.rowconfigure(row, minsize=10)

settingsColumnCount, settingsRowCount = settingsFrame.grid_size()
for column in range(settingsColumnCount):
    
    settingsFrame.columnconfigure(column, minsize=10)

for row in range(settingsRowCount):
    settingsFrame.rowconfigure(row, minsize=10)

statusColumnCount, statusRowCount = statusFrame.grid_size()
for column in range(statusColumnCount):
    
    statusFrame.columnconfigure(column, minsize=10)

for row in range(statusRowCount):
    statusFrame.rowconfigure(row, minsize=10)

aboutColumnCount, aboutRowCount = aboutFrame.grid_size()
for column in range(aboutColumnCount):
    
    aboutFrame.columnconfigure(column, minsize=10)

for row in range(aboutRowCount):
    aboutFrame.rowconfigure(row, minsize=10)


window.config(menu = menubar)
window.mainloop()
