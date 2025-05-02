import os                    #I used os for the path
from customtkinter import *  #I used customtkinter because of the UI of this is more appealing than the original
from PIL import Image        #I used pillow because the terminal recommended me to due to high pixels of the background and the image
from openpyxl import Workbook, load_workbook   #pyxl imported for storing data from python to excel
from tkinter import messagebox    #I also import the normal tkinter for the messagebox

# --------------------------------------------------------------------------OPENPYXL----------------------------------------------------------------------------------- #

# i used a variable in excel file name so i would not type the file name everytime
filename = "student_scores.xlsx"

def initialize_file():  #This function is the one who will make the excel file in this path and automaticaly type the Name, score and results
    if not os.path.exists(filename):
        wb = Workbook()
        ws = wb.active
        ws.append(["Name", "Score", "Result"])
        wb.save(filename)

def calculate_result(score): #This function will calculate if the user is pass or not and if its more than 100 or typed error it will just be invalid and return the user to the mainwindow
    try:
        score = int(score)
        if 0 <= score <= 100:
            return "Pass" if score >= 75 else "Fail"
        else:
            return "Invalid"
    except ValueError: 
        return "Invalid"
    

def add_student_score(name, score):    #This function will add student info and shows a messagebox if something is wrong like invalid output and success!
    result = calculate_result(score)
    if result == "Invalid":
        messagebox.showerror("Invalid Input", "Score must be an integer between 0 and 100.")
        return

    wb = load_workbook(filename)
    ws = wb.active
    ws.append([name, score, result])
    wb.save(filename)
    messagebox.showinfo("Success", f"Saved: {name} - {score} ({result})")

def viewData():  #This function is when the user click the button it while show the messagebox that have all infos written in the excel file
    if not os.path.exists(filename):
        messagebox.showinfo("No Data", "No records found.")
        return

    wb = load_workbook(filename)
    ws = wb.active
    records = list(ws.iter_rows(min_row=2, values_only=True))
    mainWindow.withdraw()  #I used withdraw because everytime i used destroy my code will have many error

# In this part i designed the background in canva to make it more appealing
    record_window = CTkToplevel()
    record_window.title("All Records")
    record_window.geometry("400x500")
    set_appearance_mode("light")
    record_window.resizable(False, False)

# This is the background that i used i size it to the window it is not stretch or too tight to the window
    backgroundImage = CTkImage(Image.open("Media Files\\ViewData.png"), size=(400, 500)) #i sized the image same as the window so it would be not look like cropped or stretched
    bg_label1 = CTkLabel(master=record_window, image=backgroundImage, text="") #I use label so i can import the image to the window
    bg_label1.place(x=0, y=0, relwidth=1, relheight=1)
    
# This one is also for the design i downloaded a png file and put it Beside the text to make it look good
    exam_image = CTkImage(light_image=Image.open("Media Files\\testPaper.png"), size=(50, 50)) # i manually choose the proper size to make it more look good
    exam_label = CTkLabel(master=record_window, image=exam_image, fg_color="white",text="") 
    exam_label.place(x=100, y=40)
    
# Texts in the upper part
    studentsData = CTkLabel(master=record_window, text="Student's Data", font=("Helvetica", 25, "bold"), fg_color="white")
    studentsData.place(x=160, y=50)  
    
    nameLbl = CTkLabel(master=record_window, text="Student's Name", font=("Helvetica", 14, "bold"), fg_color="white")
    nameLbl.place(x=110, y=105)

    scoreLbl = CTkLabel(master=record_window, text="Score", font=("Helvetica", 14, "bold"), fg_color="white")
    scoreLbl.place(x=242, y=105)
    
    resultLbl = CTkLabel(master=record_window, text="Result", font=("Helvetica", 14, "bold"), fg_color="white")
    resultLbl.place(x=292, y=105)
    
# In this part i declared a y and spacing so incase that the user will add more info it has good spacing and make it more easy to read
    y_offset = 140
    spacing = 40

# I make it for loop so it will continue recording and putting the info in the window until the user is done
    for name, score, result in records:
        
        nameTxt = CTkLabel(record_window, text=f"{name}", font=("Arial", 12, "bold"), fg_color="white")
        nameTxt.place(x=83, y=y_offset)
        
        scoreTxt = CTkLabel(record_window, text=f"{score}", font=("Arial", 12, "bold"), fg_color="white")
        scoreTxt.place(x=255, y=y_offset)
        
        resultTxt = CTkLabel(record_window, text=f"{result}", font=("Arial", 12, "bold"), fg_color="white")
        resultTxt.place(x=300, y=y_offset)
        
        y_offset += spacing #This part is when the user add info it will automatic calcu the position of the infos

    # This function is when the user click the button the window will destroy and will back to the main window
    def back_to_main():
        record_window.destroy()
        mainWindow.deiconify() 
    
    # i feel that this was needed so i create a function that it will reset the infos in the window and it will ask first whether the user is sure or not
    def reset():
        confirm = messagebox.askyesno("Confirm Reset", "Are you sure you want to delete all records?")
        if confirm:
            if os.path.exists(filename):
                wb = load_workbook(filename)
                ws = wb.active

                ws.delete_rows(2, ws.max_row)
                wb.save(filename)
                messagebox.showinfo("Reset Successful", "All records have been cleared. \nPress Back button to input info again.")
            
# The button at the bottom part
    backBtn = CTkButton(master=record_window, text="Back", fg_color="black", height=35, width=80, command=back_to_main)
    backBtn.place(x=125, y=435)
    resetBtn = CTkButton(master=record_window, text="Reset", fg_color="black", height=35, width=80, command=reset)
    resetBtn.place(x=240, y=435)


# ------------------------------------------------------------------------------TKINTER------------------------------------------------------------------------------- #

mainWindow = CTk()  # This part i make it outside the function so it will not cause error or loop errors

def main():
    mainWindow.geometry("400x500")
    mainWindow.title("Student's Score Tracker")
    set_appearance_mode("light") #this will change the theme of the windows
    mainWindow.resizable(False, False) # i used this so the user cannot change the window size

    # Background image to make the window more appealing
    # i edited the backgound in canva so the backgound won't be plain
    backgroundImage = CTkImage(Image.open("Media Files\\loginPage.png"), size=(400, 500)) # same reason like the other one. I sized the image same as the window so it would be not look like cropped or stretched
    bg_label = CTkLabel(master=mainWindow, image=backgroundImage, text="") #I use label so i can import the image to the window
    bg_label.place(x=0, y=0, relwidth=1, relheight=1)

    jabi_image = CTkImage(light_image=Image.open("Media Files\\miniPicture.png"), size=(75, 75)) 
    jabi_label = CTkLabel(master=mainWindow, image=jabi_image, text="")
    jabi_label.place(x=160, y=61)

    
    # In this part i didn't make their variable because they dont have use in others
    CTkLabel(mainWindow, text="Student's", fg_color="white", font=("Helvetica", 28, "bold")).place(x=135, y=140)
    CTkLabel(mainWindow, text="Score Tracker", fg_color="white", font=("Helvetica", 28, "bold")).place(x=100, y=170)
    CTkLabel(mainWindow, text="Made by Jabiii", fg_color="white", font=("Helvetica", 12), text_color="grey").place(x=160, y=200)

    # In this part i used placeholder text to guide the user on the format. Even though it is not strict like the month/date/year but still functional :)
    CTkLabel(mainWindow, text="Student Name", fg_color="white", font=("Arial", 12), text_color="grey").place(x=105, y=235)
    nameEntry = CTkEntry(master=mainWindow, placeholder_text='Surname, Name, Middle In.', height=35, width=180, fg_color="#c4c3be", border_color="#c4c3be")
    nameEntry.place(x=105, y=260)

    # Same as this one :) but this one is strict so the user will only input integers and not strings
    CTkLabel(mainWindow, text="Quiz/Test Score", fg_color="white", font=("Arial", 12), text_color="grey").place(x=105, y=295)
    scoreEntry = CTkEntry(master=mainWindow, placeholder_text="Enter your score (100 Max)", height=35, width=180, fg_color="lightgray", border_color="lightgray")
    scoreEntry.place(x=105, y=320)

    def save_data():
        name = nameEntry.get()
        score = scoreEntry.get()

        if name.strip() == "" or score.strip() == "":
            messagebox.showwarning("Missing Info", "Please enter both name and score.")
            return

        add_student_score(name, score)
        nameEntry.delete(0, END)
        scoreEntry.delete(0, END)

    saveBtn = CTkButton(master=mainWindow, text="Save", fg_color="black", height=35, width=80, command=save_data)
    saveBtn.place(x=105, y=365)

    viewBtn = CTkButton(master=mainWindow, text="View Data", fg_color="black", height=35, width=80, command=viewData)
    viewBtn.place(x=205, y=365)

initialize_file()
main()
mainWindow.mainloop()
