import os
from customtkinter import *
from PIL import Image
from openpyxl import Workbook, load_workbook
from tkinter import messagebox

# --------------------------------------------------------------------------OPENPYXL----------------------------------------------------------------------------------- #

filename = "student_scores.xlsx"

def initialize_file():
    if not os.path.exists(filename):
        wb = Workbook()
        ws = wb.active
        ws.append(["Name", "Score", "Result"])
        wb.save(filename)

def calculate_result(score):
    try:
        score = int(score)
        if 0 <= score <= 100:
            return "Pass" if score >= 75 else "Fail"
        else:
            return "Invalid"
    except ValueError:
        return "Invalid"
    

def add_student_score(name, score):
    result = calculate_result(score)
    if result == "Invalid":
        messagebox.showerror("Invalid Input", "Score must be an integer between 0 and 100.")
        return

    wb = load_workbook(filename)
    ws = wb.active
    ws.append([name, score, result])
    wb.save(filename)
    messagebox.showinfo("Success", f"Saved: {name} - {score} ({result})")

def viewData():
    if not os.path.exists(filename):
        messagebox.showinfo("No Data", "No records found.")
        return

    wb = load_workbook(filename)
    ws = wb.active
    records = list(ws.iter_rows(min_row=2, values_only=True))
    mainWindow.withdraw()  

    record_window = CTkToplevel()
    record_window.title("All Records")
    record_window.geometry("400x500")
    set_appearance_mode("light")
    record_window.resizable(False, False)

    backgroundImage = CTkImage(Image.open("Media Files\\ViewData.png"), size=(400, 500))
    bg_label1 = CTkLabel(master=record_window, image=backgroundImage, text="")
    bg_label1.place(x=0, y=0, relwidth=1, relheight=1)
    
    exam_image = CTkImage(light_image=Image.open("Media Files\\testPaper.png"), size=(50, 50))
    exam_label = CTkLabel(master=record_window, image=exam_image, fg_color="white",text="")
    exam_label.place(x=100, y=40)
    
    studentsData = CTkLabel(master=record_window, text="Student's Data", font=("Helvetica", 25, "bold"), fg_color="white")
    studentsData.place(x=160, y=50)  
    
    nameLbl = CTkLabel(master=record_window, text="Student's Name", font=("Helvetica", 14, "bold"), fg_color="white")
    nameLbl.place(x=110, y=105)

    scoreLbl = CTkLabel(master=record_window, text="Score", font=("Helvetica", 14, "bold"), fg_color="white")
    scoreLbl.place(x=242, y=105)
    
    resultLbl = CTkLabel(master=record_window, text="Result", font=("Helvetica", 14, "bold"), fg_color="white")
    resultLbl.place(x=292, y=105)
    
    y_offset = 140
    spacing = 40

    for name, score, result in records:
        
        nameTxt = CTkLabel(record_window, text=f"{name}", font=("Arial", 12, "bold"), fg_color="white")
        nameTxt.place(x=83, y=y_offset)
        
        scoreTxt = CTkLabel(record_window, text=f"{score}", font=("Arial", 12, "bold"), fg_color="white")
        scoreTxt.place(x=255, y=y_offset)
        
        resultTxt = CTkLabel(record_window, text=f"{result}", font=("Arial", 12, "bold"), fg_color="white")
        resultTxt.place(x=300, y=y_offset)
        
        y_offset += spacing

    def back_to_main():
        record_window.destroy()
        mainWindow.deiconify() 
        
    def reset():
        confirm = messagebox.askyesno("Confirm Reset", "Are you sure you want to delete all records?")
        if confirm:
            if os.path.exists(filename):
                wb = load_workbook(filename)
                ws = wb.active

                ws.delete_rows(2, ws.max_row)
                wb.save(filename)
                messagebox.showinfo("Reset Successful", "All records have been cleared. \nPress Back button to input info again.")
            

    backBtn = CTkButton(master=record_window, text="Back", fg_color="black", height=35, width=80, command=back_to_main)
    backBtn.place(x=125, y=435)
    resetBtn = CTkButton(master=record_window, text="Reset", fg_color="black", height=35, width=80, command=reset)
    resetBtn.place(x=240, y=435)


# ------------------------------------------------------------------------------TKINTER------------------------------------------------------------------------------- #

mainWindow = CTk()  

def main():
    mainWindow.geometry("400x500")
    mainWindow.title("Student's Score Tracker")
    set_appearance_mode("light")
    mainWindow.resizable(False, False)

    backgroundImage = CTkImage(Image.open("Media Files\\loginPage.png"), size=(400, 500))
    bg_label = CTkLabel(master=mainWindow, image=backgroundImage, text="")
    bg_label.place(x=0, y=0, relwidth=1, relheight=1)

    jabi_image = CTkImage(light_image=Image.open("Media Files\\miniPicture.png"), size=(75, 75))
    jabi_label = CTkLabel(master=mainWindow, image=jabi_image, text="")
    jabi_label.place(x=160, y=61)

    CTkLabel(mainWindow, text="Student's", fg_color="white", font=("Helvetica", 28, "bold")).place(x=135, y=140)
    CTkLabel(mainWindow, text="Score Tracker", fg_color="white", font=("Helvetica", 28, "bold")).place(x=100, y=170)
    CTkLabel(mainWindow, text="Made by Jabiii", fg_color="white", font=("Helvetica", 12), text_color="grey").place(x=160, y=200)

    global nameEntry
    CTkLabel(mainWindow, text="Student Name", fg_color="white", font=("Arial", 12), text_color="grey").place(x=105, y=235)
    nameEntry = CTkEntry(master=mainWindow, placeholder_text='Surname, Name, Middle In.', height=35, width=180, fg_color="#c4c3be", border_color="#c4c3be")
    nameEntry.place(x=105, y=260)

    global scoreEntry
    CTkLabel(mainWindow, text="Quiz/Test Score", fg_color="white", font=("Arial", 12), text_color="grey").place(x=105, y=295)
    scoreEntry = CTkEntry(master=mainWindow, placeholder_text="Enter your score", height=35, width=180, fg_color="lightgray", border_color="lightgray")
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
