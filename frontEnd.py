import customtkinter as ct                                      # pip install customtkinter
import tkinter as tk
import threading
from PIL import Image, ImageTk
import os
import backEnd

# Updates the progress bar based on the multiple flags throughout the code
def update_progressbar(progressBar):
    global progress
    progressBar.set(progress)
    progressBar.update()
    if progress < 1:
        progressBar.after_idle(update_progressbar, progressBar)
    else:
        update_window()

# Starts running the code after the run button is pressed
def start_processing(sheetName, progressBar, studentResponses, hostResponses, studentName, hostName):
    global progress
    progress = 0
    threading.Thread(target = backEnd.runCode, args = (sheetName, studentResponses, hostResponses, studentName, hostName)).start()
    update_progressbar(progressBar)

# Input frame, houses the tabview that shows the inputs needs from Student and Host responses
def inputFrame():
    global studentInputs
    global hostInputs
    global inputFrame
    inputFrame = ct.CTkTabview(scrollableFrame, segmented_button_selected_hover_color = "#1F6AA5", border_width = 1, width = 550)
    inputFrame.pack(pady = (0, 10), padx = 10)
    inputFrame.add("Student Responses")
    instructions1 = ct.CTkLabel(inputFrame.tab("Student Responses"), text = "INPUT INSTRUCTIONS: \n\nPlease enter the LETTER of the column where the following \ncolumns are" \
                                " located in the sheet. For example, if the Student \nChoice 1 is in column K, enter the corresponding letter \"K\".")
    instructions1.grid(pady = 20, padx = 100)
    inputFrame.add("Host Responses")
    instructions2 = ct.CTkLabel(inputFrame.tab("Host Responses"), text = "INPUT INSTRUCTIONS: \n\nPlease enter the LETTER of the column where the following \ncolumns are" \
                                " located in the sheet. For example, if the Host \nIndustry is in column M, enter the corresponding letter \"M\".")
    instructions2.grid(pady = 20, padx = 100)
    studentInputs = []
    hostInputs = []
    studentQ = ["\nStudent UD Email Address :\n", 
            "Student Choice 1 :\n", 
            "Student Choice 2 :\n", 
            "Student Choice 3 :\n", 
            "Student Industry Interests :\n", 
            "Student School Year :\n", 
            "Student Major :\n"]
    
    hostQ = ["\nJob Shadow ID :\n", 
            "Host Email Address :\n", 
            "Host Industry :\n", 
            "Host Preferred Number of Matches (Capacity) :\n", 
            "Host Modality (Virtual/In Person) :\n",
            "Alumni Host Program of Study :\n",
            "Host Job Title :\n",
            "Host Company/Organization Name :\n",
            "Host Student Major Preferences :\n"]
    for i in range(7):
        input = ct.CTkLabel(inputFrame.tab("Student Responses"), text = studentQ[i])
        input.grid(padx = 100)
        inputID = ct.CTkEntry(input)
        inputID.grid(padx = 100, pady = (0, 40))
        studentInputs.append(inputID)
    for i in range(9):
        input = ct.CTkLabel(inputFrame.tab("Host Responses"), text = hostQ[i])
        input.grid(padx = 100)
        inputID = ct.CTkEntry(input)
        inputID.grid(padx = 100, pady = (0, 40))
        hostInputs.append(inputID)

# Host frame, houses the two tables containing filled hosts and unfilled hosts
# Need to fix double scrolling bug
def hostFrame():
    global rightFrame
    global leftFrame
    hostFrame = ct.CTkFrame(scrollableFrame, fg_color="#242424", bg_color="#242424")
    hostFrame.pack(pady = 40)
    leftFrame = ct.CTkScrollableFrame(hostFrame, border_width = 3, label_text = "Filled Hosts", label_fg_color="#1F6AA5", height = 350, width = 300, bg_color="#2B2B2B", fg_color="#2B2B2B")
    leftFrame.pack(side = ct.LEFT, padx = (0, 50))
    leftFrame._scrollbar.grid(padx = (0, 5))
    innerFrame = ct.CTkFrame(leftFrame, fg_color = "#2B2B2B", bg_color="#2B2B2B")
    innerFrame.pack()
    rightFrame = ct.CTkScrollableFrame(hostFrame, border_width = 3, label_text = "Unfilled Hosts", label_fg_color="#1F6AA5", height = 350, width = 300)
    rightFrame.pack()
    rightFrame._scrollbar.grid(padx = (0, 5))
    innerFrame2 = ct.CTkFrame(rightFrame, fg_color = "#2B2B2B", bg_color="#2B2B2B")
    innerFrame2.pack()
    fullHosts = []
    availableHosts = []
    for host in backEnd.result:
        if (len(backEnd.result[host]) == backEnd.capacities[int(host.name)]):
            fullHosts.append({"name":str(host.name), "hosting":str(len(backEnd.result[host])), "cap":str(backEnd.capacities[int(host.name)])})
        else:
            availableHosts.append({"name":str(host.name), "hosting":str(len(backEnd.result[host])), "cap":str(backEnd.capacities[int(host.name)])})
    
    for host in fullHosts:
        name = host["name"]
        hosting = host["hosting"]
        max_students = host["cap"]
        label_text = f"{name}:\t {hosting} out of {max_students}"
        employer_label = ct.CTkLabel(innerFrame, text=label_text, bg_color="#2B2B2B")
        employer_label.pack(pady=5)
    
    for host in availableHosts:
        name = host["name"]
        hosting = host["hosting"]
        max_students = host["cap"]
        label_text = f"{name}:\t {hosting} out of {max_students}"
        employer_label = ct.CTkLabel(innerFrame2, text=label_text, bg_color="#2B2B2B")
        employer_label.pack(pady=5)

# Stat frame, houses statistics of the matching program
def statFrame():
    statFrame = ct.CTkFrame(scrollableFrame, fg_color = "#242424", bg_color = "#242424")
    statFrame.pack(pady = 30)
    frame = ct.CTkFrame(statFrame, bg_color = "#242424", corner_radius = 20, border_width = 2)
    frame.pack(side = ct.LEFT, padx = 5)
    frame2 = ct.CTkFrame(statFrame, bg_color = "#242424", corner_radius = 20, border_width = 2)
    frame2.pack(side = ct.LEFT, padx = 5)
    frame3 = ct.CTkFrame(statFrame, bg_color = "#242424", corner_radius = 20, border_width = 2)
    frame3.pack(side = ct.LEFT, padx = 5)
    frame4 = ct.CTkFrame(statFrame, bg_color = "#242424", corner_radius = 20, border_width = 2)
    frame4.pack(side = ct.LEFT, padx = 5)
    frame5 = ct.CTkFrame(statFrame, bg_color = "#242424", corner_radius = 20, border_width = 2)
    frame5.pack(side = ct.LEFT, padx = 5)
    frame6 = ct.CTkFrame(statFrame, bg_color = "#242424", corner_radius = 20, border_width = 2)
    frame6.pack(padx = 5)

    label_frame = ct.CTkFrame(frame, bg_color = "#242424")
    label_frame.pack(padx = 10, pady = 10)
    matches_label = ct.CTkLabel(label_frame, bg_color = "#2B2B2B", wraplength=110, 
            text = f"Number of successful student matches: \n\n{backEnd.studentsWith} of {backEnd.totStudents}\n{round((backEnd.studentsWith/backEnd.totStudents)*100)}%")
    matches_label.pack()

    label_frame2 = ct.CTkFrame(frame2, bg_color = "#242424")
    label_frame2.pack(padx = 10, pady = 10)
    matches_label2 = ct.CTkLabel(label_frame2, bg_color = "#2B2B2B", wraplength=110, 
            text = f"Number of unsuccessful student matches: \n\n{backEnd.studentsWithout} of {backEnd.totStudents}\n{round((backEnd.studentsWithout/backEnd.totStudents)*100)}%")
    matches_label2.pack()

    label_frame3 = ct.CTkFrame(frame3, bg_color = "#242424")
    label_frame3.pack(padx = 10, pady = 10)
    matches_label3 = ct.CTkLabel(label_frame3, bg_color = "#2B2B2B", wraplength=110, 
            text = f"Number of successful host matches: \n\n{backEnd.hostsWith} of {backEnd.totHosts}\n{round((backEnd.hostsWith/backEnd.totHosts)*100)}%")
    matches_label3.pack()

    label_frame4 = ct.CTkFrame(frame4, bg_color = "#242424")
    label_frame4.pack(padx = 10, pady = 10)
    matches_label4 = ct.CTkLabel(label_frame4, bg_color = "#2B2B2B", wraplength=110, 
            text = f"Number of unsuccessful host matches: \n\n{backEnd.hostsWithout} of {backEnd.totHosts}\n{round((backEnd.hostsWithout/backEnd.totHosts)*100)}%")
    matches_label4.pack()

    label_frame5 = ct.CTkFrame(frame5, bg_color = "#242424")
    label_frame5.pack(padx = 10, pady = 10)
    matches_label5 = ct.CTkLabel(label_frame5, bg_color = "#2B2B2B", wraplength = 125, 
            text = f"Green List Students:\n {backEnd.solidStudents} of {backEnd.studentsWith}\n\n\n Yellow List Students: \n{backEnd.maybeStudents} of {backEnd.studentsWith}")
    matches_label5.pack()

    totCapStr = "Out of " + str(backEnd.totalCap)  + " positions\n\n\n " + str(backEnd.totalFill) + " spots were filled"
    label_frame6 = ct.CTkFrame(frame6, bg_color = "#242424")
    label_frame6.pack(padx = 10, pady = 10)
    matches_label6 = ct.CTkLabel(label_frame6, bg_color = "#2B2B2B", wraplength = 110, text = totCapStr)
    matches_label6.pack()

# Updates window after the matching code has finished, shows stats
def update_window(): # Nicks BLANK of 2022 Summer Job Shadow Program
    inputQ.destroy()
    sheetName.destroy()
    runButton.configure(text = "Close", command = window.destroy)
    progressBar.destroy()
    inputQ2.destroy()
    inputQ3.destroy()
    hostName.destroy()
    studentName.destroy()
    preReq.destroy()
    inputFrame.destroy()
    hostFrame()
    statFrame()

# Function for UD top logo properties, fixes background issue and matches it with the window color
def open_image(image_path):
    image = Image.open(image_path)
    return ImageTk.PhotoImage(image)

# Sets window theme and progress bar value to 0
ct.set_appearance_mode("dark")
progress = 0.0

# Creates image pathing for .exe bundle
exe_dir2 = os.path.dirname(os.path.abspath(__file__))
data_path2 = os.path.join(exe_dir2, 'UD.ico')   
exe_dir3 = os.path.dirname(os.path.abspath(__file__))
data_path3 = os.path.join(exe_dir3, 'UD.png')   

# Creates popup window
global window
global scrollableFrame
window = ct.CTk()
window.iconbitmap(data_path2)
window.title("Job Shadow Matching Program")
window.geometry("900x800")
window.configure(fg_color="#242424", bg_color="#242424")
scrollableFrame = ct.CTkScrollableFrame(window, fg_color="#242424", bg_color="#242424")
scrollableFrame.pack(fill="both", expand=True)
pause_execution = tk.BooleanVar()
pause_execution.set(False)

# Sizing and properties of UD top logo
canvas = ct.CTkCanvas(scrollableFrame, width = 340, height = 160, bg = "#242424", highlightthickness = 0)
image_path = data_path3
photo_image = open_image(image_path)
image_width, image_height = photo_image.width(), photo_image.height()
canvas.create_image(image_width//2, image_height//2 + 10, image=photo_image)
canvas.pack(pady = 20)

preReq = ct.CTkLabel(scrollableFrame, text = "Please share the google doc with this email:\n\n matching-prog@cisc498-group2.iam.gserviceaccount.com")
preReq.pack(padx = 0.5, pady = (10, 25))

# Gets the name of the Google Sheet
inputQ = ct.CTkLabel(scrollableFrame, text = "Name of Google Sheet :")
inputQ.pack(padx = 0.5, pady = (10, 5))
sheetName = ct.CTkEntry(scrollableFrame, width = 300)
sheetName.pack(padx = 0.5, pady = (0, 40))

inputQ2 = ct.CTkLabel(scrollableFrame, text = "Name of Student Responses Sheet :")
inputQ2.pack(padx = 0.5, pady = (10, 5))
studentName = ct.CTkEntry(scrollableFrame, width = 300)
studentName.pack(padx = 0.5, pady = (0, 40))

inputQ3 = ct.CTkLabel(scrollableFrame, text = "Name of Host Responses Sheet :")
inputQ3.pack(padx = 0.5, pady = (10, 5))
hostName = ct.CTkEntry(scrollableFrame, width = 300)
hostName.pack(padx = 0.5, pady = (0, 40))

# Initializes the input frame
inputFrame()

# Additon of run button for multi-input
runButton = ct.CTkButton(scrollableFrame, corner_radius = 75, hover_color = "#488dc2", text = "Run", command = lambda: start_processing(sheetName.get(), progressBar, studentInputs, hostInputs, studentName.get(), hostName.get()))
runButton.pack(padx=0.5, pady=(40,20))


# Addition of progress bar
progressBar = ct.CTkProgressBar(scrollableFrame, border_width = 1, height = 15, width = 400, orientation = 'horizontal', mode = 'determinate')
progressBar.set(0)
progressBar.pack(pady = (30, 70))

window.mainloop() 