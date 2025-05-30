import customtkinter as ct                                      # pip install customtkinter
import threading
from PIL import Image, ImageTk
import os
import backEnd

# -----------------------------------------------------------------------------
# Job Shadow Matching Program GUI
# -----------------------------------------------------------------------------
# Author: Michael Litvin
# Date: 05/30/2025
# Description:
#   A CTkinter-based application that collects student and host data from a
#   Google Sheet, executes a matching algorithm in the background, and displays
#   results including filled and unfilled host slots and detailed statistics.
#   The app leverages a progress bar, tabbed input interface, and scrollable
#   result panels for a responsive user experience.
# -----------------------------------------------------------------------------

"""
Load an image file and convert it into a format compatible with CTkinter.

Args:
    image_path (str): Absolute or relative path to an image file on disk.

Returns:
    ImageTk.PhotoImage: Image object ready for display in CTkCanvas or Label.
"""
def open_image(image_path):
    image = Image.open(image_path)
    return ImageTk.PhotoImage(image)


"""
Periodically update the progress bar to reflect the current progress of
the backend matching algorithm. This function re-schedules itself
until progress reaches completion (1.0).

Args:
    progressBar (CTkProgressBar): The progress bar widget to update.
"""
def update_progressbar(progressBar):
    global progress
    progressBar.set(progress)
    progressBar.update()
    # Continue polling as long as processing is incomplete
    if progress < 1:
        progressBar.after_idle(update_progressbar, progressBar)
    else:
        # All backend work is complete; switch to results view
        update_window()

"""
Initialize and launch the matching algorithm in a separate thread to
avoid blocking the UI. Reset progress tracking and begin updating the bar.

Args:
    sheet_name (str): Title of the Google Sheet containing raw response data.
    progressBar (CTkProgressBar): Widget to display algorithm progress.
    student_entries (list[CTkEntry]): Column-letter inputs for student data.
    host_entries (list[CTkEntry]): Column-letter inputs for host data.
    student_sheet (str): Name of the worksheet tab for student responses.
    host_sheet (str): Name of the worksheet tab for host responses.
"""
def start_processing(sheetName, progressBar, studentResponses, hostResponses, studentName, hostName):
    global progress
    progress = 0
    # Launch backend processing thread with user-provided parameters
    threading.Thread(target = backEnd.runCode, args = (sheetName, studentResponses, hostResponses, studentName, hostName)).start()
    update_progressbar(progressBar)

"""
Create a tabbed interface for mapping spreadsheet columns to data fields.
Tab 1: Student response columns
Tab 2: Host response columns

Populates global lists: student_inputs and host_inputs
"""
def inputFrame():
    global studentInputs
    global hostInputs
    global inputFrame
    inputFrame = ct.CTkTabview(scrollableFrame, segmented_button_selected_hover_color = "#1F6AA5", border_width = 1, width = 550)
    inputFrame.pack(pady = (0, 10), padx = 10)
    
    # Student tab setup
    inputFrame.add("Student Responses")
    instructions1 = ct.CTkLabel(inputFrame.tab("Student Responses"), 
                                text = "INPUT INSTRUCTIONS: \n\nPlease enter the LETTER of the column where the following \ncolumns are" \
                                " located in the sheet. For example, if the Student \nChoice 1 is in column K, enter the corresponding letter \"K\".")
    instructions1.grid(pady = 20, padx = 100)
    
    # Host tab setup
    inputFrame.add("Host Responses")
    instructions2 = ct.CTkLabel(inputFrame.tab("Host Responses"), 
                                text = "INPUT INSTRUCTIONS: \n\nPlease enter the LETTER of the column where the following \ncolumns are" \
                                " located in the sheet. For example, if the Host \nIndustry is in column M, enter the corresponding letter \"M\".")
    instructions2.grid(pady = 20, padx = 100)
    
    # Define field prompts
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

"""
Render two scrollable panels listing hosts by capacity status:
    - Filled Hosts: hosts with no remaining slots
    - Unfilled Hosts: hosts with available capacity

Uses backEnd.result and backEnd.capacities for data.
"""
def hostFrame():
    global rightFrame
    global leftFrame
    hostFrame = ct.CTkFrame(scrollableFrame, fg_color="#242424", bg_color="#242424")
    hostFrame.pack(pady = 40)
    leftFrame = ct.CTkScrollableFrame(hostFrame, border_width = 3, label_text = "Filled Hosts", 
                                        label_fg_color="#1F6AA5", height = 350, width = 300, bg_color="#2B2B2B", fg_color="#2B2B2B")
    leftFrame.pack(side = ct.LEFT, padx = (0, 50))
    leftFrame._scrollbar.grid(padx = (0, 5))
    innerFrame = ct.CTkFrame(leftFrame, fg_color = "#2B2B2B", bg_color="#2B2B2B")
    innerFrame.pack()
    rightFrame = ct.CTkScrollableFrame(hostFrame, border_width = 3, label_text = "Unfilled Hosts", label_fg_color="#1F6AA5", height = 350, width = 300)
    rightFrame.pack()
    rightFrame._scrollbar.grid(padx = (0, 5))
    innerFrame2 = ct.CTkFrame(rightFrame, fg_color = "#2B2B2B", bg_color="#2B2B2B")
    innerFrame2.pack()
    
    # Segregate hosts based on capacity
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

"""
Present a summary of matching outcomes in six statistic cards:
    1. Successful Student Matches
    2. Unsuccessful Student Matches
    3. Successful Host Matches
    4. Unsuccessful Host Matches
    5. Breakdown of Green vs. Yellow Student Lists
    6. Overall Capacity Utilization
"""
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

"""
Remove input widgets and progress bar, then display match results and stats.
"""
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

"""
Initialize the main window, load resources, and assemble the initial UI.
Enter the CTkinter event loop.
"""
def main():
    global scrollableFrame, progressBar, inputQ, sheetName, window, \
    inputQ2, studentName, inputQ3, hostName, preReq, runButton
    
    # Determine application directory for resource paths
    exe_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(exe_dir, 'UD.ico')
    logo_path = os.path.join(exe_dir, 'UD.png')

    # Create main window
    window = ct.CTk()
    window.iconbitmap(icon_path)
    window.title("Job Shadow Matching Program")
    window.geometry("900x800")
    window.configure(fg_color="#242424", bg_color="#242424")
    ct.set_appearance_mode("dark")

    # Scrollable container for dynamic content
    scrollableFrame = ct.CTkScrollableFrame(window, fg_color="#242424", bg_color="#242424")
    scrollableFrame.pack(fill="both", expand=True)

    # Display UD logo at top
    canvas = ct.CTkCanvas(scrollableFrame, width=340, height=160, bg="#242424", highlightthickness=0)
    logo_img = open_image(logo_path)
    w, h = logo_img.width(), logo_img.height()
    canvas.create_image(w//2, h//2 + 10, image=logo_img)
    canvas.pack(pady=20)

    # Instruction for sharing the sheet
    preReq = ct.CTkLabel(scrollableFrame, 
                        text="Please share the Google Sheet with this email:\n matching-prog@cisc498-group2.iam.gserviceaccount.com")
    preReq.pack(pady=(10, 25))

    # Google Sheet name input
    inputQ = ct.CTkLabel(scrollableFrame, text="Name of Google Sheet:")
    inputQ.pack(pady=(10, 5))
    sheetName = ct.CTkEntry(scrollableFrame, width=300)
    sheetName.pack(pady=(0, 40))

    # Student responses worksheet name input
    inputQ2 = ct.CTkLabel(scrollableFrame, text="Name of Student Responses Sheet:")
    inputQ2.pack(pady=(10, 5))
    studentName = ct.CTkEntry(scrollableFrame, width=300)
    studentName.pack(pady=(0, 40))

    # Host responses worksheet name input
    inputQ3 = ct.CTkLabel(scrollableFrame, text="Name of Host Responses Sheet:")
    inputQ3.pack(pady=(10, 5))
    hostName = ct.CTkEntry(scrollableFrame, width=300)
    hostName.pack(pady=(0, 40))

    # Initialize input tabs
    inputFrame()

    # Run button starts backend matching
    runButton = ct.CTkButton(
        scrollableFrame,
        corner_radius=75,
        hover_color="#488dc2",
        text="Run",
        command=lambda: start_processing(
            sheetName.get(), progressBar,
            studentInputs, hostInputs,
            studentName.get(), hostName.get()
        )
    )
    runButton.pack(pady=(40, 20))

    # Progress bar to visualize algorithm status
    progressBar = ct.CTkProgressBar(
        scrollableFrame,
        border_width=1,
        height=15,
        width=400,
        orientation='horizontal',
        mode='determinate'
    )
    progressBar.set(0)
    progressBar.pack(pady=(30, 70))

    # Launch the UI loop
    window.mainloop()

if __name__ == "__main__":
    main()