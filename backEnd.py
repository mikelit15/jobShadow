#https://matching.readthedocs.io/en/latest/discussion/hospital_resident/
#https://pypi.org/project/matching/

import time     
from matching.games import HospitalResident                     # pip install matching
from google.oauth2.service_account import Credentials           # pip install google-auth
import gspread                                                  # pip install gspread
import pandas as pd
import sys
import customtkinter as ct                                      # pip install customtkinter
import os
from more_itertools import collapse                             # pip install more-itertools
import frontEnd

def runCode(sheetName, stdResponses, hostResponses, studentName, hostName):

    global progress  

    # Progress bar flag #0.075
    for _ in range(30):
        time.sleep(0.02)
        progress += .0025
    
    start_time = time.time()
    sys.setrecursionlimit(5000)

    # Path for .json for .exe bundle
    exe_dir = os.path.dirname(os.path.abspath(__file__))
    data_path = os.path.join(exe_dir, 'cisc498-group2-4f0a0c27d47d.json')   

    scopes = ['https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive']
    credentials = Credentials.from_service_account_file(
        data_path, scopes=scopes)
    client = gspread.authorize(credentials)
    spreadsheet = client.open(sheetName) # Nicks BLANK of 2022 Summer Job Shadow Program
    
    # Returns boolean if a worksheet with title (sheet_name) exists
    def doesExist(spreadsheet, sheet_name):
        try:
            spreadsheet.worksheet(sheet_name)
            return True
        except gspread.WorksheetNotFound:
            return False
        
    def converterStart(start):
        num = 0
        for c in start:
            if c.isalpha():
                num = num * 26 + (ord(c.upper()) - ord('A')) + 1
        return num
    
    def converterOffset(letter, start):
        num = 0
        for c in letter:
            if c.isalpha():
                num = num * 26 + (ord(c.upper()) - ord('A')) + 1
                num = num - (start - 1)
        return num
    #Draws from given Google Sheet, make new worksheet (NEWER STUDENT) to later draw from
    '''

    NEWER STUDENTS 

    '''
    name = 'NEWER STUDENTS:'
    if(doesExist(spreadsheet, name)):
        spreadsheet.del_worksheet(spreadsheet.worksheet(name))
        if(doesExist(spreadsheet, "NEWER MENTORS:")):
            spreadsheet.del_worksheet(spreadsheet.worksheet("NEWER MENTORS:"))
        if(doesExist(spreadsheet, "Multi-Round: Students")):
            spreadsheet.del_worksheet(spreadsheet.worksheet("Multi-Round: Students"))
        if(doesExist(spreadsheet, "Multi-Round: Hosts")):
            spreadsheet.del_worksheet(spreadsheet.worksheet("Multi-Round: Hosts"))
        if(doesExist(spreadsheet, "Multi-Round Summary")):
            spreadsheet.del_worksheet(spreadsheet.worksheet("Multi-Round Summary"))
        spreadsheet.add_worksheet(rows=200,cols=7,title=name)
        sheet_instances = spreadsheet.worksheet(name)
    else:
        spreadsheet.add_worksheet(rows=200,cols=7,title=name)
        sheet_instances = spreadsheet.worksheet(name)

    # Progress bar flag #0.075
    for _ in range(30):
        time.sleep(0.02)
        progress += .0025

    sheet_instances.update('A1', 'EMAILS')
    sheet_instances.update('B1', 'Choice 1')
    sheet_instances.update('C1', 'Choice 2')
    sheet_instances.update('D1', 'Choice 3')
    sheet_instances.update('E1', 'Industry Areas')##WOULD BE H ANOTHER HOST SHOULD BE E
    sheet_instances.update('F1', 'SCHOOL YEAR')
    sheet_instances.update('G1', 'MAJOR')
    
    #stuemails=input('what letter column is UD emailss:\n')
    stuEmails=stdResponses[0].get()
    sheet_instances.update_cell(2, 1, '=unique(filter(\'{}\'!${}2:${}1000,\'{}\'!${}2:${}1000<>CHAR(34)))'.format(studentName, stuEmails, stuEmails, studentName, stuEmails, stuEmails))
    #stuIndu=input('What is the # of columns from UD email to Choice 1 ? example Email is D and Industry is J you should enter 3:\n')
    stuStart = converterStart(stuEmails)
    stuChoice1 = converterOffset(stdResponses[1].get(), stuStart)
    sheet_instances.update_cell(2, 2, '=IFERROR(LEFT(VLOOKUP($A2,\'{}\'!${}$2:$Z$999,{},FALSE), FIND(":", VLOOKUP($A2,\'{}\'!${}$2:$Z$999,{},FALSE))-1),0)'.format(studentName, stuEmails, stuChoice1, studentName, stuEmails, stuChoice1))
    #stuIndu=input('What is the # of columns from UD email to Choice 2? Just Add 1 from Choice 1 response:\n')
    stuChoice2 = converterOffset(stdResponses[2].get(), stuStart)
    sheet_instances.update_cell(2, 3, '=IFERROR(LEFT(VLOOKUP($A2,\'{}\'!${}$2:$Z$999,{},FALSE), FIND(":", VLOOKUP($A2,\'{}\'!${}$2:$Z$999,{},FALSE))-1),0)'.format(studentName, stuEmails, stuChoice2, studentName, stuEmails, stuChoice2))
    #stuIndu=input('What is the # of columns from UD email to Choice 3? Just Add 2 from Choice 1 response:\n')
    stuChoice3 = converterOffset(stdResponses[3].get(), stuStart)
    sheet_instances.update_cell(2, 4, '=IFERROR(LEFT(VLOOKUP($A2,\'{}\'!${}$2:$Z$999,{},FALSE), FIND(":", VLOOKUP($A2,\'{}\'!${}$2:$Z$999,{},FALSE))-1),0)'.format(studentName, stuEmails, stuChoice3, studentName, stuEmails, stuChoice3))
    #stuIndu=input('What is the # of columns from UD email to Industry Area ? example Email is D and Industry is J you should enter 3:\n')
    stuIndu = converterOffset(stdResponses[4].get(), stuStart)
    sheet_instances.update_cell(2, 5, '=VLOOKUP($A2,\'{}\'!${}$2:$Z$999,{},FALSE)'.format(studentName, stuEmails, stuIndu))
    #stuYear=input('What is the # of columns from UD email to School Year? example Email is D and Industry is J you should enter 3:\n')
    stuYear = converterOffset(stdResponses[5].get(), stuStart)
    sheet_instances.update_cell(2, 6, '=VLOOKUP($A2,\'{}\'!${}$2:$Z$999,{},FALSE)'.format(studentName, stuEmails, stuYear))
    #stuMajor=input('What is the # of columns from UD email to Major? example Email is D and Industry is J you should enter 3:\n')
    stuMajor = converterOffset(stdResponses[6].get(), stuStart)
    sheet_instances.update_cell(2, 7, '=VLOOKUP($A2,\'{}\'!${}$2:$Z$999,{},FALSE)'.format(studentName, stuEmails, stuMajor))
    print("TEST")
    
    #Draws from given Google Sheet, make new worksheet (NEWER MENTORS) to later draw from
    '''

    NEWER MENTORS
    
    '''
    
    name = 'NEWER MENTORS:'
    if(doesExist(spreadsheet, name)):
        spreadsheet.del_worksheet(spreadsheet.worksheet(name))
        if(doesExist(spreadsheet, 'NEWER STUDENTS:')):
            spreadsheet.del_worksheet(spreadsheet.worksheet('NEWER STUDENTS:'))
        if(doesExist(spreadsheet, "Matching Round One: Students")):
            spreadsheet.del_worksheet(spreadsheet.worksheet("Matching Round One: Students"))
        if(doesExist(spreadsheet, "Matching Round One: Hosts")):
            spreadsheet.del_worksheet(spreadsheet.worksheet("Matching Round One: Hosts"))
        if(doesExist(spreadsheet, "Matching Round One Summary")):
            spreadsheet.del_worksheet(spreadsheet.worksheet("Matching Round One Summary"))
        spreadsheet.add_worksheet(rows=200,cols=9,title=name)
        sheet_instances = spreadsheet.worksheet(name)
    else:
        spreadsheet.add_worksheet(rows=200,cols=9,title=name)
        sheet_instances = spreadsheet.worksheet(name)

    sheet_instances = spreadsheet.worksheet(name)
    sheet_instances.update('A1', 'ID')
    sheet_instances.update('B1', 'EMAILS')
    sheet_instances.update('C1', 'Industry Catergory')
    sheet_instances.update('D1', 'Capacity')
    sheet_instances.update('E1', 'Virtual/Inperson')
    sheet_instances.update('F1', 'PROGRAM OF STUDY')
    sheet_instances.update('G1', 'JOB TITLE')
    sheet_instances.update('H1', 'COMPANY NAME')
    sheet_instances.update('I1', 'INTERESTED MAJORS')
    
    #hostID=input('what letter column is Job IDs:\n')
    hostID = hostResponses[0].get()
    sheet_instances.update_cell(2, 1, '=unique(filter(\'{}\'!A2:A,\'{}\'!A2:A<>""))'.format(hostName, hostName))
    #hostEmail=input('What is the # of columns from jobid to email? example ID is A and Email is C you should enter 3:\n')
    hostStart = converterStart(hostID)
    hostEmail = converterOffset(hostResponses[1].get(), hostStart)
    sheet_instances.update_cell(2, 2, '=VLOOKUP($A2,\'{}\'!${}$2:$AI$999,{},FALSE)'.format(hostName, hostID, hostEmail))
    #hostIndustry=input('What is the # of columns from jobid to Industry? example ID is A and Email is C you should enter 3:\n')
    hostIndustry = converterOffset(hostResponses[2].get(), hostStart)
    sheet_instances.update_cell(2, 3, '=VLOOKUP($A2,\'{}\'!${}$2:$AI$999,{},FALSE)'.format(hostName, hostID, hostIndustry))
    #hostCapacity=input('What is the # of columns from jobid to Capacity? example ID is A and Email is C you should enter 3:\n')
    hostCapacity = converterOffset(hostResponses[3].get(), hostStart)
    sheet_instances.update_cell(2, 4, '=VLOOKUP($A2,\'{}\'!${}$2:$AI$999,{},FALSE)'.format(hostName, hostID, hostCapacity))
    #hostVirtInper=input('What is the # of columns from jobid to Virtual/Inperson? example ID is A and Email is C you should enter 3:\n')
    hostVirtInper = converterOffset(hostResponses[4].get(), hostStart)
    sheet_instances.update_cell(2, 5, '=VLOOKUP($A2,\'{}\'!${}$2:$AI$999,{},FALSE)'.format(hostName, hostID, hostVirtInper))
    #hostPrgmStudy=input('What is the # of columns from jobid to Program of Study? example ID is A and Email is C you should enter 3:\n')
    hostPrgmStudy = converterOffset(hostResponses[5].get(), hostStart)
    sheet_instances.update_cell(2, 6, '=VLOOKUP($A2,\'{}\'!${}$2:$AI$999,{},FALSE)'.format(hostName, hostID, hostPrgmStudy))
    #hostJobTit=input('What is the # of columns from jobid to Job Title? example ID is A and Email is C you should enter 3:\n')
    hostJobTit = converterOffset(hostResponses[6].get(), hostStart)
    sheet_instances.update_cell(2, 7, '=VLOOKUP($A2,\'{}\'!${}$2:$AI$999,{},FALSE)'.format(hostName, hostID, hostJobTit))
    #hostCompName=input('What is the # of columns from jobid to Company Name? example ID is A and Email is C you should enter 3:\n')
    hostCompName = converterOffset(hostResponses[7].get(), hostStart)
    sheet_instances.update_cell(2, 8, '=VLOOKUP($A2,\'{}\'!${}$2:$AI$999,{},FALSE)'.format(hostName, hostID, hostCompName))
    #hostIntMajors=input('What is the # of columns from jobid to Interested Majors? example ID is A and Email is C you should enter 3:\n')
    hostIntMajors = converterOffset(hostResponses[8].get(), hostStart)
    sheet_instances.update_cell(2, 9, '=VLOOKUP($A2,\'{}\'!${}$2:$AI$999,{},FALSE)'.format(hostName, hostID, hostIntMajors))

    '''

    '''
    # Progress bar flag #0.25
    for _ in range(50):
        time.sleep(0.04)
        progress += .005

    # Opens up a pause window with a continute button, pauses the code so the google sheets can be adjusted before running the rest of the code
    frontEnd.pause_execution.set(True)
    def continue_code():
        popup.destroy()
        frontEnd.pause_execution.set(False)
    popup = ct.CTkToplevel(frontEnd.window)
    popup.title("Pause")
    label = ct.CTkLabel(popup, text="Please go to the Google Sheet, and expand the columns\n of the 2 new sheets that were created. \n\nClick 'Continue' to proceed.")
    label.pack(padx=20, pady=20)
    continue_button = ct.CTkButton(popup, text="Continue", command=continue_code)
    continue_button.pack(pady=10)
    frontEnd.window.wait_variable(frontEnd.pause_execution)

    #From new Google sheet sheets, take in actual data we will use to match
    sheet_instance = spreadsheet.worksheet('NEWER STUDENTS:')  # 14 FOR 2022 SUMMER, 18 NEW STUDENTS FOR 2023 Winter
    records_data = sheet_instance.get_all_records()
    student_df = pd.DataFrame.from_dict(records_data)
    sheet_instance = spreadsheet.worksheet('NEWER MENTORS:') # 15 FOR 2022 SUMMER, 19 NEW HOSTS FOR 2023 Winter
    records_data = sheet_instance.get_all_records()
    host_df = pd.DataFrame.from_dict(records_data)

    student_test = student_df.apply(lambda df : True
                if "@" in df['EMAILS'] else False, axis = 1)
    student_indeces = student_test[student_test == True].index
    student_df = student_df.loc[student_indeces, :]
    student_df['Virtual/Inperson'] = ''
    student_df['Choice 1 Industry Catergory'] = ''
    student_df['Choice 2 Industry Catergory'] = ''
    student_df['Choice 3 Industry Catergory'] = ''
    student_df.set_index('EMAILS', inplace=True)
    host_test = host_df.apply(lambda df : True
                if isinstance(df['ID'], int) else False, axis = 1)
    host_indeces = host_test[host_test == True].index
    host_df = host_df.loc[host_indeces, :]
    host_df = pd.concat([host_df], ignore_index=True)
    host_df.set_index('ID', inplace=True)
    num_students, num_hosts = len(student_df), len(host_df)
    
    #preset empty host preferences
    student_prefs = {
        s: []
        for s in student_df.index
    }
    host_prefs = {
        h: []
        for h in host_df.index
    }
    
    #Find and take host capacities
    global capacities
    capacities = {h: host_df.at[h, 'Capacity'] for h in host_prefs}
    print(capacities)
    
    
    def det_delivery_pref(pref, current_pref):
        if (current_pref == 'Both' or pref == current_pref):
            return current_pref
        elif (not current_pref):
            return pref
        else:
            return 'Both'

    student_inds = {
        s: []
        for s in student_df.index
    }
    #Go over each student by email, set student and hosts preferences
    for email in student_df.index:
        for choice in ['Choice 1', 'Choice 2', 'Choice 3']:
            if isinstance(int(student_df.at[email, choice]), int) and (int(student_df.at[email, choice]) >= 0 and int(student_df.at[email, choice]) not in student_prefs[email]):
                student_prefs[email].append(int(student_df.at[email, choice]))
                student_df.at[email, 'Virtual/Inperson'] = det_delivery_pref(host_df.at[student_df.at[email, choice], 'Virtual/Inperson'], student_df.at[email, 'Virtual/Inperson'])
                student_df.at[email, (choice + " Industry Catergory")] = host_df.at[student_df.at[email, choice], 'Industry Catergory']
                if(host_df.at[student_df.at[email, choice], 'Industry Catergory'] not in student_inds[email]):
                    student_inds[email].append(host_df.at[student_df.at[email, choice], 'Industry Catergory'])
                if (email not in host_prefs[student_df.at[email, choice]]):
                    host_prefs[student_df.at[email, choice]].append(email)
    
    #For ranking student senority, can be changed, this is similar to what they got when they did it by hand
    years = {'First Year': 0, 'Sophomore': 1, 'Junior': 2, 'Senior': 4, 'Graduate Student': 1}
    for host in host_prefs:
        host_prefs[host] = sorted(host_prefs[host], key=lambda x: years[student_df.at[x, 'SCHOOL YEAR']], reverse=True)

    #Using HospitalResident from the "" library create and run game which matches based on Hospital residency problem (see link at top)
    game1 = HospitalResident.create_from_dictionaries(
        student_prefs, host_prefs, capacities
    )
    print(student_prefs)
    roundone = game1.solve()
    print(roundone)
   
    #Set empty lists for sorting how the "game1" sorted the students
    gameOne_green_host_list = []
    gameOne_yellow_host_list = []
    gameOne_red_host_list = []
    gameOne_green_student_list = []
    gameOne_yellow_student_list = []
    gameOne_red_student_list = []
    
    #This is for setting students and hosts that were matches well together (green; obvious matches), not as well and might need a second look (yellow; less obvious matches),  and ones that werent matched (red)
    global gameOne_totalCapacity
    gameOne_totalCapacity = 0
    
    for host in roundone:
        gameOne_totalCapacity = gameOne_totalCapacity + capacities[int(host.name)]
    for host in roundone:
        if (host.matching):
            gameOne_green_host_list.append(host.name)
            for student in roundone[host]:
                if (roundone[host][-1] == student):
                    gameOne_yellow_student_list.append(student.name)
                else:
                    gameOne_green_student_list.append(student.name)
        else:
            gameOne_red_host_list.append(host.name)
    for student in game1.residents:
        if (student.matching == None):
            gameOne_red_student_list.append(student.name)
    
    #finding gameones total capacity, for later use in round two
    for host in roundone:
        gameOne_totalCapacity = gameOne_totalCapacity + capacities[int(host.name)]

    def corresponding_key(val, dictionary):
        for k, v in dictionary.items():
            if val in v:
                return k
    print("Round One", time.time() - start_time)     
    print("___________________________________________________________")
    print("___________________________________________________________")
    print("Round Two")
    print("___________________________________________________________")
    #Beginning of round two
    completedhosts = []
    
    #For updating what hosts/students are left and their capacities

    for host in roundone:
        if capacities[int(host.name)] == len(host.matching):
            completedhosts.append(int(host.name))    

    popme = host_prefs.copy()

    for x in completedhosts:
        del popme[x]
    for_capacities = popme.copy()
    global new_capacities
    new_capacities = {h: host_df.at[h, 'Capacity'] for h in for_capacities}

    for host in roundone:
        if host.name in host_prefs:
            if host.name in for_capacities:
                new_capacities[int(host.name)] = new_capacities[int(host.name)] - len(host.matching)

    #############
    
    
    new_student_prefs = {
        s: []
        for s in student_df.index
    }

    new_student_inds = {
        s: []
        for s in student_df.index
    }
      
    new_a = {}
    for host in roundone:
        new_students = []
        students = host.matching
        for c in students:
            new_students.append(c.name)
        new_a[host.name] = new_students
    
    #Basically same as round one, just without completed students and hosts
    for email in student_df.index:
        for choice in ['Choice 1', 'Choice 2', 'Choice 3']:
            if isinstance(int(student_df.at[email, choice]), int) and (int(student_df.at[email, choice]) >= 0 and int(student_df.at[email, choice]) not in new_student_prefs[email]) and (int(student_df.at[email, choice]) not in completedhosts) and ((int(student_df.at[email, choice])) != 0):
                if (corresponding_key(email, new_a) == int(student_df.at[email, choice])):
                    continue
                else:
                    new_student_prefs[email].append(int(student_df.at[email, choice]))
                    student_df.at[email, 'Virtual/Inperson'] = det_delivery_pref(host_df.at[student_df.at[email, choice], 'Virtual/Inperson'], student_df.at[email, 'Virtual/Inperson'])
                    student_df.at[email, (choice + " Industry Catergory")] = host_df.at[student_df.at[email, choice], 'Industry Catergory']
                    if(host_df.at[student_df.at[email, choice], 'Industry Catergory'] not in new_student_inds[email]):
                        new_student_inds[email].append(host_df.at[student_df.at[email, choice], 'Industry Catergory'])

    new_student_prefs = {k:v for k,v in new_student_prefs.items() if v}
    new_host_prefs = {}
    for student in new_student_prefs:
        #print(student)
        host = new_student_prefs.get(student)
        for a in host:
            if a in list(new_host_prefs):
                new_host_prefs[a].append(student)
            else:   
                new_host_prefs[a] = [student]
    #print(new_capacities)
    game2 = HospitalResident.create_from_dictionaries(
        new_student_prefs, new_host_prefs, new_capacities
    )
    roundtwo = game2.solve()
    print(roundtwo)

    gameTwo_green_host_list = []
    gameTwo_yellow_host_list = []
    gameTwo_red_host_list = []
        
    gameTwo_green_student_list = []
    gameTwo_yellow_student_list = []
    gameTwo_red_student_list = []

    for host in roundtwo:
        if (host.matching):
            gameTwo_green_host_list.append(host.name)
            for student in roundtwo[host]:
                if (roundtwo[host][-1] == student):
                    gameTwo_yellow_student_list.append(student.name)
                else:
                    gameTwo_green_student_list.append(student.name)
        else:
            gameTwo_red_host_list.append(host.name)
    for student in game1.residents:
        if (student.matching == None):
            gameTwo_red_student_list.append(student.name)

    print("Round Two", time.time() - start_time)     
    #print(gameTwo_yellow_student_list)
    print("___________________________________________________________")
    print("Round Three")
    print("___________________________________________________________")
    
    #Round three is basically the same as the previous rounds in function, just the completed from round 2 (along with 1) are now excluded, and the final few are updated
    #Progress flag #0.225
    for _ in range(45):
        time.sleep(0.04)
        progress += .005
   
    for host in roundtwo:
        if new_capacities[int(host.name)] == len(host.matching):
            completedhosts.append(int(host.name))  

    popme = new_host_prefs.copy()

    for x in completedhosts:
        #print(x)
        if x in popme:
            del popme[x]
    for_capacities = popme.copy()
    global newer_capacities
    newer_capacities = {h: host_df.at[h, 'Capacity'] for h in new_host_prefs}

    for host in roundtwo:
        if host.name in host_prefs:
            if host.name in for_capacities:
                newer_capacities[int(host.name)] = new_capacities[int(host.name)] - len(host.matching)

    newer_student_prefs = {
        s: []
        for s in student_df.index
    }

    newer_student_inds = {
        s: []
        for s in student_df.index
    }
        
    new_b = {}
    for host in roundtwo:
        new_students = []
        students = host.matching
        for c in students:
            new_students.append(c.name)
        new_b[host.name] = new_students
            
    for email in student_df.index:
        #print(email)
        for choice in ['Choice 1', 'Choice 2', 'Choice 3']:
            if isinstance(int(student_df.at[email, choice]), int) and (int(student_df.at[email, choice]) >= 0) and (int(student_df.at[email, choice]) not in newer_student_prefs[email]) and (int(student_df.at[email, choice]) not in completedhosts) and ((int(student_df.at[email, choice])) != 0):
                if (corresponding_key(email, new_a) == int(student_df.at[email, choice])) or (corresponding_key(email, new_b) == int(student_df.at[email, choice])):
                    continue
                else:
                    newer_student_prefs[email].append(int(student_df.at[email, choice]))
                    student_df.at[email, 'Virtual/Inperson'] = det_delivery_pref(host_df.at[student_df.at[email, choice], 'Virtual/Inperson'], student_df.at[email, 'Virtual/Inperson'])
                    student_df.at[email, (choice + " Industry Catergory")] = host_df.at[student_df.at[email, choice], 'Industry Catergory']
                    if(host_df.at[student_df.at[email, choice], 'Industry Catergory'] not in newer_student_inds[email]):
                        newer_student_inds[email].append(host_df.at[student_df.at[email, choice], 'Industry Catergory'])

                    


    newer_student_prefs = {k:v for k,v in newer_student_prefs.items() if v}
    newer_host_prefs = {}
    for student in newer_student_prefs:
        #print(student)
        host = newer_student_prefs.get(student)
        for each in host:
            if each in list(newer_host_prefs):
                newer_host_prefs[each].append(student)
            else:   
                newer_host_prefs[each] = [student]
    #print(newer_student_prefs)
    #print(newer_host_prefs)

    game3 = HospitalResident.create_from_dictionaries(
        newer_student_prefs, newer_host_prefs, newer_capacities
    )
    roundthree = game3.solve()
    # print(roundthree)

    gameThree_green_host_list = []
    gameThree_yellow_host_list = []
    gameThree_red_host_list = []
        
    gameThree_green_student_list = []
    gameThree_yellow_student_list = []
    gameThree_red_student_list = []

    for host in roundthree:
        if (host.matching):
            gameThree_green_host_list.append(host.name)
            for student in roundthree[host]:
                if (roundthree[host][-1] == student):
                    gameThree_yellow_student_list.append(student.name)
                else:
                    gameThree_green_student_list.append(student.name)
        else:
            gameThree_red_host_list.append(host.name)
    for student in game1.residents:
        if (student.matching == None):
            gameThree_red_student_list.append(student.name)
        
    #Any and all matching is done, but data will still later be manipulated    
    print("Round Three", time.time() - start_time)     
    print("___________________________________________________________")
    print("Final Round")
    print("___________________________________________________________")

    #For combing and manipulating data to be best sent into Google Sheet
    finaldict = {**roundone,**roundtwo, **roundthree}
    truefinal= {}

    matched = []

    for match in finaldict:
        if match.name in truefinal.keys():
            for c in match.matching:
                truefinal[match.name] = truefinal[match.name],c
        else:
            truefinal[match.name] = match.matching
        
    absolutefinal = {}
    for what in finaldict.keys():
        for yes in truefinal.keys():
            if what.name == yes:
                absolutefinal[what] = truefinal[yes]
                
    
    """
    truefinal= {}
    for match in finaldict:
        truefinal[match] = match.matching
        if match in truefinal:
            print("A")
    """
    for done in absolutefinal:
        absolutefinal[done] = list(collapse(absolutefinal[done]))
    #print(absolutefinal)

    global result
    result = {}

    for key,value in absolutefinal.items():
        if value not in result.values() or bool(value) == False:
            result[key] = value
    #Matching and data/result manipulation done
    ######################        
            
    print(result)
    print("Round all together now", time.time() - start_time)     
    print("_______________________________________________")
    print("For printing")
    print("________________________________________________")
    
    #Not used but scared to delete/comment
    final_green_host_list = []
    final_yellow_host_list = []
    final_red_host_list = []
        
    final_green_student_list = []
    final_yellow_student_list = []
    final_red_student_list = []
    ############
    
    check = []
    for a in game1.hospitals:
        if a.name not in check and bool(a.matching) == True:
            check.append(a.name)
    for b in game2.hospitals:
        if b.name not in check and bool(b.matching) == True:
            check.append(b.name)
    for c in game3.hospitals:
        if c.name not in check and bool(c.matching) == True:
            check.append(c.name)
    
    
    #make new worksheet for student matching results in Google Sheet
    def worksheet_handle(spreadsheet, title, rows, cols):
        if title in [worksheet.title for worksheet in spreadsheet.worksheets()]:
            spreadsheet.del_worksheet(spreadsheet.worksheet(title))
        return spreadsheet.add_worksheet(title=title, rows=str(rows), cols=str(cols))
    # for hospital in game.hospitals:
    # sheet_instance.batch_update()
    student_sheet_round1 = worksheet_handle(spreadsheet, "Multi-Round: Students", gameOne_totalCapacity + 1, 2)
    test_liststu1 = []
    test_liststu2 = []
    test_liststu3 = []

    formats = []
    for student in game1.residents:
        test_liststu1.append([student.name, str(student.matching)])
    for student in game2.residents:
        test_liststu2.append([student.name, str(student.matching)])
    for student in game3.residents:
        test_liststu3.append([student.name, str(student.matching)])
    for a in test_liststu1:
        for b in test_liststu2:
            if(a[0] == b[0]):
                a.append(b[1])
        for c in test_liststu3:
            if(a[0] == c[0]):
                a.append(c[1])
    #print(num_students)
    #For coloring the student matching results worksheet's cells and adding their results
    for num in range(2, num_students+2):
        if game1.residents[num-2].name in gameOne_green_student_list:
            formats.append({
                "range": "B" + str(num) + ":B" + str(num),
                "format": {
                    "backgroundColor": {
                        "red": 0,
                        "green": 1,
                        "blue": 0
                    },
                },
            },
            )
        elif game1.residents[num-2].name in gameOne_red_student_list:
            formats.append({
                "range": "B" + str(num) + ":B" + str(num),
                "format": {
                    "backgroundColor": {
                        "red": 1,
                        "green": 0,
                        "blue": 0
                    },
                },
            },
            )
        else:
            formats.append({
                "range": "B" + str(num) + ":B" + str(num),
                "format": {
                    "backgroundColor": {
                        "red": 1,
                        "green": 1,
                        "blue": 0
                    },
                },
            },
            )


    student_sheet_round1.update('A1', 'STUDENT')  
    student_sheet_round1.update('B1', 'Round 1 : HOST ID')

    #Progress flag #0.225
    for _ in range(45):
        time.sleep(0.04)
        progress += .005
        
    #print(test_liststu1)
    #print(gameOne_totalCapacity)
    student_sheet_round1.batch_update([{
        'range': 'A2:D' + str(1000),
        'values': test_liststu1,
    }])
    #print(game2.residents)
    val = student_sheet_round1.col_values(3)
    student_sheet_round1.update('C1', 'Round 2 : HOST ID')
    intgameTwo_green_host_list = []
    intgameTwo_red_host_list= []
    #print(gameTwo_yellow_student_list)
    #print(gameTwo_green_host_list)
    for a in gameTwo_green_host_list:
        intgameTwo_green_host_list.append(str(a))
    for num in range(1, len(val)+1):
        if (val[num-1] in intgameTwo_green_host_list):
            formats.append({
                "range": "C" + str(num) + ":C" + str(num),
                "format": {
                    "backgroundColor": {
                        "red": 1,
                        "green": 1,
                        "blue": 0
                    },
                },
            },
            )
        elif (val[num-1] == 'None'):
            formats.append({
                "range": "C" + str(num) + ":C" + str(num),
                "format": {
                    "backgroundColor": {
                        "red": 1,
                        "green": 0,
                        "blue": 0
                    },
                },
            },
            )
        """
        else:
            formats.append({
                "range": "C" + str(num) + ":C" + str(num),
                "format": {
                    "backgroundColor": {
                        "red": 1,
                        "green": 1,
                        "blue": 0
                    },
                },
            },
            )
            """
        
    #for a in range(219):
    #    print(student_sheet_round1.cell(a+2, 3).value)
    student_sheet_round1.batch_format(formats)

    val = student_sheet_round1.col_values(4)
    student_sheet_round1.update('D1', 'Round 3 : HOST ID')
    intgameThree_green_host_list= []

    for a in gameThree_green_host_list:
        intgameThree_green_host_list.append(str(a))
    #print(len(val))
    for num in range(1, len(val)+1):
        if (val[num-1] in intgameThree_green_host_list):
            formats.append({
                "range": "D" + str(num) + ":D" + str(num),
                "format": {
                    "backgroundColor": {
                        "red": 1,
                        "green": 1,
                        "blue": 0
                    },
                },
            },
            )
        elif (val[num-1] == 'None'):
            formats.append({
                "range": "D" + str(num) + ":D" + str(num),
                "format": {
                    "backgroundColor": {
                        "red": 1,
                        "green": 0,
                        "blue": 0
                    },
                },
            },
            )
            """
        else:
            formats.append({
                "range": "D" + str(num) + ":D" + str(num),
                "format": {
                    "backgroundColor": {
                        "red": 1,
                        "green": 1,
                        "blue": 0
                    },
                },
            },
            )
            """
        
    #for a in range(219):
    #    print(student_sheet_round1.cell(a+2, 3).value)
    student_sheet_round1.batch_format(formats)
    #make new worksheet for host matching results in Google Sheet
    hosts_sheet_round1 = worksheet_handle(spreadsheet, "Multi-Round: Hosts", num_hosts + 1, 2)
    test_listgame1 = []
    test_listgame2 = []
    test_listgame3 = []

    formats2 = []
    for host in game1.hospitals:
        test_listgame1.append([host.name, ', '.join(str(student) for student in host.matching)])
    for host in game2.hospitals:
        test_listgame2.append([host.name, ', '.join(str(student) for student in host.matching)])
    for host in game3.hospitals:
        test_listgame3.append([host.name, ', '.join(str(student) for student in host.matching)])
    for a in test_listgame1:
        for b in test_listgame2:
            if(a[0] == b[0]):
                a.append(b[1])
        for c in test_listgame3:
            if(a[0] == c[0]):
                a.append(c[1])

    #For coloring the host matching results worksheet's cells and adding their results
    for num in range(2, num_hosts+2):
        if game1.hospitals[num-2].name in gameOne_green_host_list:
            formats2.append({
                "range": "B" + str(num) + ":B" + str(num),
                "format": {
                    "backgroundColor": {
                        "red": 0,
                        "green": 1,
                        "blue": 0
                    },
                },
            },
            )
        elif game1.hospitals[num-2].name in gameOne_red_host_list:
            formats2.append({
                "range": "B" + str(num) + ":B" + str(num),
                "format": {
                    "backgroundColor": {
                        "red": 1,
                        "green": 0,
                        "blue": 0
                    },
                },
            },
            )
        else:
            formats2.append({
                "range": "B" + str(num) + ":B" + str(num),
                "format": {
                    "backgroundColor": {
                        "red": 1,
                        "green": 1,
                        "blue": 0
                    },
                },
            },
            )
    hosts_sheet_round1.update('A1', 'HOST ID')
    hosts_sheet_round1.update('B1', 'STUDENT(S)')
    hosts_sheet_round1.batch_update([{
        'range': 'A2:D' + str(1000),
        'values': test_listgame1,
    }])
    hosts_sheet_round1.batch_format(formats2)

    val = hosts_sheet_round1.col_values(3)
    #print(len(val))
    for num in range(1, len(val)+1):
        if (val[num-1] == ''):
            formats2.append({
                "range": "C" + str(num) + ":C" + str(num),
                "format": {
                    "backgroundColor": {
                        "red": 1,
                        "green": 0,
                        "blue": 0
                    },
                },
            },
            )
        else:
            formats2.append({
                "range": "C" + str(num) + ":C" + str(num),
                "format": {
                    "backgroundColor": {
                        "red": 0,
                        "green": 1,
                        "blue": 0
                    },
                },
            },
            )
        formats2.append({
            "range": "C" + str(1) + ":C" + str(1),
                "format": {
                    "backgroundColor": {
                        "red": 1,
                        "green": 1,
                        "blue": 1
                    },
                },
            },
        )
        
    #for a in range(219):
    #    print(student_sheet_round1.cell(a+2, 3).value)
    hosts_sheet_round1.batch_format(formats2)

    val = hosts_sheet_round1.col_values(4)
    #print(len(val))
    for num in range(1, len(val)+1):
        if (val[num-1] == ''):
            formats2.append({
                "range": "D" + str(num) + ":D" + str(num),
                "format": {
                    "backgroundColor": {
                        "red": 1,
                        "green": 0,
                        "blue": 0
                    },
                },
            },
            )
        else:
            formats2.append({
                "range": "D" + str(num) + ":D" + str(num),
                "format": {
                    "backgroundColor": {
                        "red": 0,
                        "green": 1,
                        "blue": 0
                    },
                },
            },
            )
        formats2.append({
            "range": "D" + str(1) + ":D" + str(1),
                "format": {
                    "backgroundColor": {
                        "red": 1,
                        "green": 1,
                        "blue": 1
                    },
                },
            },
        )
        
    #for a in range(219):
    #    print(student_sheet_round1.cell(a+2, 3).value)
    hosts_sheet_round1.batch_format(formats2)


    test_list3 = [["Multi-Round Summary"], 
    ["Number of Student with a match", len(gameOne_green_student_list) + len(gameOne_yellow_student_list)],
    ["Number of Student without a match", len(gameOne_red_student_list)],
    ["Number of Hosts with a match", len(check)],
    ["Number of Hosts without a match", len(game1.hospitals) - len(check)]]
    summary_sheet_round1 = worksheet_handle(spreadsheet, "Multi-Round Summary", 5, 2)
    summary_sheet_round1.batch_update([{
        'range': 'A2:B' + str(num_hosts + 1),
        'values': test_list3,
    }])
    print("All done", time.time() - start_time)     
    
    # Creates variables to call in updateWindow(), required for stat frame
    global studentsWithout
    global studentsWith
    global totStudents
    global hostsWith
    global hostsWithout
    global totHosts
    global solidStudents
    global maybeStudents
    global totalCap
    global totalFill
    studentsWith = len(gameOne_green_student_list) + len(gameOne_yellow_student_list)
    studentsWithout = len(gameOne_red_student_list)
    solidStudents = len(gameOne_green_student_list)
    maybeStudents = len(gameOne_yellow_student_list)
    hostsWith = len(check)
    hostsWithout = len(game1.hospitals) - len(check)
    totStudents = len(game1.residents)
    totHosts = len(game1.hospitals)
    totalCap = gameOne_totalCapacity//2
    totalFill = len(gameOne_yellow_student_list + gameOne_green_student_list + gameTwo_yellow_student_list +
                        gameTwo_green_student_list + gameThree_yellow_student_list + gameThree_green_student_list)
    
    print("My program took", time.time() - start_time, "to run")
    
    # Progress bar flag #0.15
    for _ in range(30):
        time.sleep(0.02)
        progress += .005