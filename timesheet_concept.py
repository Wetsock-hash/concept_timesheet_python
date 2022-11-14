'''
This program is a concept design for an adroid app that will be made
to assist someone with their job. They currently hand write their timesheet
each day. This concept accepts user input and writes to xlsx file however it
does not filter or wash the user input, as it is for their benefit.
'''
# import xlsxwriter module
import xlsxwriter
# import datetime module
import datetime as dt

class timesheet:
    # Three global variables that each function works with
    start = []
    jobs = []
    eod = []

    # This function is simplistic in that it will ask a series of questions and append the answers to a list
    # which is then appended to a list (a list of lists for multiple clients).
    def add_client(jobs):
        job = []
        print("Client Name: ")
        client_name = input()
        print("Client Address: ")
        client_address = input()
        print("Suburb: ")
        suburb = input()
        print("Job type:")
        job_type = input()
        print("job start time:")
        job_start = input()
        print("job finish time:")
        job_finish = input()
        print("Extra charges")
        extra_charges = input()
        print("money type (cash, credit card, cheque)")
        money_type = input()
        if extra_charges == None:
            extra_charges = 0.0
        print("Job comments: ")
        job_comments = input()
        
        job.append(client_name)
        job.append(client_address)
        job.append(suburb)
        job.append(job_type)
        job.append(job_start)
        job.append(job_finish)
        job.append(extra_charges)
        job.append(money_type)
        job.append(job_comments)

        jobs.append(job)
        
    # This function prepares the eod variable and performs some simple maths to aquire the total hours
    # before appendng to the list
    def end_of_day(eod, odometer_1):
        print("Enter your odometer: ")
        odometer = int(input())
        kms = odometer_1 - odometer
        print("Enter your start time: (24hr format 00:00:00)")
        strt_time = input()
        start_time = strt_time.split(":")
        print("Enter your finish time: (24hr format 00:00:00)")
        fnsh_time = input()
        finish_time = fnsh_time.split(":")
        total_start_minutes = (int(start_time[0]) * 60) + int(start_time[1])
        total_finish_minutes = (int(finish_time[0]) * 60) + int(finish_time[1])
        total_hours = (total_finish_minutes - total_start_minutes) / 60
        #print(total_start_minutes, total_finish_minutes, total_hours)
        print("Enter comments for timesheet: ")
        timesheet_comments = input()
        eod.append(kms)
        eod.append(strt_time)
        eod.append(fnsh_time)
        eod.append(total_hours)
        eod.append(timesheet_comments)
        
    # This is the writing function where the three global variables start, jobs, eod are written to an xlsx sheet
    # that the user names
    def write_xlsx(start, jobs, eod):
        usefull_int = 4
        total_money = 0
        print("Enter the name of the timesheet")
        timesheet_name = str(input())
        # Workbook() takes one, non-optional, argument
        # which is the filename that we want to create.
        workbook = xlsxwriter.Workbook(timesheet_name + '.xlsx')
         
        # The workbook object is then used to add new
        # worksheet via the add_worksheet() method.
        worksheet = workbook.add_worksheet()
         
        # Use the worksheet object to write
        # data via the write() method.
        # ROW 1
        worksheet.write('A1', 'job completion sheet - domestic/commercial')
        worksheet.write('F1', 'technician: ' + start[0])
        worksheet.write('K1', 'Day: ' + start[1])
        worksheet.write('N1', 'Date: ' + start[2])
        # ROW 2
        worksheet.write('A2', 'Left home at:' + start[3])
        # ROW 3
        worksheet.write('A3', 'Client')
        worksheet.write('C3', 'Address')
        worksheet.write('F3', 'Suburb')
        worksheet.write('H3', 'Job type')
        worksheet.write('J3', 'ON')
        worksheet.write('K3', 'OFF')
        worksheet.write('L3', 'Money')
        worksheet.write('M3', 'Money type')
        worksheet.write('N3', 'Job comments')
        # ROW 4 to 14 job data
        for item in jobs:
            worksheet.write('A' + str(usefull_int), item[0])
            worksheet.write('C' + str(usefull_int), item[1])
            worksheet.write('F' + str(usefull_int), item[2])
            worksheet.write('H' + str(usefull_int), item[3])
            worksheet.write('J' + str(usefull_int), item[4])
            worksheet.write('K' + str(usefull_int), item[5])
            worksheet.write('L' + str(usefull_int), item[6])
            worksheet.write('M' + str(usefull_int), item[7])
            worksheet.write('N' + str(usefull_int), item[8])
            usefull_int = usefull_int + 1
            total_money = total_money + item[6]
        # ROW 15
        worksheet.write('A15', 'EoD kms: ' + str(eod[0]))
        worksheet.write('K15', 'Total: ' + total_money)
        # ROW 16
        worksheet.write('A16', 'Start time: ' + str(eod[1]))
        # ROW 17
        worksheet.write('A17', 'Finish time: ' + str(eod[2]))
        # ROW 18
        worksheet.write('A18', 'HOURS')
        worksheet.write('B18', 'hre: ' + str(eod[3]))
        worksheet.write('C18', '')
        worksheet.write('D18', '30min Break')
        worksheet.write('E18', 'Total: ' + str(eod[3] - 0.5))
        worksheet.write('F18', 'Timesheet comments' + eod[4]) 
        # Finally, close the Excel file
        # via the close() method.
        workbook.close()

    # pre menu questions or pre while loop program
    print("Please enter your name:")
    user_name = input()
    print("Day of the week: ")
    weekday = input()
    print("Enter your odometer: ")
    odometer_1 = int(input())
    print("enter 1 if you are leaving now\nEnter 2 to input different time")
    leave_time_choice = input()
    if leave_time_choice == "1":
        time_Now = dt.datetime.now()
        time_now = str(time_Now.hour) + ":" + str(time_Now.minute) + ":" + str(time_Now.second)
        print(time_now)
    else:
        print("Enter the time you left in the format 00:00:00")
        time_now = input()
    today = dt.date.today()
    start.append(user_name)
    start.append(weekday)
    start.append(str(today))
    start.append(time_now)

    # main menu
    while True:
        print("To add a job enter 1\nTo end the day enter 2\nTo save xlsx file enter 3\nTo exit the program enter 4\n")
        user_selection = input()
        if user_selection == "1":
            add_client(jobs)
        elif user_selection == "2":
            end_of_day(eod, odometer_1)
        elif user_selection == "3":
            write_xlsx(start, jobs, eod)
        elif user_selection == "4":
            exit()
        else:
            print("Invalid entry, please try again")

'''
Data to be collected
row 1 = [(a1)title], [(f1)tecnician], [(k1)day], [(n1)date]
row 2 = [(a2)first job start time], [(f2)leave time]
row 3 = [(a3)client], [(c3)address], [(f3)suburb], [(h3)job type], [(j3)on?], [(k3)off?], [(l3)money], [(m3)money type], [(n3)comments]
row 4 to 14 are row 3' data
row 15 = [(a15)eod kms], [(k15)total money]
row 16 = [(a16)start time]
row 17 = [(a17)finish time]
row 18 = [(a18)hours], [(b18)hrs], [(c18)minutes], [(d18)mins], [(e18)total hrs -30mins] [(f18)timesheet comments]
'''
