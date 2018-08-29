from tkinter import filedialog
import pandas as pd


def convert_file(filename):
    data_scores = []  # reads the file specified and returns the username and attended value, unfortunately
    file_data = pd.read_excel(filename, sheet_name=1)  # the email address column isn't populated
    usernames = file_data['username']
    usernames_fixed = []
    for x in usernames:
        if x[-4] != ".":
            usernames_fixed.append(x + "@ncsu.edu")
        else:
            usernames_fixed.append(x)
    data_scores.append(usernames_fixed)
    attend_list = []
    for i in range(0, len(usernames_fixed)):
        if file_data['status'][i] == "Attended":
            attend_list.append(1)
        else:
            attend_list.append(0)
    data_scores.append(attend_list)
    return data_scores


def write_wa_file(data, attend_date):
    wa_file_path = filedialog.asksaveasfile(title="Select a Save Location", filetypes=(("csv", "*.csv"),
                                                                                       ("all files", "*.*")),
                                            defaultextension='.csv')
    wa_file = open(wa_file_path.name, 'w+')
    wa_file.write("Assignment , " + attend_date + "\n")
    wa_file.write("Category, Attendance" + "\n")
    wa_file.write("Description, " + attend_date + "\n")
    wa_file.write("Points , 1 " + "\n")
    wa_file.write("Due" + "\n")
    wa_file.write("Available , Yes , Y" + "\n")
    print(len(data[0]))
    for i in range(0, len(data[0])):
        wa_file.write(str(data[0][i]) + "," + str(data[1][i])+"\n")
    wa_file.close()


def find_date(file):
    month = file[-15:-13]
    print(month)
    day = file[-13:-11]
    year = file[-19:-15]
    attend_date = month + " " + str(day) + " " + str(year)
    print(attend_date)
    return attend_date


def main_loop():
    file_names = filedialog.askopenfilenames(title="Select a File for Conversion", filetypes=(
        ("xlsx files", "*.xlsx"), ("all files", "*.*")))
    for file in file_names:
        data_to_write = convert_file(file)
        attend_date = find_date(file)
        write_wa_file(data_to_write, attend_date)


main_loop()