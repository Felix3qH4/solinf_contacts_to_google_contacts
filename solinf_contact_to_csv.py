import PySimpleGUI as sg
import pandas as pd
import subprocess
import os
from datetime import *

## DEPENDENCIES:
#   - pandas
#   - openpyxl
#   - PySimpleGUI


LABEL = "NEW CONTACTS" # Label displayed on google contacts
CHANGE_TO = "Wellefcher -> Scouten"
FONT_SIZE = 96
CHANGE_TO_LIST = ["Ignore", "Beaver -> Wellefcher", "Wellefcher -> Scouten", "Scouten -> Explorer", "Explorer -> Rover", "Rover -> Tembo"]

csv_output = {
                "Name": [],                 # Max Mustermann (= Displayed Name)
                "Given Name": [],           # Max 
                "Family Name": [],          # Mustermann
                "Group Membership": [],     # 1 Joer Scout
                "E-mail 1 - Value": [],     # Max.Mustermann@mymail.com
                "Phone 1 - Value": [],      # 621 123 456
                }


def add_child(row, output) -> None:
    """
    Add a child as new contact

    Parameters:
    -----------
        - row: the row containing the data for one child -> type list
        - output:  the dictionary which will be written to the csv file -> type dict
    """
    output["Name"].append(str(str(row[6]) + " " + str(row[3]) + " (Kand)"))
    output["Given Name"].append(row[6])
    output["Family Name"].append(row[3])
    output["Group Membership"].append(LABEL)
    output["E-mail 1 - Value"].append(row[14])

    ## Check if there are 2 phone numbers associated to one person, if yes, add both separated by ':::', if only one, add only one
    if not pd.isna(row[22]): # If phone number under "office phone"
        if not pd.isna(row[23]): # If phone number under "private phone" too (= 2 phone numbers)
            phone_number = str(row[22]) + " ::: " + str(row[23])
        else:
            phone_number = str(row[22])
    else: # If only one phone number under "private phone"
        if not pd.isna(row[23]):
            phone_number = str(row[23])
        else:
            phone_number = ""

    output["Phone 1 - Value"].append(phone_number)



def add_tutor_1(row, output) -> None:
    """
    Add a new contact as 'Tutor 1' of the child

    Parameters:
    -----------
        - row: the row containing the data for one child -> type list
        - output:  the dictionary which will be written to the csv file -> type dict
    """
    output["Name"].append(str(str(row[6]) + " " + str(row[3]) + " (Tutor 1)"))
    output["Given Name"].append(row[6])
    output["Family Name"].append(row[3])
    output["Group Membership"].append(LABEL)
    
    ## Check of email field is not empty
    if not pd.isna(row[46]):
        output["E-mail 1 - Value"].append(row[46])
    else:
        output["E-mail 1 - Value"].append("")

    ## Check if phone number field is not empty
    if not pd.isna(row[47]):
        output["Phone 1 - Value"].append(row[47])
    else:
        output["Phone 1 - Value"].append("")



def add_tutor_2(row, output) -> None:
    """
    Add a new contact as 'Tutor 2' of the child

    Parameters:
    -----------
        - row: the row containing the data for one child -> type list
        - output:  the dictionary which will be written to the csv file -> type dict
    """
    output["Name"].append(str(str(row[6]) + " " + str(row[3]) + " (Tutor 2)"))
    output["Given Name"].append(row[6])
    output["Family Name"].append(row[3])
    output["Group Membership"].append(LABEL)

    ## Check if email field is not empty
    if not pd.isna(row[57]):
        output["E-mail 1 - Value"].append(row[57])
    else:
        output["E-mail 1 - Value"].append("")

    ## Check if phone number field is not empty
    if not pd.isna(row[58]):
        output["Phone 1 - Value"].append(row[58])
    else:
        output["Phone 1 - Value"].append("")



## Calculate age the person has on the first of September this year
def calculate_age(birthdate):
    birthday = datetime.strptime(birthdate, "%Y-%m-%d %H:%M:%S")
    this_year = "01/09/" + str(datetime.now().year)
    first_september = datetime.strptime(this_year, "%d/%m/%Y")

    age = first_september.year - birthday.year - ((first_september.month, first_september.day) < (birthday.month, birthday.day))
    
    return age


## Parse file
def parse_xlsx_data(xlsx_file, csv_output_folder, change_to):
    """
    Parse the given xlsx file and extract the needed data to a csv file
    """

    solinf_data = pd.DataFrame(pd.read_excel(xlsx_file))

    for row in solinf_data.itertuples():
        ## If user wants to filter by age group
        if change_to != CHANGE_TO_LIST[0]:

            if not pd.isna(row[26]): # Check if filed is not empty/None
                age = calculate_age(str(row[26]))

                ## Beaver -> Wellefcher
                if change_to == CHANGE_TO_LIST[1]:
                    if age >= 8:
                        add_child(row, csv_output)
                        add_tutor_1(row, csv_output)
                        add_tutor_2(row, csv_output)
                ## Wellefcher -> Scouten
                elif change_to == CHANGE_TO_LIST[2]:
                    if age >= 11:
                        add_child(row, csv_output)
                        add_tutor_1(row, csv_output)
                        add_tutor_2(row, csv_output)
                ## Scouten -> Explorer
                elif change_to == CHANGE_TO_LIST[3]:
                    if age >= 14:
                        add_child(row, csv_output)
                        add_tutor_1(row, csv_output)
                        add_tutor_2(row, csv_output)
                # Explorer -> Rover
                elif change_to == CHANGE_TO_LIST[4]:
                    if age >= 16:
                        add_child(row, csv_output)
                        add_tutor_1(row, csv_output)
                        add_tutor_2(row, csv_output)
                # Rover -> Tembo
                elif change_to == CHANGE_TO_LIST[5]:
                    if age >= 21:
                        add_child(row, csv_output)
                        add_tutor_1(row, csv_output)
                        add_tutor_2(row, csv_output)
            else:
                raise Exception(f"Error while reading row, especially row[26]: {row}")

        ## If not filtered by age group
        else:
            add_child(row, csv_output)
            add_tutor_1(row, csv_output)
            add_tutor_2(row, csv_output)
        

    ## Save csv to file
    csv_output_file = csv_output_folder + "\OUTPUT.csv"
    dataFrame = pd.DataFrame(csv_output)
    dataFrame.to_csv(csv_output_file, index=False)

    ## Display success
    sg.PopupOK("The data has successfully been converted to a csv file!", title="SUCCESS", keep_on_top=True, auto_close=False, font=FONT_SIZE)
    print(csv_output_folder)
    subprocess.Popen(f'explorer {os.path.abspath(csv_output_folder)}')






## GUI
layout = [
    [sg.Text("Convert .xlsx files from Solinf to .csv files for google contacts", font=FONT_SIZE)],
    [sg.Push(), sg.Text(text="", font=FONT_SIZE), sg.Push()],
    [sg.FileBrowse("Choose .xlsx file", font=FONT_SIZE, file_types=[(".xlsx files ", "*.xlsx")], key="-XLSX_FILE-"), sg.Text("No input file chosen!", font=FONT_SIZE, key="-XLSX_FILE-")],
    [sg.FolderBrowse("Choose output folder", font=FONT_SIZE, key="-OUTPUT_FOLDER-"), sg.Text("No output folder chosen!", font=FONT_SIZE, key="-OUTPUT_FOLDER-")],
    [sg.Combo(CHANGE_TO_LIST, default_value=CHANGE_TO_LIST[0], font=FONT_SIZE, key="-CHANGE_TO_LIST_VALUE-"), sg.Text("Choose which age class you want (ex.: only children old enough for Scouts and above: Wellefcher -> Scouten) \n If 'Ignore' it will convert all contacts.", font=FONT_SIZE)],
    [sg.InputText(key="-LABEL_INPUT-", default_text=LABEL), sg.Text("Set the text for the label displayed in google contacts:", font=FONT_SIZE)],
    [sg.Push(), sg.Text(text="", font=FONT_SIZE), sg.Push()],
    [sg.Push(), sg.Button("Start Converting", font=FONT_SIZE), sg.Push()],
    []
    ]

    
window = sg.Window("Converter", layout)

while True:
    event, values = window.read()


    if event == "OK" or event == sg.WIN_CLOSED:
        break

    elif event == "Start Converting": # If user presses start button check if he provided an input file and an output folder
        xlsx_file = values["-XLSX_FILE-"]
        output_folder = values["-OUTPUT_FOLDER-"]
        change_to = values["-CHANGE_TO_LIST_VALUE-"]
        LABEL = values["-LABEL_INPUT-"]
        if xlsx_file:
            if output_folder:
                parse_xlsx_data(xlsx_file, output_folder, change_to)
            else:
                sg.PopupOK("You did not provide an output folder!", title="ERROR", keep_on_top=True, auto_close=False, font=FONT_SIZE)
        else:
            sg.PopupOK("You did not provide an input file!", title="ERROR", keep_on_top=True, auto_close=False, font=FONT_SIZE)
            


window.close()

