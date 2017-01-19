import csv, os, easygui, sys, datetime, decimal

#----------------------------------------
#file setup functions 

#checks that the file time picked for a save to file is correct (run by save_file)
def type_check(valid_name, file_type, title):
    output_name = easygui.filesavebox("Where would you like to save the report? Please save as a " + file_type + " file", title, file_type)
    try:
        period_index = output_name.index(".")
    except ValueError:
        return False
    except AttributeError:
        sys.exit()
    output_type = str(output_name[period_index:])
    if output_type == file_type:
        valid_name = True
        return output_name
    else:
        return False

#creates a save to file and checks that it is the correct type; file_type is ".csv"; title is window title
def save_file(file_type, title):
    valid_name = False
    output_name = type_check(valid_name, file_type, title)
    while output_name is False:
        easygui.msgbox("Please save the report as a " + file_type + " file.", title)
        output_name = type_check(valid_name, file_type, title)
    return output_name

#--------------------------------------------------------------------------
#funciton definitions

#adds a message to a list to be printed in a text document, will be called if a record contains an error
def add_error_list(name, message):
    error_list.append(name + ": " + message)

#detects whether a character pulled from a string is an integer
def is_int(character):
    try:
        int(character)
        return True
    except ValueError:
        return False

#takes an input row, and creates separate rows for each individual parcel number
def parseParcel(row):
    val = row[0].strip()
    val = sciNotation(val)
    if is_int(val) is True:
        final_list = []
        final_list.append(row)
        return final_list
    else:
        list0 = row[0].split(",")
        list1 = splitFurther(list0, ";")
        list2 = splitFurther(list1, ":")
        list3 = splitFurther(list2, "&")
        list4 = splitFurther(list3, " ")
         
        n = len(list4) - 1
        list5 = []

        while n >= 0:
            temp = list4[n]
            list4.remove(list4[n])
            entry = intsOnly(temp)
            if len(entry) >=1 :
                list5.append(entry)
            else:
                pass
            n = n - 1

        list6 = []
        list6 = testLength(list5)
       
        final_list = []
            
        if len(list6) > 0:
            for x in list6:
                newList = []
                newList.append(x)
                newList.append(row[1].strip())
                newList.append(row[2].strip())
                final_list.append(newList)
        else:
            newList = []
            newList.append("")
            newList.append(row[1])
            newList.append(row[2])
            final_list.append(newList)

        return final_list
        
#strips a string, then splits into a list by a deliminter
def splitFurther(parcels, delim):
    master = []
    for x in parcels:
        x = x.strip()
        y = x.split(delim)
        for z in y:
            master.append(z)
    return master

#removes all characters from a string that aren't integers
def intsOnly(parcel):
    for char in parcel:
        if is_int(char) is True:
            pass
        else:
            parcel = parcel.replace(char, "")
    return parcel 

#removes entries in a list that are not 12 characters long
def testLength(parcel_list):
    for x in parcel_list:
        if len(x) != 12 or x == "":
            parcel_list.remove(x)
        else:
            pass
    return parcel_list

#converts from excel's scientific notation
def sciNotation(val):
    if "E+" in val:
        print val
        e_index = val.find("E+")
        plus_index = e_index + 2
        exponent = val[plus_index:]
        print exponent
        base = val[:e_index]
        print base
        num = int((base)*(10**int(exponent)))
        print num
        print type(num)
        return num
    else:
        return val
        
#--------------------------------------------------------------------------
#status message
print "This window will remain open while the Scrubber is running. \n\nClosing it will close the program."

#file setup

#define title of easygui boxes
box_title = "Scrubber"

#display intro message, and allow the user to continue or cancel
if easygui.buttonbox("Welcome to the Scrubber", box_title, ("Scrub a Spreadsheet", "Cancel")) is "Scrub a Spreadsheet":
    pass
else:
    sys.exit()

enter_message = "Pick a .csv file to scrub"

#define the path to the file that the user wants to format 
input_file = easygui.fileopenbox(enter_message, box_title, "C:\\")
if os.path.isfile(input_file) is True:
    pass
else:
    sys.exit()

#check if the file entered is a csv file
while input_file[-3:] != "csv":
    if os.path.isfile(input_file) is True:
        if easygui.ccbox("Error: program can only format .csv files", box_title) is True:
            input_file = easygui.fileopenbox(enter_message, box_title, "C:\\")
            if os.path.isfile(input_file) is True:
                pass
            else:
                sys.exit()
        else:
            sys.exit()
    else:
        if easygui.ccbox("Error: please select a .csv file", title) is True:
            input_file = easygui.fileopenbox(enter_message, box_title, "C:\\")
            if os.path.isfile(input_file) is True:
                pass
            else:
                sys.exit()
        else:
            sys.exit()

valid_file = False

while valid_file is False:
    
    #create a save to file
    output_name = save_file(".csv", box_title)

    #create csv writer using output_file as the filename if the file is not already open
    try:
        output_writer = csv.writer(open(output_name, "wb"))
        valid_file = True
    except IOError:
        if easygui.ccbox("Error: " + output_name + " is open. \n\nPlease close before saving to this location.\n\nNOTE: THE FILE WILL BE OVERWRITTEN IF YOU CONTINUE.", box_title) is True:
            pass
        else:
            sys.exit()
    

#create csv reader from input_file
input_reader = csv.reader(open(input_file, "rb"))


#create list in which to store errors for error report, and an integer to count the number of errors
data_list = []
output_list = []
error_list = []
error_count = 0

#---------------------------------------------------------------------------
#formatting

#input rows of data from original spreadsheet
for row in input_reader:
    data_list.append([row[0], row[1], row[2]])

#remove , characterse from property name field
for row in data_list:
    row[1] = row[1].replace(",", "")

#run parseParcel for each row, and append results to output_list
for row in data_list:
    tempList = []
    tempList = parseParcel(row)

    for x in tempList:
        output_list.append(x)

#---------------------------------------------------------------------------
#file output

#ouput comma delimited string to text file
name_index = output_name.find(".csv")

#writes each row of output_list to output_writer
for row in output_list:
    output_writer.writerow(row)

easygui.msgbox("Scrubber complete \n\nResults saved to: \n" + output_name, box_title)
