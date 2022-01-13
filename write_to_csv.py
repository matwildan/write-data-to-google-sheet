import sys
import pycountry
import time
import gspread
import pandas as pd
import csv
#from oauth2client.service_account import ServiceAccountCredentials
from google.oauth2.service_account import Credentials

def write_data(csv_filename, sheetnum):
    id = df_col['TRUE']
    print(id)
    email = df_col['Email']
    print(email)
    phone = df_col['Billing Phone']
    print(phone)
    name = df_col['Billing Name']
    print(name)
    zip = df_col['Billing Zip']
    print(zip)
    city = df_col['Billing City']
    print(city)
    country = df_col['Billing Country']
    print(country)

    email_column = 1
    phone_column = 2
    first_name_column = 3
    last_name_column = 4
    zip_column = 5
    city_column = 6
    country_column = 7

    firstRow = len(sheet1.col_values(first_name_column))
    print(firstRow)
    #projCol = 2
    projCol = firstRow

    count = 0
    row_before = ''
    row_num = 0
    col_num = 0
    data = ""

    for i in range(len(id[1:])):
        count = count + 1
        projCol = projCol + 1
        #data = "data: "+email[i]+", "+phone[i]+", "+name[i]+", "+zip[i]+", "+city[i]+", "+country[i]
        if name[i] == "":
            projCol = projCol - 1
            continue
        email_fix = email[i]
        phone_fix = edit_phone(phone[i])
        first_name_fix = edit_first_name(name[i])
        last_name_fix = edit_last_name(name[i])
        zip_code_fix = edit_zip(zip[i])
        city_fix = city[i]
        country_fix = edit_country(country[i])

        print(data)
        print(email_fix)
        print(phone_fix)
        print(first_name_fix)
        print(last_name_fix)
        print(zip_code_fix)
        print(city_fix)
        print(country_fix)

        cells = "A%s:G%s"%(projCol,projCol)
        print(cells)
        try:
            #sheet1.update(cells,[[email_fix, phone_fix, first_name_fix, last_name_fix, zip_code_fix, city_fix, country_fix]])
            print("[INFO] Datas have been writed!")
        except Exception as e:
            print("[ERROR] Error: ", e)
        email_fix = ""
        phone_fix = ""
        first_name_fix = ""
        last_name_fix = ""
        zip_code_fix = ""
        city_fix = ""
        country_fix = ""
        time.sleep(1)

def edit_country(data_country):
    data_country = pycountry.countries.get(alpha_2=data_country)
    data_country = data_country.name
    #print(data_country)
    return data_country

def edit_zip(data_zip):
    chars = "+'- "
    for char in chars:
        data_zip = data_zip.replace(char,"")
    #print(data_zip)
    return data_zip

def edit_first_name(data_name):
    first, *last = data_name.split()
    first = first
    last = " ".join(last)
    #print(first + "," + last)
    return first

def edit_last_name(data_name):
    first, *last = data_name.split()
    first = first
    last = " ".join(last)
    #print(first + "," + last)
    return last

def edit_phone(phone_num):
    chars = "+'- "
    for char in chars:
        phone_num = phone_num.replace(char,"")
    #print(phone_num)
    index = phone_num.find("62")
    find_zero = phone_num.find("0")
    find_eight = phone_num.find("8")
    print(index)
    print(find_zero)
    if index == -1 and find_zero < 1:
        phone_num = phone_num.replace("0","62",1)
        print(phone_num)
    if find_eight == 0:
        phone_num = "62"+phone_num
    return phone_num

def write_email():
    email_column = 1
    email = df_col['Email']
    #print(email)

    firstRow = len(sheet1.col_values(email_column))
    print(firstRow)
    #projCol = 2
    projCol = firstRow
    count = 0
    row_before = ''

    for row in email:
        projCol = projCol + 1
        #if row == '' and name.empty == False:
        #    projCol = projCol - 1
        #    continue
        if row == row_before and row_before != '':
            projCol = projCol - 1
            continue
        row_before = row
        count = count + 1
        print(projCol, row)
        sheet1.update_cell(projCol,email_column,row)
        print("[INFO] Datas have been writed!")
        time.sleep(1)

def write_phone():
    phone_column = 2
    phone = df_col['Billing Phone']
    print(phone)

    firstRow = len(sheet1.col_values(phone_column))
    print(firstRow)
    #projCol = 2
    projCol = firstRow
    count = 0
    row_before = ''

    for row in phone:
        projCol = projCol + 1
        if row == '':
            #projCol = projCol - 1
            continue
        #if count > 0:
        if row == row_before:
            projCol = projCol - 1
            continue
        chars = "+'- "
        for char in chars:
            row = row.replace(char,"")
        print(row)
        index = row.find("62")
        find_zero = row.find("0")
        find_eight = row.find("8")
        if index == -1 and find_zero < 1:
            row = row.replace("0","62",1)
            print(row)
        if find_eight == 0:
            row = "62"+row
        row_before = row
        #count = count + 1
        print(projCol, row)
        #sheet1.update_cell(projCol,phone_column,row)
        print("[INFO] Datas have been writed!")
        time.sleep(1)

def write_name():
    first_name_column = 3
    last_name_column = 4
    name = df_col['Billing Name']
    print(name)

    firstRow = len(sheet1.col_values(first_name_column))
    secondRow = len(sheet1.col_values(last_name_column))
    print(firstRow)
    print(secondRow)
    #projCol = 2
    projCol = firstRow
    count = 0
    row_before = ''

    for row in name:
        projCol = projCol + 1
        if row == '':
            projCol = projCol - 1
            continue
        #if count > 0:
        if row == row_before:
            projCol = projCol - 1
            continue
        first, *last = row.split()
        first = first
        last = " ".join(last)
        print(first + "," + last)
        row_before = row
        #count = count + 1
        print(projCol, row)
        sheet1.update_cell(projCol,first_name_column,first)
        time.sleep(1)
        #sheet1.update_cell(projCol,last_name_column,last)
        print("[INFO] Datas have been writed!")
        time.sleep(1)

def write_zip():
    zip_column = 5
    zip = df_col['Billing Zip']
    print(zip)

    firstRow = len(sheet1.col_values(zip_column))
    print(firstRow)
    #projCol = 2
    projCol = firstRow
    count = 0
    row_before = ''

    for row in zip:
        projCol = projCol + 1
        if row == '':
            projCol = projCol - 1
            continue
        #if count > 0:
        if row == row_before:
            projCol = projCol - 1
            continue
        chars = "+'- "
        for char in chars:
            row = row.replace(char,"")
        print(row)
        print(projCol, row)
        #sheet1.update_cell(projCol,zip_column,row)
        print("[INFO] Datas have been writed!")
        time.sleep(1)

def write_city():
    city_column = 6
    city = df_col['Billing City']
    print(city)

    firstRow = len(sheet1.col_values(city_column))
    print(firstRow)
    #projCol = 2
    projCol = firstRow
    count = 0
    row_before = ''

    for row in city:
        projCol = projCol + 1
        if row == '':
            projCol = projCol - 1
            continue
        #if count > 0:
        if row == row_before:
            projCol = projCol - 1
            continue
        print(row)
        print(projCol, row)
        #sheet1.update_cell(projCol,city_column,row)
        print("[INFO] Datas have been writed!")
        time.sleep(1)

def write_country():
    country_column = 7
    country = df_col['Billing Country']
    #country = df_col['Country']
    print(country)

    firstRow = len(sheet1.col_values(country_column))
    print(firstRow)
    #projCol = 2
    projCol = firstRow
    count = 0
    row_before = ''

    for row in country:
        projCol = projCol + 1
        if row == '':
            projCol = projCol - 1
            continue
        #if count > 0:
        if row == row_before:
            projCol = projCol - 1
            continue
        print(row)
        row_country = pycountry.countries.get(alpha_2=row)
        row_country = row_country.name
        print(projCol, row_country)
        #sheet1.update_cell(projCol,country_column,row_country)
        print("[INFO] Datas have been writed!")
        time.sleep(1)

def main(column):
    if column == 'all':
        write_data(csv_filename, sheetnum)
    elif column == 'email':
        print("[INFO] Write email from CSV")
        write_email()
    elif column == 'phone':
        print("[INFO] Write phone from CSV")
        write_phone()
    elif column == 'name':
        print("[INFO] Write name from CSV")
        write_name()
    elif column == 'zip':
        print("[INFO] Write zip from CSV")
        write_zip()
    elif column == 'city':
        print("[INFO] Write city from CSV")
        write_city()
    elif column == 'country':
        print("[INFO] Write country from CSV")
        write_country()
    else:
        print("[ERROR] No column suitable!")
        print("Usage: python3 <file name> <sheet number> <column which you want to write>")

if __name__ == "__main__":
    try:
        csv_filename = sys.argv[1]
        print("[INFO] File name: ",csv_filename)
        sheetnum = int(sys.argv[2])
        print("[INFO] Sheet number: ",sheetnum)
        column = sys.argv[3]
        print("[INFO] Column which you want to write: ",column)

        #define the scope
        scope = ['https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive']

        #add creds to the acc
        creds = Credentials.from_service_account_file("../client_secret.json", scopes=scope)

        # authorize the clientsheet 
        client = gspread.authorize(creds)

        # get the instance of the Spreadsheet
        sheet = client.open('Data december')

        # get the sheet num of the Spreadsheet
        sheet1 = sheet.get_worksheet(sheetnum)

        df = pd.DataFrame(data=sheet1.get_all_records())
        #print(df)

        col_list = ["TRUE", "Email", "Billing Phone", "Billing Name", "Billing Zip", "Billing Zip", "Billing City", "Billing Country"]
        df_col = pd.read_csv(csv_filename, usecols=col_list, na_filter=False)

        #if len(sys.argv) != 2 or len(sys.argv) > 2 or len(sys.argv) == 0:
        #    print("Usage: python3 <file name> <sheet number>")
        #else:
        main(column)

    except Exception as e:
        print("Usage: python3 <file name> <sheet number> <column which you want to write>")
        print(e)
