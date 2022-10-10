import os
import PySimpleGUI as sg
import pandas as pd

def switching_headers(df): #Using the first row as headers and deleting it
    df.columns = df.iloc[0]
    df.drop(index=0, inplace=True)
    return df

def convert_to_date(text): # converting date column with text into datetime type
    text = str(text)
    year = text[-4:]
    month = text[0]
    day = "1"
    date_text = day + "/" + month + "/" + year
    date = pd.to_datetime(arg=date_text, format= "%d/%m/%Y" )
    return date


data = pd.read_excel("data/sample.xlsx") # opening file with data export



duplikat = data.iloc[:,4].duplicated(keep= False) # checking for duplicates
dt = data[duplikat] # checking, if there are any duplicated project definitions

# the data will be split for one df which will contain info about projects and anothe which will conatin revenue and operating margin
project_data = data.iloc[:,:15] # selecting columns with project data
column_to_move = data.pop("Unnamed: 4")

data.insert(14, "ID ", column_to_move) # moving the column with project definition for it to be next to the numbers

project_numbers = data.iloc[:, 14:] # selecting columns for the financial data


numbers_columns = project_numbers.columns # getting a list of columns to be able to split them into Revenue and OM
# this is necessary, as if we get an export with same data, but with additional months, the new months for revenue will be shown after the old ones, but before the months with OM
#  therefore we have to use the loop below to filter them


columns_list_rev = ['ID ']
columns_list_om = ['ID ']
for name in numbers_columns:
    if name[0] == "R":
        columns_list_rev.append(name)
    elif name[0] == "O":
        columns_list_om.append(name)


project_numbers_rev = project_numbers[columns_list_rev].copy()# Spliting data for Revenue and OM
project_numbers_om = project_numbers[columns_list_om].copy()



switching_headers(project_numbers_rev) # we use the first row as a header and delete it afterwards
switching_headers(project_numbers_om)

project_numbers_rev = project_numbers_rev.melt(id_vars="Project Definition.", value_name= "Revenue", var_name= "Date") # unpivoting the financial data
project_numbers_om = project_numbers_om.melt(id_vars="Project Definition.", value_name= "OM", var_name= "Date")



merged = project_numbers_rev.merge(project_numbers_om, how = "outer", on = ["Project Definition.", "Date"]).copy() # merging the Revenue and OM data

merged["Date"] = merged["Date"].apply(convert_to_date) # converting text in the Date column to datetime type

merged["OM"].fillna(value=0, inplace= True) # filling out the NaN values with 0
merged["Revenue"].fillna(value=0, inplace= True)

# here we change the names of the columns, as some of them were incorrect or missing
project_data.columns = ['Segment', 'Company_Code', "Company_name", 'WBS_Activity',
       "Project_Definition", "Project_Text", 'Industry', 'Subindustry',
       'Contract_Admin_1', 'Project_Manager', 'Final_Customer', "Final_Customer_ID",
       'IR_Code', 'Corp_Customer', "IR_ID"]

project_data.drop(index=0, inplace= True) # deleting the first row, as it contained column names in the export

excel_file = pd.ExcelWriter("Data cleaning1.xlsx") # putting data into an excel file
project_data.to_excel(excel_file, sheet_name= "Data", index= False)
merged.to_excel(excel_file, sheet_name= "Rev and Om", index= False)


excel_file.save()


print(merged.info())