import pandas as pd
import win32com.client as win32
import os
import shutil

# Get the current working directory
cwd = os.getcwd()

# Define a function to get the path of the input CSV file
def get_path():

    #Make sure file is of type csv.
    val = cwd + "\\Book1.csv"

    return val

# Define a function to read the data from the input CSV file and save it as a dataframe
def get_data(path):

    df = pd.read_csv(path)

    return df

def change_data_PRLU(prlu):

    #Insert Column Names
    prlu.columns=['Date','Time','Ticket','AgentNo','RouteGroup',
    'RequestedGroup', 'AgentID', 'AnsweringAgent', 'TotalCallDuration',
    'CallRoutingTime', 'QueueTime', 'TelephoneRinging', 'ConvTime',
    'Event', 'Code', 'Calling', 'Called','Default']

    #Save rows with Ticket value A
    prlu = prlu[prlu.Ticket == 'A']

    #Save rows with group value 3 and 4
    prlu = prlu[prlu['RequestedGroup'].isin(['3','4'])]

    #Saving count of the different call types
    #A = Abondoned, D = Dettered, S = Answered, T = Voicemail, length of dataset = Recieved
    #Others can be different letters, Others = Recieved - (A + D + S + T)   
     
    recieved = len(prlu)
    abondoned = (prlu.Event == 'A').sum()
    dettered = (prlu.Event == 'D').sum()
    answered = (prlu.Event == 'S').sum()
    voicemail = (prlu.Event == 'T').sum()
    others = recieved - (abondoned + dettered + answered + voicemail)

    #Save count types in list 
    list_values = ['PRLU',recieved,abondoned,dettered,answered,voicemail,others] 

    #Change settings for excel sheet
    prlu = prlu.style.set_properties(**
    {
        'font-family' : 'Calibri', 'text-align' : 'right', 'font-size' : '16px'
    })

    #Save excel sheet with specific sheet name and file name. Remove index column, not useful
    prlu.to_excel('PRLU.xlsx', sheet_name='PRLU', index=False)

      
    return list_values 

def change_data_CU(customer):

    #Insert Column Names
    customer.columns=['Date','Time','Ticket','AgentNo','RouteGroup',
    'RequestedGroup', 'AgentID', 'AnsweringAgent', 'TotalCallDuration',
    'CallRoutingTime', 'QueueTime', 'TelephoneRinging', 'ConvTime',
    'Event', 'Code', 'Calling', 'Called','Default']

    #Save rows with Ticket value A
    customer = customer[customer.Ticket == 'A']

    #Save rows with group value 3 and 4
    customer = customer[customer['RequestedGroup'].isin(['1'])]

    #Saving count of the different call types
    #A = Abondoned, D = Dettered, S = Answered, T = Voicemail, length of dataset = Recieved
    #Others can be different letters, Others = Recieved - (A + D + S + T)   
    recieved = len(customer)
    abondoned = (customer.Event == 'A').sum()
    dettered = (customer.Event == 'D').sum()
    answered = (customer.Event == 'S').sum()
    voicemail = (customer.Event == 'T').sum()
    others = recieved - (abondoned + dettered + answered + voicemail)

    #Save count types in list 
    list_values = ['Customer Care',recieved,abondoned,dettered,answered,voicemail,others]   

    #Change settings for excel sheet
    customer = customer.style.set_properties(**
    {
        'font-family' : 'Calibri', 'text-align' : 'right', 'font-size' : '16px'
    })

    #Save excel sheet with specific sheet name and file name. Remove index column, not useful
    customer.to_excel('Customer Care.xlsx', sheet_name='Customer Care', index=False)  

    return list_values 

def main():

    #Call functions 
    path = get_path()
    df = get_data(path)
    listP = change_data_PRLU(df)
    listC = change_data_CU(df)

    #Insert Columns for the count excel
    count = pd.DataFrame (columns=['type','recieved','abondoned','dettered','answered','voicemail','others'])

    #Add call counter of Customer Care and PRLU 
    count.loc[len(count.index)] = listC
    count.loc[len(count.index)] = listP

    #Change settings for excel sheet
    count = count.style.set_properties(**
    {
        'font-family' : 'Calibri', 'text-align' : 'center', 'font-size' : '18.667px'
    })

    #Save excel sheet with specific sheet name and file name. Remove index column, not useful
    count.to_excel('Counter.xlsx', sheet_name='Counter', index=False)


main()

def customer_care():

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    #Call workbook
    str = cwd + "\\Customer Care.xlsx"
    wb = excel.Workbooks.Open(str)
    #Call sheet
    ws = wb.Worksheets("Customer Care")
    #Change font and size or first row
    ws.Rows(1).RowHeight = 40
    ws.Rows(1).Font.Size = 16 
    #Apply colour to names
    for i in range(1,20):

        #Colour 33 = light blue
        #Colour scheme can be found here https://docs.microsoft.com/en-us/office/vba/api/excel.colorindex
        ws.Cells(1,i).Interior.ColorIndex = 33

    #Auto size the columns to fit   
    ws.Columns.AutoFit()
    wb.Close(SaveChanges=True)
    excel.Quit()

def prlu():

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    #Call workbook
    str = cwd + "\\PRLU.xlsx"
    wb = excel.Workbooks.Open(str)
    #Call sheet
    ws = wb.Worksheets("PRLU")
    #Change font and size or first row
    ws.Rows(1).RowHeight = 40
    ws.Rows(1).Font.Size = 16 
    #Apply colour to names
    for i in range(1,20):

        #Colour 33 = light blue
        #Colour scheme can be found here https://docs.microsoft.com/en-us/office/vba/api/excel.colorindex
        ws.Cells(1,i).Interior.ColorIndex = 33

    #Auto size the columns to fit   
    ws.Columns.AutoFit()
    wb.Close(SaveChanges=True)
    excel.Quit()

def counter():

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    #Call workbook
    str = cwd + "\\Counter.xlsx"
    wb = excel.Workbooks.Open(str)
    #Call sheet
    ws = wb.Worksheets("Counter")
    #Auto size the columns to fit   
    ws.Columns.AutoFit()
    wb.Close(SaveChanges=True)
    excel.Quit()

def main2():

    #Call main functions
    customer_care()
    prlu()
    counter()

main2()