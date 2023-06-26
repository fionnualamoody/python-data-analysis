import csv
import xlsxwriter

# Opens the given file, loads the data, and fills a list from the given field name
def openFile(fileName, fieldName):
    # Opening given file in read mode
    with open(fileName,'r')as csv_file:
        # reading in data from the csv file
        spreadsheet=csv.DictReader(csv_file)
        # creating an empty list to store data in
        data=[]
        # looping over all the rows in the spreadsheet
        for row in spreadsheet:
            # getting the coinciding cell for the given fieldName
            cell_value=row[fieldName]
            # adding in the cell value to our data list
            data.append(cell_value)
    # return the filled data list
    return data

# Creates a work sheet, takes in a list of titles and values and adds them to work sheet
def write_to_workbook(workbook, titles, values):
    # Creating a new worksheet
    worksheet=workbook.add_worksheet()
    # Defining Column counter
    column=0

    # Loop through the titles we have
    for number in range (len(titles)):
        # In each row...
        # we write the title in the first cell
        worksheet.write(number,column,titles[number]);
        # We write the value in the second cell
        worksheet.write(number,column+1,values[number]);

    return workbook

# Calculates the % change in sales from the previous month
def calculate_monthly_changes(sales):
    values=[]
    # loop over our sales
    for number in range(len(sales)):
        # need to start after january, as we cant compare it to anythinng
        if(number >0):
            # calculation for % change
            values.append((sales[number-1] - sales[number]) / sales[number] * 100)
    # return % changes
    return values

# Adds a worksheet to our work book that dislays the monthly changes
def add_monthly_changes_worksheet(workbook,months,values):
    # Adding our worksheet
    worksheet = workbook.add_worksheet()

    # looping over our months (January should aready be removed)
    for number in range(len(months)):
        # In each row...
        # we write the title in the first cell
        worksheet.write(number, 0, months[number]);
        # We write the value in the second cell
        worksheet.write(number, 1, values[number]);

def run():
    # getting our sales values from our csv
    sales = []

    # looping through out returned sales, making sure we turn the strings into numbers
    for sale in openFile('sales.csv', 'sales'):
        sales.append(int(sale))

    # getting our month names from our csv
    months = openFile('sales.csv', 'month')

    # deleting january from our list as we cant use it in our monthly changes,theres nothing to compare it to
    del months[0]

    # getting total sales
    total_sales = sum(sales)

    # Creating our titles and values list for outputitng to excel
    titles = ['Total sales', 'Highest sales', 'Lowest sales', 'Average sales']
    values = [total_sales, max(sales), min(sales), (total_sales / len(sales))]

    # creating a workbook to write to
    workbook = xlsxwriter.Workbook('summary.xlsx')

    # Writing out data to our excel file
    sales_workbook = write_to_workbook(workbook, titles, values)

    # Calculating the values for monthly changes
    monthly_changes=calculate_monthly_changes(sales)

    # Adding our monthly changes work sheet to our work book
    add_monthly_changes_worksheet(workbook,months,monthly_changes)

    # Closing and Saving our workbook
    workbook.close()

run()


