#! /usr/bin/env python
# A simply script to take excel files from my employer and calculate the P11D
# tax liability.
# Ciarán Mooney
# 2019

import datetime
import os
from xlrd import open_workbook

EXCEL_DIRECTORY = 'C:\something'
FIRST_DAY_TAX = (2017,4,6)
LAST_DAY_TAX = (2018,4,5)

def excel_date_conv(excel_date):
    ''' Date from Excel is an integer for the number of days since
        1st January 1900.
    '''
    epoch = datetime.date(1900,1,1)
    return epoch + datetime.timedelta(days=excel_date - 2) # -2 due to bug in Excel


def excel_date_parse(excel_file):
    ''' Parses an excel Expenses file from Domino to produce the list of days
        that an engineer has travelled.
    '''
    dates = []
    with open_workbook(excel_file, 'rb') as wb:
        backSheet = wb.sheet_by_index(1)

    for cell in backSheet.col_slice(0,3,35):
        if cell.value != '':
            date = excel_date_conv(cell.value)
            dates.append(date)

    return dates

    
if __name__ == '__main__':
    today = datetime.date.today()
    travel_days = []
    EXCEL_DIRECTORY = r'c:\Users\ciaran.mooney@domino-uk.com\Documents\Domino\Service\Expenses\2017-2018'
    for expenses in os.listdir(EXCEL_DIRECTORY):
        if expenses.split('.')[1] == 'xls':
            path = EXCEL_DIRECTORY + '\\' + expenses
            for day in excel_date_parse(path):
                if today < datetime.date(*LAST_DAY_TAX):
                    pass # calculate just up til today

                elif datetime.date(*FIRST_DAY_TAX) < day < datetime.date(*LAST_DAY_TAX):
                    travel_days.append(day)

            

    no_of_weekdays = 52*5
    non_travel_days = no_of_weekdays - len(travel_days)
    for each in travel_days:
        print(each.isoformat())
    print("Total Working Days: ", no_of_weekdays)
    print("200 Days minimum travel for frequent travel allowance.")
    print("Total Days Travelled: ", len(travel_days))
    print("Total Days Taxed: ", 200-len(travel_days))
    print("P11D Value: ", 5*(200-len(travel_days)))
