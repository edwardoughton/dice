"""
Build the DICE workbook.

"""
# import argparse
import os
import configparser
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.datavalidation import DataValidation
# from openpyxl.reader.excel import load_workbook
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter

CONFIG = configparser.ConfigParser()
CONFIG.read(os.path.join(os.path.dirname(__file__), 'script_config.ini'))
BASE_PATH = CONFIG['file_locations']['base_path']

DATA_RAW = os.path.join(BASE_PATH, 'raw')
DATA_INTERMEDIATE = os.path.join(BASE_PATH, 'intermediate')

def generate_workbook():
    """
    Generate the workbook and associated worksheets.

    """
    #import the Workbook class
    wb = Workbook()

    # #create the countries workbook sheet
    countries = wb.create_sheet("Countries", (2-1))
    # set_border(countries, 'A1:AZ1000', "thin", "00FFFFFF")
    countries = add_countries_sheet(countries)

    # #create the population workbook sheet
    capacity = wb.create_sheet("Capacity", (3-1))
    capacity = add_capacity_sheet(capacity)

    # #create the population workbook sheet
    population = wb.create_sheet("Population", (4-1))
    population = add_population_sheet(population)

    # #create the settings workbook sheet
    settings = wb.active
    set_border(settings, 'A1:AZ1000', "thin", "00FFFFFF")
    settings = add_settings_sheet(settings)
    # settings = adjust_column_widths(settings)
    settings.column_dimensions['A'].width = 15

    wb.save('Oughton et al. (2021) DICE (v0.1).xlsx')

    return print("Generated workbook")


def add_settings_sheet(ws):
    """
    Add the settings sheet.

    """
    ws.title = "Settings"
    ws.sheet_properties.tabColor = "FFFF00"

    c = ws['A1']
    c.font = openpyxl.styles.Font(size=14, color="000066CC")

    ws['A1'] = "Welcome to the Digital Infrastructure Cost Estimator (DICE) Model!"
    ws['A2'] = "------------------------------------------------------------------"
    ws['A3'] = "Quick Start instructions are provided here."
    ws['A4'] = "For detailed instructions, please see the model documentation: www.github.com/edwardoughton/dice "
    ws['A5'] = ""

    ws['A6'] = "Parameter"
    ws['B6'] = "Option"

    ws['A7'] = "Country"
    data_val = DataValidation(type="list", formula1='=Countries!A2:A251')
    ws.add_data_validation(data_val)
    data_val.add(ws["B7"])

    ws['A8'] = "Speed Target"
    # data_val = DataValidation(type="list", formula1='=Countries!A2:A251')
    # ws.add_data_validation(data_val)
    # data_val.add(ws["B7"])

    ws['A9'] = "Scenario"

    ws['A10'] = "Strategy"

    set_border(ws, 'A6:B10', "thin", "000000")

    return ws


def set_border(ws, cell_range, style, color):

    thin = Side(border_style=style, color=color)
    for row in ws[cell_range]:
        for cell in row:
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)


def add_countries_sheet(ws):
    """
    Add the countries sheet.

    """
    ws.sheet_properties.tabColor = "FFFF00"

    data = pd.read_csv(os.path.join(BASE_PATH, 'global_information.csv'), encoding = "ISO-8859-1")

    for r in dataframe_to_rows(data, index=False, header=True):
        ws.append(r)

    return ws


def add_capacity_sheet(ws):
    """

    """
    ws['A1'] = "Capacity Per User"
    ws['A2'] = 10
    ws['A3'] = 5
    ws['A4'] = 2


def add_population_sheet(ws):
    """
    Add the population sheet.

    """
    ws.sheet_properties.tabColor = "FFFF00"

    path = os.path.join(DATA_INTERMEDIATE, 'all_pop_data.csv')
    population = pd.read_csv(path)

    for r in dataframe_to_rows(population, index=False, header=True):
        ws.append(r)

    return ws


# for i in range(1,11+1):

#     cell = "A{}".format(i)

#     if i == 1:
#         ws[cell] = "=Title!$A$2:$A$10"
#     else:
#         ws[cell] = "=SUM(1, {})".format(i-2)


# wb = Workbook()

# ws = wb.create_sheet('New Sheet')

# for number in range(2,11): #Generates 99 "ip" address in the Column A;
#     ws['A{}'.format(number)].value= "{}".format(number)

# data_val = DataValidation(type="list",formula1='=$A:$A') #You can change =$A:$A with a smaller range like =A1:A9
# ws.add_data_validation(data_val)

# data_val.add(ws["A1"]) #If you go to the cell B1 you will find a drop down list with all the values from the column A


# wb.save('Oughton et al. (2021) DICE (v0.1).xlsx')

if __name__ == "__main__":

    generate_workbook()
