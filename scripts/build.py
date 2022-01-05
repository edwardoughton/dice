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

    settings = wb.active
    settings = add_settings(settings)

    options = wb.create_sheet("Options", (2-1))
    options = add_options(options)

    population = wb.create_sheet("Population", (3-1))
    population = add_population_sheet(population)

    demand = wb.create_sheet("Demand", (4-1))
    demand = add_demand_sheet(demand)

    supply = wb.create_sheet("Supply", (5-1))
    supply = add_supply_sheet(supply)

    # costs = wb.create_sheet("Costs", (6-1))

    # # #edit population workbook sheet




    wb.save('Oughton et al. (2021) DICE (v0.1).xlsx')

    return print("Generated workbook")


def add_settings(ws):
    """

    """
    ##Color white
    set_border(ws, 'A1:AZ1000', "thin", "00FFFFFF")

    ##Set column width
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 15

    ##Allocate title
    ws.title = "Settings"
    ws.sheet_properties.tabColor = "009900"

    ##Introductory note
    c = ws['A1'] #Set font for A1
    c.font = openpyxl.styles.Font(size=14, color="000066CC")
    ws['A1'] = "Welcome to the Digital Infrastructure Cost Estimator (DICE) Model!"
    ws['A2'] = "------------------------------------------------------------------"
    ws['A3'] = "Quick Start instructions are provided here."
    ws['A4'] = "For detailed instructions, please see the model documentation: www.github.com/edwardoughton/dice "
    ws['A5'] = ""

    ## Add parameters box
    set_border(ws, 'A6:B16', "thin", "000000")
    ws['A6'] = "Parameter"
    ws['B6'] = "Option"

    ws['A7'] = "Country"
    data_val = DataValidation(type="list", formula1='=Options!A2:A251')
    ws.add_data_validation(data_val)
    data_val.add(ws["B7"])
    ws['B7'] = "Afghanistan"

    ws['A8'] = "ISO3"
    ws['B8'] = "=IFERROR(INDEX(Options!B2:B1611,MATCH(Settings!B7, Options!A2:A1611)), \"\")"

    ws['A9'] = "Speed Target (Mbps)"
    ws['B9'] = 10

    ws['A10'] = "Adoption Scenario"
    data_val = DataValidation(type="list", formula1='=Options!C2:C4')
    ws.add_data_validation(data_val)
    data_val.add(ws["B10"])

    ws['A11'] = "Market Share (%)"
    ws['B11'] = 25

    ws['A12'] = "Active Users (%)"
    ws['B12'] = 5

    ws['A13'] = "Infrastructure Strategy"
    data_val = DataValidation(type="list", formula1='=Options!D2:D6')
    ws.add_data_validation(data_val)
    data_val.add(ws["B13"])

    ws['A14'] = "Spectrum Availability"
    data_val = DataValidation(type="list", formula1='=Options!E2:E4')
    ws.add_data_validation(data_val)
    data_val.add(ws["B14"])

    ws['A15'] = "Sites Availability"
    data_val = DataValidation(type="list", formula1='=Options!F2:F4')
    ws.add_data_validation(data_val)
    data_val.add(ws["B15"])

    ########Deciles
    ws['D6'] = "Decile"
    for row in range(1, 11):
        cell = "D{}".format(6+row)
        ws[cell] = row * 10

    ws['E6'] = "Population"
    ws.column_dimensions['E'].width = 20
    for row in range(1, 11):
        cell = "E{}".format(6+row)
        part1 = "=INDEX(Population!$D$2:$D$1611,MATCH(1,INDEX(($B$8=Population!$A$2:$A$1611)"
        part2 = "*($D${}=Population!$C$2:$C$1611), 0,1),0))".format(6+row)
        ws[cell] = part1 + part2
        ws[cell].style = 'Comma'

    ws['F6'] = "Area (km^2)"
    ws.column_dimensions['F'].width = 20
    for row in range(1, 11):
        cell = "F{}".format(6+row)
        part1 = "=INDEX(Population!$E$2:$E$1611,MATCH(1,INDEX(($B$8=Population!$A$2:$A$1611)"
        part2 = "*($D${}=Population!$C$2:$C$1611), 0,1),0))".format(6+row)
        ws[cell] = part1 + part2
        ws[cell].style = 'Comma'

    ws['G6'] = "Pop Density (km^2)"
    ws.column_dimensions['G'].width = 20
    for row in range(1, 11):
        cell = "G{}".format(6+row)
        part1 = "=INDEX(Population!$F$2:$F$1611,MATCH(1,INDEX(($B$8=Population!$A$2:$A$1611)"
        part2 = "*($D${}=Population!$C$2:$C$1611), 0,1),0))".format(6+row)
        ws[cell] = part1 + part2
        ws[cell].style = 'Comma'

    ##Set deciles box
    set_border(ws, 'D6:J16', "thin", "000000")

    return ws


def set_border(ws, cell_range, style, color):

    thin = Side(border_style=style, color=color)
    for row in ws[cell_range]:
        for cell in row:
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)


def add_options(ws):
    """

    """
    ws.sheet_properties.tabColor = "0099ff"

    data = pd.read_csv(os.path.join(BASE_PATH, 'global_information.csv'), encoding = "ISO-8859-1")

    data = data.rename({
        'country': 'Country',
        'ISO_3digit': 'ISO3'
        },
        axis='columns'
    )

    data = data[[
        'Country',
        'ISO3',
    ]]

    for r in dataframe_to_rows(data, index=False, header=True):
        ws.append(r)

    ws['C1'] = "Scenario"
    ws['C2'] = "High"
    ws['C3'] = "Baseline"
    ws['C4'] = "Low"

    ws['D1'] = "Strategy"
    ws['D2'] = "4G (Wireless)"
    ws['D3'] = "4G (Fiber)"
    ws['D4'] = "5G (Wireless)"
    ws['D5'] = "5G (Fiber)"

    ws['E1'] = "Spectrum"
    ws['E2'] = "High"
    ws['E3'] = "Baseline"
    ws['E4'] = "Low"

    ws['F1'] = "Sites"
    ws['F2'] = "High"
    ws['F3'] = "Baseline"
    ws['F4'] = "Low"

    return ws


def add_population_sheet(ws):
    """
    Add the population sheet.

    """
    ws.sheet_properties.tabColor = "666600"

    path = os.path.join(DATA_INTERMEDIATE, 'all_pop_data.csv')
    population = pd.read_csv(path)

    population = population.rename({
        'GID_0': 'iso3',
        }, axis='columns')

    population = population[[
        'iso3',
        'country_name',
        'decile',
        'population',
        'area_km2',
        'population_km2',
        ]]

    for r in dataframe_to_rows(population, index=False, header=True):
        ws.append(r)

    return ws


def add_demand_sheet(ws):
    """

    """
    ##Color white
    set_border(ws, 'A1:AZ1000', "thin", "00FFFFFF")

    ##Set column width
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['E'].width = 25
    ws.column_dimensions['F'].width = 25
    ws.column_dimensions['G'].width = 25
    ws.column_dimensions['H'].width = 25
    ws.column_dimensions['I'].width = 25
    ws.column_dimensions['J'].width = 25

    ##Allocate title
    ws.title = "Demand"
    ws.sheet_properties.tabColor = "66ff99"

    ## Add parameters box
    set_border(ws, 'A1:J11', "thin", "000000")

    for i in range(1, 12): #Decile
        cell = "A{}".format(i)
        ws[cell] = "=Settings!D{}".format(5+i)

    for i in range(1, 12): #Population
        cell = "B{}".format(i)
        ws[cell] = "=Settings!E{}".format(5+i)

    for i in range(1, 12): #Area
        cell = "C{}".format(i)
        ws[cell] = "=Settings!F{}".format(5+i)

    for i in range(1, 12): #Pop Density
        cell = "D{}".format(i)
        ws[cell] = "=Settings!G{}".format(5+i)

    ws['E1'] = 'Users'
    for i in range(1, 11): #Users
        cell = "E{}".format(i+1)
        ws[cell] = "=(Settings!$B$11/100)*B{}".format(i+1)

    ws['F1'] = 'User Density (km^2)'
    for i in range(1, 11): #User Density
        cell = "F{}".format(i+1)
        ws[cell] = "=(Settings!$B$11/100)*D{}".format(i+1)

    ws['G1'] = 'Active Users'
    for i in range(1, 11): #Users
        cell = "G{}".format(i+1)
        ws[cell] = "=(Settings!$B$12/100)*E{}".format(i+1)

    ws['H1'] = 'Active User Density (km^2)'
    for i in range(1, 11): #User Density
        cell = "H{}".format(i+1)
        ws[cell] = "=(Settings!$B$12/100)*F{}".format(i+1)

    ws['I1'] = 'Total Demand (Mbps)'
    for i in range(1, 11): #Users
        cell = "I{}".format(i+1)
        ws[cell] = "=Settings!$B$9*G{}".format(i+1)

    ws['J1'] = 'Demand Density (Mbps)'
    for i in range(1, 11): #Users
        cell = "J{}".format(i+1)
        ws[cell] = "=Settings!$B$9*F{}".format(i+1)

    return ws


def add_supply_sheet(ws):
    """

    """
    ##Color white
    set_border(ws, 'A1:AZ1000', "thin", "00FFFFFF")

    ##Set column width
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['E'].width = 25
    ws.column_dimensions['F'].width = 25
    ws.column_dimensions['G'].width = 25
    ws.column_dimensions['H'].width = 25
    ws.column_dimensions['I'].width = 25
    ws.column_dimensions['J'].width = 25

    ##Allocate title
    ws.title = "Supply"
    ws.sheet_properties.tabColor = "0000ff"

    ## Add parameters box
    set_border(ws, 'A1:J11', "thin", "000000")

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
