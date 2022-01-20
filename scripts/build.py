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

    capacity = wb.create_sheet("Capacity", (6-1))
    capacity = add_capacity_sheet(capacity)

    costs = wb.create_sheet("Costs", (7-1))
    costs = add_cost_sheet(costs)

    gdp = wb.create_sheet("GDP", (8-1))
    gdp = add_gdp_sheet(gdp)

    wb.save('Oughton et al. (2022) DICE.xlsx')

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
    set_border(ws, 'A6:B15', "thin", "000000")
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
    ws['B10'] = "Baseline"

    ws['A11'] = "Market Share (%)"
    ws['B11'] = 25

    ws['A12'] = "Active Users (%)"
    ws['B12'] = 5

    ws['A13'] = "Infrastructure Strategy"
    data_val = DataValidation(type="list", formula1='=Options!D2:D6')
    ws.add_data_validation(data_val)
    data_val.add(ws["B13"])
    ws['B13'] = "4G"

    ws['A14'] = "Spectrum Availability"
    data_val = DataValidation(type="list", formula1='=Options!E2:E4')
    ws.add_data_validation(data_val)
    data_val.add(ws["B14"])
    ws['B14'] = "Baseline"

    ws['A15'] = "Sites Availability"
    data_val = DataValidation(type="list", formula1='=Options!F2:F4')
    ws.add_data_validation(data_val)
    data_val.add(ws["B15"])
    ws['B15'] = "Baseline"

    ws['A16'] = "Sites Availability"
    data_val = DataValidation(type="list", formula1='=Options!F2:F4')
    ws.add_data_validation(data_val)
    data_val.add(ws["B15"])
    ws['B16'] = "Baseline"

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

    ws['H6'] = "Active User Density (km^2)"
    ws.column_dimensions['H'].width = 25
    for row in range(1, 11):
        cell = "H{}".format(6+row)
        ws[cell] = "=Demand!H{}".format(row+1)
        ws[cell].style = 'Comma'

    ws['I6'] = "Demand Density (Mbps/km^2)"
    ws.column_dimensions['I'].width = 25
    for row in range(1, 11):
        cell = "I{}".format(6+row)
        ws[cell] = "=Demand!J{}".format(row+1)
        ws[cell].style = 'Comma'

    ws['I6'] = "Demand Density (Mbps/km^2)"
    ws.column_dimensions['I'].width = 25
    for row in range(1, 11):
        cell = "I{}".format(6+row)
        ws[cell] = "=Demand!J{}".format(row+1)
        ws[cell].style = 'Comma'

    ws['J6'] = "Required New Sites"
    ws.column_dimensions['J'].width = 25
    for row in range(1, 11):
        cell = "J{}".format(6+row)
        ws[cell] = "=Supply!L{}".format(row+1)
        ws[cell].style = 'Comma'

    ws.column_dimensions['K'].width = 25
    for row in range(1, 12):
        cell = "K{}".format(5+row)
        ws[cell] = "=Costs!H{}".format(row)
        ws[cell].style = 'Comma'

    ws.column_dimensions['L'].width = 25
    for row in range(1, 12):
        cell = "L{}".format(5+row)
        ws[cell] = "=Costs!I{}".format(row)
        ws[cell].style = 'Comma'

    ws.column_dimensions['M'].width = 25
    for row in range(1, 12):
        cell = "M{}".format(5+row)
        ws[cell] = "=Costs!J{}".format(row)
        ws[cell].style = 'Comma'

    ws.column_dimensions['N'].width = 25
    ws['N6'] = "Share of Annual GDP (%)"
    for row in range(1, 11):
        cell = "N{}".format(6+row)
        ws[cell] = "=VLOOKUP($B$8,GDP!$A$2:$B$266,2)/M{}".format(6+row)
        # ws[cell].style = 'Comma'

    ##Set deciles box
    set_border(ws, 'D6:N16', "thin", "000000")

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

    population['population'] = round(population['population'])

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
    ws.column_dimensions['K'].width = 25
    ws.column_dimensions['L'].width = 25

    ##Allocate title
    ws.title = "Supply"
    ws.sheet_properties.tabColor = "0000ff"

    ## Add parameters box
    set_border(ws, 'A1:L11', "thin", "000000")

    for i in range(1, 12): #Decile
        cell = "A{}".format(i)
        ws[cell] = "=Demand!A{}".format(i)

    for i in range(1, 12): #Active Users
        cell = "B{}".format(i)
        ws[cell] = "=Demand!G{}".format(i)

    for i in range(1, 12): #Active User Density
        cell = "C{}".format(i)
        ws[cell] = "=Demand!H{}".format(i)

    for i in range(1, 12): #Total Demand
        cell = "D{}".format(i)
        ws[cell] = "=Demand!I{}".format(i)

    for i in range(1, 12): #Demand Density
        cell = "E{}".format(i)
        ws[cell] = "=Demand!J{}".format(i)

    ws['F1'] = 'Total Sites'
    for i in range(2, 12): #Total Sites
        cell = "F{}".format(i)
        ws[cell] = "=Demand!C{}/20".format(i)

    ws['G1'] = 'Total Site Density (km^2)'
    for i in range(2, 12): #Total Sites Density
        cell = "G{}".format(i)
        ws[cell] = "=F{}/Demand!C{}".format(i, i)

    ws['H1'] = 'MNO Sites'
    for i in range(2, 12): #Total Sites
        cell = "H{}".format(i)
        ws[cell] = "=F{}/Settings!$B$11".format(i)

    ws['I1'] = 'MNO Site Density (km^2)'
    for i in range(2, 12): #Total Sites
        cell = "I{}".format(i)
        ws[cell] = "=G{}/Settings!$B$11".format(i)

    ws['J1'] = 'Capacity (Mbps/km^2)'
    for i in range(2, 12): #Capacity
        cell = "J{}".format(i)
        part1 = "=I{}*VLOOKUP(Settings!$B$14,Capacity!$A$3:$B$5, 2)".format(i)
        part2 = "*VLOOKUP(Settings!$B$13,Capacity!$D$9:$E$11, 2)".format(i)
        ws[cell] = part1 + part2

    ws['K1'] = 'Required Site Density (Sites/km^2)'
    for i in range(2, 12): #Capacity
        cell = "K{}".format(i)
        ws[cell] = "=MIN(IF(Capacity!$L$3:$L$10>Demand!J{},Capacity!$K$3:$K$10))".format(i)
        ws.formula_attributes[cell] = {'t': 'array', 'ref': "{}:{}".format(cell, cell)}

    ws['L1'] = 'Required Sites'
    for i in range(2, 12): #Capacity
        cell = "L{}".format(i)
        ws[cell] = "=IF(I{}<K{},(K{}-I{})*Demand!C{},0)".format(i,i,i,i,i)

    return ws


def add_capacity_sheet(ws):
    """

    """
    ws.sheet_properties.tabColor = "66ffcc"

    ##Spectrum Portfolio Box
    ws.merge_cells('A1:B1')
    ws['A1'] = "Spectrum Portfolio"

    ws['A2'] = 'Scenario'
    ws['A3'] = 'High'
    ws['A4'] = 'Baseline'
    ws['A5'] = 'Low'

    ws['B2'] = 'Bandwidth (MHz)'
    ws['B3'] = 20
    ws['B4'] = 10
    ws['B5'] = 5

    set_border(ws, 'A1:B5', "thin", "000000")
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 18

    ##Spectral Efficiency Box
    ws.merge_cells('D1:G1')
    ws['D1'] = "Spectral Effiency (b/s/Hz)"

    ws['D2'] = 'Distance'
    ws['D3'] = 5
    ws['D4'] = 10
    ws['D5'] = 15
    ws['D6'] = 20

    ws['E2'] = '3G'
    ws['E3'] = 2
    ws['E4'] = 2
    ws['E5'] = 2
    ws['E6'] = 2

    ws['F2'] = '4G'
    ws['F3'] = 4
    ws['F4'] = 4
    ws['F5'] = 4
    ws['F6'] = 4

    ws['G2'] = '5G'
    ws['G3'] = 6
    ws['G4'] = 6
    ws['G5'] = 6
    ws['G6'] = 6

    set_border(ws, 'D1:G6', "thin", "000000")

    ws['D8'] = "Tech"
    ws['D9'] = "3G"
    ws['D10'] = "4G"
    ws['D11'] = "5G"
    ws['E8'] = "Mean (b/s/Hz)"
    ws['E9'] = "=AVERAGE(E3:E6)"
    ws['E10'] = "=AVERAGE(F3:F6)"
    ws['E11'] = "=AVERAGE(G3:G6)"

    ws.column_dimensions['E'].width = 15

    set_border(ws, 'D8:E11', "thin", "000000")

    ##Density Lookup Table
    ws.merge_cells('I1:L1')
    ws['I1'] = "Density Lookup Table"

    ws['I2'] = 'Sites'
    ws['I3'] = 4
    ws['I4'] = 2
    ws['I5'] = 1
    ws['I6'] = 0.5
    ws['I7'] = 0.25
    ws['I8'] = 0.1
    ws['I9'] = 0.05

    ws['J2'] = 'Area'
    ws['J3'] = 1
    ws['J4'] = 1
    ws['J5'] = 1
    ws['J6'] = 1
    ws['J7'] = 1
    ws['J8'] = 1
    ws['J9'] = 1

    ws['K2'] = 'Density (sites/km^2)'
    ws['K3'] = 4
    ws['K4'] = 2
    ws['K5'] = 1
    ws['K6'] = 0.5
    ws['K7'] = 0.25
    ws['K8'] = 0.1
    ws['K9'] = 0.05

    ws['L2'] = 'Capacity (Mbps/km^2)'
    ws['L3'] = 2500
    ws['L4'] = 500
    ws['L5'] = 100
    ws['L6'] = 50
    ws['L7'] = 25
    ws['L8'] = 1
    ws['L9'] = 0.5

    set_border(ws, 'I1:L9', "thin", "000000")
    ws.column_dimensions['K'].width = 22
    ws.column_dimensions['L'].width = 22

    return ws

def add_cost_sheet(ws):
    """

    """
    ws.sheet_properties.tabColor = "9966ff"

    #Column dimensions
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 15
    ws.column_dimensions['G'].width = 15
    ws.column_dimensions['H'].width = 15
    ws.column_dimensions['I'].width = 15
    ws.column_dimensions['J'].width = 15

    ##Spectrum Portfolio Box
    set_border(ws, 'A14:C18', "thin", "000000")
    ws.merge_cells('A14:C14')
    ws['A14'] = "Equipment Costs"

    ws['A15'] = 'Asset'
    ws['A16'] = 'RAN'
    ws['A17'] = 'Fiber'
    ws['A18'] = 'Towers'

    ws['B15'] = 'Cost ($)'
    ws['B16'] = 30000
    ws['B17'] = 10
    ws['B18'] = 30000

    ws['C15'] = 'Unit'
    ws['C16'] = "Per Tower"
    ws['C17'] = "Per Meter"
    ws['C18'] = "Per Tower"

    #Deciles
    set_border(ws, 'A1:J11', "thin", "000000")

    for i in range(1, 12): #Decile
        cell = "A{}".format(i)
        ws[cell] = "=Demand!A{}".format(i)

    for i in range(1, 12): #Population
        cell = "B{}".format(i)
        ws[cell] = "=Demand!B{}".format(i)

    for i in range(1, 12): #Area
        cell = "C{}".format(i)
        ws[cell] = "=Demand!C{}".format(i)

    for i in range(1, 12): #Sites
        cell = "D{}".format(i)
        ws[cell] = "=Supply!L{}".format(i)

    ws["E1"] = "RAN"
    for i in range(2, 12): #RAN
        cell = "E{}".format(i)
        ws[cell] = "=Supply!L{}*VLOOKUP($E$1, $A$14:$B$24, 2, FALSE)".format(i)

    ws["F1"] = "Fiber"
    for i in range(2, 12): #Fiber
        cell = "F{}".format(i)
        ws[cell] = "=Supply!L{}*VLOOKUP($F$1, $A$14:$B$24, 2, FALSE)".format(i)

    ws["G1"] = "Towers"
    for i in range(2, 12): #Towers
        cell = "G{}".format(i)
        ws[cell] = "=Supply!L{}*VLOOKUP($G$1, $A$14:$B$24, 2, FALSE)".format(i)

    ws["H1"] = "Total MNO Cost"
    for i in range(2, 12): #Towers
        cell = "H{}".format(i)
        ws[cell] = "=SUM(E{}, F{}, G{})".format(i,i,i)

    ws["I1"] = "Cost Per User"
    for i in range(2, 12): #Towers
        cell = "I{}".format(i)
        ws[cell] = "=H{}/Demand!E{}".format(i,i)

    ws["J1"] = "Total Market Cost"
    for i in range(2, 12): #Towers
        cell = "J{}".format(i)
        ws[cell] = "=I{}*B{}".format(i,i)

    return ws


def add_gdp_sheet(ws):
    """

    """
    ws.sheet_properties.tabColor = "ffff33"

    path = os.path.join(DATA_RAW, 'gdp.csv')
    gdp = pd.read_csv(path)

    # population = population.rename({
    #     'GID_0': 'iso3',
    #     }, axis='columns')

    gdp = gdp[[
        'iso3',
        # 'country',
        'gdp',
        'income_group',
        'source',
        'Date',
        ]]

    gdp = gdp.sort_values('iso3')

    for r in dataframe_to_rows(gdp, index=False, header=True):
        ws.append(r)

    return gdp

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
