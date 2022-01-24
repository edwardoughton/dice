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

    population = wb.create_sheet("Population", (2-1))
    population = add_country_data(population, 'population')

    area = wb.create_sheet("Area", (3-1))
    area = add_country_data(area, 'area_km2')

    pop_density = wb.create_sheet("Pop_Density", (4-1))
    pop_density = add_country_data(pop_density, 'population_km2')

    cols = ['A','B','C','D','E','F','G','H','I','J','K','L']
    users = wb.create_sheet("Users", (5-1))
    users = add_users(users, cols)
    users.sheet_properties.tabColor = "66ff99"

    data_demand = wb.create_sheet("Data Demand", (6-1))
    data_demand = add_data_demand(data_demand, cols)
    data_demand.sheet_properties.tabColor = "66ff99"

    coverage = wb.create_sheet("Coverage", (7-1))
    coverage = add_coverage_sheet(coverage, cols)
    coverage.sheet_properties.tabColor = "0000ff"

    towers = wb.create_sheet("Towers", (8-1))
    towers = add_towers_sheet(towers, cols)
    towers.sheet_properties.tabColor = "0000ff"

    lookups = wb.create_sheet("Lookups", (9-1))
    lookups = add_lookups_sheet(lookups)
    lookups.sheet_properties.tabColor = "0000ff"

    capacity = wb.create_sheet("Capacity", (10-1))
    capacity = add_capacity_sheet(capacity, cols)
    capacity.sheet_properties.tabColor = "0000ff"

    sites = wb.create_sheet("Total Sites", (11-1))
    sites = add_sites_sheet(sites, cols)
    sites.sheet_properties.tabColor = "0000ff"

    new = wb.create_sheet("New Sites", (12-1))
    new = add_new_sites_sheet(new, cols)
    new.sheet_properties.tabColor = "0000ff"

    costs = wb.create_sheet("Costs", (13-1))
    costs = add_cost_sheet(costs, cols)

    gdp = wb.create_sheet("GDP", (14-1))
    gdp = add_gdp_sheet(gdp)

    options = wb.create_sheet("Options", (15-1))
    options = add_options(options)

    wb.save('Oughton et al. (2022) DICE.xlsx')

    return print("Generated workbook")


def add_settings(ws):
    """

    """
    ##Color white
    set_border(ws, 'A1:AZ1000', "thin", "00FFFFFF")

    ##Set column width
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 15
    ws.column_dimensions['G'].width = 30
    ws.column_dimensions['H'].width = 15

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
    set_border(ws, 'A6:B11', "thin", "000000")
    ws.merge_cells('A6:B6')
    ws['A6'] = "Country Parameters"

    ws['A7'] = "Parameter"
    ws['B7'] = "Option"

    ws['A8'] = "Country"
    data_val = DataValidation(type="list", formula1='=Options!A2:A251')
    ws.add_data_validation(data_val)
    data_val.add(ws["B8"])
    ws['B8'] = "Afghanistan"

    ws['A9'] = "ISO3"
    ws['B9'] = "=IFERROR(INDEX(Options!B2:B1611,MATCH(Settings!B8, Options!A2:A1611)), \"\")"

    ws['A10'] = "Start Year"
    ws['B10'] = 2022

    ws['A11'] = "End Year"
    ws['B11'] = 2030

    set_border(ws, 'D6:E13', "thin", "000000")
    ws.merge_cells('D6:E6')
    ws['D6'] = "Demand Parameters"

    ws['D7'] = "Parameter"
    ws['E7'] = "Option"

    ws['D8'] = "Pop. Growth Annually (%)"
    ws['E8'] = 3

    ws['D9'] = "Smartphone adoption (%)"
    ws['E9'] = 95

    ws['D10'] = "Market Share (%)"
    ws['E10'] = 25

    ws['D11'] = "Active Users (%)"
    ws['E11'] = 5

    ws['D12'] = "Minimum capacity per user (Mbps)"
    ws['E12'] = 10

    ws['D13'] = "Data demand per month (GB) "
    ws['E13'] = 2

    set_border(ws, 'G6:H9', "thin", "000000")
    ws.merge_cells('G6:H6')
    ws['G6'] = "Supply Parameters"

    ws['G7'] = "Parameter"
    ws['H7'] = "Option"

    ws['G8'] = "Infrastructure Strategy"
    data_val = DataValidation(type="list", formula1='=Options!D2:D6')
    ws.add_data_validation(data_val)
    data_val.add(ws["H8"])
    ws['H8'] = "4G"

    ws['G9'] = "Spectrum Availability"
    data_val = DataValidation(type="list", formula1='=Options!E2:E4')
    ws.add_data_validation(data_val)
    data_val.add(ws["H9"])
    ws['H9'] = "Baseline"

    # ws['A15'] = "Sites Availability"
    # data_val = DataValidation(type="list", formula1='=Options!F2:F4')
    # ws.add_data_validation(data_val)
    # data_val.add(ws["B15"])
    # ws['B15'] = "Baseline"


    # ws['G8'] = "Spectrum Availability"
    # ws['H8'] =

    # ws['G9'] = "Smartphone adoption (%)"
    # ws['H9'] = 95

    # ws['G10'] = "Market Share (%)"
    # ws['H10'] = 25

    # ws['G11'] = "Active Users (%)"
    # ws['H11'] = 5





    # ws['A16'] = "Sites Availability"
    # data_val = DataValidation(type="list", formula1='=Options!F2:F4')
    # ws.add_data_validation(data_val)
    # data_val.add(ws["B15"])
    # ws['B16'] = "Baseline"

    # ########Deciles
    # ws['D6'] = "Decile"
    # for row in range(1, 11):
    #     cell = "D{}".format(6+row)
    #     ws[cell] = row * 10

    # ws['E6'] = "Population"
    # ws.column_dimensions['E'].width = 20
    # for row in range(1, 11):
    #     cell = "E{}".format(6+row)
    #     part1 = "=INDEX(Population!$D$2:$D$1611,MATCH(1,INDEX(($B$8=Population!$A$2:$A$1611)"
    #     part2 = "*($D${}=Population!$C$2:$C$1611), 0,1),0))".format(6+row)
    #     ws[cell] = part1 + part2
    #     ws[cell].style = 'Comma'

    # ws['F6'] = "Area (km^2)"
    # ws.column_dimensions['F'].width = 20
    # for row in range(1, 11):
    #     cell = "F{}".format(6+row)
    #     part1 = "=INDEX(Population!$E$2:$E$1611,MATCH(1,INDEX(($B$8=Population!$A$2:$A$1611)"
    #     part2 = "*($D${}=Population!$C$2:$C$1611), 0,1),0))".format(6+row)
    #     ws[cell] = part1 + part2
    #     ws[cell].style = 'Comma'

    # ws['G6'] = "Pop Density (km^2)"
    # ws.column_dimensions['G'].width = 20
    # for row in range(1, 11):
    #     cell = "G{}".format(6+row)
    #     part1 = "=INDEX(Population!$F$2:$F$1611,MATCH(1,INDEX(($B$8=Population!$A$2:$A$1611)"
    #     part2 = "*($D${}=Population!$C$2:$C$1611), 0,1),0))".format(6+row)
    #     ws[cell] = part1 + part2
    #     ws[cell].style = 'Comma'

    # ws['H6'] = "Active User Density (km^2)"
    # ws.column_dimensions['H'].width = 25
    # for row in range(1, 11):
    #     cell = "H{}".format(6+row)
    #     ws[cell] = "=Demand!H{}".format(row+1)
    #     ws[cell].style = 'Comma'

    # ws['I6'] = "Demand Density (Mbps/km^2)"
    # ws.column_dimensions['I'].width = 25
    # for row in range(1, 11):
    #     cell = "I{}".format(6+row)
    #     ws[cell] = "=Demand!J{}".format(row+1)
    #     ws[cell].style = 'Comma'

    # ws['I6'] = "Demand Density (Mbps/km^2)"
    # ws.column_dimensions['I'].width = 25
    # for row in range(1, 11):
    #     cell = "I{}".format(6+row)
    #     ws[cell] = "=Demand!J{}".format(row+1)
    #     ws[cell].style = 'Comma'

    # ws['J6'] = "Required New Sites"
    # ws.column_dimensions['J'].width = 25
    # for row in range(1, 11):
    #     cell = "J{}".format(6+row)
    #     ws[cell] = "=Supply!L{}".format(row+1)
    #     ws[cell].style = 'Comma'

    # ws.column_dimensions['K'].width = 25
    # for row in range(1, 12):
    #     cell = "K{}".format(5+row)
    #     ws[cell] = "=Costs!H{}".format(row)
    #     ws[cell].style = 'Comma'

    # ws.column_dimensions['L'].width = 25
    # for row in range(1, 12):
    #     cell = "L{}".format(5+row)
    #     ws[cell] = "=Costs!I{}".format(row)
    #     ws[cell].style = 'Comma'

    # ws.column_dimensions['M'].width = 25
    # for row in range(1, 12):
    #     cell = "M{}".format(5+row)
    #     ws[cell] = "=Costs!J{}".format(row)
    #     ws[cell].style = 'Comma'

    # ws.column_dimensions['N'].width = 25
    # ws['N6'] = "Share of Annual GDP (%)"
    # for row in range(1, 11):
    #     cell = "N{}".format(6+row)
    #     ws[cell] = "=VLOOKUP($B$8,GDP!$A$2:$B$266,2)/M{}".format(6+row)
    #     # ws[cell].style = 'Comma'

    # ##Set deciles box
    # set_border(ws, 'D6:N16', "thin", "000000")

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


def add_country_data(ws, metric):
    """
    Add the country data sheets.

    """
    ws.sheet_properties.tabColor = "666600"

    filename = 'all_pop_data_{}.csv'.format(metric)
    path = os.path.join(DATA_INTERMEDIATE, filename)
    data = pd.read_csv(path)

    data = data.rename({
        'GID_0': 'iso3',
        }, axis='columns')

    for r in dataframe_to_rows(data, index=False, header=True):
        ws.append(r)

    return ws


def add_users(ws, cols):
    """
    Calculate the total number of users.

    """
    for col in cols:
            cell = "{}1".format(col)
            ws[cell] = "='Pop_Density'!{}".format(cell)

    for col in cols[:2]:
        for i in range(2, 250):
            cell = "{}{}".format(col, i)
            ws[cell] = "='Pop_Density'!{}".format(cell)

    for col in cols[2:]:
        for i in range(2, 250): #Decile
            cell = "{}{}".format(col, i)
            part1 = "='Pop_Density'!{}".format(cell) #
            part2 = "*(POWER(1+(Settings!$E$8/100), Settings!$B$11-Settings!$B$10))"
            part3 = "*(Settings!$E$9/100)*(Settings!$E$10/100)*(Settings!$E$11/100)"
            ws[cell] = part1 + part2 + part3

    return ws


def add_data_demand(ws, cols):
    """
    Calculate the total data demand.

    """
    for col in cols:
            cell = "{}1".format(col)
            ws[cell] = "='Users'!{}".format(cell)

    for col in cols[:2]:
        for i in range(2, 250):
            cell = "{}{}".format(col, i)
            ws[cell] = "='Users'!{}".format(cell)

    for col in cols[2:]:
        for i in range(2, 250): #Decile
            cell = "{}{}".format(col, i)
            part1 = "=(Users!{}*MAX(Settings!$E$12,".format(cell)
            part2 = "(Settings!$E$13*1024*8*(1/30)*(15/100)*(1/3600))))"
            ws[cell] = part1 + part2

    return ws


def add_coverage_sheet(ws, cols):
    """

    """
    for col in cols[:2]:
        for i in range(1, 250):
            cell = "{}{}".format(col, i)
            ws[cell] = "='Data Demand'!{}".format(cell)

    ws['C1'] = 'Towers'
    for i in range(2, 250): #Decile
        cell = "C{}".format(i)
        ws[cell] = 1000

    ws['D1'] = 'Coverage'
    for i in range(2, 250): #Decile
        cell = "D{}".format(i)
        ws[cell] = 50

    ws['E1'] = 'Sites per covered population'
    for i in range(2, 250): #Decile
        cell = "E{}".format(i)
        ws[cell] = "=(SUM(Population!C{}:L{})*(D{}/100)/C{})".format(i,i,i,i)

    return ws


def add_towers_sheet(ws, cols):
    """

    """
    for col in cols:
        cell = "{}1".format(col)
        ws[cell] = "='Data Demand'!{}".format(cell)

    for col in cols[:2]:
        for i in range(1, 250):
            cell = "{}{}".format(col, i)
            ws[cell] = "='Coverage'!{}".format(cell)

    for i in range(2, 250): #Decile
        cell = "C{}".format(i)
        ws[cell] = "=Population!C{}/Coverage!$E${}".format(i,i)

    for i in range(2, 250): #Decile
        cell = "D{}".format(i)
        ws[cell] = "=IF(SUM(C2)<Coverage!C3, Population!D{}/Coverage!$E${},0)".format(i,i)

    for i in range(2, 250): #Decile
        cell = "E{}".format(i)
        ws[cell] = "=IF(SUM(C2:D2)<Coverage!C3, Population!E{}/Coverage!$E${},0)".format(i,i)

    for i in range(2, 250): #Decile
        cell = "F{}".format(i)
        ws[cell] = "=IF(SUM(C2:E2)<Coverage!C3, Population!F{}/Coverage!$E${},0)".format(i,i)

    for i in range(2, 250): #Decile
        cell = "G{}".format(i)
        ws[cell] = "=IF(SUM(C2:F2)<Coverage!C3, Population!G{}/Coverage!$E${},0)".format(i,i)

    for i in range(2, 250): #Decile
        cell = "H{}".format(i)
        ws[cell] = "=IF(SUM(C2:G2)<Coverage!C3, Population!H{}/Coverage!$E${},0)".format(i,i)

    for i in range(2, 250): #Decile
        cell = "I{}".format(i)
        ws[cell] = "=IF(SUM(C2:H2)<Coverage!C3, Population!I{}/Coverage!$E${},0)".format(i,i)

    for i in range(2, 250): #Decile
        cell = "J{}".format(i)
        ws[cell] = "=IF(SUM(C2:I2)<Coverage!C3, Population!J{}/Coverage!$E${},0)".format(i,i)

    for i in range(2, 250): #Decile
        cell = "K{}".format(i)
        ws[cell] = "=IF(SUM(C2:J2)<Coverage!C3, Population!K{}/Coverage!$E${},0)".format(i,i)

    for i in range(2, 250): #Decile
        cell = "L{}".format(i)
        ws[cell] = "=IF(SUM(C2:K2)<Coverage!C3, Population!L{}/Coverage!$E${},0)".format(i,i)

    return ws


def add_lookups_sheet(ws):
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

    ##Cost Information
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


    return ws


def add_capacity_sheet(ws, cols):
    """

    """
    # ##Color white
    # set_border(ws, 'A1:AZ1000', "thin", "00FFFFFF")

    for col in cols:
            cell = "{}1".format(col)
            ws[cell] = "='Data Demand'!{}".format(cell)

    for col in cols[:2]:
        for i in range(2, 250):
            cell = "{}{}".format(col, i)
            ws[cell] = "='Data Demand'!{}".format(cell)

    for col in cols[2:]:
        for i in range(2, 250): #Total Sites Density
            cell = "{}{}".format(col, i)
            part1 = "=Towers!{}/Area!{}*(Settings!$E$10/100)".format(cell, cell)
            part2 = "*VLOOKUP(Settings!$H$9, Lookups!$A$3:$B$5, 2)".format(i)
            part3 = "*VLOOKUP(Settings!$H$8,Lookups!$D$9:$E$11, 2)".format(i)
            ws[cell] = part1 + part2 + part3

    return ws


def add_sites_sheet(ws, cols):
    """

    """
    # ##Color white
    # set_border(ws, 'A1:AZ1000', "thin", "00FFFFFF")

    for col in cols:
            cell = "{}1".format(col)
            ws[cell] = "='Capacity'!{}".format(cell)

    for col in cols[:2]:
        for i in range(2, 250):
            cell = "{}{}".format(col, i)
            ws[cell] = "='Capacity'!{}".format(cell)

    for col in cols[2:]:
        for i in range(2, 250):
            cell = "{}{}".format(col, i)
            # part1 = "=MIN(IF(Lookups!$L$3:$L$10>Data Demand!{},Lookups!$K$3:$K$10))".format(cell)
            part1 = "=MIN(IF('Lookups'!$L$3:$L$10>'Data Demand'!{},'Lookups'!$K$3:$K$10))".format(cell)
            # part2 = "*VLOOKUP(Settings!$H$9, Capacity!$A$3:$B$5, 2)".format(i)
            # part3 = "*VLOOKUP(Settings!$H$8,Capacity!$D$9:$E$11, 2)".format(i)
            ws[cell] = part1 #+ part2 + part3
            ws.formula_attributes[cell] = {'t': 'array', 'ref': "{}:{}".format(cell, cell)}


    # ws['L1'] = 'Required Sites'
    # for i in range(2, 12): #Capacity
    #     cell = "L{}".format(i)
    #     ws[cell] = "=IF(I{}<K{},(K{}-I{})*Demand!C{},0)".format(i,i,i,i,i)

    return ws


def add_new_sites_sheet(ws, cols):
    """

    """
    # ##Color white
    # set_border(ws, 'A1:AZ1000', "thin", "00FFFFFF")

    for col in cols:
            cell = "{}1".format(col)
            ws[cell] = "='Total Sites'!{}".format(cell)

    for col in cols[:2]:
        for i in range(2, 250):
            cell = "{}{}".format(col, i)
            ws[cell] = "='Total Sites'!{}".format(cell)

    for col in cols[2:]:
        for i in range(2, 250):
            cell = "{}{}".format(col, i)
            part1 = "=IF(Towers!{}/Area!{}*(Settings!$E$10/100)<'Total Sites'!{},".format(cell, cell, cell)
            part2 = "('Total Sites'!{}-Towers!{}/Area!{}*(Settings!$E$10/100)),0)".format(cell, cell, cell)
            ws[cell] = part1 + part2 #+ part3
            ws.formula_attributes[cell] = {'t': 'array', 'ref': "{}:{}".format(cell, cell)}

    return ws


def add_cost_sheet(ws, cols):
    """

    """
    ws.sheet_properties.tabColor = "9966ff"

    #Column dimensions
    # ws.column_dimensions['A'].width = 15
    # ws.column_dimensions['B'].width = 15
    # ws.column_dimensions['C'].width = 15
    # ws.column_dimensions['D'].width = 15
    # ws.column_dimensions['E'].width = 15
    # ws.column_dimensions['F'].width = 15
    # ws.column_dimensions['G'].width = 15
    # ws.column_dimensions['H'].width = 15
    # ws.column_dimensions['I'].width = 15
    # ws.column_dimensions['J'].width = 15

    #Deciles
    # set_border(ws, 'A1:J11', "thin", "000000")

    for col in cols:
            cell = "{}1".format(col)
            ws[cell] = "='New Sites'!{}".format(cell)

    for col in cols[:2]:
        for i in range(2, 250):
            cell = "{}{}".format(col, i)
            ws[cell] = "='New Sites'!{}".format(cell)

    for col in cols[2:]:
        for i in range(2, 250):
            cell = "{}{}".format(col, i)
            part1 = "='New Sites'!{}*VLOOKUP('Lookups'!$A$16, 'Lookups'!$A$14:'Lookups'!$B$24, 2, FALSE)".format(cell)
            part2 = "+'New Sites'!{}*VLOOKUP('Lookups'!$A$17, 'Lookups'!$A$17:'Lookups'!$B$24, 2, FALSE)".format(cell)
            part3 = "+'New Sites'!{}*VLOOKUP('Lookups'!$A$18, 'Lookups'!$A$18:'Lookups'!$B$24, 2, FALSE)".format(cell)
            ws[cell] = part1 + part2 + part3

    return ws


def add_gdp_sheet(ws):
    """

    """
    ws.sheet_properties.tabColor = "ffff33"

    path = os.path.join(DATA_RAW, 'gdp.csv')
    gdp = pd.read_csv(path)

    gdp = gdp[[
        'iso3',
        'gdp',
        'income_group',
        'source',
        'Date',
        ]]

    gdp = gdp.sort_values('iso3')

    for r in dataframe_to_rows(gdp, index=False, header=True):
        ws.append(r)

    return gdp


if __name__ == "__main__":

    generate_workbook()
