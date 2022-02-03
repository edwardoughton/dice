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

    population = wb.create_sheet("Pop", (2-1))
    population, lnth = add_country_data(population, 'population')

    area = wb.create_sheet("Area", (3-1))
    area, lnth = add_country_data(area, 'area_km2')

    pop_density = wb.create_sheet("P_Density", (4-1))
    pop_density, lnth = add_country_data(pop_density, 'population_km2')

    cols = ['A','B','C','D','E','F','G','H','I','J','K','L']
    users = wb.create_sheet("Users", (5-1))
    users = add_users(users, cols, lnth)
    users.sheet_properties.tabColor = "66ff99"

    data_demand = wb.create_sheet("Data", (6-1))
    data_demand = add_data_demand(data_demand, cols, lnth)
    data_demand.sheet_properties.tabColor = "66ff99"

    lookups = wb.create_sheet("Lookups", (7-1))
    lookups = add_lookups_sheet(lookups)
    lookups.sheet_properties.tabColor = "0000ff"

    coverage = wb.create_sheet("Coverage", (8-1))
    coverage = add_coverage_sheet(coverage, cols)
    coverage.sheet_properties.tabColor = "0000ff"

    towers = wb.create_sheet("Towers", (9-1))
    towers = add_towers_sheet(towers, cols, lnth)
    towers.sheet_properties.tabColor = "0000ff"

    capacity = wb.create_sheet("Capacity", (10-1))
    capacity = add_capacity_sheet(capacity, cols, lnth)
    capacity.sheet_properties.tabColor = "0000ff"

    sites = wb.create_sheet("Total Sites", (11-1))
    sites = add_sites_sheet(sites, cols, lnth)
    sites.sheet_properties.tabColor = "0000ff"

    new = wb.create_sheet("New Sites", (12-1))
    new = add_new_sites_sheet(new, cols, lnth)
    new.sheet_properties.tabColor = "0000ff"

    costs = wb.create_sheet("Costs", (13-1))
    costs = add_cost_sheet(costs, cols, lnth)

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
    ws.column_dimensions['E'].width = 25
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
    set_border(ws, 'A6:B9', "thin", "000000")
    ws.merge_cells('A6:B6')
    ws['A6'] = "Time Parameters"

    ws['A7'] = "Parameter"
    ws['B7'] = "Option"

    ws['A8'] = "Start Year"
    ws['B8'] = 2022

    ws['A9'] = "End Year"
    ws['B9'] = 2030

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

    set_border(ws, 'G6:H10', "thin", "000000")
    ws.merge_cells('G6:H6')
    ws['G6'] = "Supply Parameters"

    ws['G7'] = "Parameter"
    ws['H7'] = "Option"

    ws['G8'] = "Infrastructure Strategy"
    data_val = DataValidation(type="list", formula1='=Options!D2:D6')
    ws.add_data_validation(data_val)
    data_val.add(ws["H8"])
    ws['H8'] = "4G"

    ws['G9'] = "Existing Spectrum Availability"
    data_val = DataValidation(type="list", formula1='=Options!E2:E4')
    ws.add_data_validation(data_val)
    data_val.add(ws["H9"])
    ws['H9'] = "Baseline"

    ws['G10'] = "Future Spectrum Availability"
    data_val = DataValidation(type="list", formula1='=Options!E2:E4')
    ws.add_data_validation(data_val)
    data_val.add(ws["H10"])
    ws['H10'] = "Baseline"

    #Cross Country Comparisons
    ws.merge_cells('A15:E15')
    ws['A15'] = "Cross-Country Comparisons"
    ws['A16'] = "Country"
    ws['B16'] = "ISO3"
    ws['C16'] = "Total Cost"
    ws['D16'] = "Mean Annual 10-Year GDP"
    ws['E16'] = "GDP Percentage"

    ### Country 1
    data_val = DataValidation(type="list", formula1='=Options!A2:A251')
    ws.add_data_validation(data_val)
    data_val.add(ws["A17"])
    ws['A17'] = "Afghanistan"
    ws['B17'] = "=IFERROR(INDEX(Options!B2:B1611,MATCH(Settings!A17, Options!A2:A1611)), \"\")"
    ws['C17'] = "=IFERROR(INDEX(Costs!M2:M1611,MATCH(Settings!B17, Costs!A2:A1611)), \"\")"
    ws['D17'] = "=IFERROR(INDEX(GDP!L2:L1611,MATCH(Settings!B17, GDP!A2:A1611)), \"\")"
    ws['E17'] = "=(C17/D17)*100"

    ### Country 2
    data_val = DataValidation(type="list", formula1='=Options!A2:A251')
    ws.add_data_validation(data_val)
    data_val.add(ws["A18"])
    ws['A18'] = "Bhutan"
    ws['B18'] = "=IFERROR(INDEX(Options!B2:B1611,MATCH(Settings!A18, Options!A2:A1611)), \"\")"
    ws['C18'] = "=IFERROR(INDEX(Costs!M2:M1611,MATCH(Settings!B18, Costs!A2:A1611)), \"\")"
    ws['D18'] = "=IFERROR(INDEX(GDP!L2:L1611,MATCH(Settings!B18, GDP!A2:A1611)), \"\")"
    ws['E18'] = "=(C18/D18)*100"

    ### Country 3
    data_val = DataValidation(type="list", formula1='=Options!A2:A251')
    ws.add_data_validation(data_val)
    data_val.add(ws["A19"])
    ws['A19'] = "Bangladesh"
    ws['B19'] = "=IFERROR(INDEX(Options!B2:B1611,MATCH(Settings!A19, Options!A2:A1611)), \"\")"
    ws['C19'] = "=IFERROR(INDEX(Costs!M2:M1611,MATCH(Settings!B19, Costs!A2:A1611)), \"\")"
    ws['D19'] = "=IFERROR(INDEX(GDP!L2:L1611,MATCH(Settings!B19, GDP!A2:A1611)), \"\")"
    ws['E19'] = "=(C19/D19)*100"

    ### Country 4
    data_val = DataValidation(type="list", formula1='=Options!A2:A251')
    ws.add_data_validation(data_val)
    data_val.add(ws["A20"])
    ws['A20'] = "India"
    ws['B20'] = "=IFERROR(INDEX(Options!B2:B1611,MATCH(Settings!A20, Options!A2:A1611)), \"\")"
    ws['C20'] = "=IFERROR(INDEX(Costs!M2:M1611,MATCH(Settings!B20, Costs!A2:A1611)), \"\")"
    ws['D20'] = "=IFERROR(INDEX(GDP!L2:L1611,MATCH(Settings!B20, GDP!A2:A1611)), \"\")"
    ws['E20'] = "=(C20/D20)*100"

    set_border(ws, 'A15:E20', "thin", "000000")


    ### Costs by Income Group
    ws.merge_cells('A23:E23')
    ws['A23'] = "Cost by Income Group"
    ws.merge_cells('A24:B24')
    ws['A24'] = "Income Group"
    ws['C24'] = "Total Cost ($)"
    ws['D24'] = "Mean Annual 10-Year GDP ($)"
    ws['E24'] = "Cost as GDP Percentage"

    ws.merge_cells('A25:B25')
    ws['A25'] = 'Advanced Economies'
    ws['C25'] = '=SUMIF(Costs!$O$2:$O$250,Settings!A25,Costs!$M$2:$M$250)'
    ws['D25'] = '=SUMIF(GDP!$M$2:$M$250,Settings!A25,GDP!$L$2:$L$250)'
    ws['E25'] = "=(C25/D25)*100"

    ws.merge_cells('A26:B26')
    ws['A26'] = 'Emerging Market Economies'
    ws['C26'] = '=SUMIF(Costs!$O$2:$O$250,Settings!A26,Costs!$M$2:$M$250)'
    ws['D26'] = '=SUMIF(GDP!$M$2:$M$250,Settings!A26,GDP!$L$2:$L$250)'
    ws['E26'] = "=(C26/D26)*100"

    ws.merge_cells('A27:B27')
    ws['A27'] = 'Low Income Developing Countries'
    ws['C27'] = '=SUMIF(Costs!$O$2:$O$250,Settings!A27,Costs!$M$2:$M$250)'
    ws['D27'] = '=SUMIF(GDP!$M$2:$M$250,Settings!A27,GDP!$L$2:$L$250)'
    ws['E27'] = "=(C27/D27)*100"

    set_border(ws, 'A23:E27', "thin", "000000")

    ### Costs by Income Group
    ws.merge_cells('A30:E30')
    ws['A30'] = "Cost by Region"
    ws.merge_cells('A31:B31')
    ws['A31'] = "Region"
    ws['C31'] = "Total Cost ($)"
    ws['D31'] = "Mean Annual 10-Year GDP ($)"
    ws['E31'] = "Cost as GDP Percentage"

    ws.merge_cells('A32:B32')
    ws['A32'] = 'Advanced Economies'
    ws['C32'] = '=SUMIF(Costs!$P$2:$P$250,Settings!A32,Costs!$M$2:$M$250)'
    ws['D32'] = '=SUMIF(GDP!$N$2:$N$250,Settings!A32,GDP!$L$2:$L$250)'
    ws['E32'] = "=(C32/D32)*100"

    ws.merge_cells('A33:B33')
    ws['A33'] = "Caucasus and Central Asia"
    ws['C33'] = '=SUMIF(Costs!$P$2:$P$250,Settings!A33,Costs!$M$2:$M$250)'
    ws['D33'] = '=SUMIF(GDP!$N$2:$N$250,Settings!A33,GDP!$L$2:$L$250)'
    ws['E33'] = "=(C33/D33)*100"

    ws.merge_cells('A34:B34')
    ws['A34'] = "Emerging and Developing Asia"
    ws['C34'] = '=SUMIF(Costs!$P$2:$P$250,Settings!A34,Costs!$M$2:$M$250)'
    ws['D34'] = '=SUMIF(GDP!$N$2:$N$250,Settings!A34,GDP!$L$2:$L$250)'
    ws['E34'] = "=(C34/D34)*100"

    ws.merge_cells('A35:B35')
    ws['A35'] = "Emerging and Developing Europe"
    ws['C35'] = '=SUMIF(Costs!$P$2:$P$250,Settings!A35,Costs!$M$2:$M$250)'
    ws['D35'] = '=SUMIF(GDP!$N$2:$N$250,Settings!A35,GDP!$L$2:$L$250)'
    ws['E35'] = "=(C35/D35)*100"

    ws.merge_cells('A36:B36')
    ws['A36'] = "Latin America and the Caribbean"
    ws['C36'] = '=SUMIF(Costs!$P$2:$P$250,Settings!A36,Costs!$M$2:$M$250)'
    ws['D36'] = '=SUMIF(GDP!$N$2:$N$250,Settings!A36,GDP!$L$2:$L$250)'
    ws['E36'] = "=(C36/D36)*100"

    ws.merge_cells('A37:B37')
    ws['A37'] = "Middle East, North Africa, Afghanistan, and Pakistan"
    ws['C37'] = '=SUMIF(Costs!$P$2:$P$250,Settings!A37,Costs!$M$2:$M$250)'
    ws['D37'] = '=SUMIF(GDP!$N$2:$N$250,Settings!A37,GDP!$L$2:$L$250)'
    ws['E37'] = "=(C37/D37)*100"

    ws.merge_cells('A38:B38')
    ws['A38'] = "Sub-Sahara Africa"
    ws['C38'] = '=SUMIF(Costs!$P$2:$P$250,Settings!A38,Costs!$M$2:$M$250)'
    ws['D38'] = '=SUMIF(GDP!$N$2:$N$250,Settings!A38,GDP!$L$2:$L$250)'
    ws['E38'] = "=(C38/D38)*100"

    set_border(ws, 'A30:E38', "thin", "000000")

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

    path = os.path.join(BASE_PATH, 'global_information.csv')
    data = pd.read_csv(path, encoding = "ISO-8859-1")

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
    ws['D2'] = "4G" #"4G (Wireless)"
    # ws['D3'] = "4G (Fiber)"
    # ws['D4'] = "5G (Wireless)"
    # ws['D5'] = "5G (Fiber)"

    ws['E1'] = "Spectrum"
    ws['E2'] = "High"
    ws['E3'] = "Baseline"
    ws['E4'] = "Low"

    path = os.path.join(DATA_RAW, 'imf_gdp_2020_2030_real.csv')
    country_info = pd.read_csv(path, encoding = "ISO-8859-1")
    country_info.rename(columns={'isocode':'ISO3'}, inplace=True)

    country_info = country_info[[
        'ISO3',
        'ifscode',
        'region',
        'income'
    ]]

    my_list = [
        ('G', 'ISO3'),
        ('H', 'ifscode'),
        ('I', 'region'),
        ('J', 'income'),
    ]

    for item in my_list:
        col = item[0]
        metric = item[1]
        for idx, row in country_info.iterrows():

            if idx == 0:
                col_name_cell = '{}1'.format(col)
                ws[col_name_cell] = metric

            cell = '{}{}'.format(col, idx+2)
            ws[cell] = row[metric]
            ws.column_dimensions[col].width = 10

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
        'GID_0': 'ISO3',
        }, axis='columns')

    lnth = len(data)+2

    for r in dataframe_to_rows(data, index=False, header=True):
        ws.append(r)

    if metric == 'population':
        ws['M1'] = 'pop_sum'
        for i in range(2, lnth):
            cell = 'M{}'.format(i)
            ws[cell] = "=SUM(C{}:L{})".format(i,i)

    if metric == 'area_km2':
        ws['M1'] = 'area_km2_sum'
        for i in range(2, lnth):
            cell = 'M{}'.format(i)
            ws[cell] = "=SUM(C{}:L{})".format(i,i)

    return ws, lnth


def add_users(ws, cols, lnth):
    """
    Calculate the total number of users.

    """
    for col in cols:
            cell = "{}1".format(col)
            ws[cell] = "='P_Density'!{}".format(cell)

    for col in cols[:2]:
        for i in range(2, lnth):
            cell = "{}{}".format(col, i)
            ws[cell] = "='P_Density'!{}".format(cell)

    for col in cols[2:]:
        for i in range(2, lnth): #Decile
            cell = "{}{}".format(col, i)
            part1 = "='P_Density'!{}".format(cell) #
            part2 = "*(POWER(1+(Settings!$E$8/100), Settings!$B$11-Settings!$B$10))"
            part3 = "*(Settings!$E$9/100)*(Settings!$E$10/100)*(Settings!$E$11/100)"
            ws[cell] = part1 + part2 + part3

    return ws


def add_data_demand(ws, cols, lnth):
    """
    Calculate the total data demand.

    """
    for col in cols:
            cell = "{}1".format(col)
            ws[cell] = "='Users'!{}".format(cell)

    for col in cols[:2]:
        for i in range(2, lnth):
            cell = "{}{}".format(col, i)
            ws[cell] = "='Users'!{}".format(cell)

    for col in cols[2:]:
        for i in range(2, lnth): #Decile
            cell = "{}{}".format(col, i)
            part1 = "=(Users!{}*MAX(Settings!$E$12,".format(cell)
            part2 = "(Settings!$E$13*1024*8*(1/30)*(15/100)*(1/3600))))"
            ws[cell] = part1 + part2

    return ws


def add_coverage_sheet(ws, cols):
    """

    """
    path = os.path.join(DATA_RAW, 'gsma_3g_coverage.csv')
    coverage = pd.read_csv(path, encoding = "ISO-8859-1")
    coverage['coverage'] = coverage['coverage'] * 100
    coverage = coverage.sort_values('ISO3')
    coverage = coverage.dropna()
    coverage = coverage[['ISO3', 'country_name', 'coverage']]

    path = os.path.join(DATA_RAW, 'site_counts', 'site_counts.csv')
    sites = pd.read_csv(path, encoding = "ISO-8859-1")
    sites = sites[['ISO3', 'sites']]

    coverage = coverage.merge(sites, left_on='ISO3', right_on='ISO3')

    for r in dataframe_to_rows(coverage, index=False, header=True):
        ws.append(r)

    lnth = len(coverage)+2

    ws['E1'] = 'Population'
    for i in range(2, lnth):
        cell = "E{}".format(i)
        ws[cell] = "=IFERROR(INDEX(Pop!$M$2:$M$1611,MATCH(A{}, Pop!$A$2:$A$1611,0)),"")".format(i)

    ws['F1'] = 'Sites per covered population'
    for i in range(2, lnth):
        cell = "F{}".format(i)
        ws[cell] = '=IFERROR(E{}*(C{}/100)/D{}, "-")'.format(i,i,i)

    return ws#, lnth


def add_towers_sheet(ws, cols, lnth):
    """

    """
    for col in cols:
        cell = "{}1".format(col)
        ws[cell] = "='Data'!{}".format(cell)

    for col in cols[:2]:
        for i in range(1, lnth):
            cell = "{}{}".format(col, i)
            ws[cell] = "='Data'!{}".format(cell)

    for i in range(2, lnth): #Decile
        cell = "C{}".format(i)
        part1 = "=IFERROR(INDEX(Pop!$C$2:$C$1611,MATCH(A{}, Pop!$A$2:$A$1611,0)) /".format(i)
        part2 = "INDEX(Coverage!$F$2:$F$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),0)".format(i)
        ws[cell] = part1 + part2 #+ part3

    for i in range(2, lnth):
        cell = "D{}".format(i)
        part1 = "=IF(SUM(C{})<INDEX(Coverage!$D$2:$D$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),".format(i,i)
        part2 = "INDEX(Pop!$D$2:$D$1611,MATCH(A{}, Pop!$A$2:$A$1611,0))/".format(i)
        part3 = "INDEX(Coverage!$F$2:$F$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),0)".format(i)
        ws[cell] = part1 + part2 + part3

    for i in range(2, lnth):
        cell = "E{}".format(i)
        part1 = "=IF(SUM(C{}:D{})<INDEX(Coverage!$D$2:$D$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),".format(i,i,i)
        part2 = "INDEX(Pop!$E$2:$E$1611,MATCH(A{}, Pop!$A$2:$A$1611,0))/".format(i)
        part3 = "INDEX(Coverage!$F$2:$F$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),0)".format(i)
        ws[cell] = part1 + part2 + part3

    for i in range(2, lnth):
        cell = "F{}".format(i)
        part1 = "=IF(SUM(C{}:E{})<INDEX(Coverage!$D$2:$D$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),".format(i,i,i)
        part2 = "INDEX(Pop!$F$2:$F$1611,MATCH(A{}, Pop!$A$2:$A$1611,0))/".format(i)
        part3 = "INDEX(Coverage!$F$2:$F$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),0)".format(i)
        ws[cell] = part1 + part2 + part3

    for i in range(2, lnth):
        cell = "G{}".format(i)
        part1 = "=IF(SUM(C{}:F{})<INDEX(Coverage!$D$2:$D$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),".format(i,i,i)
        part2 = "INDEX(Pop!$G$2:$G$1611,MATCH(A{}, Pop!$A$2:$A$1611,0))/".format(i)
        part3 = "INDEX(Coverage!$F$2:$F$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),0)".format(i)
        ws[cell] = part1 + part2 + part3

    for i in range(2, lnth):
        cell = "H{}".format(i)
        part1 = "=IF(SUM(C{}:G{})<INDEX(Coverage!$D$2:$D$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),".format(i,i,i)
        part2 = "INDEX(Pop!$H$2:$H$1611,MATCH(A{}, Pop!$A$2:$A$1611,0))/".format(i)
        part3 = "INDEX(Coverage!$F$2:$F$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),0)".format(i)
        ws[cell] = part1 + part2 + part3

    for i in range(2, lnth):
        cell = "I{}".format(i)
        part1 = "=IF(SUM(C{}:H{})<INDEX(Coverage!$D$2:$D$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),".format(i,i,i)
        part2 = "INDEX(Pop!$I$2:$I$1611,MATCH(A{}, Pop!$A$2:$A$1611,0))/".format(i)
        part3 = "INDEX(Coverage!$F$2:$F$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),0)".format(i)
        ws[cell] = part1 + part2 + part3

    for i in range(2, lnth):
        cell = "J{}".format(i)
        part1 = "=IF(SUM(C{}:I{})<INDEX(Coverage!$D$2:$D$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),".format(i,i,i)
        part2 = "INDEX(Pop!$J$2:$J$1611,MATCH(A{}, Pop!$A$2:$A$1611,0))/".format(i)
        part3 = "INDEX(Coverage!$F$2:$F$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),0)".format(i)
        ws[cell] = part1 + part2 + part3

    for i in range(2, lnth):
        cell = "K{}".format(i)
        part1 = "=IF(SUM(C{}:J{})<INDEX(Coverage!$D$2:$D$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),".format(i,i,i)
        part2 = "INDEX(Pop!$K$2:$K$1611,MATCH(A{}, Pop!$A$2:$A$1611,0))/".format(i)
        part3 = "INDEX(Coverage!$F$2:$F$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),0)".format(i)
        ws[cell] = part1 + part2 + part3

    for i in range(2, lnth):
        cell = "L{}".format(i)
        part1 = "=IF(SUM(C{}:K{})<INDEX(Coverage!$D$2:$D$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),".format(i,i,i)
        part2 = "INDEX(Pop!$L$2:$L$1611,MATCH(A{}, Pop!$A$2:$A$1611,0))/".format(i)
        part3 = "INDEX(Coverage!$F$2:$F$1611,MATCH(A{}, Coverage!$A$2:$A$1611,0)),0)".format(i)
        ws[cell] = part1 + part2 + part3

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
    ws['B3'] = 30
    ws['B4'] = 20
    ws['B5'] = 10

    set_border(ws, 'A1:B5', "thin", "000000")
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 12

    ##Cost Information
    set_border(ws, 'A7:C11', "thin", "000000")
    ws.merge_cells('A7:C7')
    ws['A7'] = "Equipment Costs"

    ws['A8'] = 'Asset'
    ws['A9'] = 'RAN'
    ws['A10'] = 'Fiber'
    ws['A11'] = 'Towers'

    ws['B8'] = 'Cost ($)'
    ws['B9'] = 30000
    ws['B10'] = 10
    ws['B11'] = 30000

    ws['C8'] = 'Unit'
    ws['C9'] = "Per Tower"
    ws['C10'] = "Per Meter"
    ws['C11'] = "Per Tower"

    ##Density Lookup Table
    ws.merge_cells('E1:H1')
    ws['E1'] = "Density Lookup Table"

    filename = 'capacity_lut_by_frequency.csv'
    path = os.path.join(DATA_INTERMEDIATE, 'luts', filename)
    lookup = pd.read_csv(path)
    lookup.loc[lookup['frequency_GHz'] == '0.8']
    lookup = lookup[['sites_per_km2', 'capacity_mbps_km2']]
    lookup = lookup[lookup['capacity_mbps_km2'] != 0].reset_index()
    df_length = len(lookup)
    lookup = lookup.sort_values('sites_per_km2')

    my_list = [
        ('E', 'sites_per_km2'),
        ('F', 'capacity_mbps_km2'),
    ]
    for item in my_list:
        col = item[0]
        metric = item[1]
        for idx, row in lookup.iterrows():

            if idx == 0:
                col_name_cell = '{}2'.format(col)
                ws[col_name_cell] = metric

            cell = '{}{}'.format(col, idx+3)
            ws[cell] = row[metric]
            ws.column_dimensions[col].width = 18

    ws['G2'] = 'w_existing_spectrum'
    for i in range(3,250):
        col = 'G{}'.format(i)
        ws[col] = '=F{}*(VLOOKUP(Settings!$H$9, Lookups!$A$3:$B$5, 2, 0)/10)'.format(i)
    ws.column_dimensions['G'].width = 20

    ws['H2'] = 'w_future_spectrum'
    for i in range(3,250):
        col = 'H{}'.format(i)
        ws[col] = '=F{}*(VLOOKUP(Settings!$H$10, Lookups!$A$3:$B$5, 2, 0)/10)'.format(i)
    ws.column_dimensions['H'].width = 20

    set_border(ws, 'E1:H{}'.format(df_length+2), "thin", "000000")

    return ws


def add_capacity_sheet(ws, cols, lnth):
    """

    """
    # set_border(ws, 'A1:AZ1000', "thin", "00FFFFFF")

    for col in cols:
            cell = "{}1".format(col)
            ws[cell] = "='Data'!{}".format(cell)

    for col in cols[:2]:
        for i in range(2, lnth):
            cell = "{}{}".format(col, i)
            ws[cell] = "='Data'!{}".format(cell)

    for col in cols[2:]:
        for i in range(2, lnth): #Total Sites Density
            cell = "{}{}".format(col, i)
            part1 = '=IFERROR(MAX(IF(Lookups!$E$3:$E$250<Towers!{}/Area!{}'.format(cell, cell)
            part2 = '*(Settings!$E$10/100),Lookups!$GL$3:$G$250)),"-")'
            ws[cell] = part1 + part2
            ws.formula_attributes[cell] = {'t': 'array', 'ref': "{}:{}".format(cell, cell)}

    return ws


def add_sites_sheet(ws, cols, lnth):
    """

    """
    # set_border(ws, 'A1:AZ1000', "thin", "00FFFFFF")

    for col in cols:
            cell = "{}1".format(col)
            ws[cell] = "='Capacity'!{}".format(cell)

    for col in cols[:2]:
        for i in range(2, lnth):
            cell = "{}{}".format(col, i)
            ws[cell] = "='Capacity'!{}".format(cell)

    for col in cols[2:]:
        for i in range(2, lnth):
            cell = "{}{}".format(col, i)
            part1 = "=MIN(IF('Lookups'!$H$3:$H$250>'Data'!{}".format(cell)
            part2 = ",'Lookups'!$E$3:$E$250))*Area!{}".format(cell)
            ws[cell] = part1 + part2
            ws.formula_attributes[cell] = {'t': 'array', 'ref': "{}:{}".format(cell, cell)}

    return ws


def add_new_sites_sheet(ws, cols, lnth):
    """

    """
    # ##Color white
    # set_border(ws, 'A1:AZ1000', "thin", "00FFFFFF")

    for col in cols:
            cell = "{}1".format(col)
            ws[cell] = "='Total Sites'!{}".format(cell)

    for col in cols[:2]:
        for i in range(2, lnth):
            cell = "{}{}".format(col, i)
            ws[cell] = "='Total Sites'!{}".format(cell)

    for col in cols[2:]:
        for i in range(2, lnth):
            cell = "{}{}".format(col, i)
            part1 = "=IF(Towers!{}*(Settings!$E$10/100)<'Total Sites'!{},".format(cell, cell)
            part2 = "('Total Sites'!{}-Towers!{}*(Settings!$E$10/100)),0)".format(cell, cell)
            ws[cell] = part1 + part2

    return ws


def add_cost_sheet(ws, cols, lnth):
    """

    """
    ws.sheet_properties.tabColor = "9966ff"

    #Deciles
    # set_border(ws, 'A1:J11', "thin", "000000")

    for col in cols:
            cell = "{}1".format(col)
            ws[cell] = "='New Sites'!{}".format(cell)

    for col in cols[:2]:
        for i in range(2, lnth):
            cell = "{}{}".format(col, i)
            ws[cell] = "='New Sites'!{}".format(cell)

    for col in cols[2:]:
        for i in range(2, lnth):
            cell = "{}{}".format(col, i)
            part1 = "='New Sites'!{}*VLOOKUP('Lookups'!$A$9, 'Lookups'!$A$9:'Lookups'!$B$24, 2, FALSE)".format(cell)
            part2 = "+'New Sites'!{}*VLOOKUP('Lookups'!$A$10, 'Lookups'!$A$9:'Lookups'!$B$24, 2, FALSE)".format(cell)
            part3 = "+'New Sites'!{}*VLOOKUP('Lookups'!$A$11, 'Lookups'!$A$9:'Lookups'!$B$24, 2, FALSE)".format(cell)
            ws[cell] = part1 + part2 + part3

    ws['M1'] = 'Total Cost ($)'
    for i in range(2,lnth):
        cell = "M{}".format(i)
        part1 = '=IFERROR(SUMIF(C{}:L{}, "<>n/a"), "-")'.format(i, i)
        line = part1
        ws[cell] = line

    ws['N1'] = 'Cost Per User ($)'
    for i in range(2,lnth):
        cell = "N{}".format(i)
        ws[cell] = "=M{}/Pop!M{}".format(i, i)

    ws['O1'] = 'Income Group'
    for i in range(2,lnth):
        cell = "O{}".format(i)
        ws[cell] = "=IFERROR(INDEX(Options!$J$2:$J$1611,MATCH(A{}, Options!$G$2:$G$1611,0)), "")".format(i)

    ws['P1'] = 'Region'
    for i in range(2,lnth):
        cell = "P{}".format(i)
        ws[cell] = "=IFERROR(INDEX(Options!$I$2:$I$1611,MATCH(A{}, Options!$G$2:$G$1611,0)), "")".format(i)

    ws.column_dimensions['M'].width = 15
    ws.column_dimensions['N'].width = 15
    ws.column_dimensions['O'].width = 30
    ws.column_dimensions['P'].width = 30

    return ws


def add_gdp_sheet(ws):
    """

    """
    ws.sheet_properties.tabColor = "ffff33"

    path = os.path.join(DATA_RAW, 'imf_gdp_2020_2030_real.csv')
    gdp = pd.read_csv(path, encoding = "ISO-8859-1")

    gdp.rename(columns={'isocode':'ISO3'}, inplace=True)

    for i in range(2020, 2031):
        col = "GDP{}".format(i)
        if col in gdp.columns:
            gdp.rename(columns={col:i}, inplace=True)

    gdp = gdp[['ISO3',2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030]]

    gdp = gdp.sort_values('ISO3')

    lnth = len(gdp)+2

    for r in dataframe_to_rows(gdp, index=False, header=True):
        ws.append(r)

    ws['L1'] = 'Mean 10-Year GDP ($)'
    for i in range(2,lnth):
        cell = 'L{}'.format(i)
        ws[cell] = "=(SUM(B{}:K{})*1e9)/10".format(i,i)

    ws['M1'] = 'Income Group'
    for i in range(2,lnth):
        cell = "M{}".format(i)
        ws[cell] = "=IFERROR(INDEX(Options!$J$2:$J$1611,MATCH(A{}, Options!$G$2:$G$1611,0)), "")".format(i)

    ws['N1'] = 'Region'
    for i in range(2,lnth):
        cell = "N{}".format(i)
        ws[cell] = "=IFERROR(INDEX(Options!$I$2:$I$1611,MATCH(A{}, Options!$G$2:$G$1611,0)), "")".format(i)

    ws.column_dimensions['L'].width = 20
    ws.column_dimensions['M'].width = 30
    ws.column_dimensions['N'].width = 30

    return ws


if __name__ == "__main__":

    generate_workbook()
