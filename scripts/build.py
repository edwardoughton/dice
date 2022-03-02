"""
Build the DICE workbook.

"""
# import argparse
import os
import configparser
import pandas as pd
# import openpyxl
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.datavalidation import DataValidation
# from openpyxl.reader.excel import load_workbook
from openpyxl.styles import Border, Side, Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from openpyxl.chart import BarChart, Reference  #, Series
from openpyxl.chart.label import DataLabelList

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

    index = wb.active
    index = add_index(index)

    readme = wb.create_sheet("Read_Me", (2-1))
    readme = add_readme(readme)

    settings = wb.create_sheet("Settings", (3-1))
    settings = add_settings(settings)

    countries = wb.create_sheet("Countries", (4-1))
    countries = add_country_selection(countries)

    estimates = wb.create_sheet("Estimates", (5-1))
    estimates = add_estimates(estimates)

    population = wb.create_sheet("Pop", (6-1))
    population, lnth = add_country_data(population, 'population')

    area = wb.create_sheet("Area", (7-1))
    area, lnth = add_country_data(area, 'area_km2')

    pop_density = wb.create_sheet("P_Density", (8-1))
    pop_density, lnth = add_country_data(pop_density, 'population_km2')

    pop_growth = wb.create_sheet("P_Growth", (9-1))
    pop_growth, lnth2 = add_pop_growth(pop_growth)

    cols = ['A','B','C','D','E','F','G','H','I','J','K','L']
    users = wb.create_sheet("Users", (10-1))
    users = add_users(users, cols, lnth)
    users.sheet_properties.tabColor = "66ff99"

    data_demand = wb.create_sheet("Data", (11-1))
    data_demand = add_data_demand(data_demand, cols, lnth)
    data_demand.sheet_properties.tabColor = "66ff99"

    lookups = wb.create_sheet("Lookups", (12-1))
    lookups = add_lookups_sheet(lookups)
    lookups.sheet_properties.tabColor = "0000ff"

    coverage = wb.create_sheet("Coverage", (13-1))
    coverage = add_coverage_sheet(coverage, cols)
    coverage.sheet_properties.tabColor = "0000ff"

    towers = wb.create_sheet("Towers", (14-1))
    towers = add_towers_sheet(towers, cols, lnth)
    towers.sheet_properties.tabColor = "0000ff"

    capacity = wb.create_sheet("Capacity", (15-1))
    capacity = add_capacity_sheet(capacity, cols, lnth)
    capacity.sheet_properties.tabColor = "0000ff"

    sites = wb.create_sheet("Total Sites", (16-1))
    sites = add_sites_sheet(sites, cols, lnth)
    sites.sheet_properties.tabColor = "0000ff"

    new = wb.create_sheet("New Sites", (17-1))
    new = add_new_sites_sheet(new, cols, lnth)
    new.sheet_properties.tabColor = "0000ff"

    costs = wb.create_sheet("Costs", (18-1))
    costs = add_cost_sheet(costs, cols, lnth)

    gdp = wb.create_sheet("GDP", (19-1))
    gdp = add_gdp_sheet(gdp)

    options = wb.create_sheet("Options", (20-1))
    options = add_options(options)

    context = wb.create_sheet("Context", (21-1))
    add_context(context, estimates)

    # estimates.state = 'hidden'
    population.sheet_state  = 'hidden'
    area.sheet_state = 'hidden'
    pop_density.sheet_state = 'hidden'
    pop_growth.sheet_state = 'hidden'
    users.sheet_state = 'hidden'
    data_demand.sheet_state = 'hidden'
    lookups.sheet_state = 'hidden'
    coverage.sheet_state = 'hidden'
    towers.sheet_state = 'hidden'
    capacity.sheet_state = 'hidden'
    sites.sheet_state = 'hidden'
    new.sheet_state = 'hidden'
    costs.sheet_state = 'hidden'
    gdp.sheet_state = 'hidden'
    options.sheet_state = 'hidden'

    wb.save('Oughton et al. (2022) DICE.xlsx')

    return print("Generated workbook")


def add_index(ws):
    """
    Add the welcome index page.

    """
    ws.title = "Index"
    ws.sheet_properties.tabColor = "004C97"

    ws = set_font(ws, 'A1:AZ1000', 'Segoe UI')

    #Set height of row one
    ws.row_dimensions[1].height = 34.5
    ws.row_dimensions[2].height = 14.25

    ##Color white
    ws.sheet_view.showGridLines = False
    # set_border(ws, 'A1:AZ1000', "thin", "00FFFFFF")

    #Set blue and red border strips
    set_cell_color(ws, 'A1:AZ1', "004C97")
    set_cell_color(ws, 'A2:AZ2', "C00000")

    ws['B4'] = "IMF FAD Digital Infrastructure Costing Estimator (DICE)"
    ws['B4'].font = Font(size=20)
    ws['B5'] = "Developed by George Mason University and the Fiscal Affairs Department (FAD), International Monetary Fund (IMF)"

    path = os.path.join(BASE_PATH, '..', 'images', 'imf_logo.png')
    img = Image(path)
    img.height = 150 # insert image height in pixels as float or int (e.g. 305.5)
    img.width = 150# insert image width in pixels as float or int (e.g. 405.8)
    img.anchor = 'B7'
    ws.add_image(img)

    ws['B15'] = "Contact: Edward Oughton (eoughton@gmu.edu); David Amaglobeli (damaglobeli@imf.org); Mariano Moszoro (mmoszoro@imf.org)."
    ws['B16'] = '=HYPERLINK("{}")'.format('https://github.com/edwardoughton/dice')
    hyperlink = Font(underline='single', color='0563C1')
    ws['B16'].font = hyperlink

    ws['B18'] = "Capabilities/Uses:"
    ws['B18'].font = Font(bold=True)
    ws['B19'] = """    - Supporting high-level decisions pertaining to universal broadband investment by capturing the main cost drivers which affect the deployment of broadband infrastructure."""
    ws['B20'] = "    - Breaking down investment by specific country income groups and regions."

    ws['B22'] = "Caveats to using DICE:"
    ws['B22'].font = Font(bold=True)
    ws['B23'] = """    - DICE is not a replacement for detailed country-specific modeling. The aim is to provide comparative understanding across all countries globally."""
    ws['B24'] = """    - DICE is not an exact measurement tool. Accuracy and precision are commensurate with being able to make global cross-country comparisons."""
    ws['B25'] = """    - DICE is not a stand-alone policy option tool. If deep subject matter exertise are required, country teams should reach out to the DICE development team at GMU and the IMF (eoughton@gmu.edu; damaglobeli@imf.org; mmoszoro@imf.org)."""

    ws['B27'] = "Main Datasets:"
    ws['B27'].font = Font(bold=True)
    ws['B28'] = "WorldPop 2020 Unconstrained 1km Global Mosaic, UN Population Forecasts 2020-2030, IMF GDP Projections 2020-2030"

    ws['B30'] = "Reference:"
    ws['B30'].font = Font(bold=True)
    ws['B31'] = "    - Working Paper Citation TBC."
    return ws


def add_readme(ws):
    """

    """
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "004C97"
    set_cell_color(ws, 'A1:O6', "004C97")
    ws = set_font(ws, 'A1:AZ1000', 'Segoe UI')

    path = os.path.join(BASE_PATH, '..', 'images', 'imf_logo.png')
    img = Image(path)
    img.height = 133 # insert image height in pixels as float or int (e.g. 305.5)
    img.width = 133 # insert image width in pixels as float or int (e.g. 405.8)
    img.anchor = 'A1'
    ws.add_image(img)

    ws.merge_cells('A3:O4')
    ws['A3'] = "Read Me & FAQ"
    ws['A3'].font = Font(size=20, color='FFFFFF')
    ws['A3'].alignment = Alignment(horizontal='center')

    return ws


def add_settings(ws):
    """

    """
    ##Color white
    ws.sheet_view.showGridLines = False
    #Set blue and red border strips
    set_cell_color(ws, 'A1:AZ1', "004C97")
    set_cell_color(ws, 'A2:AZ2', "C00000")

    ##Allocate title
    ws.title = "Settings"
    ws.sheet_properties.tabColor = "004C97"

    # ##Set column width
    # ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 40
    # ws.column_dimensions['D'].width = 35
    # ws.column_dimensions['E'].width = 30
    # ws.column_dimensions['F'].width = 25
    # ws.column_dimensions['G'].width = 35
    # ws.column_dimensions['H'].width = 20

    # ws = format_numbers(ws, ['C','D', 'F'], (17,20), 'Comma [0]', 1)
    # ws = format_numbers(ws, ['E', 'H'], (17,20), 'Percent', 1)
    # ws = format_numbers(ws, ['C','D'], (25,27), 'Comma [0]', 1)
    # ws = format_numbers(ws, ['E'], (25,27), 'Percent', 1)
    # ws = format_numbers(ws, ['C','D'], (32,38), 'Comma [0]', 1)
    # ws = format_numbers(ws, ['E'], (32,38), 'Percent', 1)

    ## Add parameters box
    format_numbers(ws, ['C'], (11,11), 'Percent', 0)
    set_border(ws, 'B6:C11', "thin", "000000")
    ws.merge_cells('B6:C6')
    ws['B6'] = "Time Parameters"

    ws['B7'] = "Parameter"
    ws['C7'] = "Option"

    ws['B8'] = "Start Year"
    ws['C8'] = 2022

    ws['B9'] = "End Year"
    ws['C9'] = 2030

    ws['B10'] = "Total Years"
    ws['C10'] = "=C9-C8"

    ws['B11'] = "Depreciation"
    ws['C11'] = "=0.08"

    set_border(ws, 'B13:C19', "thin", "000000")
    ws.merge_cells('B13:C13')
    ws['B13'] = "Demand Parameters"

    ws['B14'] = "Parameter"
    ws['C14'] = "Option"

    ws['B15'] = "Smartphone adoption (%)"
    ws['C15'] = 95

    ws['B16'] = "Market Share (%)"
    ws['C16'] = 25

    ws['B17'] = "Active Users (%)"
    ws['C17'] = 5

    ws['B18'] = "Minimum capacity per user (Mbps)"
    ws['C18'] = 10

    ws['B19'] = "Data demand per month (GB) "
    ws['C19'] = 2

    set_border(ws, 'B21:C26', "thin", "000000")
    ws.merge_cells('B21:C21')
    ws['B21'] = "Supply Parameters"

    ws['B22'] = "Parameter"
    ws['C22'] = "Option"

    ws['B23'] = "Infrastructure Strategy"
    data_val = DataValidation(type="list", formula1='=Options!D2:D6')
    ws.add_data_validation(data_val)
    data_val.add(ws["C23"])
    ws['C23'] = "4G"

    ws['B24'] = "Existing Spectrum Availability"
    data_val = DataValidation(type="list", formula1='=Options!E2:E4')
    ws.add_data_validation(data_val)
    data_val.add(ws["C24"])
    ws['C24'] = "Baseline"

    ws['B25'] = "Future Spectrum Availability"
    data_val = DataValidation(type="list", formula1='=Options!E2:E4')
    ws.add_data_validation(data_val)
    data_val.add(ws["C25"])
    ws['C25'] = "Baseline"

    ws['B26'] = "Minimum Pop.Density to Serve (km^2)"
    ws['C26'] = 5

    #Center text
    ws = center_text(ws, 'A2:AZ1000')
    # set_border(ws, 'A1:AZ1000', "thin", "00FFFFFF")
    ws = set_font(ws, 'A1:AZ1000', 'Segoe UI')
    ws = set_bold(ws, 'B6', 'Segoe UI')
    ws = set_bold(ws, 'B7', 'Segoe UI')
    ws = set_bold(ws, 'C7', 'Segoe UI')
    ws = set_bold(ws, 'B13', 'Segoe UI')
    ws = set_bold(ws, 'B14', 'Segoe UI')
    ws = set_bold(ws, 'C14', 'Segoe UI')
    ws = set_bold(ws, 'B21', 'Segoe UI')
    ws = set_bold(ws, 'B22', 'Segoe UI')
    ws = set_bold(ws, 'C22', 'Segoe UI')

    return ws


def add_country_selection(ws):
    """
    Add country selection.

    """
    ##Color white
    ws.sheet_view.showGridLines = False
    #Set blue and red border strips
    set_cell_color(ws, 'A1:AZ1', "004C97")
    set_cell_color(ws, 'A2:AZ2', "C00000")

    ##Allocate title
    ws.title = "Countries"
    ws.sheet_properties.tabColor = "004C97"

    ## Set widths
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 40

    ## Add parameters box
    format_numbers(ws, ['C'], (11,11), 'Percent', 0)
    set_border(ws, 'B6:C11', "thin", "000000")
    ws.merge_cells('B6:C6')
    ws['B6'] = "Select Countries For Comparison"
    ws['B7'] = "Country"
    ws['C7'] = "ISO3"

    ### Country 1
    data_val = DataValidation(type="list", formula1='=Options!A2:A251')
    ws.add_data_validation(data_val)
    data_val.add(ws["B8"])
    ws['B8'] = "Afghanistan"
    ws['C8'] = "=IFERROR(INDEX(Options!B2:B1611,MATCH(B8, Options!A2:A1611)), \"\")"

    # ### Country 2
    data_val = DataValidation(type="list", formula1='=Options!A2:A251')
    ws.add_data_validation(data_val)
    data_val.add(ws["B9"])
    ws['B9'] = "Bhutan"
    ws['C9'] = "=IFERROR(INDEX(Options!B2:B1611,MATCH(B9, Options!A2:A1611)), \"\")"

    ### Country 3
    data_val = DataValidation(type="list", formula1='=Options!A2:A251')
    ws.add_data_validation(data_val)
    data_val.add(ws["B10"])
    ws['B10'] = "Bangladesh"
    ws['C10'] = "=IFERROR(INDEX(Options!B2:B1611,MATCH(B10, Options!A2:A1611)), \"\")"

    ### Country 4
    data_val = DataValidation(type="list", formula1='=Options!A2:A251')
    ws.add_data_validation(data_val)
    data_val.add(ws["B11"])
    ws['B11'] = "India"
    ws['C11'] = "=IFERROR(INDEX(Options!B2:B1611,MATCH(B11, Options!A2:A1611)), \"\")"

    #Center text
    ws = center_text(ws, 'A2:AZ1000')
    # set_border(ws, 'A1:AZ1000', "thin", "00FFFFFF")
    ws = set_font(ws, 'A1:AZ1000', 'Segoe UI')
    ws = set_bold(ws, 'B6', 'Segoe UI')
    ws = set_bold(ws, 'B7', 'Segoe UI')
    ws = set_bold(ws, 'C7', 'Segoe UI')

    return ws


def add_estimates(ws):
    """

    """
    ##Color white
    ws.sheet_view.showGridLines = False
    #Set blue and red border strips
    set_cell_color(ws, 'A1:AZ1', "004C97")
    set_cell_color(ws, 'A2:AZ2', "C00000")

    ##Allocate title
    # ws.title = "Settings"
    ws.sheet_properties.tabColor = "004C97"

    #Cross Country Comparisons
    ws.merge_cells('B6:I6')
    ws['B6'] = "Cross-Country Comparisons"
    ws['B7'] = "Country"
    ws['C7'] = "ISO3"
    ws['D7'] = "Total Cost ($Bn)"
    ws['E7'] = "Mean Annual 10-Year GDP ($Bn)"
    ws['F7'] = "GDP Growth Rate (%)(2022-2030)"
    ws['G7'] = "Initial Investment ($Bn)"
    ws['H7'] = "2022 GDP ($Bn)"
    ws['I7'] = "Annual GDP Share (%)"

    # ### Country 1
    ws['B8'] = "=IFERROR(Countries!B8, \"\")"
    ws['C8'] = "=IFERROR(Countries!C8, \"\")"
    ws['D8'] = "=IFERROR(INDEX(Costs!M2:M1611,MATCH(C8, Costs!A2:A1611)), \"\")"
    ws['E8'] = "=IFERROR(INDEX(GDP!L2:L1611,MATCH(C8, GDP!A2:A1611)), \"\")"
    ws['F8'] = "=IFERROR(INDEX(GDP!M2:M1611,MATCH(C8, GDP!A2:A1611)), )"
    ws['G8'] = "=IF(F8=0,D8/Settings!C10,(D8*(1-(1+F8)/(1-Settings!C11)))/((1-Settings!C11)^Settings!C10*(1-((1+F8)/(1-Settings!C11))^(Settings!C10+1))))"
    ws['H8'] = "=IFERROR(INDEX(GDP!C2:C1611,MATCH(C8, GDP!A2:A1611)), \"\")"
    ws['I8'] = "=IFERROR(G8/H8, \"\")"

    # ### Country 2
    ws['B9'] = "=IFERROR(Countries!B9, \"\")"
    ws['C9'] = "=IFERROR(Countries!C9, \"\")"
    ws['D9'] = "=IFERROR(INDEX(Costs!M2:M1611,MATCH(C9, Costs!A2:A1611)), \"\")"
    ws['E9'] = "=IFERROR(INDEX(GDP!L2:L1611,MATCH(C9, GDP!A2:A1611)), \"\")"
    ws['F9'] = "=IFERROR(INDEX(GDP!M2:M1611,MATCH(C9, GDP!A2:A1611)), )"
    ws['G9'] = "=IF(F8=0,D9/Settings!C10,(D9*(1-(1+F9)/(1-Settings!C11)))/((1-Settings!C11)^Settings!C10*(1-((1+F9)/(1-Settings!C11))^(Settings!C10+1))))"
    ws['H9'] = "=IFERROR(INDEX(GDP!C2:C1611,MATCH(C9, GDP!A2:A1611)), \"\")"
    ws['I9'] = "=IFERROR(G9/H9, \"\")"

    # ### Country 3
    ws['B10'] = "=IFERROR(Countries!B10, \"\")"
    ws['C10'] = "=IFERROR(Countries!C10, \"\")"
    ws['D10'] = "=IFERROR(INDEX(Costs!M2:M1611,MATCH(C10, Costs!A2:A1611)), \"\")"
    ws['E10'] = "=IFERROR(INDEX(GDP!L2:L1611,MATCH(C10, GDP!A2:A1611)), \"\")"
    ws['F10'] = "=IFERROR(INDEX(GDP!M2:M1611,MATCH(C10, GDP!A2:A1611)), )"
    ws['G10'] = "=IF(F8=0,D10/Settings!C10,(D10*(1-(1+F10)/(1-Settings!C11)))/((1-Settings!C11)^Settings!C10*(1-((1+F10)/(1-Settings!C11))^(Settings!C10+1))))"
    ws['H10'] = "=IFERROR(INDEX(GDP!C2:C1611,MATCH(C10, GDP!A2:A1611)), \"\")"
    ws['I10'] = "=IFERROR(G10/H10, \"\")"

    # ### Country 4
    ws['B11'] = "=IFERROR(Countries!B11, \"\")"
    ws['C11'] = "=IFERROR(Countries!C11, \"\")"
    ws['D11'] = "=IFERROR(INDEX(Costs!M2:M1611,MATCH(C11, Costs!A2:A1611)), \"\")"
    ws['E11'] = "=IFERROR(INDEX(GDP!L2:L1611,MATCH(C11, GDP!A2:A1611)), \"\")"
    ws['F11'] = "=IFERROR(INDEX(GDP!M2:M1611,MATCH(C11, GDP!A2:A1611)), )"
    ws['G11'] = "=IF(F8=0,D11/Settings!C10,(D11*(1-(1+F11)/(1-Settings!C11)))/((1-Settings!C11)^Settings!C10*(1-((1+F11)/(1-Settings!C11))^(Settings!C10+1))))"
    ws['H11'] = "=IFERROR(INDEX(GDP!C2:C1611,MATCH(C11, GDP!A2:A1611)), \"\")"
    ws['I11'] = "=IFERROR(G11/H11, \"\")"

    # ##Set column width
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 30
    ws.column_dimensions['F'].width = 30
    ws.column_dimensions['G'].width = 30
    ws.column_dimensions['H'].width = 30
    ws.column_dimensions['I'].width = 30

    ws = format_numbers(ws, ['D', 'E', 'G', 'H'], (8,11), 'Comma [0]', 1)
    ws = format_numbers(ws, ['D', 'E', 'G', 'H'], (17,19), 'Comma [0]', 1)
    ws = format_numbers(ws, ['D', 'E', 'G', 'H'], (24,30), 'Comma [0]', 1)
    ws = format_numbers(ws, ['F'], (8,11), 'Percent', 1)
    ws = format_numbers(ws, ['F'], (17,19), 'Percent', 1)
    ws = format_numbers(ws, ['F'], (24,30), 'Percent', 1)
    ws = format_numbers(ws, ['I'], (8,11), 'Percent', 2)
    ws = format_numbers(ws, ['I'], (17,19), 'Percent', 2)
    ws = format_numbers(ws, ['I'], (24,30), 'Percent', 2)

    # ### Costs by Income Group
    ws.merge_cells('B15:I15')
    ws['B15'] = "Cost by Income Group"
    ws.merge_cells('B16:C16')
    ws['B16'] = "Income Group"
    ws['D16'] = "Total Cost ($Bn)"
    ws['E16'] = "Mean Annual 10-Year GDP ($Bn)"
    ws['F16'] = "GDP Growth Rate (%)(2022-2030)"
    ws['G16'] = "Initial Investment ($Bn)"
    ws['H16'] = "2022 GDP ($Bn)"
    ws['I16'] = "Annual GDP Share (%)"

    ws.merge_cells('B17:C17')
    ws['B17'] = 'Advanced Economies'
    ws['D17'] = '=SUMIF(Costs!$O$2:$O$250,B17,Costs!$M$2:$M$250)'
    ws['E17'] = '=SUMIF(GDP!$N$2:$N$250,B17,GDP!$L$2:$L$250)'
    ws['F17'] = '=AVERAGEIF(GDP!$N$2:$N$250,B17,GDP!$M$2:$M$250)'
    ws['G17'] = "=IF(F17=0,D17/Settings!C10,(D17*(1-(1+F17)/(1-Settings!C11)))/((1-Settings!C11)^Settings!C10*(1-((1+F17)/(1-Settings!C11))^(Settings!C10+1))))"
    ws['H17'] = "=SUMIF(GDP!$N$2:$N$250,B17,GDP!$C$2:$C$250)"
    ws['I17'] = "=IFERROR(G17/H17, \"\")"

    ws.merge_cells('B18:C18')
    ws['B18'] = 'Emerging Market Economies'
    ws['D18'] = '=SUMIF(Costs!$O$2:$O$250,B18,Costs!$M$2:$M$250)'
    ws['E18'] = '=SUMIF(GDP!$N$2:$N$250,B18,GDP!$L$2:$L$250)'
    ws['F18'] = '=AVERAGEIF(GDP!$N$2:$N$250,B18,GDP!$M$2:$M$250)'
    ws['G18'] = "=IF(F18=0,D18/Settings!C10,(D18*(1-(1+F18)/(1-Settings!C11)))/((1-Settings!C11)^Settings!C10*(1-((1+F17)/(1-Settings!C11))^(Settings!C10+1))))"
    ws['H18'] = "=SUMIF(GDP!$N$2:$N$250,B18,GDP!$C$2:$C$250)"
    ws['I18'] = "=IFERROR(G18/H18, \"\")"

    ws.merge_cells('B19:C19')
    ws['B19'] = 'Low Income Developing Countries'
    ws['D19'] = '=SUMIF(Costs!$O$2:$O$250,B19,Costs!$M$2:$M$250)'
    ws['E19'] = '=SUMIF(GDP!$N$2:$N$250,B19,GDP!$L$2:$L$250)'
    ws['F19'] = '=AVERAGEIF(GDP!$N$2:$N$250,B19,GDP!$M$2:$M$250)'
    ws['G19'] = "=IF(F19=0,D19/Settings!C10,(D19*(1-(1+F19)/(1-Settings!C11)))/((1-Settings!C11)^Settings!C10*(1-((1+F17)/(1-Settings!C11))^(Settings!C10+1))))"
    ws['H19'] = "=SUMIF(GDP!$N$2:$N$250,B19,GDP!$C$2:$C$250)"
    ws['I19'] = "=IFERROR(G19/H19, \"\")"

    # ### Costs by Region
    ws.merge_cells('B22:I22')
    ws['B22'] = "Cost by Region"
    ws.merge_cells('B23:C23')
    ws['B23'] = "Region"
    ws['D23'] = "Total Cost ($Bn)"
    ws['E23'] = "Mean Annual 10-Year GDP ($Bn)"
    ws['F23'] = "GDP Growth Rate (%)(2022-2030)"
    ws['G23'] = "Initial Investment ($Bn)"
    ws['H23'] = "2022 GDP ($Bn)"
    ws['I23'] = "Annual GDP Share (%)"

    ws.merge_cells('B24:C24')
    ws['B24'] = 'Advanced Economies'
    ws['D24'] = '=SUMIF(Costs!$P$2:$P$250,B24,Costs!$M$2:$M$250)'
    ws['E24'] = '=SUMIF(GDP!$O$2:$O$250,B24,GDP!$L$2:$L$250)'
    ws['F24'] = '=AVERAGEIF(GDP!$O$2:$O$250,B24,GDP!$M$2:$M$250)'
    ws['G24'] = "=IF(F24=0,D24/Settings!C17,(D24*(1-(1+F24)/(1-Settings!C11)))/((1-Settings!C11)^Settings!C17*(1-((1+F24)/(1-Settings!C11))^(Settings!C10+1))))"
    ws['H24'] = "=SUMIF(GDP!$O$2:$O$250,B24,GDP!$C$2:$C$250)"
    ws['I24'] = "=IFERROR(G24/H24, "")"

    ws.merge_cells('B25:C25')
    ws['B25'] = "Caucasus and Central Asia"
    ws['D25'] = '=SUMIF(Costs!$P$2:$P$250,B25,Costs!$M$2:$M$250)'
    ws['E25'] = '=SUMIF(GDP!$O$2:$O$250,B25,GDP!$L$2:$L$250)'
    ws['F25'] = '=AVERAGEIF(GDP!$O$2:$O$250,B25,GDP!$M$2:$M$250)'
    ws['G25'] = "=IF(F25=0,D25/Settings!C17,(D25*(1-(1+F25)/(1-Settings!C11)))/((1-Settings!C11)^Settings!C17*(1-((1+F25)/(1-Settings!C11))^(Settings!C10+1))))"
    ws['H25'] = "=SUMIF(GDP!$O$2:$O$250,B25,GDP!$C$2:$C$250)"
    ws['I25'] = "=IFERROR(G25/H25, "")"

    ws.merge_cells('B26:C26')
    ws['B26'] = "Emerging and Developing Asia"
    ws['D26'] = '=SUMIF(Costs!$P$2:$P$250,B26,Costs!$M$2:$M$250)'
    ws['E26'] = '=SUMIF(GDP!$O$2:$O$250,B26,GDP!$L$2:$L$250)'
    ws['F26'] = '=AVERAGEIF(GDP!$O$2:$O$250,B26,GDP!$M$2:$M$250)'
    ws['G26'] = "=IF(F26=0,D26/Settings!C17,(D26*(1-(1+F26)/(1-Settings!C11)))/((1-Settings!C11)^Settings!C17*(1-((1+F26)/(1-Settings!C11))^(Settings!C10+1))))"
    ws['H26'] = "=SUMIF(GDP!$O$2:$O$250,B26,GDP!$C$2:$C$250)"
    ws['I26'] = "=IFERROR(G26/H26, "")"

    ws.merge_cells('B27:C27')
    ws['B27'] = "Emerging and Developing Europe"
    ws['D27'] = '=SUMIF(Costs!$P$2:$P$250,B27,Costs!$M$2:$M$250)'
    ws['E27'] = '=SUMIF(GDP!$O$2:$O$250,B27,GDP!$L$2:$L$250)'
    ws['F27'] = '=AVERAGEIF(GDP!$O$2:$O$250,B27,GDP!$M$2:$M$250)'
    ws['G27'] = "=IF(F27=0,D27/Settings!C17,(D27*(1-(1+F27)/(1-Settings!C11)))/((1-Settings!C11)^Settings!C17*(1-((1+F27)/(1-Settings!C11))^(Settings!C10+1))))"
    ws['H27'] = "=SUMIF(GDP!$O$2:$O$250,B27,GDP!$C$2:$C$250)"
    ws['I27'] = "=IFERROR(G27/H27, "")"

    ws.merge_cells('B28:C28')
    ws['B28'] = "Latin America and the Caribbean"
    ws['D28'] = '=SUMIF(Costs!$P$2:$P$250,B28,Costs!$M$2:$M$250)'
    ws['E28'] = '=SUMIF(GDP!$O$2:$O$250,B28,GDP!$L$2:$L$250)'
    ws['F28'] = '=AVERAGEIF(GDP!$O$2:$O$250,B28,GDP!$M$2:$M$250)'
    ws['G28'] = "=IF(F28=0,D28/Settings!C17,(D28*(1-(1+F28)/(1-Settings!C11)))/((1-Settings!C11)^Settings!C17*(1-((1+F28)/(1-Settings!C11))^(Settings!C10+1))))"
    ws['H28'] = "=SUMIF(GDP!$O$2:$O$250,B28,GDP!$C$2:$C$250)"
    ws['I28'] = "=IFERROR(G28/H28, "")"

    ws.merge_cells('B29:C29')
    ws['B29'] = "Middle East, North Africa, Afghanistan, and Pakistan"
    ws['D29'] = '=SUMIF(Costs!$P$2:$P$250,B29,Costs!$M$2:$M$250)'
    ws['E29'] = '=SUMIF(GDP!$O$2:$O$250,B29,GDP!$L$2:$L$250)'
    ws['F29'] = '=AVERAGEIF(GDP!$O$2:$O$250,B29,GDP!$M$2:$M$250)'
    ws['G29'] = "=IF(F29=0,D29/Settings!C17,(D29*(1-(1+F29)/(1-Settings!C11)))/((1-Settings!C11)^Settings!C17*(1-((1+F29)/(1-Settings!C11))^(Settings!C10+1))))"
    ws['H29'] = "=SUMIF(GDP!$O$2:$O$250,B29,GDP!$C$2:$C$250)"
    ws['I29'] = "=IFERROR(G29/H29, "")"

    ws.merge_cells('B30:C30')
    ws['B30'] = "Sub-Sahara Africa"
    ws['D30'] = '=SUMIF(Costs!$P$2:$P$250,B30,Costs!$M$2:$M$250)'
    ws['E30'] = '=SUMIF(GDP!$O$2:$O$250,B30,GDP!$L$2:$L$250)'
    ws['F30'] = '=AVERAGEIF(GDP!$O$2:$O$250,B30,GDP!$M$2:$M$250)'
    ws['G30'] = "=IF(F30=0,D30/Settings!C17,(D30*(1-(1+F30)/(1-Settings!C11)))/((1-Settings!C11)^Settings!C17*(1-((1+F30)/(1-Settings!C11))^(Settings!C10+1))))"
    ws['H30'] = "=SUMIF(GDP!$O$2:$O$250,B30,GDP!$C$2:$C$250)"
    ws['I30'] = "=IFERROR(G30/H30, "")"

    set_border(ws, 'B6:I11', "thin", "000000")
    set_border(ws, 'B15:I19', "thin", "000000")
    set_border(ws, 'B22:I30', "thin", "000000")

    ws = relocate(ws, 'B', 7, 11, 'B', 34)
    ws = relocate(ws, 'D', 7, 11, 'C', 34)
    ws = format_numbers(ws, ['C'], (34,40), 'Comma [0]', 1)

    ws = relocate(ws, 'B', 7, 11, 'E', 34)
    ws = relocate(ws, 'E', 7, 11, 'F', 34)
    ws = format_numbers(ws, ['F'], (34,40), 'Comma [0]', 1)

    ws = relocate(ws, 'B', 7, 11, 'H', 34)
    ws = relocate(ws, 'G', 7, 11, 'I', 34)
    ws = format_numbers(ws, ['I'], (34,40), 'Comma [0]', 1)

    ws = relocate(ws, 'B', 7, 11, 'B', 42)
    ws = relocate(ws, 'I', 7, 11, 'C', 42)
    ws = format_numbers(ws, ['C'], (42,48), 'Percent', 2)

    # ### by country
    # ws['B34'] = "=B7"
    # ws['C34'] = "=B8"
    # ws['D34'] = "=B9"
    # ws['E34'] = "=B10"
    # ws['F34'] = "=B11"
    # ws['B35'] = "=TRANSPOSE(D7:D11)"
    # ws.formula_attributes["B35"] = {"t": "array", "ref": "B35:F35"}

    # ### by income group
    # ws['AA15'] = "=B16"
    # ws['AB15'] = "=B17"
    # ws['AC15'] = "=B18"
    # ws['AD15'] = "=B19"
    # ws['AA16'] = "=TRANSPOSE(D16:I19)"
    # ws.formula_attributes["AA16"] = {"t": "array", "ref": "AA16:AD21"}

    # ### by regional group
    # ws['AA24'] = "=B23"
    # ws['AB24'] = "=B24"
    # ws['AC24'] = "=B25"
    # ws['AD24'] = "=B26"
    # ws['AE24'] = "=B27"
    # ws['AF24'] = "=B28"
    # ws['AG24'] = "=B29"
    # ws['AH24'] = "=B30"
    # ws['AA25'] = "=TRANSPOSE(D23:I30)"
    # ws.formula_attributes["AA25"] = {"t": "array", "ref": "AA25:AH30"}

    # set_border(ws, 'AA6:AE11', "thin", "000000")
    # set_border(ws, 'AA15:AD21', "thin", "000000")
    # set_border(ws, 'AA24:AH30', "thin", "000000")

    return ws



def relocate(ws, col, min_row, max_row, end_col, end_min_row):
    """
    ws, col, min_row, max_row, end_col, end_min_row
    worksheet
    starting_column
    starting_minimum_row
    starting_maximum_row
    ending_column
    ending_minimum_row

    """
    for i in range(min_row, max_row+1):
        starting_loc = '={}{}'.format(col, i)
        difference = end_min_row - min_row
        ending_loc = '{}{}'.format(end_col, i+difference)
        ws[ending_loc] = starting_loc

    return ws


def format_numbers(ws, columns, set_range, format_type, number_format):

    lower, upper = set_range

    for column in columns:
        for i in range(lower, upper+1):
            cell = '{}{}'.format(column, i)

            ws[cell].style = format_type

            if format_type == 'Percent' and number_format == 1:
                ws[cell].number_format = '0.0%'
            elif format_type == 'Percent' and number_format == 2:
                ws[cell].number_format = '0.00%'
            elif number_format == 0:
                ws[cell].number_format = '0'
            elif number_format == 1:
                ws[cell].number_format = '0.0'
            elif number_format == 2:
                ws[cell].number_format = '0.00'

            ws[cell].font = Font(size=11)

    return ws


def set_border(ws, cell_range, style, color):

    thin = Side(border_style=style, color=color)
    for row in ws[cell_range]:
        for cell in row:
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)


def set_cell_color(ws, cell_range, color):

    fill_gen = PatternFill(fill_type='solid',
                                 start_color=color,
                                 end_color=color)

    for row in ws[cell_range]:
        for cell in row:
            cell.fill = fill_gen


def set_font(ws, cell_range, font):
    """

    """
    for row in ws[cell_range]:
        for cell in row:
            cell.font = Font(name=font, size=11)

    return ws


def set_bold(ws, cell_id, font):
    """

    """
    # for row in ws[cell_range]:
    # for cell in ws[cell_range]:
    ws[cell_id].font = Font(name=font, size=11, bold=True)

    return ws


def center_text(ws, cell_range):
    """

    """
    for row in ws[cell_range]:
        for cell in row:
            cell.alignment = Alignment(horizontal='center')

    return ws


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

    set_border(ws, 'A1:L{}'.format(len(data)+1), "thin", "000000")

    if metric == 'population':
        ws['M1'] = 'Population Sum'
        for i in range(2, lnth):
            cell = 'M{}'.format(i)
            ws[cell] = "=SUM(C{}:L{})".format(i,i)
        set_border(ws, 'M1:M{}'.format(len(data)+1), "thin", "000000")
        ws.column_dimensions['M'].width = 20

    if metric == 'area_km2':
        ws['M1'] = 'area_km2_sum'
        for i in range(2, lnth):
            cell = 'M{}'.format(i)
            ws[cell] = "=SUM(C{}:L{})".format(i,i)
        set_border(ws, 'M1:M{}'.format(len(data)+1), "thin", "000000")
        ws.column_dimensions['M'].width = 20

    ws.column_dimensions['B'].width = 30

    return ws, lnth


def add_pop_growth(ws):
    """

    """
    ws.sheet_properties.tabColor = "666600"

    path = os.path.join(DATA_RAW, 'population_growth_rate_2020_2030.csv')
    p_growth = pd.read_csv(path, encoding = "ISO-8859-1")

    p_growth.rename(columns={
        'isocode':'ISO3',
        'country': 'country_name'
        }, inplace=True
    )

    p_growth = p_growth[[
        'ISO3','country_name','2020', '2021', '2022', '2023', '2024',
        '2025', '2026', '2027', '2028', '2029', '2030'
    ]]

    p_growth = p_growth.sort_values('ISO3')

    lnth = len(p_growth) + 2

    for r in dataframe_to_rows(p_growth, index=False, header=True):
        ws.append(r)

    ws['N1'] = 'Mean 10-Year Population Growth Rate (%)'
    for i in range(2,lnth):
        cell = 'N{}'.format(i)
        ws[cell] = "=SUM(C{}:M{})/11".format(i,i)

    ws.column_dimensions['N'].width = 40

    ws = format_numbers(ws, ['N'], (1, 200), 'Comma [0]', 1)

    set_border(ws, 'A1:N{}'.format(len(p_growth)+1), "thin", "000000")

    return ws, lnth


def add_context(ws, data_sheet):
    """

    """
    ws = bar_chart(ws, "Estimates!$C$34:$C$38", "Estimates!$B$35:$B$38", "Total Cost", 'Cost ($Bn)', "B10")
    ws = bar_chart(ws, "Estimates!$F$34:$F$38", "Estimates!$E$35:$E$38", "Mean Annual 10-Year GDP",'GDP ($Bn)', "L10")
    ws = bar_chart(ws, "Estimates!$I$34:$I$38", "Estimates!$H$35:$H$38", "Initial Investment",'Cost ($Bn)', "B26")
    ws = bar_chart(ws, "Estimates!$C$42:$C$46", "Estimates!$B$43:$B$46", "GDP Share",'Percent of GDP (%)', "L26")

    return


def bar_chart(ws, data, categories, title, y_axis, loc):

    chart1 = BarChart()
    chart1.type = "bar"
    chart1.title = title
    chart1.style = 10
    chart1.y_axis.title = y_axis
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(categories)
    chart1.shape = 4
    chart1.dataLabels = DataLabelList()
    chart1.dataLabels.showVal = True
    chart1.legend = None
    ws.add_chart(chart1, loc)

    return ws


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
            part2 = "*(POWER(1+(INDEX(P_Growth!$M$2:$M$1611,MATCH(A2, P_Growth!$A$2:$A$1611,0))/100), Settings!$C$9-Settings!$C$8))"
            part3 = "*(Settings!$C$15/100)*(Settings!$C$16/100)*(Settings!$C$17/100)"
            ws[cell] = part1 + part2 + part3
# P_Density!C2*(POWER(1+(INDEX(P_Growth!$M$2:$M$1611,MATCH(A2, P_Growth!$A$2:$A$1611,0))/100), Settings!$B$11-Settings!$B$10))*(Settings!$E$9/100)*(Settings!$E$10/100)*(Settings!$E$11/100)
    set_border(ws, 'A1:L{}'.format(lnth-1), "thin", "000000")
    ws.column_dimensions['B'].width = 30

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
            part1 = "=(Users!{}*MAX(Settings!$C$18,".format(cell)
            part2 = "(Settings!$C$19*1024*8*(1/30)*(15/100)*(1/3600))))"
            ws[cell] = part1 + part2

    set_border(ws, 'A1:L{}'.format(lnth-1), "thin", "000000")
    ws.column_dimensions['B'].width = 30

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
    ws.column_dimensions['B'].width = 18
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
    # lookup.loc[lookup['frequency_GHz'] == '0.8']
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
        ws[col] = '=F{}*(VLOOKUP(Settings!$C$24, Lookups!$A$3:$B$5, 2, 0)/10)'.format(i)
    ws.column_dimensions['G'].width = 20

    ws['H2'] = 'w_future_spectrum'
    for i in range(3,250):
        col = 'H{}'.format(i)
        ws[col] = '=F{}*(VLOOKUP(Settings!$C$25, Lookups!$A$3:$B$5, 2, 0)/10)'.format(i)
    ws.column_dimensions['H'].width = 20

    set_border(ws, 'E1:H{}'.format(df_length+2), "thin", "000000")

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

    set_border(ws, 'A1:F{}'.format(len(coverage)+1), "thin", "000000")

    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 30

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

    set_border(ws, 'A1:L{}'.format(lnth-1), "thin", "000000")

    return ws


def add_capacity_sheet(ws, cols, lnth):
    """

    """
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
            part2 = '*(Settings!$C$16/100),Lookups!$GL$3:$G$250)),"-")'
            ws[cell] = part1 + part2
            ws.formula_attributes[cell] = {'t': 'array', 'ref': "{}:{}".format(cell, cell)}

    set_border(ws, 'A1:L{}'.format(lnth-1), "thin", "000000")

    return ws


def add_sites_sheet(ws, cols, lnth):
    """

    """
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

    columns = ['C','D','E','F','G','H','I','J','K','L']
    ws = format_numbers(ws, columns, (1, 200), 'Comma [0]', 0)

    set_border(ws, 'A1:L{}'.format(lnth-1), "thin", "000000")

    return ws


def add_new_sites_sheet(ws, cols, lnth):
    """

    """
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
            part1 = "=IF(P_Density!{}>Settings!$C$26,".format(cell)
            part2 = "IF(Towers!{}*(Settings!$C$16/100)<'Total Sites'!{},".format(cell, cell)
            part3 = "('Total Sites'!{}-Towers!{}*(Settings!$C$16/100)),0),".format(cell, cell)
            part4 = '"-")'
            ws[cell] = part1 + part2 + part3 + part4

    columns = ['C','D','E','F','G','H','I','J','K','L']
    ws = format_numbers(ws, columns, (1, 200), 'Comma [0]', 0)

    set_border(ws, 'A1:L{}'.format(lnth-1), "thin", "000000")

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
            part1 = "=IFERROR('New Sites'!{}*VLOOKUP('Lookups'!$A$9, 'Lookups'!$A$9:'Lookups'!$B$24, 2, FALSE)".format(cell)
            part2 = "+((SQRT((1/('Total Sites'!{}/Area!{}))/2)*1000)*'New Sites'!{}".format(cell, cell, cell)
            part3 = "*VLOOKUP('Lookups'!$A$10, 'Lookups'!$A$9:'Lookups'!$B$24, 2, FALSE))"
            part4 = "+'New Sites'!{}*VLOOKUP('Lookups'!$A$11, 'Lookups'!$A$9:'Lookups'!$B$24, 2, FALSE)".format(cell)
            part5 = ',"-")'
            ws[cell] = part1 + part2 + part3 + part4 + part5

    ws['M1'] = 'Total Cost ($Bn)'
    for i in range(2,lnth):
        cell = "M{}".format(i)
        part1 = '=IFERROR(SUMIF((C{}:L{}), "<>n/a")/1e9, "-")'.format(i, i)
        line = part1
        ws[cell] = line

    ws['N1'] = 'Cost Per User ($)'
    for i in range(2,lnth):
        cell = "N{}".format(i)
        ws[cell] = "=(M{}*1e9)/Pop!M{}".format(i, i)

    ws['O1'] = 'Income Group'
    for i in range(2,lnth):
        cell = "O{}".format(i)
        ws[cell] = "=IFERROR(INDEX(Options!$J$2:$J$1611,MATCH(A{}, Options!$G$2:$G$1611,0)), "")".format(i)

    ws['P1'] = 'Region'
    for i in range(2,lnth):
        cell = "P{}".format(i)
        ws[cell] = "=IFERROR(INDEX(Options!$I$2:$I$1611,MATCH(A{}, Options!$G$2:$G$1611,0)), "")".format(i)

    ws.column_dimensions['M'].width = 20
    ws.column_dimensions['N'].width = 20
    ws.column_dimensions['O'].width = 35
    ws.column_dimensions['P'].width = 45

    ws = format_numbers(ws, ['N'], (1,200), 'Comma [0]', 0)
    ws = format_numbers(ws, ['M'], (1,200), 'Comma [0]', 1)

    set_border(ws, 'A1:P{}'.format(lnth-1), "thin", "000000")

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

    lnth = len(gdp) + 2

    for r in dataframe_to_rows(gdp, index=False, header=True):
        ws.append(r)

    ws['L1'] = 'Mean 10-Year GDP ($Bn)'
    for i in range(2,lnth):
        cell = 'L{}'.format(i)
        ws[cell] = "=((SUM(B{}:K{})*1e9)/10)/1e9".format(i,i)

    ws['M1'] = 'GDP Growth Rate (%)'
    for i in range(2,lnth):
        cell = 'M{}'.format(i)
        ws[cell] = "=IFERROR((K{}-B{})/B{},"")".format(i,i,i)
    ws = format_numbers(ws, ['M'], (2,len(gdp)+1), 'Percent', 1)

    ws['N1'] = 'Income Group'
    for i in range(2,lnth):
        cell = "N{}".format(i)
        ws[cell] = "=IFERROR(INDEX(Options!$J$2:$J$1611,MATCH(A{}, Options!$G$2:$G$1611,0)), "")".format(i)

    ws['O1'] = 'Region'
    for i in range(2,lnth):
        cell = "O{}".format(i)
        ws[cell] = "=IFERROR(INDEX(Options!$I$2:$I$1611,MATCH(A{}, Options!$G$2:$G$1611,0)), "")".format(i)

    ws.column_dimensions['L'].width = 20
    ws.column_dimensions['M'].width = 20
    ws.column_dimensions['N'].width = 35
    ws.column_dimensions['O'].width = 45

    ws = format_numbers(ws, ['L'], (1, 200), 'Comma [0]', 1)

    set_border(ws, 'A1:O{}'.format(len(gdp)+1), "thin", "000000")

    return ws


if __name__ == "__main__":

    generate_workbook()
