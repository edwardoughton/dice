"""
Extract all model data for visualization.

Written by Ed Oughton.

March 2022.

"""
import os
import configparser
import pandas as pd
import xlwings as xw

CONFIG = configparser.ConfigParser()
CONFIG.read(os.path.join(os.path.dirname(__file__), '..', 'scripts', 'script_config.ini'))
BASE_PATH = CONFIG['file_locations']['base_path']

RESULTS = os.path.join(BASE_PATH, '..', 'results')


def extract_data():
    """
    Function to extract results from DICE workbook.

    """
    if not os.path.exists(RESULTS):
        os.makedirs(RESULTS)

    path = os.path.join(BASE_PATH, '..', 'Oughton et al. (2022) DICE.xlsx')

    wb = xw.Book(path)

    extract_component_costs(wb, 'Pop', 'population')
    extract_component_costs(wb, 'Area', 'area')

    extract_component_costs(wb, 'RAN_Capex', 'capex')
    extract_component_costs(wb, 'BH_Capex', 'capex')
    extract_component_costs(wb, 'Tower_Capex', 'capex')
    extract_component_costs(wb, 'Labor_Capex', 'capex')
    extract_component_costs(wb, 'Power_Capex', 'capex')
    extract_component_costs(wb, 'RAN_Opex', 'opex')
    extract_component_costs(wb, 'BH_Opex', 'opex')
    extract_component_costs(wb, 'Tower_Opex', 'opex')
    extract_component_costs(wb, 'Power_Opex', 'opex')

    extract_total_costs(wb)

    extract_gdp(wb)

    return


def extract_component_costs(wb, component, value_name):
    """

    """
    sheet = wb.sheets[component]

    data = sheet.range('A1').options(
        pd.DataFrame,
        header=1,
        index=False,
        expand='table'
        ).value

    data = data.to_dict('records')

    output = []

    for item in data:

        iso3 = item['ISO3']
        income = item['Income Group']
        region = item['Region']

        if str(iso3) == 'nan':
            continue

        for key, value in item.items():

            if key in ['ISO3','country_name',
                'Income Group','Region', 'Population Sum', 'area_km2_sum']:
                continue

            if value == '-':
                value = 0

            output.append({
                'iso3': iso3,
                'decile': key,
                value_name: float(value),
                'income': income,
                'region': region,
            })

    output = pd.DataFrame(output)
    path = os.path.join(RESULTS, '{}.csv'.format(component.lower()))
    output.to_csv(path, index=False)

    return output


def extract_total_costs(wb):
    """

    """
    sheet = wb.sheets['total_costs']

    data = sheet.range('A1').options(
        pd.DataFrame,
        header=1,
        index=False,
        expand='table'
        ).value

    data = data.to_dict('records')

    output = []

    for item in data:

        iso3 = item['ISO3']
        income = item['Income Group']
        region = item['Region']

        if str(iso3) == 'nan':
            continue

        for key, value in item.items():

            if key in ['ISO3','country_name','Total Cost ($)',
                'Cost Per Pop ($)','Income Group','Region']:
                continue

            if value == '-':
                value = 0

            output.append({
                'iso3': iso3,
                'decile': key,
                'cost': float(value),
                'income': income,
                'region': region,
            })

    output = pd.DataFrame(output)
    path = os.path.join(RESULTS, 'total_cost.csv')
    output.to_csv(path, index=False)

    return output


def extract_gdp(wb):
    """

    """
    sheet = wb.sheets['GDP']

    data = sheet.range('A1').options(
        pd.DataFrame,
        header=1,
        index=False,
        expand='table'
        ).value

    data = data.to_dict('records')

    output = []

    for item in data:

        iso3 = item['ISO3']
        income = item['Income Group']
        region = item['Region']

        if str(iso3) == 'nan':
            continue

        for key, value in item.items():

            if key in ['ISO3','country_name','Total Cost ($)',
                'Cost Per Pop ($)','Income Group','Region']:
                continue

            if value == '-':
                value = 0

            output.append({
                'iso3': iso3,
                'year': key,
                'gdp': float(value),
                'income': income,
                'region': region,
            })

    output = pd.DataFrame(output)
    path = os.path.join(RESULTS, 'gdp.csv')
    output.to_csv(path, index=False)

    return output


if __name__ == "__main__":

    extract_data()
