"""
Collect Cost Data.

Written by Ed Oughton.

March 2022.

"""
import os
import configparser
# import json
# import csv
import pandas as pd
import geopandas as gpd
import matplotlib.pyplot as plt
import seaborn as sns
# import contextily as ctx
# import openpyxl
import xlwings as xw

CONFIG = configparser.ConfigParser()
CONFIG.read(os.path.join(os.path.dirname(__file__), '..', 'scripts', 'script_config.ini'))
BASE_PATH = CONFIG['file_locations']['base_path']

DATA_RAW = os.path.join(BASE_PATH, 'raw')
DATA_INTERMEDIATE = os.path.join(BASE_PATH, 'intermediate')
DATA_PROCESSED = os.path.join(BASE_PATH, 'processed')
VIS = os.path.join(BASE_PATH, '..', 'vis', 'figures')


def find_country_list(continent_list):
    """
    This function produces country information by continent.

    Parameters
    ----------
    continent_list : list
        Contains the name of the desired continent, e.g. ['Africa']

    Returns
    -------
    countries : list of dicts
        Contains all desired country information for countries in
        the stated continent.

    """
    glob_info_path = os.path.join(BASE_PATH, 'global_information.csv')
    countries = pd.read_csv(glob_info_path, encoding = "ISO-8859-1")

    if len(continent_list) > 0:
        data = countries.loc[countries['continent'].isin(continent_list)]
    else:
        data = countries

    output = []

    for index, country in data.iterrows():

        output.append({
            'country_name': country['country'],
            'iso3': country['ISO_3digit'],
            'iso2': country['ISO_2digit'],
            'regional_level': country['lowest'],
            'imf': country['imf']
        })

    return output


def get_country_outlines(countries):
    """

    """
    imf_iso3_codes = []

    for item in countries:
        if item['imf'] == 1:
            imf_iso3_codes.append(item['iso3'])

    path = os.path.join(DATA_RAW, 'gadm36_levels_shp', 'gadm36_0.shp')
    country_shapes = gpd.read_file(path, crs='epsg:4326')

    imf_countries = country_shapes[country_shapes['GID_0'].isin(imf_iso3_codes)]

    return imf_countries


def get_non_imf_outlines(countries):
    """

    """
    non_imf_iso3_codes = []

    for item in countries:
        if not item['imf'] == 1:
            non_imf_iso3_codes.append(item['iso3'])

    path = os.path.join(DATA_RAW, 'gadm36_levels_shp', 'gadm36_0.shp')
    country_shapes = gpd.read_file(path, crs='epsg:4326')

    non_imf = country_shapes[country_shapes['GID_0'].isin(non_imf_iso3_codes)]

    return non_imf


def collect_cost_results():
    """
    Collect results.

    """
    path = os.path.join(BASE_PATH, '..', 'vis', 'Oughton et al. (2022) DICE v3.6.xlsx')

    wb1 = xw.Book(path)
    sht1 = wb1.sheets['Total_Cost_Per_User']
    data1 = sht1.range('A1').options(
        pd.DataFrame,
        header=1,
        index=False,
        expand='table'
        ).value
    data1 = data1.to_dict('records')

    output = {}

    for item1 in data1:

        iso3 = item1['ISO3']
        income1 = item1['Income Group']
        region1 = item1['Region']
        # if not iso3 == 'AFG':
        #     continue
        if str(iso3) == 'nan':
            continue

        interim = {}

        for key1, value1 in item1.items():

            if key1 in ['ISO3','Country Name','Income Group','Region']:
                continue

            key1 = correct_decile(key1)

            if value1 == '-':
                value1 = 0

            interim[key1] = round(float(value1))

        output[iso3] = interim

    return output


def correct_decile(key1):
    """

    """
    if 'Decile 10' in key1:
        key1 = 100
    elif 'Decile 9' in key1:
        key1 = 90
    elif 'Decile 8' in key1:
        key1 = 80
    elif 'Decile 7' in key1:
        key1 = 70
    elif 'Decile 6' in key1:
        key1 = 60
    elif 'Decile 5' in key1:
        key1 = 50
    elif 'Decile 4' in key1:
        key1 = 40
    elif 'Decile 3' in key1:
        key1 = 30
    elif 'Decile 2' in key1:
        key1 = 20
    elif 'Decile 1' in key1:
        key1 = 10

    return key1


def collect_unconnected_results():
    """
    Collect results.

    """
    path = os.path.join(BASE_PATH, '..', 'vis', 'Oughton et al. (2022) DICE v3.6.xlsx')

    wb1 = xw.Book(path)
    sht1 = wb1.sheets['Unconnected_Users']
    data1 = sht1.range('A1').options(
        pd.DataFrame,
        header=1,
        index=False,
        expand='table'
        ).value
    data1 = data1.to_dict('records')

    output = {}

    for item1 in data1:

        iso3 = item1['ISO3']
        income1 = item1['Income Group']
        region1 = item1['Region']
        # if not iso3 == 'AFG':
        #     continue
        if str(iso3) == 'nan':
            continue

        interim = {}

        for key1, value1 in item1.items():

            if key1 in ['ISO3','Country Name','Income Group',
                'Region', 'country_name', 'Sum']:
                continue

            key1 = correct_decile(key1)

            interim[key1] = round(float(value1))

        output[iso3] = interim

    return output


def collect_deciles(countries):
    """
    Collect all decile information.

    """
    path_out = os.path.join(VIS, '..', 'data', 'global_deciles.csv')

    if os.path.exists(path_out):
        output = pd.read_csv(path_out)
        return output

    else:

        costs_dict = collect_cost_results()#cost_[:100]
        unconnected_dict = collect_unconnected_results()#cost_[:100]

        output = []

        for country in countries:#[:1]:

            if not country['imf'] == 1:
                continue

            iso3 = country['iso3']

            # if not iso3 == 'AFG':
            #     continue

            if not iso3 in costs.keys():
                continue

            country_costs = costs_dict[iso3]
            country_unconnected = unconnected_dict[iso3]

            filename = "regional_data_deciles.csv"
            folder = os.path.join(DATA_INTERMEDIATE, iso3)
            path = os.path.join(folder, filename)

            if not os.path.exists(path):
                continue

            decile_data = pd.read_csv(path)

            deciles = decile_data['decile'].unique()#[:1]

            for decile in deciles:

                subset = decile_data.loc[decile_data['decile'] == decile]
                pop_sum = subset['population'].sum()

                decile_costs = country_costs[decile]
                decile_unconnected = country_unconnected[decile]

                for idx, item in decile_data.iterrows():

                    if item['decile'] == decile:

                        if item['population'] == 0 or pop_sum == 0:
                            pop_share = 0
                            unconnected = 0
                            cost = 0
                        else:
                            pop_share = item['population'] /pop_sum
                            unconnected = decile_unconnected * pop_share
                            cost = (decile_unconnected * pop_share) * decile_costs

                        output.append({
                            'GID_0': item['GID_0'],
                            'GID_id': item['GID_id'],
                            'GID_level': item['GID_level'],
                            'population': item['population'],
                            # 'pop_share': pop_share,
                            'unconnected': unconnected,
                            'cost': cost,
                            'area_km2': item['area_km2'],
                            'decile': item['decile'],
                        })

        output = pd.DataFrame(output)
        output = output.drop_duplicates().reset_index()
        output.to_csv(path_out, index=False)

    return output


def get_regional_shapes():
    """
    Load regional shapes.
    """

    path = os.path.join(VIS, '..', 'data', 'regions.shp')

    if os.path.exists(path):
        output = gpd.read_file(path)
        # output = output[output['GID_id'].str.startswith('AFG')]
        return output
    else:

        output = []

        for item in os.listdir(DATA_INTERMEDIATE):#[:10]:
            if len(item) == 3: # we only want iso3 code named folders

                filename_gid3 = 'regions_3_{}.shp'.format(item)
                path_gid3 = os.path.join(DATA_INTERMEDIATE, item, 'regions', filename_gid3)

                filename_gid2 = 'regions_2_{}.shp'.format(item)
                path_gid2 = os.path.join(DATA_INTERMEDIATE, item, 'regions', filename_gid2)

                filename_gid1 = 'regions_1_{}.shp'.format(item)
                path_gid1 = os.path.join(DATA_INTERMEDIATE, item, 'regions', filename_gid1)

                if os.path.exists(path_gid3):
                    data = gpd.read_file(path_gid3)
                    data['GID_id'] = data['GID_3']
                    data = data.to_dict('records')
                elif os.path.exists(path_gid2):
                    data = gpd.read_file(path_gid2)
                    data['GID_id'] = data['GID_2']
                    data = data.to_dict('records')
                elif os.path.exists(path_gid1):
                    data = gpd.read_file(path_gid1)
                    data['GID_id'] = data['GID_1']
                    data = data.to_dict('records')
                else:
                    print('No shapefiles for {}'.format(item))
                    continue

                for datum in data:
                    output.append({
                        'geometry': datum['geometry'],
                        'properties': {
                            'GID_id': datum['GID_id'],
                            # 'area_km2': datum['area_km2']
                        },
                    })

        output = gpd.GeoDataFrame.from_features(output, crs='epsg:4326')

        output.to_file(path)

        return output


def combine_data(deciles, regions):
    """

    """
    regions['iso3'] = regions['GID_id'].str[:3]
    regions = regions[['GID_id', 'iso3', 'geometry']] #[:1000]
    regions = regions.copy()

    regions = regions.merge(deciles, how='left', left_on='GID_id', right_on='GID_id')
    regions.reset_index(drop=True, inplace=True)
    regions.to_file(os.path.join(VIS,'..','data','test3.shp'))

    return regions


def plot_regions_by_geotype(regions, path, imf_countries, non_imf):
    """
    Plot regions by geotype.

    """
    metric = 'cost'

    regions['cost'] = round(regions['cost'] / 1e6)

    regions['cost'] = regions['cost'].fillna(0)
    regions.to_file(os.path.join(VIS,'..','data','test4.shp'))

    satellite = regions[regions['GID_0'].isna()]

    regions = regions.dropna()
    zeros = regions[regions['cost'] == 0]
    regions = regions[regions['cost'] != 0]

    bins = [0,10,20,30,40,50,60,70,80,90,1e12]
    labels = ['<$10m','$20m','$30m','$40m','$50m','$60m','$70m','$80m','$90m','>$100m']

    regions['bin'] = pd.cut(
        regions[metric],
        bins=bins,
        labels=labels
    )

    sns.set(font_scale=0.9)
    fig, ax = plt.subplots(1, 1, figsize=(10, 4.5))

    minx, miny, maxx, maxy = regions.total_bounds
    # ax.set_xlim(minx+20, maxx-2)
    # ax.set_ylim(miny+2, maxy-10)
    ax.set_xlim(minx-20, maxx+5)
    ax.set_ylim(miny-5, maxy)

    base = regions.plot(column='bin', ax=ax, cmap='viridis', linewidth=0, #inferno_r
        legend=True, antialiased=False)
    # # imf_countries.plot(ax=base, facecolor="none", edgecolor='grey', linewidth=0.1)
    zeros = zeros.plot(ax=base, color='dimgray', edgecolor='dimgray', linewidth=0)
    non_imf.plot(ax=base, color='lightgrey', edgecolor='lightgrey', linewidth=0)

    handles, labels = ax.get_legend_handles_labels()

    fig.legend(handles[::-1], labels[::-1])

    # ctx.add_basemap(ax, crs=regions.crs, source=ctx.providers.CartoDB.Voyager)

    n = len(regions)
    name = 'Universal Broadband Infrastructure Investment Cost by Sub-National Region (n={})'.format(n)
    fig.suptitle(name)

    fig.tight_layout()
    fig.savefig(path)

    plt.close(fig)


if __name__ == "__main__":

    countries = find_country_list([])

    imf_countries = get_country_outlines(countries)
    non_imf = get_non_imf_outlines(countries)

    deciles = collect_deciles(countries)#[:300]
    # out = pd.DataFrame(deciles)
    # out.to_csv(os.path.join(VIS, '..', 'data.csv'))

    regions = get_regional_shapes()#[:1000]
    regions = combine_data(deciles, regions)
    regions = pd.DataFrame(regions)
    # regions = regions[['GID_id', 'cost', 'decile']]
    # regions.to_csv(os.path.join(VIS, '..', 'test.csv'))

    regions = gpd.read_file(os.path.join(VIS,'..','data','test3.shp'), crs='epsg:4326')
    path = os.path.join(VIS, 'regions_by_cost.png')

    plot_regions_by_geotype(regions, path, imf_countries, non_imf)
