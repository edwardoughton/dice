"""
Population processing script.

Written by Ed Oughton.

December 2021

"""
# import argparse
import os
import sys
import configparser
import geopandas as gpd
from shapely.geometry import Polygon, mapping, MultiPolygon
import pandas as pd
import numpy as np
import rasterio
from rasterstats import zonal_stats

CONFIG = configparser.ConfigParser()
CONFIG.read(os.path.join(os.path.dirname(__file__), 'script_config.ini'))
BASE_PATH = CONFIG['file_locations']['base_path']

DATA_RAW = os.path.join(BASE_PATH, 'raw')
DATA_INTERMEDIATE = os.path.join(BASE_PATH, 'intermediate')


def get_countries_lut():
    """
    Get country LUT.

    """
    output = []

    path = os.path.join(DATA_RAW, 'settlement_layer', 'total_population.csv')
    population_lut = pd.read_csv(path)#[:1]

    for idx, item in population_lut.iterrows():

        if item['Country Code'] in ['INX', 'ERI']:
            continue

        output.append({
            'iso3': item['Country Code'],
            'name': item['Country Name'],
            'population': int(item['2020'])
        })

    return output


def process_country_shapes(country):
    """
    Creates a single national boundary for the desired country.

    Parameters
    ----------
    country : string
        Three digit ISO country code.

    """
    iso3 = country['iso3']

    path = os.path.join(DATA_INTERMEDIATE, iso3)

    if os.path.exists(os.path.join(path, 'national_outline.shp')):
        return 'Completed national outline processing'

    if not os.path.exists(path):
        print('Creating directory {}'.format(path))
        os.makedirs(path)
    shape_path = os.path.join(path, 'national_outline.shp')

    # print('Loading all country shapes')
    path = os.path.join(DATA_RAW, 'gadm36_levels_shp', 'gadm36_0.shp')
    countries = gpd.read_file(path)

    # print('Getting specific country shape for {}'.format(iso3))
    single_country = countries[countries.GID_0 == iso3]

    # print('Excluding small shapes')
    single_country = single_country.copy()
    single_country.loc[:, 'geometry'] = single_country.apply(
        exclude_small_shapes, axis=1)

    # print('Adding ISO country code and other global information')
    glob_info_path = os.path.join(BASE_PATH, 'global_information.csv')
    load_glob_info = pd.read_csv(glob_info_path, encoding = "ISO-8859-1")
    single_country = single_country.merge(
        load_glob_info, left_on='GID_0', right_on='iso3')

    # print('Exporting processed country shape')
    single_country.to_file(shape_path, driver='ESRI Shapefile')

    return print('Processing country shape complete')


def exclude_small_shapes(x):
    """
    Remove small multipolygon shapes.

    Parameters
    ---------
    x : polygon
        Feature to simplify.

    Returns
    -------
    MultiPolygon : MultiPolygon
        Shapely MultiPolygon geometry without tiny shapes.

    """
    # if its a single polygon, just return the polygon geometry
    if x.geometry.geom_type == 'Polygon':
        return x.geometry

    # if its a multipolygon, we start trying to simplify
    # and remove shapes if its too big.
    elif x.geometry.geom_type == 'MultiPolygon':

        area1 = 0.01
        area2 = 50

        # dont remove shapes if total area is already very small
        if x.geometry.area < area1:
            return x.geometry
        # remove bigger shapes if country is really big

        if x['GID_0'] in ['CHL','IDN']:
            threshold = 0.01
        elif x['GID_0'] in ['RUS','GRL','CAN','USA']:
            threshold = 0.01

        elif x.geometry.area > area2:
            threshold = 0.1
        else:
            threshold = 0.001

        # save remaining polygons as new multipolygon for
        # the specific country
        new_geom = []
        for y in x.geometry:
            if y.area > threshold:
                new_geom.append(y)

        return MultiPolygon(new_geom)


def generate_grid(country):
    """
    Generate a 10x10km spatial grid for the chosen country.

    """
    iso3 = country['iso3']

    path_out = os.path.join(DATA_INTERMEDIATE, iso3, 'grid.shp')

    if not os.path.exists(path_out):

        filename = 'national_outline.shp'
        country_outline = gpd.read_file(os.path.join(DATA_INTERMEDIATE, iso3, filename))

        country_outline.crs = "epsg:4326"
        country_outline = country_outline.to_crs("epsg:3857")

        xmin, ymin, xmax, ymax = country_outline.total_bounds

        #10km sides, leading to 100km^2 area
        length = 1e4
        wide = 1e4

        xmin = xmin - wide
        xmax = xmax + wide
        ymin = ymin - length
        ymax = ymax + length

        cols = list(range(int(np.floor(xmin)), int(np.ceil(xmax)), int(wide)))
        rows = list(range(int(np.floor(ymin)), int(np.ceil(ymax)), int(length)))
        rows.reverse()

        polygons = []
        for x in cols:
            for y in rows:
                polygons.append(
                    Polygon([(x,y), (x+wide, y), (x+wide, y-length), (x, y-length)])
                    )

        grid = gpd.GeoDataFrame({'geometry': polygons}, crs='epsg:3857')
        # grid.to_file(os.path.join(DATA_INTERMEDIATE, iso3, 'full_grid.shp'))

        intersection = gpd.overlay(grid, country_outline, how='intersection')
        intersection.crs = "epsg:3857"
        intersection = intersection.to_crs("epsg:4326")
        # intersection = intersection[:100]

        final_grid = query_settlement_layer(iso3, intersection)

        final_grid = final_grid[final_grid.geometry.notnull()]
        final_grid.to_file(path_out)

    else:
        final_grid = gpd.read_file(path_out, crs='epsg:4326')

    pop_dict = {
        'iso3': iso3,
        'name': country['name'],
        'pop_estimate': final_grid['population'].sum(),
        'pop_true': country['population'],
        'difference_perc': (
            (final_grid['population'].sum() -country['population']) /
            final_grid['population'].sum()
        ) * 100
    }

    return pop_dict


def query_settlement_layer(iso3, grid):
    """
    Query the settlement layer to get an estimated population for each grid square.

    """
    path = os.path.join(DATA_RAW, 'settlement_layer', 'ppp_2020_1km_Aggregated.tif')

    grid['population'] = pd.DataFrame(
        zonal_stats(vectors=grid['geometry'], raster=path, stats='sum'))['sum']

    grid = grid.replace([np.inf, -np.inf], np.nan)

    grid.loc[grid['population'] < 0, 'population'] = 0
    grid['population'] = grid['population'].fillna(0)

    return grid


if __name__ == '__main__':

    # path = os.path.join(DATA_RAW, '..', 'global_information.csv')
    # countries = pd.read_csv(path, encoding = "ISO-8859-1")[:1]

    countries = get_countries_lut()

    pop_matching = []

    for country in countries[::-1]:

        # if not country['iso3'] == 'AFG':
        #     continue

        print('-- {}'.format(country['iso3']))

        process_country_shapes(country)

        pop_dict = generate_grid(country)

        pop_matching.append(pop_dict)

        print('------{} Pop. Estimated: {}m, True: {}m'.format(
            country['iso3'],
            round(pop_dict['pop_estimate'] / 1e6, 3),
            round(pop_dict['pop_true'] / 1e6, 3)
            )
        )

    pop_matching = pd.DataFrame(pop_matching)
    path = os.path.join(DATA_INTERMEDIATE, 'pop_estimates.csv')
    pop_matching.to_csv(path, index=False)
