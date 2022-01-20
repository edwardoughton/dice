"""
Preprocessing scripts.

Written by Ed Oughton.

Winter 2021

"""
import os
import configparser
import json
import csv
import pandas as pd
import geopandas as gpd
import pyproj
from shapely.geometry import (Polygon, MultiPolygon, mapping, shape,
    MultiLineString, LineString, box)
from shapely.ops import transform, unary_union, nearest_points
import fiona
import fiona.crs
import rasterio
from rasterio.mask import mask
from rasterstats import zonal_stats
import networkx as nx
from rtree import index
import numpy as np
import random
import math

CONFIG = configparser.ConfigParser()
CONFIG.read(os.path.join(os.path.dirname(__file__), 'script_config.ini'))
BASE_PATH = CONFIG['file_locations']['base_path']

DATA_RAW = os.path.join(BASE_PATH, 'raw')
DATA_INTERMEDIATE = os.path.join(BASE_PATH, 'intermediate')
DATA_PROCESSED = os.path.join(BASE_PATH, 'processed')


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
        os.makedirs(path)
    shape_path = os.path.join(path, 'national_outline.shp')

    path = os.path.join(DATA_RAW, 'gadm36_levels_shp', 'gadm36_0.shp')
    countries = gpd.read_file(path)

    single_country = countries[countries.GID_0 == iso3]

    try:
        single_country['geometry'] = single_country.apply(
            exclude_small_shapes, axis=1)
    except:
        return 'All small shapes'

    glob_info_path = os.path.join(BASE_PATH, 'global_information.csv')
    load_glob_info = pd.read_csv(glob_info_path, encoding = "ISO-8859-1")
    single_country = single_country.merge(
        load_glob_info,left_on='GID_0', right_on='ISO_3digit')

    single_country.to_file(shape_path, driver='ESRI Shapefile')

    return


def process_regions(country):
    """
    Function for processing the lowest desired subnational
    regions for the chosen country.

    Parameters
    ----------
    country : string
        Three digit ISO country code.

    """
    regions = []

    iso3 = country['iso3']
    level = country['regional_level']

    for regional_level in range(1, level + 1):

        filename = 'regions_{}_{}.shp'.format(regional_level, iso3)
        folder = os.path.join(DATA_INTERMEDIATE, iso3, 'regions')
        path_processed = os.path.join(folder, filename)

        if os.path.exists(path_processed):
            continue

        if not os.path.exists(folder):
            os.mkdir(folder)

        filename = 'gadm36_{}.shp'.format(regional_level)
        path_regions = os.path.join(DATA_RAW, 'gadm36_levels_shp', filename)
        regions = gpd.read_file(path_regions)

        regions = regions[regions.GID_0 == iso3]

        try:
            regions['geometry'] = regions.apply(exclude_small_shapes, axis=1)
        except:
            return 'All small shapes'

        try:
            regions.to_file(path_processed, driver='ESRI Shapefile')
        except:
            pass

    return


def process_settlement_layer(country):
    """
    Clip the settlement layer to the chosen country
    boundary and place in desired country folder.

    Parameters
    ----------
    country : string
        Three digit ISO country code.

    """
    iso3 = country['iso3']
    # regional_level = country['regional_level']

    path_settlements = os.path.join(DATA_RAW,'settlement_layer',
        'ppp_2020_1km_Aggregated.tif')

    settlements = rasterio.open(path_settlements, 'r+')
    settlements.nodata = 255
    settlements.crs = {"init": "epsg:4326"}

    iso3 = country['iso3']
    path_country = os.path.join(DATA_INTERMEDIATE, iso3,
        'national_outline.shp')

    if os.path.exists(path_country):
        country = gpd.read_file(path_country)
    else:
        print('Must generate national_outline.shp first' )

    path_country = os.path.join(DATA_INTERMEDIATE, iso3)
    shape_path = os.path.join(path_country, 'settlements.tif')

    # if os.path.exists(shape_path):
    #     return print('Completed settlement layer processing')

    geo = gpd.GeoDataFrame()
    minx, miny, maxx, maxy = country['geometry'].total_bounds
    bbox = box(minx, miny, maxx, maxy)
    geo = gpd.GeoDataFrame({'geometry': bbox}, index=[0])
    # bbox = country['geometry']
    # geo = gpd.GeoDataFrame({'geometry': bbox})

    coords = [json.loads(geo.to_json())['features'][0]['geometry']]

    #chop on coords
    out_img, out_transform = mask(settlements, coords, crop=True)

    # Copy the metadata
    out_meta = settlements.meta.copy()

    out_meta.update({"driver": "GTiff",
                    "height": out_img.shape[1],
                    "width": out_img.shape[2],
                    "transform": out_transform,
                    "crs": 'epsg:4326'})

    with rasterio.open(shape_path, "w", **out_meta) as dest:
            dest.write(out_img)

    return print("Written raster layer")


def get_pop_and_luminosity_data(country):
    """
    Extract regional luminosity and population data.

    Parameters
    ----------
    country : string
        Three digit ISO country code.

    """
    iso3 = country['iso3']
    level = country['regional_level']
    gid_level = 'GID_{}'.format(level)

    path_output = os.path.join(DATA_INTERMEDIATE, iso3, 'regional_data.csv')

    # if os.path.exists(path_output):
    #     return print('Regional data already exists')

    path_settlements = os.path.join(DATA_INTERMEDIATE, iso3,
        'settlements.tif')
    # path_settlements = os.path.join(DATA_RAW,'settlement_layer',
    #     'ppp_2020_1km_Aggregated.tif')

    filename = 'regions_{}_{}.shp'.format(level, iso3)
    folder = os.path.join(DATA_INTERMEDIATE, iso3, 'regions')
    path = os.path.join(folder, filename)

    regions = gpd.read_file(path)

    results = []

    for index, region in regions.iterrows():

        with rasterio.open(path_settlements) as src:

            affine = src.transform
            array = src.read(1)
            array[array <= 0] = 0

            population_summation = [d['sum'] for d in zonal_stats(
                region['geometry'],
                array,
                stats=['sum'],
                affine=affine,
                nodata=255
                )][0]

        area_km2 = round(area_of_polygon(region['geometry']) / 1e6)

        if area_km2 > 0:
            population_km2 = (
                population_summation / area_km2 if population_summation else 0)
        else:
            population_km2 = 0

        results.append({
            'GID_0': region['GID_0'],
            'GID_id': region[gid_level],
            'GID_level': gid_level,
            # 'mean_luminosity_km2': mean_luminosity_km2,
            'population': (population_summation if population_summation else 0),
            'area_km2': area_km2,
            'population_km2': population_km2,
        })

    results_df = pd.DataFrame(results)

    print(round(results_df['population'].sum()/1e6,2))

    results_df.to_csv(path_output, index=False)

    return


def area_of_polygon(geom):
    """
    Returns the area of a polygon. Assume WGS84 as crs.
    """
    geod = pyproj.Geod(ellps="WGS84")

    poly_area, poly_perimeter = geod.geometry_area_perimeter(
        geom
    )

    return abs(poly_area)


def collect_results(countries):
    """

    """
    output = []

    for country in countries:

        path = os.path.join(DATA_INTERMEDIATE, country['iso3'], 'regional_data.csv')

        if os.path.exists(path):

            data = pd.read_csv(path)

            data = data.sort_values(by='population_km2', ascending=True)#.reset_index()

            try:
                data['decile'] = pd.qcut(data['population_km2'],
                    q=10, #precision=0,
                    labels=[100,90,80,70,60,50,40,30,20,10],
                    duplicates='drop'
                    ) #[0,10,20,30,40,50,60,70,80,90,100]
            except:
                continue

            data = data[['GID_0', 'decile', 'population', 'area_km2']]

            data = data.groupby(['GID_0','decile'])[
                ['population', 'area_km2']
                ].sum().reset_index()

            data['population_km2'] = data['population'] / data['area_km2']

            data['country_name'] = country['country_name']

            data = data.to_dict('records')

            output = output + data

    output = pd.DataFrame(output)
    path = os.path.join(DATA_INTERMEDIATE, 'all_pop_data.csv')
    output.to_csv(path, index=False)


def define_deciles(regions):
    """
    Allocate deciles to regions.
    """
    regions = regions.sort_values(by='population_km2', ascending=True)

    regions['decile'] = regions.groupby([
        'GID_0',
        'scenario',
        'strategy',
        'confidence'
    ], as_index=True).population_km2.apply( #cost_per_sp_user
        pd.qcut, q=11, precision=0,
        labels=[100,90,80,70,60,50,40,30,20,10,0],
        duplicates='drop') #   [0,10,20,30,40,50,60,70,80,90,100]

    return regions

if __name__ == '__main__':

    countries = find_country_list([])

    for country in countries:

        if not country['iso3'] == 'GBR':
            continue

        # path = os.path.join(DATA_INTERMEDIATE, country['iso3'], 'regional_data.csv')

        # if not os.path.exists(path):
        #     print(country['country_name'], country['iso3'])
        # print('--Working on {}'.format(country['iso3']))

        # try:
        print('Processing country boundary')
        process_country_shapes(country)

        print('Processing regions')
        response = process_regions(country)
        if response == 'All small shapes':
            continue

        print('Processing settlement layer')
        process_settlement_layer(country)

        print('Getting population and luminosity')
        get_pop_and_luminosity_data(country)

        # except:
        #     continue

    collect_results(countries)
