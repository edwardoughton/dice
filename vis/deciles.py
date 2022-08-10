"""
Collect Decile Data.

Written by Ed Oughton.

March 2022.

"""
import os
import configparser
import pandas as pd
import geopandas as gpd
import matplotlib.pyplot as plt
import seaborn as sns

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


def collect_data(countries):
    """
    Collect all decile information.

    """
    path_out = os.path.join(VIS, '..', 'data', 'global_deciles.csv')

    if os.path.exists(path_out):
        output = pd.read_csv(path_out)
        return output

    else:
        output = []

        for country in countries:#[:1]:

            if not country['imf'] == 1:
                continue

            iso3 = country['iso3']

            # if not iso3 == 'IND':
            #     continue

            filename = "regional_data_deciles.csv"
            folder = os.path.join(DATA_INTERMEDIATE, iso3)
            path = os.path.join(folder, filename)

            if not os.path.exists(path):
                continue

            decile_data = pd.read_csv(path)

            for idx, item in decile_data.iterrows():
                output.append({
                    'GID_0': item['GID_0'],
                    'GID_id': item['GID_id'],
                    'GID_level': item['GID_level'],
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
        return output
    else:

        output = []
        print('here')
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


def plot_regions_by_geotype(data, regions, path, imf_countries, non_imf):
    """
    Plot regions by geotype.
    """
    # data = data.loc[data['area_km2'] > 100]
    data = data[['GID_id', 'decile']]
    data = data.drop_duplicates()

    regions = regions[['GID_id', 'geometry']] #[:1000]
    regions = regions.copy()

    regions = regions.merge(data, left_on='GID_id', right_on='GID_id')
    # regions.reset_index(drop=True, inplace=True)
    regions.to_file(os.path.join(VIS,'..','data','test3.shp'))
    metric = 'decile'

    bins = [0,10,20,30,40,50,60,70,80,90,100]
    labels = ['Decile 1','Decile 2','Decile 3','Decile 4','Decile 5','Decile 6','Decile 7','Decile 8','Decile 9','Decile 10']

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

    base = regions.plot(column='bin', ax=ax, cmap='viridis_r', linewidth=0, #inferno_r
        legend=True, antialiased=False)
    # imf_countries.plot(ax=base, facecolor="none", edgecolor='grey', linewidth=0.1)
    non_imf.plot(ax=base, color='lightgrey', edgecolor='grey', linewidth=0)

    handles, labels = ax.get_legend_handles_labels()

    fig.legend(handles[::-1], labels[::-1])

    # ctx.add_basemap(ax, crs=regions.crs, source=ctx.providers.CartoDB.Voyager)

    n = len(regions)
    name = 'Population Density Deciles for Sub-National Regions (n={})'.format(n)
    fig.suptitle(name)

    fig.tight_layout()
    fig.savefig(path)

    plt.close(fig)


def get_blended_hex(regions, cmap, bins, replacement_dict):
    """
    Get a blended hex.
    """
    sdn = regions.loc[regions['GID_id'] == 'SDN.17.1_1']
    ssd = regions.loc[regions['GID_id'] == 'SSD.5.4_1']

    sdn = sdn['bin'].values[0]
    ssd = ssd['bin'].values[0]

    values_hex = {}

    for i, z in zip(range(cmap.N), bins):
        rgba = cmap(i) # rgb2hex accepts rgb or rgba
        values_hex[str(z)] = str(matplotlib.colors.rgb2hex(rgba))

    for k, v in replacement_dict:
        sdn = sdn.replace(k, v)
        ssd = ssd.replace(k, v)

    if '-' in sdn:
        sdn = sdn[:2]
    if '-' in ssd:
        ssd = ssd[:2]

    if 'Viable' in sdn:
        sdn = -1000000000.0
    if 'Viable' in ssd:
        ssd = -1000000000.0

    sdn_hex = values_hex[str(sdn)].replace('#', '')
    ssd_hex = values_hex[str(ssd)].replace('#', '')

    blended_hex = jcolor_split(sdn_hex, ssd_hex)

    return blended_hex


if __name__ == "__main__":

    countries = find_country_list([])

    imf_countries = get_country_outlines(countries)

    non_imf = get_non_imf_outlines(countries)

    data = collect_data(countries)#[:300]

    regions = get_regional_shapes()#[:5000]

    path = os.path.join(VIS, 'regions_by_decile.png')

    plot_regions_by_geotype(data, regions, path, imf_countries, non_imf)
