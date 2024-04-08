import streamlit as st
import pandas as pd
import requests
import matplotlib.pyplot as plt
import os
from io import BytesIO

from docx import Document
from docx.shared import Cm, RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK


data_loaded = False

def make_table(model, type, var, aggregation, period, percentile, scenario, model_code, model_calculation, statistic, region_code, region):

    # Building the URL
    url = f'https://cckpapi.worldbank.org/cckp/v1/{model}_{type}_{var}_{type}_{aggregation}_{period}_{percentile}_{scenario}_{model_code}_{model_calculation}_{statistic}/{region_code}?_format=json'

    # Requesting the data
    response = requests.get(url)
    data = response.json()

    table = pd.DataFrame(data['data']).rename_axis('year').reset_index().rename(columns={region_code: f'{var}_{scenario}_{percentile}'})
    table['year'] = pd.to_datetime(table['year'])
    table = table.set_index('year').resample('Y').mean()

    return table

def make_plot_single(table, var, plot=True):
    # Make the plot
    fig = plt.figure(figsize=(6, 3))
    table.plot(ax=plt.gca(), legend=False, color='black', linewidth=0.5, label='Yearly value')
    table.rolling(5, center=True).mean().plot(ax=plt.gca(), legend=False, color='red', linewidth=2. , label='5yr rolling mean')
    ax = plt.gca()

    # hide axis spines
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    ax.set_ylabel(var_ref[var_ref['Code'] == var].Unit.values[0])
    ax.set_title(var_ref[var_ref['Code'] == var].Variable.values[0])
    ax.grid(True)
    #put legend outside right
    plt.legend(['Yearly value', '5yr rolling mean'], loc='center left', bbox_to_anchor=(1, 0.8), frameon=False)

    if plot:
        st.pyplot(fig)

    return fig

def make_plot_multi(historical, historical_lower, historical_upper,
                    ssp126, ssp126_lower, ssp126_upper,
                    ssp245, ssp245_lower, ssp245_upper,
                    ssp585, ssp585_lower, ssp585_upper,
                    plot=True):

    fig = plt.figure(figsize=(6, 3))
    historical.rolling(5).mean().plot(ax=plt.gca(), color='blue', linewidth=2.)
    historical_lower.rolling(5).mean().plot(ax=plt.gca(), color='blue', linewidth=0.5, linestyle=':')
    historical_upper.rolling(5).mean().plot(ax=plt.gca(), color='blue', linewidth=0.5, linestyle='--')

    ssp126.rolling(5).mean().plot(ax=plt.gca(),  color='green', linewidth=2.)
    ssp126_lower.rolling(5).mean().plot(ax=plt.gca(), color='green', linewidth=0.5, linestyle=':')
    ssp126_upper.rolling(5).mean().plot(ax=plt.gca(),  color='green', linewidth=0.5, linestyle='--')

    ssp245.rolling(5).mean().plot(ax=plt.gca(),  color='orange', linewidth=2.)
    ssp245_lower.rolling(5).mean().plot(ax=plt.gca(), color='orange', linewidth=0.5, linestyle=':')
    ssp245_upper.rolling(5).mean().plot(ax=plt.gca(), color='orange', linewidth=0.5, linestyle='--')

    ssp585.rolling(5).mean().plot(ax=plt.gca(), color='red', linewidth=2. )
    ssp585_lower.rolling(5).mean().plot(ax=plt.gca(), color='red', linewidth=0.5, linestyle=':')
    ssp585_upper.rolling(5).mean().plot(ax=plt.gca(), color='red', linewidth=0.5, linestyle='--')

    ax = plt.gca()

    ax.plot([],[], color='black', linewidth=0, label=' ')
    ax.plot([],[], color='black', linewidth=2., label='Median')
    ax.plot([],[], color='black', linewidth=0.5, linestyle=':', label='Lower')
    ax.plot([],[], color='black', linewidth=0.5, linestyle='--', label='Upper')
    # hide axis spines
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    ax.set_ylabel(var_ref[var_ref['Code'] == var].Unit.values[0])
    ax.set_title(var_ref[var_ref['Code'] == var].Variable.values[0])
    ax.grid(True)

    #put legend outside right
    handles, labels = ax.get_legend_handles_labels()
    ax.legend(list( handles[i] for i in [0,3,6,9,12,13,14,15] ), ['Historical','SSP 1-2.6','SSP 2-4.5','SSP 5-8.5',' ','Median','Lower','Upper'], loc='center left', bbox_to_anchor=(1, 0.5), frameon=False)

    if plot:
        st.pyplot(fig)

    return fig


# make wide screen
st.set_page_config(layout="wide")

geo_ref = pd.read_excel('geonames.xlsx', sheet_name='Regions')
var_ref = pd.read_excel('geonames.xlsx', sheet_name='Variables')

era_var_code = ['tas','tasmax','tasmin','tnn','tr','txx','fd','pr','rx1day','rx5day']
era_var = [var_ref[var_ref['Code'] == var].Variable.values[0] for var in era_var_code]

cmip_var_code = ['tas', 'tasmax', 'tasmin', 'tnn', 'tr', 'txx', 'fd','hd30', 'hd35', 'hd40', 'hd45', 'hdd65', 'id',  'cdd65', 'sd',  'tr23', 'tr26', 'tr29', 'pr', 'rx1day', 'rx5day', 'cdd', 'cwd',    'prpercnt', 'r20mm', 'r50mm']
cmip_var = [var_ref[var_ref['Code'] == var].Variable.values[0] for var in cmip_var_code]


# Defining region
st.subheader('Select region')
col1, col2 = st.columns(2)

with col1:
    country = st.selectbox('Country', geo_ref['Country'].unique())
    all_reg = geo_ref[geo_ref['Country'] == country]['State'].unique()

with col2:
    region = st.selectbox('Region', all_reg)
    region_code = geo_ref[(geo_ref['Country'] == country) & (geo_ref['State'] == region)]['State Code'].values[0]


# Defining variable
st.subheader('Select variables')

with st.expander('ERA5'):
    variable_era = []

    for var_e in era_var:
        with st.container():
            col1, col2= st.columns(2)
            with col1:
                if st.checkbox(var_e, True, key=f'{var_e}_ERA'):
                    variable_era.append(var_e)
            with col2:
                st.write(var_ref[var_ref['Variable'] == var_e].Description.values[0])

    variable_era_code = [var_ref[var_ref['Variable'] == var].Code.values[0] for var in variable_era]

with st.expander('CMIP6'):
    variable_cmip = []

    for var_c in cmip_var:
        with st.container():
            col1, col2 = st.columns(2)
            with col1:
                if var_c in era_var:
                    if st.checkbox(var_c, True, key=f'{var_c}_CMIP'):
                        variable_cmip.append(var_c)
                else:
                    if st.checkbox(var_c, False, key=f'{var_c}_CMIP'):
                        variable_cmip.append(var_c)
            with col2:
                st.write(var_ref[var_ref['Variable'] == var_c].Description.values[0])
    
    variable_cmip_code = [var_ref[var_ref['Variable'] == var].Code.values[0] for var in variable_cmip]





doc = Document()
style = doc.styles['Normal']
style.paragraph_format.line_spacing = 1.15
style.font.color.rgb = RGBColor(0, 0, 0)
style.font.name = "Calibri"
style.font.size = Pt(11)

title = doc.add_heading('Climate and Climate Change', level=1)
title.style.font.color.rgb = RGBColor(0, 0, 0)
title.bold = True
title.style.font.name = "Calibri"
title.style.font.size = Pt(11)
title.add_run().add_break(WD_BREAK.LINE)


if st.button('Get data'):

    title = doc.add_heading('ERA5', level=2)
    title.style.font.color.rgb = RGBColor(0, 0, 0)
    title.bold = True
    title.style.font.name = "Calibri"
    title.style.font.size = Pt(11)
    title.add_run().add_break(WD_BREAK.LINE)

    for var in variable_era_code:

        tab = make_table('era5-x0.5', 'timeseries', var, 'annual', '1950-2020', 'mean', 'historical', 'era5', 'era5', 'mean', region_code, region)
        fig = make_plot_single(tab, var, False)
        fig.savefig('tmp.png', bbox_inches='tight', dpi=300)

        p = doc.add_paragraph().add_run(var_ref[var_ref['Code'] == var].Variable.values[0])

        p.bold = True
        p.italic = True
        doc.add_picture('tmp.png')
        os.remove('tmp.png')
        doc.add_paragraph().add_run(var_ref[var_ref['Code'] == var].Description.values[0])

        # Creating a table object
        table = doc.add_table(rows=1, cols=2)

        # Adding heading in the 1st row of the table
        row = table.rows[0].cells
        row[0].text = 'Year'
        row[1].text = f'ERA5 value [{var_ref[var_ref["Code"] == var].Unit.values[0]}]'


        # Adding data from the list to the table
        for y in range(1950,2021,10):

            # Adding a row and then adding data in it.
            row = table.add_row().cells
            # Converting id to string as table can only take string input
            row[0].text = str(y)
            row[1].text = f'{tab[tab.index.year==y].values[0][0]:6.2f}'
    

        table.style = 'Colorful List'


    title = doc.add_heading('CMIP6', level=2)
    title.style.font.color.rgb = RGBColor(0, 0, 0)
    title.bold = True
    title.style.font.name = "Calibri"
    title.style.font.size = Pt(11)
    title.add_run().add_break(WD_BREAK.LINE)

    for var in variable_cmip_code:

        tab_historical = make_table('cmip6-x0.25', 'timeseries', var, 'annual', '1950-2014', 'median', 'historical', 'ensemble', 'all', 'mean', region_code, region)
        tab_historical_lower = make_table('cmip6-x0.25', 'timeseries', var, 'annual', '1950-2014', 'p10', 'historical', 'ensemble', 'all', 'mean', region_code, region)
        tab_historical_upper = make_table('cmip6-x0.25', 'timeseries', var, 'annual', '1950-2014', 'p90', 'historical', 'ensemble', 'all', 'mean', region_code, region)
        
        tab_ssp126 = make_table('cmip6-x0.25', 'timeseries', var, 'annual', '2015-2100', 'median', 'ssp126', 'ensemble', 'all', 'mean', region_code, region)
        tab_ssp126_lower = make_table('cmip6-x0.25', 'timeseries', var, 'annual', '2015-2100', 'p10', 'ssp126', 'ensemble', 'all', 'mean', region_code, region)
        tab_ssp126_upper = make_table('cmip6-x0.25', 'timeseries', var, 'annual', '2015-2100', 'p90', 'ssp126', 'ensemble', 'all', 'mean', region_code, region)

        tab_ssp245 = make_table('cmip6-x0.25', 'timeseries', var, 'annual', '2015-2100', 'median', 'ssp245', 'ensemble', 'all', 'mean', region_code, region)
        tab_ssp245_lower = make_table('cmip6-x0.25', 'timeseries', var, 'annual', '2015-2100', 'p10', 'ssp245', 'ensemble', 'all', 'mean', region_code, region)
        tab_ssp245_upper = make_table('cmip6-x0.25', 'timeseries', var, 'annual', '2015-2100', 'p90', 'ssp245', 'ensemble', 'all', 'mean', region_code, region)

        tab_ssp585 = make_table('cmip6-x0.25', 'timeseries', var, 'annual', '2015-2100', 'median', 'ssp585', 'ensemble', 'all', 'mean', region_code, region)
        tab_ssp585_lower = make_table('cmip6-x0.25', 'timeseries', var, 'annual', '2015-2100', 'p10', 'ssp585', 'ensemble', 'all', 'mean', region_code, region)
        tab_ssp585_upper = make_table('cmip6-x0.25', 'timeseries', var, 'annual', '2015-2100', 'p90', 'ssp585', 'ensemble', 'all', 'mean', region_code, region)

        tab_tot = pd.concat([tab_historical, tab_historical_lower, tab_historical_upper, tab_ssp126, tab_ssp126_lower, tab_ssp126_upper, tab_ssp245, tab_ssp245_lower, tab_ssp245_upper, tab_ssp585, tab_ssp585_lower, tab_ssp585_upper], axis=1)
        tab_tot = tab_tot.rolling(5).mean()
        
        fig = make_plot_multi(tab_historical, tab_historical_lower, tab_historical_upper,
                        tab_ssp126, tab_ssp126_lower, tab_ssp126_upper,
                        tab_ssp245, tab_ssp245_lower, tab_ssp245_upper,
                        tab_ssp585, tab_ssp585_lower, tab_ssp585_upper, False)
        
        fig.savefig('tmp.png', bbox_inches='tight', dpi=300)

        p = doc.add_paragraph().add_run(var_ref[var_ref['Code'] == var].Variable.values[0])

        p.bold = True
        p.italic = True
        doc.add_picture('tmp.png')
        os.remove('tmp.png')
        doc.add_paragraph().add_run(var_ref[var_ref['Code'] == var].Description.values[0])

        # Creating a table object
        table = doc.add_table(rows=1, cols=4)

        # Adding heading in the 1st row of the table
        row = table.rows[0].cells
        row[0].text = 'Year'
        row[1].text = f'SSP 1-2.6 [{var_ref[var_ref["Code"] == var].Unit.values[0]}]'
        row[2].text = f'SSP 2-4.5 [{var_ref[var_ref["Code"] == var].Unit.values[0]}]'
        row[3].text = f'SSP 5-8.5 [{var_ref[var_ref["Code"] == var].Unit.values[0]}]'

        # Adding data from the list to the table
        for y in range(2020,2101,10):
            # Adding a row and then adding data in it.
            row = table.add_row().cells
            # Converting id to string as table can only take string input
            row[0].text = str(y)

            row[1].text = f"{tab_tot[tab_tot.index.year==y][f'{var}_ssp126_median'].values[0]:6.2f}"
            row[2].text = f"{tab_tot[tab_tot.index.year==y][f'{var}_ssp245_median'].values[0]:6.2f}"
            row[3].text = f"{tab_tot[tab_tot.index.year==y][f'{var}_ssp585_median'].values[0]:6.2f}"
    

        table.style = 'Colorful List'
        
    st.success('Data loaded successfully')

    data_loaded = True



if data_loaded == True:

    try:
        bio = BytesIO()
        doc.save(bio)

        st.download_button(
                label="Click here to download",
                data=bio.getvalue(),
                file_name=f"{country}_{region}.docx",
                mime="docx"
            )
    
    except:
        pass


        