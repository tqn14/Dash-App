
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from dash import Dash, dcc, html, Input, Output, callback, State

order = pd.read_excel("global_superstore_2016.xlsx", sheet_name = "Orders")
order['Order Date'] = pd.to_datetime(order['Order Date'])
order['Ship Date'] = pd.to_datetime(order['Ship Date'])
order['Year'] = order['Ship Date'].dt.year

tmp = order.groupby(['Customer ID', 'Country', 'Year']).sum('Sales').sort_values(by = 'Sales', ascending=False).reset_index()
tmp['Postal Code'] = tmp['Postal Code'].astype('int', errors='ignore')

limits = [(0,1000),(1000,5000),(5000,10000),(10000,30000)]
colors = ["lightgrey","crimson","lightseagreen","royalblue"]
YEARS = tmp['Year'].unique().tolist()

def plot_bb(df, limits): 
    fig = go.Figure()
    for i in range(len(limits)):
        lim = limits[i]
        df_sub = df[(df['Sales'] >= int(lim[0])) & (df['Sales'] <= int(lim[1]))][['Country', 'Sales']].reset_index(drop=True)
        fig.add_trace(go.Scattergeo(
            locationmode = 'country names',
            locations=df_sub['Country'],
            hoverinfo='all',
            hovertext = df_sub['Sales'],
            marker = dict(
                size = df_sub['Sales']/100,
                color = colors[i],
                line_color='rgb(40,40,40)',
                line_width=0.5,
                sizemode = 'area'
            ),
            name = '{0} - {1}'.format(lim[0],lim[1])))
    
    fig.update_layout(
        showlegend = True,
        geo = dict(
            scope = 'world',
            landcolor = 'rgb(217, 217, 217)',
        ))

    return fig

app = Dash(__name__)

app.layout = html.Div(children=[
    html.H1("Exploring Worldwide Retail Trends: Global Superstore", style={'text-align': 'center'}),
    html.Div(dcc.Slider(id="slct_year",
					min=min(YEARS),
					max=max(YEARS),
					value=min(YEARS),
					marks={str(year): {"label": str(year), "style": {"color": "#7fafdf"}} for year in YEARS})),

    html.Div(children = [html.P("Order by Country in {0}".format(min(YEARS)), id = 'worldmap_name', style={'text-align': 'center'}),
    		dcc.Graph(id = 'world_map', figure=go.Figure())]), 
])

@app.callback(
    Output('world_map', 'figure'),
    Input('slct_year', 'value')
)

def update_map(slected_year): 
    slected_year1 = int(slected_year)
    df_tmp = tmp[(tmp['Year'] == slected_year1)][['Country', 'Sales']].reset_index(drop=True)
    return plot_bb(df_tmp, limits)

@app.callback(
    Output('worldmap_name', 'children'),
    Input('slct_year', 'value')
)
def update_mapname(year):
    return "Order by Country in {0}".format(year)


if __name__ == '__main__':
    app.run_server(debug=True)


