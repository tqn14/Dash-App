
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from dash import Dash, dcc, html, Input, Output, callback, State
import warnings
warnings.filterwarnings('ignore')

order = pd.read_excel("global_superstore_2016.xlsx", sheet_name = "Orders")
order['Order Date'] = pd.to_datetime(order['Order Date'])
order['Ship Date'] = pd.to_datetime(order['Ship Date'])
order['Year'] = order['Ship Date'].dt.year

tmp = order.groupby(['Customer ID', 'Country', 'Year']).sum('Sales').sort_values(by = 'Sales', ascending=False).reset_index()
tmp['Postal Code'] = tmp['Postal Code'].astype('int', errors='ignore')

tmp2 = order.groupby(['Product ID', 'Category','Sub-Category', 'Product Name']).agg({'Sales': sum, 'Profit': sum, 'Quantity': sum}).sort_values(by = ['Sales', 'Profit', 'Quantity'], ascending=False).reset_index()
tmp2['Product_ID_Cat'] = tmp2['Product ID'].astype('str') + '_' + tmp2['Sub-Category'].astype('str')

tmp3 = tmp2.iloc[:10, :][['Product_ID_Cat', 'Sales', 'Profit']].set_index('Product_ID_Cat').unstack().to_frame().reset_index()
tmp3.rename({"level_0": "Index", 0: "Values"}, axis=1, inplace=True)
tmp3 = tmp3.sort_values(by = 'Index').reset_index(drop=True)

df1 = order.groupby(['Category', 'Year']).agg({"Sales": sum}).reset_index()
df2 = order.groupby(['Sub-Category', 'Year']).agg({"Sales": sum}).reset_index()

limits = [(0,1000),(1000,5000),(5000,10000),(10000,30000)]
colors = ["lightgrey","crimson","lightseagreen","royalblue"]
YEARS = tmp['Year'].unique().tolist()

def plot_bar(df): 
    fig = px.bar(df, y = 'Product_ID_Cat', x= 'Values', color = 'Index', color_discrete_map={"Sales": "#ffaf4a", 'Profit': "#45b7c2"})
    fig.update_layout(yaxis_title = "Product ID & Sub-Category")
    return fig

def plot_pie_cat(df): 
    fig = go.Figure(data = [go.Pie(
        values = df["Sales"], 
        labels = df['Category']
    )])
    fig.update_traces(marker = dict(colors = ["#2dc2c2", "#f4b80f", "#de663e"]), texttemplate = "%{percent:.1%}")
    fig.update_layout(width=500, height=400)
    return fig

def plot_pie_sub(df): 
    fig = go.Figure(data = [go.Pie(
            values = df['Sales'],
            labels = df['Sub-Category'])],
            )
    fig.update_traces(marker = dict(colors = px.colors.qualitative.T10), texttemplate = "%{percent:.1%}")
    fig.update_layout(width=500, height=400)
    return fig

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
            name = '${0} - ${1}'.format(lim[0],lim[1])))
    
    fig.update_layout(
        showlegend = True,
        geo = dict(
            scope = 'world',
            landcolor = 'rgb(217, 217, 217)',
        ))

    return fig

app = Dash(__name__)

app.layout = html.Div(
    children = [
        html.Div(children=[
            html.H4("Exploring Worldwide Retail Trends: Global Superstore", style={'text-align': 'center'})
        ]),
        html.Div(
            children = [
                html.Div(id = 'left-column', 
                    children= [
                        html.Div(
                            children = [
                                html.P(
                                        id="slider-text",
                                        children="Drag the slider to change the year:",
                                    ),
                                dcc.Slider(id="slct_year",
                                            min=min(YEARS),
                                            max=max(YEARS),
                                            value=min(YEARS),
                                            step = 1,
                                            marks={str(year): {"label": str(year), "style": {"color": "#7fafdf"}} for year in YEARS}),
                                html.Div(children = [html.P("Order by Country in {0}".format(min(YEARS)), id = 'worldmap_name', style={'text-align': 'center', "font-size": "16px"}),
                                    dcc.Graph(id = 'world_map', figure=go.Figure())])
                            ], className= "pretty_container"
                        ),
                        html.Div(children = [
                            html.H3("Top 10 Selling Products Globally 2012-2016"),
                            dcc.Graph(figure = plot_bar(tmp3))], className="pretty_container"
                        )
                ]),
                html.Div(id='right-column',  
                    children = [
                    html.Div(children= [
                        dcc.Dropdown(id = 'chart-dropdown', options = [{'label': year, 'value': year} for year in sorted(YEARS)], value=YEARS[0]), 
                        html.Div(children = [
                            dcc.Graph(id = 'pie-chart-cat'),
                            dcc.Graph(id = 'pie-chart-sub')
                        ], style = {"display": "flex", "flex-direction": "column", "align-items": "center"})
                    ], className="pretty_container")
                ])
            ], style = {"display": "grid", "grid-template-columns": "minmax(280px, 1fr) minmax(280px, 1fr)"}
        )
    ]
)


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
    return "Total Sales by Country in {0}".format(year)

@app.callback(
    Output('pie-chart-cat', 'figure'),
    Input('chart-dropdown', 'value')
)
def update_pie_cat(year): 
    tmp = df1[df1['Year'] == year].sort_values(by ='Category').reset_index(drop=True)
    return plot_pie_cat(tmp)

@app.callback(
    Output('pie-chart-sub', 'figure'),
    Input('chart-dropdown', 'value')
)
def update_pie_sub(year): 
    tmp = df2[df2['Year'] == year].sort_values(by ='Sub-Category').reset_index(drop=True)
    return plot_pie_sub(tmp)


if __name__ == '__main__':
    app.run_server(debug=True)


