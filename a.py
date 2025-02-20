import pandas as pd
import dash
from dash import dcc, html, Input, Output
import plotly.express as px

# Load Data from Excel File
file_path = "Assignment_Task F.xlsx"

df_missed_calls = pd.read_excel(file_path, sheet_name='Missed Call Data')
df_sms = pd.read_excel(file_path, sheet_name='SMS Report')
df_ivr = pd.read_excel(file_path, sheet_name='IVR Report')
df_whatsapp = pd.read_excel(file_path, sheet_name='WhstsApp Response Data')
df_orders = pd.read_excel(file_path, sheet_name='Purchase Data')

# Convert Date Columns
df_missed_calls['Date'] = pd.to_datetime(df_missed_calls['Date'])
df_sms['Time'] = pd.to_datetime(df_sms['Time'])
df_ivr['Date'] = pd.to_datetime(df_ivr['Date'])
df_whatsapp['DATETIME'] = pd.to_datetime(df_whatsapp['DATETIME'])
df_orders['Date'] = pd.to_datetime(df_orders['Date'])

# Aggregate Data for Charts
df_missed_calls_grouped = df_missed_calls[df_missed_calls['In-Out'] == 'IN'].groupby(df_missed_calls['Date'].dt.date).size().reset_index(name='missed_calls')
df_sms_grouped = df_sms.groupby('Status').size().reset_index(name='Count')
df_state_orders = df_orders.groupby(['State', 'Confirmation'])['Order ID'].count().reset_index()
df_state_orders.columns = ['State', 'Status', 'Total Orders']
df_retailers = df_orders.groupby(['Retailer Name', 'Confirmation'])['Number of Packets'].sum().reset_index()
df_farmers = df_orders.groupby(['Farmer Name', 'Confirmation'])['Number of Packets'].sum().reset_index()
df_fa_performance = df_orders.groupby(['FA NAME', 'Confirmation'])['Order ID'].count().reset_index()

# Initialize Dash App
app = dash.Dash(__name__, suppress_callback_exceptions=True)
app.title = "Agriculture Order Analytics"

# Define Layout
app.layout = html.Div([
    html.H1("Agriculture Order Analytics Dashboard", style={'textAlign': 'center'}),

    dcc.Tabs(id="tabs", value='missed_calls', children=[
        dcc.Tab(label='Missed Call Trends', value='missed_calls'),
        dcc.Tab(label='SMS Delivery Analysis', value='sms_delivery'),
        dcc.Tab(label='State-wise Orders', value='state_orders'),
        dcc.Tab(label='Retailer & FA Orders', value='retailer_fa_orders')
    ]),

    html.Div(id='tabs-content')
])

@app.callback(
    Output('tabs-content', 'children'),
    Input('tabs', 'value')
)
def update_tab(tab_name):
    if tab_name == 'missed_calls':
        return html.Div([
            dcc.DatePickerRange(
                id='date_picker',
                start_date=df_missed_calls_grouped['Date'].min(),
                end_date=df_missed_calls_grouped['Date'].max()
            ),
            dcc.Graph(id='missed_calls_chart')
        ])
    elif tab_name == 'sms_delivery':
        return html.Div([
            dcc.Graph(id='sms_chart')
        ])
    elif tab_name == 'state_orders':
        return html.Div([
            dcc.Dropdown(
                id='order_status',
                options=[{'label': s, 'value': s} for s in df_state_orders['Status'].unique()],
                value=df_state_orders['Status'].unique()[0],
                clearable=False
            ),
            dcc.Graph(id='state_orders_chart')
        ])
    elif tab_name == 'retailer_fa_orders':
        return html.Div([
            dcc.Dropdown(
                id='order_type',
                options=[{'label': 'Retailers', 'value': 'Retailers'},
                         {'label': 'Farmers', 'value': 'Farmers'},
                         {'label': 'FA Performance', 'value': 'FA'}],
                value='Retailers',
                clearable=False
            ),
            dcc.Graph(id='retailer_fa_chart')
        ])

@app.callback(
    Output('missed_calls_chart', 'figure'),
    Input('date_picker', 'start_date'),
    Input('date_picker', 'end_date')
)
def update_missed_calls(start_date, end_date):
    filtered_df = df_missed_calls_grouped[(df_missed_calls_grouped['Date'] >= pd.to_datetime(start_date)) &
                                          (df_missed_calls_grouped['Date'] <= pd.to_datetime(end_date))]
    fig = px.line(filtered_df, x='Date', y='missed_calls', title="Missed Calls Over Time", markers=True)
    return fig

@app.callback(
    Output('sms_chart', 'figure'),
    Input('sms_chart', 'id')
)
def update_sms_chart(_):
    fig = px.bar(df_sms_grouped, x='Status', y='Count', title="SMS Delivery Success vs Failure", color='Status')
    return fig

@app.callback(
    Output('state_orders_chart', 'figure'),
    Input('order_status', 'value')
)
def update_state_orders(order_status):
    filtered_df = df_state_orders[df_state_orders['Status'] == order_status]
    fig = px.bar(filtered_df, x='State', y='Total Orders', title=f"Orders per State ({order_status})", color='Total Orders')
    return fig

@app.callback(
    Output('retailer_fa_chart', 'figure'),
    Input('order_type', 'value')
)
def update_retailer_fa_chart(order_type):
    if order_type == 'Retailers':
        fig = px.bar(df_retailers, x='Retailer Name', y='Number of Packets', title="Top Retailers by Order Volume", color='Number of Packets')
    elif order_type == 'Farmers':
        fig = px.bar(df_farmers, x='Farmer Name', y='Number of Packets', title="Top Farmers by Order Volume", color='Number of Packets')
    else:
        fig = px.bar(df_fa_performance, x='FA NAME', y='Order ID', title="FA Performance by Order Count", color='Order ID')
    return fig

if __name__ == '__main__':
    app.run_server(debug=True)
