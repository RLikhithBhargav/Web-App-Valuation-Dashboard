import dash
from dash import dcc, html
import pandas as pd
from dash.dependencies import Input, Output, State
import urllib.parse  # To handle URL decoding

# Load mobile app company names from an Excel file
file_path = "data/Mobile Apps.xlsx" 
df = pd.read_excel(file_path, sheet_name="Companies")  
# Extract company names from the first column
company_list = df.iloc[:, 0].dropna().tolist()

# Initialize Dash app
app = dash.Dash(__name__, suppress_callback_exceptions=True) #Allows Handling Pages Dynamically
app.title = 'Web App Valuation Dashboard'

# Define financial metrics
metrics_list = ["KPIs Assumptions", "Revenue Assumptions", 
                "Costs of Goods Sold Assumption", "Expenses Assumption", 
                "Capital Expenditure & Depreciation Assumptions", "Long Term Debt Assumptions",
                "Working Capital Assumptions", "Cost of Capital Assumptions", "Capital Structure",
                "Sustainable Growth Rate Assumptions", "Total number of Shares Outstanding"]

# App layout with multi-page support
app.layout = html.Div([
    dcc.Location(id='url', refresh=False),  # Handles page navigation
    html.Div(id='page-content')  # Content container for different pages
])

# Layout of the homepage
def home_page(): 
    return html.Div([

    html.Div([
        html.Img(src="/assets/FCAT-image.png", style={'height': '80px', 'marginRight': '15px'}),  # Your company logo
        html.H1("Web App Valuation Dashboard", style={'textAlign': 'center', 'flex': '1', 'margin': '0'})
    ], style={'display': 'flex', 'alignItems': 'center', 'justifyContent': 'center', 'gap': '20px'}),

    # Description Section
    html.P(
        "Introduction of our Purpose",
        style={'fontSize': '18px', 'lineHeight': '1.5', 'maxWidth': '800px'}
    ),

    html.Br(),

    # Dropdown for selecting a company
    html.Label("Select a Mobile App Company:"),
    dcc.Dropdown(
        id='company-dropdown',
        options=[{'label': company, 'value': company} for company in company_list],
        placeholder="Choose a company",
        style={'width': '50%'}
    ),

    html.Br(),

    # Display selected company
    html.Div(id='selected-company-output', style={'fontSize': 20, 'marginTop': 20}),

    html.Br(),

    # "Next" button (Hidden initially)
    html.Div([
        dcc.Link(html.Button("Next", id="next-button", n_clicks=0), href="", id="next-page-link"),
    ], style={'textAlign': 'center', 'marginTop': '20px'}),

    # Footer with company's website link
    html.Hr(),  # Separator line
    html.Div(
        html.A("Visit Our Company Website", href="https://fcatalyst.com/overview", target="_blank",
               style={'color': 'blue', 'textDecoration': 'none', 'fontSize': '18px'}),
        style={'textAlign': 'center', 'marginTop': '20px'}
    )
])

# Callback to show selected company and reveal "Next" button
@app.callback(
    [Output('selected-company-output', 'children'),
     Output('next-page-link', 'href')],
    Input('company-dropdown', 'value'))

def update_output(selected_company):
    if selected_company:
        return f"You selected: {selected_company}", f"/metrics?company={urllib.parse.quote(selected_company)}"
    return "Please select a company.", "/"

# Callback for "Next" button to navigate to Metrics Selection Page
# @app.callback(
#     Output('url', 'pathname'),
#     Input('next-button', 'n_clicks'),
#     State('company-dropdown', 'value'),
#     prevent_initial_call=True)

# def navigate_to_metrics(n_clicks, selected_company):
#     if selected_company:
#         return f"/metrics?company={urllib.parse.quote(selected_company)}"
#     return "/"

def metrics_selection_page(selected_company):
    return html.Div([
        html.H1(f"Select Metrics for {selected_company}", style={'textAlign': 'center'}),

       dcc.Checklist(
            id='metrics-checklist',
            options=[{'label': metric, 'value': metric} for metric in metrics_list],
            value=[],
            style={'margin': '20px'}
        ),
        
        html.Button("Next", id="next-metrics-button", n_clicks=0, style={'display': 'block', 'margin': 'auto'}),
       
        # html.Div([
        #     html.Button(metric, id={'type': 'metric-button', 'index': metric}, n_clicks=0, 
        #                 style={'display': 'block', 'width': '60%', 'margin': '10px auto', 'padding': '10px'})
        #     for metric in metrics_list
        # ], style={'textAlign': 'center'}),

        html.Br(),

        html.Div([
            dcc.Link("Back to Home", href="/", style={'color': 'blue', 'textDecoration': 'none'})
        ], style={'textAlign': 'center', 'marginTop': '20px'})
    ])

@app.callback(
    Output('page-content', 'children'),
    Input('url', 'pathname'),
    State('url', 'search'))

def display_page(pathname, search):
    if pathname == '/metrics' and search:
        query_params = urllib.parse.parse_qs(search.lstrip('?'))
        selected_company = query_params.get("company", [None])[0]
        if selected_company:
            return metrics_selection_page(selected_company)
    return home_page()  # Default: Show home page  # Default: Show home page


# Run the Dash app
if __name__ == '__main__':
    app.run_server(debug=True)