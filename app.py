import dash
from dash import dcc, html
from dash.dependencies import Input, Output, State, ALL
from dash.exceptions import PreventUpdate
from dash import dash_table
import json
import re
from openpyxl import load_workbook
import os
import pythoncom
import win32com.client as win32
import csv
import pandas as pd
import numpy as np
import plotly.express as px
import urllib.parse
from dash_table.Format import Format, Scheme

# ==========================
# Load Configuration Files
# ==========================
file_path_json = "data/Combined_Assumptions.json"
file_path_mapping_json = "data/input_mapping.json"
file_path_excel = "data/Approximation Valuation Model.xlsx"
file_path_output = "data/Valuation Model.xlsx"

try:
    with open(file_path_json, "r") as json_file:
        assumptions_data = json.load(json_file)
except Exception as e:
    assumptions_data = {}
    print(f"Error loading assumptions file: {e}")

try:
    with open(file_path_mapping_json, "r") as json_file:
        input_mappings = json.load(json_file)
except Exception as e:
    input_mappings = {}
    print(f"Error loading mapping file: {e}")

# ==========================
# Initialize Dash App
# ==========================
app = dash.Dash(__name__, suppress_callback_exceptions=True)
app.title = "Web App Valuation Dashboard"
server = app.server

industries = ["Mobile App", "SaaS"]
categories = list(assumptions_data.keys())

# ==========================
# Helper Functions
# ==========================

def normalize_key(key):

    """ Standardizing categories names """

    return re.sub(r'[^A-Za-z0-9]', '', key.lower())

def flatten_mapping(mapping, parent_key=""):

    """ Turning the nested dictionairy into a flat dictionary """

    flattened = {}
    for key, value in mapping.items():
        full_key = f"{parent_key} - {key}".strip() if parent_key else key.strip()
        if isinstance(value, dict):
            flattened.update(flatten_mapping(value, full_key))
        else:
            flattened[normalize_key(full_key)] = value
    return flattened

def write_dynamic_assumptions_to_excel(category, user_inputs):
    """
    This function is reading an Excel file, mapping user input values to 
    specific cell references (from a flattened mapping), and writing those values dynamically.
    """
    wb = load_workbook(file_path_excel)
    ws = wb["Assumption Table"]
    flattened_mappings = flatten_mapping(input_mappings.get(category, {}))
    for user_key, value in user_inputs.items():
        normalized_user_key = normalize_key(user_key)
        if normalized_user_key in flattened_mappings:
            cell_ref = flattened_mappings[normalized_user_key]
            if not isinstance(cell_ref, str):
                continue
            cell_obj = ws[cell_ref]
            if isinstance(cell_obj, tuple):
                if cell_obj and hasattr(cell_obj[0], '__iter__'):
                    for row in cell_obj:
                        for c in row:
                            c.value = value
                else:
                    for c in cell_obj:
                        c.value = value
            else:
                cell_obj.value = value
    wb.save(file_path_excel)
    wb.close()

def recursive_generate_inputs(data, category, prefix=""):
    input_fields = []
    for key, value in data.items():
        field_key = (prefix + "_" + key).strip("_").replace(" ", "_")
        if isinstance(value, str):
            continue
        if isinstance(value, dict):
            input_fields.append(html.H3(key))
            input_fields.extend(recursive_generate_inputs(value, category, prefix=field_key))
        else:
            input_fields.append(
                html.Div([
                    html.Label(f"{prefix} - {key}".strip("-"), style={'marginRight': '10px'}),
                    dcc.Input(
                        id={"type": "input-field", "category": category, "field": field_key},
                        type="number",
                        value=value,
                        style={'width': '20%'}
                    )
                ], style={'marginBottom': '10px'})
            )
    return input_fields

def generate_all_tab_contents():
    tab_contents = []
    for cat in categories:
        tab_contents.append(
            html.Div(
                id={'type': 'tab-content', 'category': cat},
                children=recursive_generate_inputs(assumptions_data.get(cat, {}), cat),
                style={'display': 'block' if cat == categories[0] else 'none'}
            )
        )
    return tab_contents

def read_entire_sheet(excel_path, sheet_name):
    wb = load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name]
    values = []
    for row in ws.iter_rows(values_only=True):
        values.append(list(row))
    wb.close()
    return values

def force_recalculation_win32(excel_path):
    try:
        pythoncom.CoInitialize()
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        abs_path = os.path.abspath(excel_path)
        wb = excel.Workbooks.Open(abs_path)
        wb.Application.CalculateFullRebuild()
        wb.Save()
        wb.Close(False)
        excel.Quit()
        pythoncom.CoUninitialize()
    except Exception as e:
        print("Error in force recalculation with win32com:", e)

def export_sheets_to_excel(sheets_dict, filename="data/Valuation Model.xlsx"):
    with pd.ExcelWriter(filename) as writer:
        for sheet_name, data in sheets_dict.items():
            df = pd.DataFrame(data)
            df.to_excel(writer, sheet_name=sheet_name, index=False)

def clean_sheet_data(data):
    """
    Removes rows and columns that are completely empty from a 2D list.
    A row is considered empty if all its cells are None or empty strings.
    A column is removed if every row in that column is empty.
    """
    cleaned_rows = [row for row in data if any(cell not in [None, ""] for cell in row)]
    if not cleaned_rows:
        return cleaned_rows

    num_cols = max(len(row) for row in cleaned_rows)
    padded_rows = [row + [None]*(num_cols - len(row)) for row in cleaned_rows]

    valid_cols = []
    for col in range(num_cols):
        col_values = [row[col] for row in padded_rows]
        if any(cell not in [None, ""] for cell in col_values):
            valid_cols.append(col)

    cleaned_data = [[row[col] for col in valid_cols] for row in padded_rows]
    return cleaned_data

def insert_year_row(data, year_labels, insert_index=0, start_col=2):
    """
    Inserts a new row at 'insert_index' in the 2D list 'data'.
    Only columns from start_col..start_col+len(year_labels)-1 get the year labels;
    other cells remain blank.
    """
    if not data:
        return data
    # Determining max row length
    max_len = max(len(row) for row in data)
    # Pad rows
    for row in data:
        if len(row) < max_len:
            row.extend([None]*(max_len - len(row)))

    # Building the new row
    new_row = ["" for _ in range(max_len)]
    for i, label in enumerate(year_labels):
        col_idx = start_col + i
        if col_idx < max_len:
            new_row[col_idx] = label

    data.insert(insert_index, new_row)
    return data

def build_data_table(values, title="Sheet", skip_clean=False):
    """
    Builds a Dash DataTable from a 2D list, optionally skipping the cleaning step.
    """
    if not skip_clean:
        values = clean_sheet_data(values)
    if not values:
        return html.Div([html.H2(title), html.P("No data found.")])

    df = pd.DataFrame(values)

    # build columns with formatting
    cols = []
    for col in df.columns:
        cols.append({
            "name": str(col),
            "id": str(col),
            "type": "numeric",
            "format": Format(
                precision=2,
                scheme=Scheme.fixed
            ).group(True)
        })

    return html.Div([
        html.H2(title, style={'textAlign': 'center'}),
        dash_table.DataTable(
            data=df.to_dict('records'),
            columns=cols,
            style_table={'overflowX': 'auto', 'margin': 'auto'},
            style_cell={'textAlign': 'center'},
            page_size=10
        )
    ])

# ==========================
# Layouts
# ==========================

def home_page():
    return html.Div([
        html.Div([
            html.Img(src="/assets/FCAT-image.png", style={'height': '80px', 'marginRight': '15px'}),
            html.H1("Web App Valuation Dashboard", style={'textAlign': 'center', 'flex': '1', 'margin': '0'})
        ], style={'display': 'flex','alignItems': 'center','justifyContent': 'center','gap': '20px'}),
        html.Br(),
        html.Label("Select an Industry:"),
        dcc.Dropdown(
            id='industry-dropdown',
            options=[{'label': ind, 'value': ind} for ind in industries],
            placeholder="Choose an industry",
            style={'width': '50%'}
        ),
        html.Br(),
        html.Label("Select Assumption Category:"),
        dcc.Tabs(
            id="category-tabs",
            value=categories[0] if categories else None,
            children=[dcc.Tab(label=cat, value=cat) for cat in categories]
        ),
        html.Div(id="all-tab-contents", children=generate_all_tab_contents()),
        html.Br(),
        html.Button("Submit All", id="submit-all", n_clicks=0),
        html.Br(),
        html.Hr(),
        html.Div(
            html.A("Visit Our Company Website", href="https://fcatalyst.com/overview", target="_blank",
                   style={'color': 'blue','textDecoration': 'none','fontSize': '18px'}),
            style={'textAlign': 'center','marginTop': '20px'}
        )
    ])

def valuation_selector_page():
    return html.Div([
        html.H1("Select Valuation Type", style={'textAlign': 'center'}),
        dcc.Dropdown(
            id="valuation-type-dropdown",
            options=[
                {"label": "FCF to Equity", "value": "fcf_equity"},
                {"label": "FCF to the Firm", "value": "fcf_firm"}
            ],
            placeholder="Choose a valuation type",
            style={'width': '50%', 'margin': 'auto'}
        ),
        html.Br(),
        html.Button("Back to Home", id={"type": "back-home", "index": "selector"}, n_clicks=0, style={'display': 'block','margin': 'auto'}),
        # Hidden dummy change button to ensure callback inputs exist
        html.Button("Dummy", id="change-valuation-type", n_clicks=0, style={'display': 'none'})
    ])

def valuation_page(valuation_type):
    # Read data from the "Valuation Model" sheet from file_path_output
    data = read_entire_sheet(file_path_output, "Valuation Model")
    # print("Raw data from Valuation Model sheet:", data)
    
    # Slice the data based on valuation_type
    if valuation_type == "fcf_equity":
        filtered_data = data[7:21]   
        title = "FCF to Equity"
    else:
        filtered_data = data[23:33]  
        title = "FCF to the Firm"
    
    # Clean the data
    cleaned_data = clean_sheet_data(filtered_data)
    
    # Insert a custom header row for Year labels into the cleaned data
    year_labels = ["Year 1", "Year 2", "Year 3", "Year 4", "Year 5"]
    cleaned_data = insert_year_row(cleaned_data, year_labels, insert_index=0, start_col=2)
    # print("Data after inserting Year row:", cleaned_data)
    
    # Build the DataTable using the cleaned data (skip cleaning again)
    table_component = build_data_table(cleaned_data, title=title, skip_clean=True)
    
    # Create a list of row labels from column 0 (skipping blank/None)
    row_labels = []
    for row in cleaned_data:
        if row and row[0] not in [None, ""] and isinstance(row[0], str):
            row_labels.append(row[0])
    
    # Store the final cleaned_data in a hidden dcc.Store so callbacks can read it
    data_store = dcc.Store(id="valuation-data-store", data=cleaned_data)

    # Build a dropdown that enumerates the row labels
    label_dropdown = dcc.Dropdown(
        id="label-dropdown",
        options=[{"label": lbl, "value": lbl} for lbl in row_labels],
        placeholder="Choose a row label to visualize",
        style={'width': '50%', 'margin': 'auto'}
    )

    # A Graph that will show the bar chart
    graph_component = dcc.Graph(id="valuation-graph")

    hidden_dropdown = dcc.Dropdown(
        id="valuation-type-dropdown",
        value=valuation_type,
        style={'display': 'none'}
    )

    # ─── Monte Carlo run on page load ───
    arr, mean, std, lo, hi = run_monte_carlo(n_trials=1000)

    mc_store = dcc.Store(id="mc-mean-store", data={"mean": mean})

    # build summary
    mc_summary = html.Div([
        html.P(f"Estimated Final Value: ${mean:,.0f}", style={'margin':'4px 0'}),
        html.P(f"±1 σ variation: ±${std:,.0f}", style={'margin':'4px 0'}),
        html.P(f"95% CI: ${lo:,.0f} – ${hi:,.0f}", style={'margin':'4px 0'})
    ], style={'textAlign':'center', 'marginBottom':'20px'})

    # build histogram
    df_mc = pd.DataFrame({"Final Value": arr})
    mc_fig = px.histogram(
        df_mc, x="Final Value", nbins=30,
        title=f"Final Equity Value Distribution\nMean=${mean:,.0f} ±${std:,.0f}"
    )
    mc_fig.update_layout(
        xaxis_title="Final Equity Value ($)",
        yaxis_title="Frequency",
        xaxis_tickformat=",.0f"
    )
    mc_fig.update_traces(
        marker_color="skyblue",
        marker_line_color="black", marker_line_width=1, opacity=0.8
    )
    mc_graph = dcc.Graph(figure=mc_fig)

    # pull the Sensitivity table
    sens_df = read_sensitivity_table()
    sens_table = dash_table.DataTable(
        id="sensitivity-table",
        data=sens_df.to_dict("records"),
        columns=[
            {
              "name": col,
              "id": col,
              "editable": (col=="Input Change"),  # only this column editable
              "type": "numeric",
              "format": Format(precision=6, scheme=Scheme.fixed).group(True)
            }
            for col in sens_df.columns
        ],
        style_table={'overflowX': 'auto'},
        style_cell={'textAlign': 'left'},
        page_action="none",
        style_header={'fontWeight':'bold'},
        fixed_rows={'headers': True},
        style_as_list_view=True,
    )
    
    return html.Div([
        hidden_dropdown,
        data_store,
        html.H1("Valuation Results", style={'textAlign': 'center'}),
        html.Div([
            html.Button("Change Valuation Type", id="change-valuation-type", n_clicks=0),
            html.Button("Back to Home", id={"type": "back-home", "index": "valuation"}, n_clicks=0, style={'marginLeft': '20px'})
        ], style={'textAlign': 'center'}),
        html.Br(),
        table_component,
        html.Br(),
        html.H3("Pick a row label to visualize:", style={'textAlign': 'center'}),
        label_dropdown,
        html.Br(),
        graph_component,
        mc_store,
        html.H3("Monte Carlo Simulation", style={'textAlign':'center','marginTop':'30px'}),
        mc_summary,
        mc_graph,
        html.H3("Sensitivity Analysis", style={'textAlign':'center'}),
        sens_table
    ])

def run_monte_carlo(n_trials=1000):
    """Returns (final_values_array, mean, std, ci_lower, ci_upper)."""
    # 1) launch Excel
    pythoncom.CoInitialize()                      
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    path = os.path.abspath(file_path_excel)
    wb   = excel.Workbooks.Open(path)
    sht  = wb.Sheets["Valuation Model"]

    final_vals = []
    for _ in range(n_trials):
        # force *every* volatile formula (RAND, NORMINV, etc) to reevaluate
        excel.CalculateFullRebuild()
        raw = sht.Range("E18").Value              # your NPV cell
        final_vals.append(float(raw))

    # cleanup
    wb.Close(False)
    excel.Quit()
    pythoncom.CoUninitialize()

    arr   = np.array(final_vals)
    mean  = arr.mean()
    std   = arr.std(ddof=1)
    ci_lo = mean - 1.96*std
    ci_hi = mean + 1.96*std
    return arr, mean, std, ci_lo, ci_hi

def read_sensitivity_table():
    """
    Reads the “Sensitivity” block from the Valuation Model sheet,
    using the fact that the header row always starts at L6 and runs
    through Q6 (6 columns: Inputs → Final Estimated Value).
    Returns a pandas.DataFrame.
    """
    wb = load_workbook(file_path_excel, data_only=True)
    ws = wb["Valuation Model"]

    header_row = 38
    start_col = 4  # column L
    end_col   = 9  # column Q

    # Read header cells L6:Q6
    headers = [
        ws.cell(row=header_row, column=col).value
        for col in range(start_col, end_col+1)
    ]

    # Now read the data rows immediately below until an empty “Inputs” cell
    data = []
    for r in range(header_row+1, ws.max_row+1):
        first = ws.cell(row=r, column=start_col).value
        if first in (None, ""):
            break
        row_vals = [
            ws.cell(row=r, column=col).value
            for col in range(start_col, end_col+1)
        ]
        data.append(row_vals)

    wb.close()

    # Build and return a DataFrame
    return pd.DataFrame(data, columns=headers)

# ==========================
# Main Layout
# ==========================
app.layout = html.Div([
    dcc.Location(id='url', refresh=False),
    html.Div(id='page-content', children=home_page())
])

# ==========================
# Callbacks
# ==========================

@app.callback(
    Output({'type': 'tab-content', 'category': ALL}, 'style'),
    Input("category-tabs", "value"),
    State({'type': 'tab-content', 'category': ALL}, 'id')
)
def toggle_tab_visibility(selected_tab, ids):
    return [{"display": "block" if item["category"] == selected_tab else "none"} for item in ids]

@app.callback(
    Output("url", "pathname", allow_duplicate=True),
    Input("submit-all", "n_clicks"),
    State({'type': 'input-field', 'category': ALL, 'field': ALL}, "value"),
    State({'type': 'input-field', 'category': ALL, 'field': ALL}, "id"),
    prevent_initial_call='initial_duplicate'
)
def submit_all(n_clicks, values, ids):
    if n_clicks < 1:
        raise dash.exceptions.PreventUpdate
    all_inputs = {}
    for v, id_dict in zip(values, ids):
        cat = id_dict["category"]
        field = id_dict["field"]
        if cat not in all_inputs:
            all_inputs[cat] = {}
        all_inputs[cat][field] = v
    for cat, inputs in all_inputs.items():
        write_dynamic_assumptions_to_excel(cat, inputs)
    force_recalculation_win32(file_path_excel)
    assumption_data = read_entire_sheet(file_path_excel, "Assumption Table")
    projection_data = read_entire_sheet(file_path_excel, "Projection Model")
    valuation_data = read_entire_sheet(file_path_excel, "Valuation Model")
    sheets_dict = {
        "Assumption Table": assumption_data,
        "Projection Model": projection_data,
        "Valuation Model": valuation_data
    }
    export_sheets_to_excel(sheets_dict, "data/Valuation Model.xlsx")
    return "/valuation-selector"

@app.callback(
    Output("url", "pathname", allow_duplicate=True),
    Output("url", "search", allow_duplicate=True),
    Input("valuation-type-dropdown", "value"),
    Input("change-valuation-type", "n_clicks"),
    prevent_initial_call='initial_duplicate'
)
def handle_valuation_navigation(val_type, change_clicks):
    ctx = dash.callback_context
    if not ctx.triggered:
        raise dash.exceptions.PreventUpdate
    triggered_id = ctx.triggered[0]["prop_id"].split(".")[0]
    if triggered_id == "valuation-type-dropdown":
        if val_type in ("fcf_equity", "fcf_firm"):
            return "/valuation", f"?type={val_type}"
        else:
            raise dash.exceptions.PreventUpdate
    elif triggered_id == "change-valuation-type":
        return "/valuation-selector", ""
    raise dash.exceptions.PreventUpdate

@app.callback(
    Output("valuation-graph", "figure"),
    Input("label-dropdown", "value"),
    State("valuation-data-store", "data")
)
def update_valuation_chart(selected_label, data):
    if not selected_label or not data:
        raise dash.exceptions.PreventUpdate

    # 1) Find the row whose column-0 value matches selected_label
    numeric_row = None
    for row in data:
        if row and row[0] == selected_label:
            # 2) Convert columns 1..n to floats, ignoring None/empty
            try:
                numeric_row = [float(cell) for cell in row[1:] if cell not in [None, ""]]
            except Exception:
                numeric_row = None
            break

    # 3) If no numeric row found or conversion failed, return an empty figure
    if not numeric_row:
        return px.scatter(title=f"No numeric data found for {selected_label}")

    # 4) Build a bar chart from numeric_row
    x_labels = [f"Year {i+1}" for i in range(len(numeric_row))]
    df = pd.DataFrame({"Time Period": x_labels, "Value in $": numeric_row})
    fig = px.bar(df, x="Time Period", y="Value in $", title=f"{selected_label} Over Time Period of 5 Years")

    # 5) Annotate each bar with its value, formatted
    fig.update_traces(
        texttemplate="%{y:,.2f}",
        textposition="outside"
    )

    return fig

@app.callback(
    Output("page-content", "children"),
    Input("url", "pathname"),
    State("url", "search")
)
def display_page(pathname, search):
    if pathname == "/":
        return home_page()
    elif pathname == "/valuation-selector":
        return valuation_selector_page()
    elif pathname == "/valuation":
        query_params = dict(urllib.parse.parse_qsl(search.lstrip("?") or ""))
        val_type = query_params.get("type", None)
        if val_type not in ("fcf_equity", "fcf_firm"):
            return html.Div([
                html.H2("Unknown valuation type"),
                html.Button("Back to Home", id={"type": "back-home", "index": "error"}, n_clicks=0)
            ])
        return valuation_page(val_type)
    else:
        return html.Div([
            html.H2("404 - Page not found"),
            html.Button("Back to Home", id={"type": "back-home", "index": "404"}, n_clicks=0)
        ])

# Callback to catch edits, write back to Excel, recalc, and refresh the table
@app.callback(
    Output("sensitivity-table", "data"),
    Input("sensitivity-table", "data_timestamp"),  # fires on any edit
    State("sensitivity-table", "data"),
    State("mc-mean-store", "data"),
    prevent_initial_call=True
)

def update_sensitivity_table(data_timestamp, rows, mc_store):
    if data_timestamp is None:
        raise PreventUpdate

    wb = load_workbook(file_path_excel)
    ws = wb["Valuation Model"]

    header_row = 38
    start_col  = 4
    end_col    = 9

    # grab the headers from D38:I38
    headers = [
        ws.cell(row=header_row, column=c).value
        for c in range(start_col, end_col+1)
    ]
    input_col = headers.index("Input Change") + start_col

    # write each edited row back by matching column D
    for rec in rows:
        label = rec[headers[0]]
        new_val = rec["Input Change"]
        for r in range(header_row+1, ws.max_row+1):
            if ws.cell(row=r, column=start_col).value == label:
                ws.cell(row=r, column=input_col).value = new_val
                break

    wb.save(file_path_excel)
    wb.close()

    force_recalculation_win32(file_path_excel)
    updated_df = read_sensitivity_table()

    # Assume your DF has a column named exactly "Total Percent Change"
    mc_mean = mc_store['mean']
    if "Total Percent Change" in updated_df.columns:
        # compute new final value = mc_mean * (1 + total_pct)
        updated_df["Final Estimated Value"] = (
            mc_mean * (1 + updated_df["Total Percent Change"])
        )

    return updated_df.to_dict("records")

@app.callback(
    Output("url", "pathname", allow_duplicate=True),
    Input({"type": "back-home", "index": ALL}, "n_clicks"),
    prevent_initial_call='initial_duplicate'
)

def go_home(btns):
    ctx = dash.callback_context
    if not ctx.triggered:
        raise dash.exceptions.PreventUpdate
    return "/"

# ==========================
# Run the App
# ==========================
if __name__ == '__main__':
    app.run_server(debug=True)
