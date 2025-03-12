import dash
from dash import dcc, html, callback_context
from dash.dependencies import Input, Output, State, ALL
import json
import re
from openpyxl import load_workbook

# ==========================
# Load Configuration Files
# ==========================
file_path_json = "data/Combined_Assumptions.json"
file_path_mapping_json = "data/input_mapping.json"
file_path_excel = "data/Approximation Valuation Model.xlsx"

# Load assumptions data
try:
    with open(file_path_json, "r") as json_file:
        assumptions_data = json.load(json_file)
except Exception as e:
    assumptions_data = {}
    print(f"Error loading assumptions file: {e}")

# Load input mappings for Excel cell references
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

# Define industries and assumption categories
industries = ["Mobile App", "SaaS"]
categories = list(assumptions_data.keys())

# ==========================
# Helper Functions
# ==========================
def normalize_key(key):
    """Standardize keys to alphanumeric lowercase."""
    return re.sub(r'[^A-Za-z0-9]', '', key.lower())

def flatten_mapping(mapping, parent_key=""):
    """Flatten nested mapping dict to a single-level dict using normalized keys."""
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
    Update Excel cells based on the provided user inputs and mapping.
    If ws[cell_ref] returns a tuple, determine whether it is a tuple of tuples or a simple tuple,
    then iterate accordingly to assign the value.
    """
    wb = load_workbook(file_path_excel)
    ws = wb["Assumption Table"]

    flattened_mappings = flatten_mapping(input_mappings.get(category, {}))
    print(f"Flattened Mappings for {category}: {flattened_mappings}")

    for user_key, value in user_inputs.items():
        normalized_user_key = normalize_key(user_key)
        if normalized_user_key in flattened_mappings:
            cell_ref = flattened_mappings[normalized_user_key]
            # Check that cell_ref is a string; if not, skip updating this input
            if not isinstance(cell_ref, str):
                print(f"Invalid cell reference for {user_key}: {cell_ref}")
                continue
            cell_obj = ws[cell_ref]
            if isinstance(cell_obj, tuple):
                # Check if the first element is iterable (tuple of tuples)
                if cell_obj and hasattr(cell_obj[0], '__iter__'):
                    for row in cell_obj:
                        for c in row:
                            c.value = value
                else:
                    for c in cell_obj:
                        c.value = value
            else:
                cell_obj.value = value
            print(f"Writing {value} to {cell_ref} for {category}")
        else:
            print(f"No mapping found for: {user_key}")
    wb.save(file_path_excel)
    wb.close()
    print(f"Excel file updated for category: {category}")

def recursive_generate_inputs(data, category, prefix=""):
    """
    Recursively generate input fields from the assumptions data.
    Each numeric input is given a pattern‐matching ID containing:
      - type: "input-field"
      - category: the assumption category
      - field: a unique key for the field (built from the nested keys)
    """
    input_fields = []
    for key, value in data.items():
        # Build a field key based on prefix and current key
        field_key = (prefix + "_" + key).strip("_").replace(" ", "_")
        # Skip non-numeric (string) values
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
                        id={
                            "type": "input-field",
                            "category": category,
                            "field": field_key
                        },
                        type="number",
                        value=value,  # Default value from JSON
                        style={'width': '20%'}
                    )
                ], style={'marginBottom': '10px'})
            )
    return input_fields

def generate_all_tab_contents():
    """
    Generate a Div for each assumption category containing its dynamic inputs.
    All Divs are rendered (so input values persist) but only the active tab’s Div is visible.
    """
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

# ==========================
# Home Page Layout
# ==========================
def home_page():
    return html.Div([
        html.Div([
            html.Img(src="/assets/FCAT-image.png", style={'height': '80px', 'marginRight': '15px'}),
            html.H1("Web App Valuation Dashboard", style={'textAlign': 'center', 'flex': '1', 'margin': '0'})
        ], style={'display': 'flex', 'alignItems': 'center', 'justifyContent': 'center', 'gap': '20px'}),

        html.Br(),

        html.Label("Select an Industry:", style={'textAlign': 'left'}),
        dcc.Dropdown(
            id='industry-dropdown',
            options=[{'label': industry, 'value': industry} for industry in industries],
            placeholder="Choose an industry",
            style={'width': '50%'}
        ),

        html.Br(),

        html.Label("Select Assumption Category:", style={'textAlign': 'left'}),
        dcc.Tabs(
            id="category-tabs",
            value=categories[0] if categories else None,
            children=[dcc.Tab(label=cat, value=cat) for cat in categories]
        ),

        # Render all tabs' input fields so that values are preserved
        html.Div(id="all-tab-contents", children=generate_all_tab_contents()),

        html.Br(),

        # Final submission button to update Excel with inputs from all tabs
        html.Button("Submit All", id="submit-all", n_clicks=0),

        html.Br(),
        html.Hr(),
        html.Div(
            html.A("Visit Our Company Website", href="https://fcatalyst.com/overview", target="_blank",
                   style={'color': 'blue', 'textDecoration': 'none', 'fontSize': '18px'}),
            style={'textAlign': 'center', 'marginTop': '20px'}
        )
    ])

# Set the default layout (using a page-content container)
app.layout = html.Div([
    dcc.Location(id='url', refresh=False),
    html.Div(id='page-content', children=home_page())
])

# ==========================
# Callback: Toggle Tab Visibility
# ==========================
@app.callback(
    Output({'type': 'tab-content', 'category': ALL}, 'style'),
    Input("category-tabs", "value"),
    State({'type': 'tab-content', 'category': ALL}, 'id')
)
def toggle_tab_visibility(selected_tab, ids):
    """
    When the user selects a tab, set that category's Div to display and hide the others.
    """
    styles = []
    for item in ids:
        if item["category"] == selected_tab:
            styles.append({"display": "block"})
        else:
            styles.append({"display": "none"})
    return styles

# ==========================
# Callback: Submit All Inputs
# ==========================
@app.callback(
    Output("page-content", "children"),
    Input("submit-all", "n_clicks"),
    State({'type': 'input-field', 'category': ALL, 'field': ALL}, "value"),
    State({'type': 'input-field', 'category': ALL, 'field': ALL}, "id")
)
def submit_all(n_clicks, values, ids):
    """
    Collect all input values from all tabs (even those hidden), organize them by category,
    and update the Excel file for each category.
    """
    if n_clicks < 1:
        raise dash.exceptions.PreventUpdate

    # Organize the input values by category
    all_inputs = {}
    for v, id_dict in zip(values, ids):
        cat = id_dict["category"]
        field = id_dict["field"]
        if cat not in all_inputs:
            all_inputs[cat] = {}
        all_inputs[cat][field] = v

    print("Collected inputs for all categories:", all_inputs)

    # Update Excel for each category
    for cat, inputs in all_inputs.items():
        write_dynamic_assumptions_to_excel(cat, inputs)

    return html.Div([
        html.H1("All Assumptions Successfully Written to Excel!", style={'textAlign': 'center'}),
        html.Button("Back to Home", id="back-home", n_clicks=0, style={'display': 'block', 'margin': 'auto'})
    ])

# ==========================
# Callback: Back to Home Navigation
# ==========================
@app.callback(
    Output("url", "pathname"),
    Input("back-home", "n_clicks"),
    prevent_initial_call=True
)
def navigate_home(n_clicks):
    return "/"

# ==========================
# Run the Dash App
# ==========================
if __name__ == '__main__':
    app.run_server(debug=True)
