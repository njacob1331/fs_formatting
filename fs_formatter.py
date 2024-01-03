#!/usr/bin/env python
# coding: utf-8

# In[67]:


import dash
import dash_bootstrap_components as dbc
from dash import Input, Output, State, dcc, html, dash_table
from dash.exceptions import PreventUpdate

import base64
import io
from io import BytesIO
import os

import numpy as np
import pandas as pd

import openpyxl
from xlrd import open_workbook
import xlsxwriter
#from xlrd import open_workbook
#import xlsxwriter

from datetime import datetime

app = dash.Dash(external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server

# In[68]:


def find_header_row(df):
    
    for test in range(100):
        
        filtered = df[test:]
        
        if not any(pd.isna(filtered.iloc[0])):
            
            filtered.columns = filtered.iloc[0].values
            filtered = filtered[1:]
            filtered.reset_index(inplace=True, drop=True)
            
            return filtered
    
    return df


# In[69]:


def format_fs(data, cpt, mod, sos):
    
    # remove total row if one exists
    data = data.dropna(how='all', subset=[cpt, mod])
    
    # backfill cpt if blank for 26 and TC component
    data[cpt] = data[cpt].ffill()
    
    # replace NA modifiers with blank str
    data[mod].fillna("", inplace=True)
    
    # convert cpt and mod cols to str for easier concatenation
    data[cpt] = data[cpt].astype(str)
    data[mod] = data[mod].astype(str)
    
    # remove any extra whitespace from modifiers
    data[mod] = data[mod].apply(lambda x: x.strip())
    
    if sos == 'NA':
        
        
        cpt_mod = list(zip(data[cpt], data[mod]))
        cpt_mod = [cpt if any([mod == "", mod == '00']) else cpt + "-" + mod for cpt, mod in cpt_mod]
        
        # insert cpt-mod col
        insert_index = data.columns.get_loc(mod)
        data.insert(loc = insert_index + 1, column = 'CPT-MOD', value = cpt_mod)
        
        return data
        
    else:
        
        data[sos].fillna("", inplace=True)
        
        sos_values = sorted(list(data[sos].unique()))
        
        if len(sos_values) == 3:
            
            data_fac = data[ (data[mod] == sos_values[0]) | (data[mod] == sos_values[1]) ]
            data_nf = data[ (data[mod] == sos_values[0]) | (data[mod] == sos_values[2]) ]
            
        else:
            
            data_fac = data[data[mod] == sos_values[0]]
            data_nf = data[data[mod] == sos_values[1]]
            
        data_fac.reset_index(drop=True, inplace=True)
        data_nf.reset_index(drop=True, inplace=True)
            
        cpt_mod_fac = list(zip(data_fac[cpt], data_fac[mod]))
        cpt_mod_fac = [cpt if any([mod == "", mod == '00']) else cpt + "-" + mod for cpt, mod in cpt_mod_fac]
        
        cpt_mod_nf = list(zip(data_nf[cpt], data_nf[mod]))
        cpt_mod_nf = [cpt if any([mod == "", mod == '00']) else cpt + "-" + mod for cpt, mod in cpt_mod_nf]
        
        # insert cpt-mod col
        insert_index = data_fac.columns.get_loc(mod)
        data_fac.insert(loc = insert_index + 1, column = 'CPT-MOD', value = cpt_mod_fac)
        data_nf.insert(loc = insert_index + 1, column = 'CPT-MOD', value = cpt_mod_nf)
        
        return data_fac, data_nf


# ### Header for all pages

# In[70]:


header = html.H2("Fee Schedule Formatter", style={'textAlign': 'center'})


# ### Sidebar for page navigation

# In[71]:


# the style arguments for the sidebar. We use position:fixed and a fixed width
SIDEBAR_STYLE = {
    "position": "fixed",
    "top": 0,
    "left": 0,
    "bottom": 0,
    "width": "14rem",
    "padding": "2rem 1rem",
    "background-color": "#f8f9fa",
}

# the styles for the main content position it to the right of the sidebar and
# add some padding.
CONTENT_STYLE = {
    "margin-left": "16rem",
    "margin-right": "2rem",
    "padding": "2rem 1rem",
}

sidebar = html.Div(
    [
        html.H2("Menu", style={'textAlign': 'center'}),
        html.Hr(),
        dbc.Nav(
            [
                dbc.NavLink("Select Files", href="/", active="exact"),
                dbc.NavLink("Select Columns", href="/page-1", active="exact"),
                dbc.NavLink("Format Fee Schedules", href="/page-2", active="exact"),
                dbc.NavLink("Download Files", href="/page-3", active="exact"),
                
            ],
            vertical=True,
            pills=True,
        ),
    ],
    style=SIDEBAR_STYLE,
)


# ### Local memory storage for uploaded contents

# In[72]:


filename_store = dcc.Store(id='files', storage_type='local')
data_store = dcc.Store(id='fee-schedules', storage_type='local')
colname_store = dcc.Store(id='cols', storage_type='local')
col_selections_store = dcc.Store(id='dropdowns', storage_type='local')


# ### File upload page

# In[73]:


page1 = html.Div([
    header,
    html.Hr(),
    dcc.Upload(
        id='upload-data',
        children=html.Div([
            html.A('Select Files')
        ]),
        style={
            'width': '100%',
            'height': '60px',
            'lineHeight': '60px',
            'borderWidth': '1px',
            'borderStyle': 'dashed',
            'borderRadius': '5px',
            'textAlign': 'center',
            'margin': '10px'
        },
        # Allow multiple files to be uploaded
        multiple=True
    ),
    html.Br(),
    dcc.Loading([html.Div(id='output-data-upload')]),
    
])


# ### Column selection page

# In[74]:


page2 = html.Div([
    header,
    html.Hr(),
    html.Br(),
    html.Div([
        html.Label('CPT:'),
        dcc.Dropdown(id='cpt'),
    ], style={'margin-bottom': '20px'}),

    html.Div([
        html.Label('Modifier:'),
        dcc.Dropdown(id='mod'),
    ], style={'margin-bottom': '20px'}),

    html.Div([
        html.Label('Site of Service (if applicable):'),
        dcc.Dropdown(id='sos'),
    ], style={'margin-bottom': '20px'}),
])


# ### Formatting Page

# In[75]:


table_cols= ['Filename', 'Status']
return_frame = pd.DataFrame().to_dict('records')


page3 = html.Div([
    header,
    html.Hr(),
    html.Br(),
    
    html.Button("Process Fee Schedules", id="fs-process", disabled=False),
    
    html.Div([
        html.Br(), 
        dash_table.DataTable(
            id = 'format-status',
            data = return_frame,
            editable = True,
            columns = [{'name': i, 'id': i} for i in table_cols],
            style_cell = {
                'textAlign': 'center',
                'whiteSpace': 'normal',  
                'height': 'auto',
                'minWidth': 95, 
                'maxWidth': 95, 
                'width': 95,
            },
            style_as_list_view=False)
    ])
])


# ### Download Page

# In[76]:


page4 = html.Div(
    [
        header,
        html.Hr(),
        html.Br(),
        html.Div(
            [
                html.H5('Click Here to Download Your Files:'),
                html.Br(),
                html.Button("Download Fee Schedules", id="download-button"),
                dcc.Download(id="push-download"),
                html.Br(),
                html.Div(dcc.Loading(children=[html.Br(), html.Div(id="saved-message")]), style={'textAlign': 'center'})
                
            ], #style={'display': 'inline-block'},
        )
    ],
)


# ### Function for processing uploaded files

# In[77]:


def load_content(rc, rf):
    
    uploaded_f = []
    uploaded_c = []
    
    # load content
    for c, f in zip(rc, rf):

        try:

            content_type, content_string = c.split(',')
            decoded = base64.b64decode(content_string)

            xls = pd.ExcelFile(io.BytesIO(decoded))

            if xls.engine == 'openpyxl':

                sheets = xls.book.worksheets
                fs = [sheet.title for sheet in sheets if sheet.sheet_state == 'visible'][0]

            else:

                sheets = xls.book.sheets()
                fs = [sheet.name for sheet in sheets if sheet.visibility == 0][0]

            data = pd.read_excel(xls, sheet_name=fs, engine=xls.engine, thousands=',')
            data = find_header_row(data)
            
            result = data.to_dict('records')
            uploaded_f.append(f)
            uploaded_c.append(result)

        except:

            f_error = f'Error loading {f}. This file will be skipped and requires manual formatting.'
            c_error = pd.DataFrame().to_dict('records')
            uploaded_f.append(f_error)
            uploaded_c.append(c_error)
            
        
        
    return uploaded_c, uploaded_f


# ### Function for getting column names for dropdown

# In[78]:


def get_cols(data):

    fs_list = [pd.DataFrame(fs) for fs in data]
    column_names = [fs.columns.tolist() for fs in fs_list]
    column_names = list(set(element for sublist in column_names for element in sublist))
    column_names.append('NA')
    
    return column_names


# ### Function for uploading + displaying files

# In[79]:


@app.callback(
    [Output('files', 'data'),
     Output('fee-schedules', 'data'),
     Output('cols', 'data')
    
    ],
    
    # uploads through button
    [Input('upload-data', 'filename')],
    [Input('upload-data', 'contents')],
    
    # data stored in dcc store
    [State('files', 'data')],
    [State('fee-schedules', 'data')]
)
def update_uploaded_files(selected_files, selected_contents, stored_files, stored_contents):
    
    if selected_files is None:
        
        if stored_files is None:
            
            return_files = 'No Files Have Been Uploaded.'
            return_contents = pd.DataFrame().to_dict('records')
            cols = get_cols(return_contents)
            
            return return_files, return_contents, cols
            
        else:
            
            return_files = stored_files
            return_contents = stored_contents
            cols = get_cols(return_contents)
            
            return return_files, return_contents, cols
    
    else:
        
        return_files = selected_files
        return_contents = selected_contents
        
        return_contents, return_files = load_content(return_contents, return_files)
        cols = get_cols(return_contents)
    
    return return_files, return_contents, cols
    
@app.callback(
    Output('output-data-upload', 'children'),

    [Input('files', 'data')],
    [State('fee-schedules', 'data')],
)
def display_uploaded_files(uploaded_files, stored_data):
    
    if not uploaded_files:
        return html.Div()
    
    return html.Div([html.H5(f'Uploaded Files ({len(uploaded_files)} Total):'), html.Br(), html.Ul([html.Li(i) for i in uploaded_files])])


# In[80]:


@app.callback(
    [Output('cpt', 'options'),
     Output('mod', 'options'),
     Output('sos', 'options')],
     

    [Input('files', 'data')],
    [State('cols', 'data')]
)

def pop_drop(uploaded_files, cols):
    
    if not uploaded_files:
        return []
    
    return cols, cols, cols


# In[81]:


@app.callback(
    Output('dropdowns', 'data'),
    
    [Input('cpt', 'value')],
    [Input('mod', 'value')],
    [Input('sos', 'value')],
)
def store_dropdown_selections(cpt, mod, sos):
    
    return cpt, mod, sos


# In[82]:


@app.callback(
    
    [Output('cpt', 'value'),
     Output('mod', 'value'),
     Output('sos', 'value')],
    
    [Input('dropdowns', 'data')],
)
def track_dropdown_selections(stored_selections):
    
    if stored_selections is None:
        
        return None, None, None
    
    return stored_selections[0], stored_selections[1], stored_selections[2]


# In[83]:


@app.callback(
    Output('format-status', 'data'),
    
    [Input('files', 'data')],
    [State('format-status', 'data')],
)

def display_formatting_status_table(uploaded_files, current_frame):
    
    if not uploaded_files:
           
        current_frame = []
        
        
    current_frame = pd.DataFrame(uploaded_files, columns=['Filename'])
    current_frame['Status'] = ['Pending']*len(current_frame)
    
    return current_frame.to_dict('records')


# In[84]:


@app.callback(
    
    [Output('format-status', 'data', allow_duplicate=True),
     Output('fee-schedules', 'data', allow_duplicate=True)],
     
    
    [Input('fs-process', 'n_clicks')],
    [State('files', 'data')],
    [State('fee-schedules', 'data')],
    [State('dropdowns', 'data')],
    [State('format-status', 'data')],
    
    prevent_initial_call=True
)
def perform_formatting(process_running, filenames, fee_schedules, cols, status_frame):
    
    if process_running:
        
        downloads = []
        
        for fs in fee_schedules:
            
            fs = pd.DataFrame(fs)
            
            formatted = format_fs(fs, cols[0], cols[1], cols[2])
            output = formatted.to_dict('records')
            downloads.append(output)
        
        return_frame = pd.DataFrame(status_frame)
        return_frame['Status'] = ['Complete']*len(return_frame)
        
        return return_frame.to_dict('records'), downloads
    
    raise PreventUpdate


# In[85]:


@app.callback(
    Output("saved-message", "children"),
    [Input("download-button", "n_clicks")],
    [State('files', 'data')],
    [State('fee-schedules', 'data')]
)
def download_files(n_clicks, filenames, fee_schedules):
    
    if n_clicks is None:
        raise PreventUpdate

    parent = r"\\CorpDPT02\\MCSShare\\PSG\\BAs\\NICK\\py script\\app_output"
    child = datetime.now().strftime("%m%d%Y_%H%M%S")
    path = os.path.join(parent, child)
    
    os.makedirs(path, exist_ok=True)
    os.chdir(path)

    file_paths = []

    for name, fs in zip(filenames, fee_schedules):
        
        df = pd.DataFrame(fs)
        file_name = f"{name.split('.')[0]}_format.xlsx"
        
        with pd.ExcelWriter(file_name, 
                            engine='xlsxwriter',
                            engine_kwargs={'options': {'strings_to_numbers': True}}) as writer:
            
            df.to_excel(writer, index=False)
        
        file_paths.append(file_name)

    return f'{len(file_paths)} file(s) saved to {path}.'


# In[86]:


content = html.Div(id="page-content", style=CONTENT_STYLE)
app.layout = html.Div([dcc.Location(id="url"), sidebar, filename_store, data_store, colname_store, col_selections_store, content])


# In[87]:


@app.callback(Output("page-content", "children"), [Input("url", "pathname")])

def render_page_content(pathname):
    
    if pathname == "/":
        return page1
    elif pathname == "/page-1":
        return page2
    elif pathname == "/page-2":
        return page3
    elif pathname == "/page-3":
        return page4

    return html.Div(
        [
            html.H1("404: Not found", className="text-danger"),
            html.Hr(),
            html.P(f"The pathname {pathname} was not recognised..."),
        ],
        className="p-3 bg-light rounded-3",
    )


# In[88]:


if __name__ == "__main__":
    app.run_server(debug=False)

