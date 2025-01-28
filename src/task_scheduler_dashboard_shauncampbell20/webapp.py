import dash
from dash import dash_table, dcc, html
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output
import pandas as pd
import os
import sqlite3
import re
from core import PROCESS_AUTOMATION_HOME, DB_NAME, set_config, get_config
import sys
import argparse
from subprocess import Popen, CREATE_NEW_CONSOLE

process_automation_logs = os.path.join(PROCESS_AUTOMATION_HOME, 'logs')
process_automation_db = os.path.join(PROCESS_AUTOMATION_HOME, DB_NAME)

def format_home_table(df):
    # Helper function to transform Last Run Table into dash DataTable
    d = {col: df[col].tolist() for col in df.columns}
    url = get_config('HOST')+':'+get_config('PORT')
    d['LogFile'] = ['[%s](/%s)' % (i, i) for i in df['LogFile'].tolist()]
    d['Task'] = ['[%s](/%s)' % (i, i) for i in df['Task'].tolist()]
    df = pd.DataFrame(d)
    df['LastRunTime'] = df['LastRunTime'].apply(
        lambda x: pd.Timestamp(x).strftime('%Y-%m-%d %H:%M:%S') if x != '' else '')
    df['NextRunTime'] = df['NextRunTime'].apply(
        lambda x: pd.Timestamp(x).strftime('%Y-%m-%d %H:%M:%S') if x != '' else '')
    df = df.sort_values('StartTime', ascending=False)
    table = dash_table.DataTable(
        data=df.to_dict(orient='records'),
        columns=[{'id': x, 'name': x, 'presentation': 'markdown'} if x in ['LogFile', 'Task'] else {'id': x, 'name': x}
                 for x in df.columns],
        style_table={'position': 'relative', 'top': '5vh', 'left': '5vw', 'width': '60vw'},
        style_cell={
            'overflow': 'hidden',
            'textOverflow': 'ellipsis',
            'maxWidth': 300,
            'padding-right': '20px',
            'padding-left': '20px',
            'fontSize': 13
        },
        style_cell_conditional=[
            {
                'if': {'column_id': c},
                'textAlign': 'left'
            } for c in ['Executor', 'LastRunResult', 'Status', 'Result']
        ],

        style_data_conditional=[
            {
                'if': {
                    'filter_query': '{Result} = success',  # matching rows of a hidden column with the id, `id`
                    'column_id': 'Result'
                },
                'backgroundColor': 'green',
                'color': 'white'
            },
            {
                'if': {
                    'filter_query': '{Result} = error',  # matching rows of a hidden column with the id, `id`
                    'column_id': 'Result'
                },
                'backgroundColor': '#ff9696',
                'color': 'white'
            },
            {
                'if': {
                    'filter_query': "{Result} = 'no records'",  # matching rows of a hidden column with the id, `id`
                    'column_id': 'Result'
                },
                'backgroundColor': '#f1f1f1',
                'color': 'black'
            },
            {
                'if': {
                    'filter_query': "{Result} = 'critical'",  # matching rows of a hidden column with the id, `id`
                    'column_id': 'Result'
                },
                'backgroundColor': 'red',
                'color': 'white'
            },
            {
                'if': {
                    'filter_query': "{Result} = 'warning'",  # matching rows of a hidden column with the id, `id`
                    'column_id': 'Result'
                },
                'backgroundColor': '#ffc64d',
                'color': 'white'
            }
        ],
        style_as_list_view=True, fill_width=False, sort_action="native", )
    return table

def format_hist_table(df):
    table = dash_table.DataTable(
        data=df.to_dict(orient='records'),
        columns=[
            {'id': x, 'name': x, 'presentation': 'markdown'} if x in ['log_file', 'script_id'] else {'id': x, 'name': x}
            for x in df.columns],
        # style_table={'position': 'relative', 'top': '5vh', 'left': '5vw', 'width': '60vw'},
        style_cell={
            'overflow': 'hidden',
            'textOverflow': 'ellipsis',
            'maxWidth': 300,
            'padding-right': '20px',
            'padding-left': '20px',
            'fontSize': 13
        },
        style_cell_conditional=[
            {
                'if': {'column_id': c},
                'textAlign': 'left'
            } for c in ['Executor', 'LastRunResult', 'Status']
        ],
        style_as_list_view=True, fill_width=False)
    return table

def last_run_table():
    # Returns table with information for the last run of each task in Tasks that has run
    with sqlite3.connect(process_automation_db) as local:
        run_table = pd.read_sql_query('''
            WITH tab as (
            SELECT * FROM Runs WHERE run_id IN (
            SELECT
            MAX(run_id) as run_id
            FROM Runs
            GROUP BY script_id
            ORDER BY end_time DESC)
            ) 

            SELECT 
            Tasks.script_id as Task,
            tab.start_time as StartTime,
            tab.end_time as EndTime,
            tab.result as Result,
            tab.records as Records,
            tab.errors as Errors,
            tab.warnings as Warnings,
            tab.log_file as LogFile,
            tab.user as RanBy,
            tab.machine as Machine,
            Executors.name as Executor,
            Executors.state as Status,
            Executors.last_run_time as LastRunTime,
            Executors.next_run_time as NextRunTime
            FROM Tasks
            LEFT JOIN tab ON Tasks.script_id = tab.script_id 
            LEFT JOIN Executors ON Tasks.command = Executors.command
            --WHERE Executors.state <> 'Disabled'
        ''', local).fillna('')
    return run_table

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.layout = html.Div(
    children=[html.Div(children=[
        dcc.Location(id='url', refresh=False),
        html.Button('Execute', id='execute-button', n_clicks=0, style={'display':'none'}),
        html.Div(id='page-content'),
        html.Div(id='hidden-div', style={'display':'none'})
        ]),
    html.Footer([html.P('Version 0.5.4', style={'font':'10px Arial, sans-serif',"padding": '0px', 'position':'absolute', 'left':'-0px', 'bottom':'0px'})])]
)

@app.callback([Output('page-content', 'children'), Output('execute-button','style')],
              [Input('url', 'pathname')])
def display_page(pathname):
    path = os.path.split(pathname)[-1]
    dat = last_run_table()
    url = get_config('HOST')+':'+get_config('PORT')
    
    # Home Page
    if path == 'home' or path == '':
        return (format_home_table(dat), {'display':'none'})
    
    # Task View
    elif path in dat['Task'].tolist():
        with sqlite3.connect(process_automation_db) as local_con:
            hist = pd.read_sql_query(
                f'''SELECT script_id, start_time, end_time, records, errors, warnings, result, 
                log_file, user, machine FROM Runs WHERE script_id = '{path}' ORDER BY start_time DESC''',
                local_con)
            info = pd.read_sql_query(f'''SELECT * FROM Tasks WHERE script_id = '{path}' ''', local_con).to_dict('index')[0]
        hist['log_file'] = ['[%s](/%s)' % (i, i) for i in hist['log_file'].tolist()]
        return (html.Div([
            html.H3(path),
            html.P('Script Location: %s' % info['script']),
            html.P('Executor Location: %s' % info['command']),
            format_hist_table(hist) ], style={"padding": '35px'}), {'display':'block'})
    
    # Log View
    else:
        with open(os.path.join(process_automation_logs, path)) as f:
            txt = f.read()
            script_name = re.search('starting execution for .+', txt).group(0).replace('starting execution for ', '')
            if re.search('[0-9]{4}-[0-9]{2}-[0-9]{2}\s[0-9]{2}:[0-9]{2}:[0-9]{2}', txt):
                date = re.search('[0-9]{4}-[0-9]{2}-[0-9]{2}\s[0-9]{2}:[0-9]{2}:[0-9]{2}', txt).group(0)
            elif re.search('[0-9]{2}-\w{3}-[0-9]{2}\s[0-9]{2}:[0-9]{2}:[0-9]{2}', txt):
                date = re.search('[0-9]{2}-\w{3}-[0-9]{2}\s[0-9]{2}:[0-9]{2}:[0-9]{2}', txt).group(0)
            else:
                date = ''
        return (html.Div([
            html.H3('Log for %s | %s' % (script_name, date)),
            html.Div(txt, style={'whiteSpace': 'pre-line', "border":"2px #D0D0D0   solid", "background-color":'#F8F8F8', "padding": '15px', 'font':'15px Arial, sans-serif'})
            ], style={"padding": '35px'}), {'display':'none'})

@app.callback(Output('hidden-div', 'children'),
[Input('execute-button', 'n_clicks'), Input('url', 'pathname')])
def run_script_on_click(n_clicks, pathname):
    
    if n_clicks:
        with sqlite3.connect(process_automation_db) as local:
            path = os.path.split(pathname)[-1]
            info = pd.read_sql_query(f'''SELECT * FROM Tasks WHERE script_id = '{path}' ''', local).to_dict('index')[0]
        command = info['execution_command']
        cdir = None
        if info['run_dir'] != '':
            cdir = info['run_dir']
        info['execution_command']
        print(command)
        Popen(command, creationflags=CREATE_NEW_CONSOLE, cwd=cdir)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--host', type=str)
    parser.add_argument('--port', type=str)
    parser.add_argument('--debug', action='store_true')
    parser.add_argument('--update', '-u', action='store_true')
    args = parser.parse_args()
    host = '127.0.0.1'
    port = '8050'
    debug = False
    if args.host:
        set_config('HOST', args.host)
    if args.port:
        set_config('PORT', args.port)
        port = args.port
    if args.debug:
        debug = args.debug
    if args.update:
        print('updating')
        from task_scheduler_dashboard_shauncampbell20.config import build
        build()
    host = get_config('HOST')
    port = get_config('PORT')
    app.run_server(host=host, port=port, debug=debug)
