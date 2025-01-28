import os
import sqlite3
import re
import pandas as pd
import sys
from core import resultCodes, _loc, set_config
import argparse
import json
import warnings

def set_home_directory(config_value):
    set_config('PROCESS_AUTOMATION_HOME', config_value)

def set_scheduler_folder(config_value):
    set_config('SCHEDULER_FOLDER', config_value)
    
def set_db_name(config_value):
    set_config('DB_NAME', config_value)

def set_port(config_value):
    set_config('PORT', config_value)
    
def set_host(config_value):
    set_config('HOST', config_value)

def list_configs():
    with open(os.path.join(_loc, 'config.json'), 'r') as config:
        configs = json.load(config)
    for k, v in configs.items():
        print(k,":",v)

def parse_task_scheduler(SCHEDULER_FOLDER):
    import win32com.client
    import pythoncom
    TASK_ENUM_HIDDEN = 1
    TASK_STATE = {0: 'Unknown',
                  1: 'Disabled',
                  2: 'Queued',
                  3: 'Ready',
                  4: 'Running'}

    scheduler = win32com.client.Dispatch('Schedule.Service')
    scheduler.Connect()
    
    # Get folders and tasks
    try:
        folders = [scheduler.GetFolder(SCHEDULER_FOLDER)]
    except pythoncom.com_error:
        raise pythoncom.com_error('"'+SCHEDULER_FOLDER+'" folder not found in Task Scheduler')
        quit
    d = {}
    tasks = []
    while folders:
        folder = folders.pop(0)
        folders += list(folder.GetFolders(0))
        tasks += list(folder.GetTasks(TASK_ENUM_HIDDEN))
        
        # Loop through each task
        for task in tasks:
            settings = task.Definition.Settings
            taskName = os.path.split(task.Path)[-1]
            taskFolder = os.path.split(task.Path)[-2]
            d[taskName] = {'Hidden': settings.Hidden, 'State': TASK_STATE[task.State], 'Last Run': task.LastRunTime,
                           'Next Run': task.NextRunTime, 'Folder':taskFolder}
            d[taskName]['Command'] = re.search('<Command>.+</Command>', task.Xml).group(0)[9:-10]
            try:
                d[taskName]['Last Result'] = resultCodes[task.LastTaskResult]
            except KeyError:
                d[taskName]['Last Result'] = task.LastTaskResult
            d[taskName]['Machine'] = os.environ['COMPUTERNAME']
    return d

def build(update=True):
    ## Main function for building and updating the database
    
    def create_run_table():
        # Create run table if it does not exist
        if ('Runs',) not in cursor.execute('''SELECT name FROM sqlite_master WHERE type = 'table' ''').fetchall():
            cursor.execute('''DROP TABLE IF EXISTS Runs''')
            cursor.execute('''
            CREATE TABLE Runs (
            run_id INTEGER PRIMARY KEY,
            script_id VARCHAR,
            log_file VARCHAR,
            start_time VARCHAR,
            end_time VARCHAR,
            records INT,
            result VARCHAR,
            errors INT,
            warnings INT,
            user VARCHAR,
            machine VARCHAR
            )
            ''')
            local.commit()
        else:
            warnings.warn('Runs table exists, skipping create run table')
        
    def create_executors_table():
        # Creates Executors table
        cursor.execute('''DROP TABLE IF EXISTS Executors;''')
        cursor.execute('''
        CREATE TABLE Executors (
        name VARCHAR,
        state VARCHAR,
        next_run_time VARCHAR,
        last_run_time VARCHAR,
        last_run_result VARCHAR,
        hidden VARCHAR,
        command VARCHAR,
        folder VARCHAR,
        machine VARCHAR
        )''')
        local.commit()
        
    def create_tasks_table():
        # Creates Tasks table
        cursor.execute('''DROP TABLE IF EXISTS Tasks;''')
        cursor.execute('''
        CREATE TABLE Tasks (
        script_id VARCHAR,
        command VARCHAR,
        script VARCHAR,
        run_dir VARCHAR,
        execution_command VARCHAR,
        machine VARCHAR
        )''')
        
    # Load configs
    with open(os.path.join(_loc, 'config.json'), 'r') as config:
        configs = json.load(config)
    PROCESS_AUTOMATION_HOME = configs['PROCESS_AUTOMATION_HOME']
    SCHEDULER_FOLDER = configs['SCHEDULER_FOLDER']
    DB_NAME = configs["DB_NAME"]
    machine = os.environ['COMPUTERNAME']
    
    # Create directory PROCESS_AUTOMATION_HOME if doesn't exist
    if not os.path.exists(PROCESS_AUTOMATION_HOME):
        x = input(f'Create directory {PROCESS_AUTOMATION_HOME}? [Y/N]')
        if x == 'Y':
            os.mkdir(PROCESS_AUTOMATION_HOME)
            warnings.warn(f'Created directory {PROCESS_AUTOMATION_HOME}')
        else:
            print('**Set home folder with --home argument')
            quit()
    
    # Create logs directory in PROCESS_AUTOMATION_HOME if doesn't exist
    process_automation_logs = os.path.join(PROCESS_AUTOMATION_HOME, 'logs')
    if not os.path.exists(process_automation_logs):
        os.mkdir(process_automation_logs)
    
    # If update = False, i.e. initializing for the first time, create the Runs, Executors, and Tasks tables.
    if not os.path.exists(os.path.join(PROCESS_AUTOMATION_HOME, DB_NAME)):
        update = False
    local = sqlite3.connect(os.path.join(PROCESS_AUTOMATION_HOME, DB_NAME))
    cursor = local.cursor()
    if not update:
        create_run_table()
        create_executors_table()
        create_tasks_table()
    else:
        if ('Runs',) not in cursor.execute('''SELECT name FROM sqlite_master WHERE type = 'table' ''').fetchall():
            create_run_table()
        if ('Executors',) not in cursor.execute('''SELECT name FROM sqlite_master WHERE type = 'table' ''').fetchall():
            create_executors_table()
        if ('Tasks',) not in cursor.execute('''SELECT name FROM sqlite_master WHERE type = 'table' ''').fetchall():
            create_tasks_table()

    # Get Task Scheduler information
    d = parse_task_scheduler(SCHEDULER_FOLDER)

    # Insert Task Scheduler Information into Executors table
    cursor.execute(f'''DELETE FROM Executors WHERE machine = '{machine}' ''')
    for pname in d.keys():
            cursor.execute(f'''
            INSERT INTO Executors 
            VALUES ('{pname}', '{d[pname]['State']}', '{str(d[pname]['Next Run'])}', 
            '{str(d[pname]['Last Run'])}', '{d[pname]['Last Result']}', 
            '{d[pname]['Hidden']}', '{d[pname]['Command']}', '{d[pname]['Folder']}', '{d[pname]['Machine']}')
            ''')
    local.commit()

    # Parse batch files and insert into Tasks table
    cursor.execute(f'''DELETE FROM Tasks WHERE machine = '{machine}' ''')
    df = pd.read_sql_query(f'''SELECT * FROM Executors WHERE machine = '{machine}' ''', local)
    for batchFile, taskfolder in zip(df['command'], df['folder']):
        try:
            with open(batchFile, 'r') as f:
                batchContents = f.readlines()
            tasks = {}
            runDir = ''
            for line in batchContents:
                if line[:3] == 'cd ':
                    runDir = line.replace('cd ','').replace('"','').strip()
                if line[:2] != '::' and 'python.exe' in line:
                    script = line.split('" "')[-1].replace('"','').strip()
                    scriptID = os.path.splitext(os.path.split(script)[-1])[0]
                    executionCommand = line.replace('"','').strip()
                    tasks[script] = {'command':batchFile, 'script_id':scriptID, 'execution_command':executionCommand, 'run_dir':runDir, 'folder':taskfolder}
            for tname in tasks.keys():
                cursor.execute(f'''
                INSERT INTO Tasks 
                VALUES ('{tasks[tname]['script_id']}', '{tasks[tname]['command']}', '{tname}', '{tasks[tname]['run_dir']}', '{tasks[tname]['execution_command']}', '{machine}')
                ''')
            local.commit()
        except Exception as e:
            warnings.warn('Exception in '+batchFile+': '+str(e))
    local.commit()
    
    # Add tasks that trigger other tasks to Executors table
    pd.read_sql_query(f'''
    SELECT DISTINCT
    t1.command as name,
    'Ready' as state,
    (SELECT next_run_time FROM Executors WHERE command = t2.command) as next_run_time,
    (SELECT end_time FROM Runs WHERE script_id = t2.script_id ORDER BY end_time DESC) as last_run_time,
    (SELECT result FROM Runs WHERE script_id = t2.script_id ORDER BY end_time DESC ) as last_run_result,
    'False' as hidden,
    t1.command as command,
    '{machine}' as machine
    FROM Tasks t1 
    LEFT JOIN Tasks t2 ON t1.command = t2.script_id 
    WHERE t1.command NOT IN (SELECT command FROM Executors)
    ''', local).to_sql('Executors', local, if_exists='append', index=False)

    if update == False:
        print('DB Initialized Successfully')


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--home', type=str)
    parser.add_argument('--folder', type=str)
    parser.add_argument('--dbname', type=str)
    parser.add_argument('--host', type=str)
    parser.add_argument('--port', type=str)
    parser.add_argument('--debug', action='store_true')
    parser.add_argument('--list', '-l' , action='store_true')
    parser.add_argument('--update', '-u' , action='store_true')
    parser.add_argument('--reset', '-r' , action='store_true')
    parser.add_argument('--run', action='store_true')
    args = parser.parse_args()
    if args.home:
        set_config('PROCESS_AUTOMATION_HOME', args.home)
    if args.folder:
        set_config('SCHEDULER_FOLDER', args.folder)
    if args.dbname:
        set_config('DB_NAME', args.dbname)
    if args.host:
        set_config('HOST', args.host)
    if args.port:
        set_config('PORT', args.port)
    if args.list:
        list_configs()
    if args.reset:
        build(update=False)
    elif args.update:
        build(update=True)
    if args.run:
        build(update=True)
        from webapp import *
        debug = args.debug
        host = get_config('HOST')
        port = get_config('PORT')
        app.run_server(host=host, port=port, debug=debug)
        
