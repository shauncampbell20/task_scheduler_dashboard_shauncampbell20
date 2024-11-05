import os
import sqlite3
import re
import pandas as pd
import sys
from task_scheduler_dashboard_shauncampbell20.core import resultCodes, _loc, set_config
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


def build(update=True):
    ## Main function for building and updated the database
    with open(os.path.join(_loc, 'config.json'), 'r') as config:
        configs = json.load(config)
    PROCESS_AUTOMATION_HOME = configs['PROCESS_AUTOMATION_HOME']
    SCHEDULER_FOLDER = configs['SCHEDULER_FOLDER']
    DB_NAME = configs["DB_NAME"]
    
    # Create directory PROCESS_AUTOMATION_HOME if doesn't exist
    if not os.path.exists(PROCESS_AUTOMATION_HOME):
        os.mkdir(PROCESS_AUTOMATION_HOME)
        warnings.warn('Created directory '+PROCESS_AUTOMATION_HOME)
    
    # Create logs directory in PROCESS_AUTOMATION_HOME if doesn't exist
    process_automation_logs = os.path.join(PROCESS_AUTOMATION_HOME, 'logs')
    if not os.path.exists(process_automation_logs):
        os.mkdir(process_automation_logs)
    process_automation_db = os.path.join(PROCESS_AUTOMATION_HOME, DB_NAME)
    
    # Create Executors and Tasks tables. If update is true, it will clear and reset these tables to be up to date
    # If update = False, i.e. initializing for the first time, it will also create the Runs table.
    if not os.path.exists(os.path.join(PROCESS_AUTOMATION_HOME, DB_NAME)):
        update = False

    local = sqlite3.connect(os.path.join(PROCESS_AUTOMATION_HOME, DB_NAME))
    cursor = local.cursor()

    # --- Get Task Scheduler information ---#
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

    # --- Create Executors Table ---#
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
    folder VARCHAR
    )''')

    # --- Insert Task Scheduler Information into Executors Table ---#
    for pname in d.keys():
        cursor.execute(f'''
        INSERT INTO Executors 
        VALUES ('{pname}', '{d[pname]['State']}', '{str(d[pname]['Next Run'])}', 
        '{str(d[pname]['Last Run'])}', '{d[pname]['Last Result']}', 
        '{d[pname]['Hidden']}', '{d[pname]['Command']}', '{d[pname]['Folder']}')
        ''')
    local.commit()

    # --- Create and Populate Tasks Table ---#
    cursor.execute('''DROP TABLE IF EXISTS Tasks;''')
    cursor.execute('''
    CREATE TABLE Tasks (
    script_id VARCHAR,
    command VARCHAR,
    script VARCHAR,
    run_dir VARCHAR,
    execution_command VARCHAR
    )''')
    
    # Parse batch files
    df = pd.read_sql_query('''SELECT * FROM Executors ''', local)
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
                VALUES ('{tasks[tname]['script_id']}', '{tasks[tname]['command']}', '{tname}', '{tasks[tname]['run_dir']}', '{tasks[tname]['execution_command']}')
                ''')
            local.commit()
        except Exception as e:
            warnings.warn('Exception in '+batchFile+': '+str(e))
    local.commit()
    
    def create_run_table():
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
    
    if not update:
        # --- Create Runs Table ---#
        # --- Does not run if update is true ---#
        create_run_table()
    else:
        if ('Runs',) not in cursor.execute('''SELECT name FROM sqlite_master WHERE type = 'table' ''').fetchall():
            create_run_table()
            

    # --- Add tasks that trigger other tasks to Executors table ---#
    pd.read_sql_query('''
    SELECT DISTINCT
    t1.command as name,
    'Ready' as state,
    (SELECT next_run_time FROM Executors WHERE command = t2.command) as next_run_time,
    (SELECT end_time FROM Runs WHERE script_id = t2.script_id ORDER BY end_time DESC) as last_run_time,
    (SELECT result FROM Runs WHERE script_id = t2.script_id ORDER BY end_time DESC ) as last_run_result,
    'False' as hidden,
    t1.command as command
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
    parser.add_argument('--list', '-l' , action='store_true')
    parser.add_argument('--build', '-b' , action='store_true')
    args = parser.parse_args()
    if args.home:
        set_config('PROCESS_AUTOMATION_HOME', args.home)
    if args.folder:
        set_config('SCHEDULER_FOLDER', args.folder)
    if args.dbname:
        set_config('DB_NAME', args.dbname)
    if args.list:
        list_configs()
    if args.build:
        build()
        
