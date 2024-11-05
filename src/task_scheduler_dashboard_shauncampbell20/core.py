import os
import sqlite3
import datetime
import logging
from logging import Logger
import json
import inspect
import warnings

_loc = os.path.split(__file__)[0]

def create_config_file():
    if 'config.json' not in os.listdir(_loc):
        default = os.path.join(os.path.expanduser('~'),'Process Dashboard')
        configs = {'PROCESS_AUTOMATION_HOME': default, 'SCHEDULER_FOLDER': '\\Automation', 'DB_NAME':'process_automation.db','HOST':'127.0.0.1', 'PORT':'8050'}
        with open (os.path.join(_loc, 'config.json'), 'w') as f:
            json.dump(configs, f)


def get_config(config_name):
    if 'config.json' not in os.listdir(_loc):
        create_config_file()
    with open(os.path.join(_loc, 'config.json'), 'r') as config:
        configs = json.load(config)
    return configs[config_name]
    
    
def set_config(config_name, config_value):
    with open(os.path.join(_loc, 'config.json'), 'r') as config:
        configs = json.load(config)
    configs[config_name] = config_value
    with open (os.path.join(_loc, 'config.json'), 'w') as f:
        json.dump(configs, f)


if 'config.json' not in os.listdir(_loc):
    create_config_file()


with open(os.path.join(_loc, 'config.json'), 'r') as config:
    configs = json.load(config)


PROCESS_AUTOMATION_HOME = configs["PROCESS_AUTOMATION_HOME"]
SCHEDULER_FOLDER = configs["SCHEDULER_FOLDER"]
DB_NAME = configs["DB_NAME"]


        
resultCodes = {0: 'The operation completed successfully.',
               1: '',
               10: 'The environment is incorrect.',
               267008: 'Task is ready to run at its next scheduled time.',
               267009: 'Task is currently running.',
               267010: 'The task will not run at the scheduled times because it has been disabled.',
               267011: 'Task has not yet run.',
               267012: 'There are no more runs scheduled for this task.',
               267013: 'One or more of the properties that are needed to run this'
                       ' task on a schedule have not been set.',
               267014: 'The last run of the task was terminated by the user.',
               267015: 'Either the task has no triggers or the existing triggers are disabled or not set.',
               2147750671: 'Credentials became corrupted.',
               2147750687: 'An instance of this task is already running.',
               2147943645: 'The service is not available (is "Run only when an user is logged on" checked?).',
               3221225786: 'The application terminated as a result of a CTRL+C.',
               3228369022: 'Unknown software exception.',
               -2147020576: 'The operator or administrator has refused the request.'}


class ProcessLogger(Logger):
    def __init__(self, name=None):
        if not name:
            name = os.path.splitext(os.path.split(inspect.stack()[1][1])[-1])[0]
        super().__init__(name)
        self.start_time = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%d %H:%M:%S')
        self.records = 0
        self.errors = 0
        self.warnings = 0
        self.criticals = 0
        self.result = ''
        self.script_id = name
        self.user = os.getlogin()
        self.machine = os.environ['COMPUTERNAME']
        self.process_automation_db = os.path.join(PROCESS_AUTOMATION_HOME, DB_NAME)
        self.process_automation_logs = os.path.join(PROCESS_AUTOMATION_HOME, 'logs')
        if not os.path.exists(self.process_automation_logs):
            os.mkdir(self.process_automation_logs)
        if len(os.listdir(self.process_automation_logs)) == 0:
            self.log_file = '1000000'
        else:
            self.log_file = str(int(sorted(os.listdir(self.process_automation_logs))[-1]) + 1)
        self.log_path = os.path.join(self.process_automation_logs, self.log_file)
        handler = logging.FileHandler(self.log_path)
        handler.setFormatter(logging.Formatter('%(levelname)s:%(asctime)s - %(message)s'))
        self.addHandler(handler)
        self.info('starting execution for %s' % self.script_id)
        try:
            with sqlite3.connect(self.process_automation_db) as local:
                cursor = local.cursor()
                cursor.execute('''INSERT INTO Runs (script_id, log_file, start_time, 
                end_time, records, result, errors, warnings) VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                                    (self.script_id, self.log_file, self.start_time, '', 0, 'running', 0, 0))
                local.commit()
        except:
            from task_scheduler_dashboard_shauncampbell20.config import build
            build(update=False)
            with sqlite3.connect(self.process_automation_db) as local:
                cursor = local.cursor()
                cursor.execute('''INSERT INTO Runs 
                (script_id, log_file, start_time, end_time, records, result, errors, warnings) VALUES 
                (?, ?, ?, ?, ?, ?, ?, ?)''', (self.script_id, self.log_file, self.start_time, '', 0, 'running', 0, 0))
                local.commit()
        
        with sqlite3.connect(self.process_automation_db) as local:
            cursor = local.cursor()
            self.run_id = cursor.execute(
                f'''SELECT * FROM Runs WHERE script_id = '{self.script_id}' ORDER BY start_time DESC ''').fetchone()[0]


    def error(self, msg, *args, **kwargs):
        self.errors += 1
        self.log(40, msg, *args, **kwargs)

    def warning(self, msg, *args, **kwargs):
        self.warnings += 1
        self.log(30, msg, *args, **kwargs)

    def critical(self, msg, *args, **kwargs):
        self.criticals += 1
        self.log(50, msg, *args, **kwargs)
    
    def last_run(self):
        with sqlite3.connect(self.process_automation_db) as local:
            cursor = local.cursor()
            last_ran = cursor.execute(f'''SELECT start_time FROM Runs WHERE script_id = '{self.script_id}' ORDER BY start_time DESC ''').fetchone()[0]
            if last_ran == None:
                return '1/1/1900'
            else:
                return last_ran
    
    def progress(self, iterable, records = True):
        if inspect.isgenerator(iterable):
            iterable = [i for i in iterable]
        total = len(iterable)
        num = 1
        for item in iterable:
            s = '+PROGRESS |'
            s += '-'*int(num/total*10)
            s += ' '*(10-int(num/total*10))
            s += '| '
            s += '{:.1%}'.format(num/total)
            with open(self.log_path, 'r', encoding='UTF-8') as f:
                cur_log = f.readlines()
            if '+PROGRESS' in cur_log[-1]:
                cur_log[-1] = s
            else:
                cur_log.extend(s)
            with open(self.log_path, 'w', encoding='UTF-8') as f:
                f.write(''.join(cur_log))
                
            yield item
            num += 1
            if records == True:
                self.records += 1
        with open(self.log_path, 'r', encoding='UTF-8') as f:
            cur_log = f.read()
        cur_log += '\n'
        with open(self.log_path, 'w', encoding='UTF-8') as f:
            f.write(cur_log)
                
    def complete(self):
        end_time = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%d %H:%M:%S')
        self.info('execution for %s completed.' % self.script_id)
        if self.criticals > 0:
            self.result = 'critical'
        elif self.errors > 0:
            self.result = 'error'
        elif self.warnings > 0:
            self.result = 'warning'
        elif self.records == 0:
            self.result = 'no records'
        else:
            self.result = 'success'
        with sqlite3.connect(self.process_automation_db) as local:
            cursor = local.cursor()
            cursor.execute(f'''
                UPDATE Runs 
                SET end_time = '{end_time}', 
                records = {self.records}, 
                result = '{self.result}', 
                errors = {self.errors},
                warnings = {self.warnings},
                user = '{self.user}',
                machine = '{self.machine}'
                WHERE run_id = {self.run_id}
            ''')
            local.commit()

