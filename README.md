# task_scheduler_dashboard_shauncampbell20

The purpose of this project is to create a lightweight dashboard to monitor scripts scheduled and executed by Windows Task Scheduler.

## Installation

```pip install -i https://test.pypi.org/simple/ task-scheduler-dashboard-shauncampbell20==1.0```

## Configuration

### Command Line Interface

Add install location to PATH to use `task_scheduler` command.

`--home` set the directory where logs and database will be stored. Default is ~User\Process Dashboard

`--folder` set the folder containing tasks in Windows Task Scheduler. Default is \\Automation

`--dbname` set the name of the database. Default is process_automation.db

`--list` or `-l` list the current configuration settings

`--update` or `-u` Update the database with most recent Task Scheduler information, or build the database if initializing for the first time.

`--reset` or `-r` force reset of database (clear tables)

```
task_scheduler --home "C:\Users\Me\Dashboard" --folder "\Automated Tasks" --update
```

### Using config module

```
from task_scheduler_dashboard_shauncampbell20 import config
config.set_home_directory(r"C:\Users\Me\Dashboard")
config.set_scheduler_folder("\\Automated Tasks")
config.build()
```

## Using ProcessLogger

This package contains a ProcessLogger class that's a wrapper for logging.Logger. It writes log files to the \logs directory in the specified home folder. The log files are displayed in the dashboard when viewing execution details.

Example usage:

```
from task_scheduler_dashboard_shauncampbell20 import ProcessLogger

pl = ProcessLogger()
pl.info('starting execution')
print('Hello World!')
try:
	1/0
except Exeption as e:
	pl.error(e, exec_info=True)
pl.complete()
```

## Running the Web Application

### Command Line Interface

`--host` set the host domain for the webapp. Default is 127.0.0.1

`--port` set the port for the webapp. Default is 8050

`--debug` Start the webapp in debug mode

`--run` Run the webapp

```
task_scheduler --host 127.0.0.1 --port 8050 --debug --run
```

### Using webapp module

```
from task_scheduler_dashboard_shauncampbell20.webapp import app
from task_scheduler_dashboard_shauncampbell20.config import build
build() # Updates the database with most recent Task Scheduler information
app.run_server(host="127.0.0.1", port="8050", debug=True)
```

## Scheduling Tasks

Tasks should be created in Windows Task Scheduler within the folder specified in the configuration. Each task's action should be executing a batch file.

The batch files should be formatted as "python path" "script path". Each batch file can have multiple lines that execute different python scripts.
