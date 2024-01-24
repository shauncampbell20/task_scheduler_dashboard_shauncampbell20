# task_scheduler_dashboard_shauncampbell20

The purpose of this project is to create a lightweight dashboard to monitor scripts scheduled and executed by Windows Task Scheduler.

## Installation

```pip install -i https://test.pypi.org/simple/ task-scheduler-dashboard-shauncampbell20==0.5```

## Configuration

### Command Line Interface

`--home` set the directory where logs and database will be stored. Default is ~User\Process Dashboard

`--folder` set the folder containing tasks in Windows Task Scheduler. Default is \\Automation

`--list` or `-l` list the current configuration settings

`--build` or `-b` build the database

```
python config.py --home "C:\Users\Me\Dashboard" --folder "\Automated Tasks" --build
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

`--update` or `-u` Update the database with most recent Task Scheduler information

```
python webapp.py --host 127.0.0.1 --port 8050 --debug --update
```

### Using webapp module

```
from task_scheduler_dashboard_shauncampbell20.webapp import app
from task_scheduler_dashboard_shauncampbell20.config import build
build() # Updates the database with most recent Task Scheduler information
app.run_server(host="127.0.0.1", port="8050", debug=True)
```
