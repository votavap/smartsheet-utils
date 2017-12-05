# Collection of SmartSheet utility tools
------

## 1. SmartSheet backup with History

This tool backs up a SmartSheet including the history of each cell and stores the output in a JSON file

### Prerequisites/Dependencies

##### 1. Python3
##### 2. [smartsheet-python-sdk](https://github.com/smartsheet-platform/smartsheet-python-sdk) (depends on requests, requests-toolbelt, six, python-dateutil, certifi, urllib3, idna - these will be installed during pip install)


### Installing

Just clone this repo. 

### Running the code

```
export SMARTSHEET_ACCESS_TOKEN="your access token"
python3 smartsheets-backup.py --sheet-name="some smartsheet name" --backup-dir="backup directory"
```

### Output

The tool creates a JSON back up of the SmartSheet in the designated directory together with a log file. The SmartSheet backup file includes the history of each cell.

