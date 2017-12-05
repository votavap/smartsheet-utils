#!/usr/bin/env python
"""
This module backs up a SmartSheet via the REST API and saves it in a JSON file. The backup includes
the history of each cell = who did what when.

Dependencies:
   1. Python3
	2. smartsheet-python-sdk (https://github.com/smartsheet-platform/smartsheet-python-sdk)

Example:
   export SMARTSHEET_ACCESS_TOKEN="your access token"
   python3 smartsheets-backup.py --sheet-name="some smartsheet name" --backup-dir="backup directory"

"""


import smartsheet
import logging
import json
import pprint
import datetime
import argparse
import sys
import traceback
import os

ACCESS_TOKEN = os.environ["SMARTSHEET_ACCESS_TOKEN"];

logging.basicConfig(level=logging.WARN);

def setup_file_log(log_file):
	""" Setup file logging (for later use)"""
	
	fileh = logging.FileHandler(log_file, 'a')
	formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
	fileh.setFormatter(formatter)

	log = logging.getLogger()  
	for hdlr in log.handlers[:]: 
		log.removeHandler(hdlr);
	log.addHandler(fileh);


class SmartsheetBackup(object):
	def __init__(self):
		""" Constructor - initialze access using ACCESS_TOKEN"""
		try:
			self.smart_sheets = smartsheet.Smartsheet(ACCESS_TOKEN);
		except Exception as e:
			logging.error(traceback.format_exc());
			sys.exit(1);
			
		self.smart_sheets.errors_as_exceptions(True);


	def get_all_sheets(self):
		""" Return list of all user's SmartSheets"""
		return self.smart_sheets.Sheets.list_sheets(include_all=True);

	
	def get_sheet_id_from_name(self, sheet_name):
		""" Given a name of a SmartSheet, return the SmartSheet ID
		
		Args:
		    sheet_name (str): Name of a SmartSheet
			 
		Returns:
		    str: ID of the SmartSheet as a string
		"""
		sheet_id = None;
	
		all_sheets = self.get_all_sheets();
		for _s in all_sheets.data:
			if _s.name == sheet_name:
				sheet_id = _s.id;
				break;
		return sheet_id;	


	def backup_smart_sheet(self, sheet_name, directory=None, file_name=None):
		""" Gets history of each cell of the SmartSheet and writes it into a JSON file
		
		Args:
		    sheet_name (str): Name of the SmartSheet
			 directory  (str): Directory where the JSON outputh should go
			 file_name  (str): Name of the JSON file. Defaults to modified sheet_name+timestamp
		""" 

		timestamp = datetime.datetime.now();
		sheet_id = self.get_sheet_id_from_name(sheet_name);	

		current_sheet = self.smart_sheets.Sheets.get_sheet(sheet_id, include='attachments,discussions,ownerInfo,source,rowWriterInfo');
		
		if not file_name:
			file_name = '_'.join(current_sheet.name.split())+"."+timestamp.strftime("%Y-%m-%d-%H%M%S")+".json";
		out_file = os.path.join(directory, file_name);
		log_file = out_file.replace(".json", ".log");
		setup_file_log(log_file);
		
		logging.info("Loaded " + str(len(current_sheet.rows)) + " rows and " + str(len(current_sheet.columns)) + " columns from sheet: " + current_sheet.name);
	
		column_map = {};
		for column in current_sheet.columns:
			column_map[column.title] = column.id;

	
	
		backup_date_time = timestamp.strftime("%a, %d %b %Y %H:%M:%S");
		backup_sheet = {};
		backup_sheet['smartsheet'] = {};
		backup_sheet['backed_up'] = backup_date_time;
		backup_sheet['smartsheet']['sheet_name'] = sheet_name;
		backup_sheet['smartsheet']['columns'] = {}
	
		for column in current_sheet.columns:
			backup_sheet['smartsheet']['columns'][column.title] = {}
			backup_sheet['smartsheet']['columns'][column.title]['column_id'] = column.id;
			backup_sheet['smartsheet']['columns'][column.title]['rows'] = [];
			for row in current_sheet.rows:
				#print(row.id, row.row_number, column.id, column.title);
				logging.warn("Processing column:'"+column.title+"' row:"+str(row.row_number));
				#history = test_sheet.Cells.get_cell_history(sheet_id, row.id, column.id);
				#print(smart_sheets.Cells.get_cell_history(sheet_id, row.id, column.id));
				row_data = {}
				row_data['row_number'] = row.row_number;
				row_data['row_id'] = row.id;
				try:
					row_data['content_and_history'] = self.smart_sheets.Cells.get_cell_history(sheet_id, row.id, column.id).to_dict();
				except Exception as e:
					logging.error(traceback.format_exc());
					sys.exit(1);
				backup_sheet['smartsheet']['columns'][column.title]['rows'].append(row_data)
			
		with open(out_file, 'w') as out_json_file:
			json.dump(backup_sheet, out_json_file);
		
		completed_timestamp = datetime.datetime.now();
		print("COMPLETED SmartSheet Backup:");
		print("SmartSheet:"+current_sheet.name);
		print("Elapsed Time:"+str(completed_timestamp-timestamp));
		
		return;
	

if __name__ == '__main__':
	
	parser = argparse.ArgumentParser(description='Arguments for smartsheets-backup tool')
	parser.add_argument("-s", "--sheet-name", help="Name of the smartsheet", required=True);
	parser.add_argument("-d", "--backup-dir", help="Directory where the smarthseet should be backed up", required=True);
	parser.add_argument("-f", "--backup-file", help="Name of the file, where the smartsheet JSON should be backed up (default: derived from name + timestamp)", required=False);
	args = parser.parse_args();
	
	sb = SmartsheetBackup();
	
	sb.backup_smart_sheet(args.sheet_name, directory=args.backup_dir, file_name=args.backup_file);
	
	
		
