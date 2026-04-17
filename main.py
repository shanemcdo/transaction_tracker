#!/usr/bin/env python3

from utils import *
from DataLoader import DataLoader
from ExcelWriter import ExcelWriter
from SheetsWriter import SheetsWriter
import calendar

STARTING_YEAR = int(getenv('STARTING_YEAR'))
GOOGLE_SHEETS_ENABLED = getenv('GOOGLE_SHEETS_ENABLED') == 'true'

def write_excel(loader: DataLoader):
	calendar.setfirstweekday(calendar.SUNDAY)
	ExcelWriter(loader).write_excel()

def write_google_sheets(loader: DataLoader):
	SheetsWriter(loader).write_google_sheets()

def main():
	'''Driver Code'''
	loader = DataLoader()
	loader.load(STARTING_YEAR)
	write_excel(loader)
	# if GOOGLE_SHEETS_ENABLED:
	# 	write_google_sheets(loader)

if __name__ == '__main__':
	main()
