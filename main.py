#!/usr/bin/env python3

from utils import *
from DataLoader import DataLoader
from ExcelWriter import ExcelWriter
from SheetsWriter import SheetsWriter
import calendar

STARTING_YEAR = int(getenv('STARTING_YEAR'))
DISPLAY_METHOD = getenv('DISPLAY_METHOD')

def write_excel(loader: DataLoader):
	calendar.setfirstweekday(calendar.SUNDAY)
	ExcelWriter(loader).write_excel()

def write_google_sheets(loader: DataLoader):
	SheetsWriter(loader).write_google_sheets()

def main():
	'''Driver Code'''
	loader = DataLoader()
	loader.load(STARTING_YEAR)
	match DISPLAY_METHOD:
		case 'Excel':
			write_excel(loader)
		case 'Sheets':
			write_google_sheets(loader)
		case 'Both':
			write_excel(loader)
			write_google_sheets(loader)
		case method:
			raise ValueError(f'Unexpected DISPLAY_METHOD: "{method}"')

if __name__ == '__main__':
	main()
