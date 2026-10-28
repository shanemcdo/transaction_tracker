#!/usr/bin/env python3

from utils import *
from CsvLoader import CsvLoader, Loader
from SqliteLoader import SqliteLoader
from ExcelWriter import ExcelWriter
from SheetsWriter import SheetsWriter
from sys import argv
import calendar

STARTING_YEAR = int(getenv('STARTING_YEAR'))
DISPLAY_METHOD = getenv('DISPLAY_METHOD')
LOAD_METHOD = getenv('LOAD_METHOD')

def write_excel(loader: Loader):
	calendar.setfirstweekday(calendar.SUNDAY)
	ExcelWriter(loader).write_excel()

def write_google_sheets(loader: Loader):
	SheetsWriter(loader).write_google_sheets()

def main():
	'''Driver Code'''
	match LOAD_METHOD:
		case 'CSV':
			print('Loading from CSV')
			loader = CsvLoader()
		case 'SQLite':
			print('Loading from SQLite database')
			loader = SqliteLoader()
		case method:
			raise ValueError(f'Unexpected LOAD_METHOD: "{method}"')
	loader.load(STARTING_YEAR)
	display_method = argv[1] if len(argv) > 1 else DISPLAY_METHOD
	match display_method:
		case 'Excel':
			print('Writing to Excel')
			write_excel(loader)
		case 'Sheets':
			print('Writing to Google Sheets')
			write_google_sheets(loader)
		case 'Both':
			print('Writing to Excel and Google Sheets')
			print('Writing to Excel')
			write_excel(loader)
			print('Writing to Google Sheets')
			write_google_sheets(loader)
		case method:
			raise ValueError(f'Unexpected DISPLAY_METHOD: "{method}"')

if __name__ == '__main__':
	main()
