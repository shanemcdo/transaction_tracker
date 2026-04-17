#!/usr/bin/env python3

from utils import *
from DataLoader import DataLoader
from ExcelWriter import ExcelWriter
from SheetsWriter import SheetsWriter
import calendar
import datetime

STARTING_YEAR = int(getenv('STARTING_YEAR'))
GOOGLE_SHEETS_ENABLED = getenv('GOOGLE_SHEETS_ENABLED') == 'true'


def write_excel(loader: DataLoader):
	now = datetime.datetime.now()
	datestring = now.strftime('%Y%m%d %H%M%S')
	calendar.setfirstweekday(calendar.SUNDAY)
	current_year = get_year()
	writer = ExcelWriter(os.path.join(TRANSACTION_REPORTS_DIR, f'transactions {datestring}.xlsx'))
	for year in range(STARTING_YEAR, current_year + 1):
		writer.set_year(year)
		writer.get_starting_balances()
		any_success = False
		for month in range(1,13):
			write_month = current_year == year and month + 3 >= now.month
			any_success |= writer.handle_month(month, write_month)
			if write_month:
				writer.hide(month)
		if any_success:
			writer.reset_balances()
			writer.write_summary()
	writer.set_year(STARTING_YEAR)
	writer.reset_balances()
	writer.write_summary_all()
	writer.write_all_transactions()
	writer.set_year(current_year)
	writer.focus(now.month)
	writer.full_screen()
	writer.save()

def write_google_sheets(loader: DataLoader):
	SheetsWriter(loader).write_google_sheets()

def main():
	'''Driver Code'''
	loader = DataLoader()
	loader.load(STARTING_YEAR)
	print(loader.starting_balances)
	if GOOGLE_SHEETS_ENABLED:
		write_google_sheets(loader)

if __name__ == '__main__':
	main()
