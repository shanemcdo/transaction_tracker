#!/usr/bin/env python3

import os
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import pandas as pd
from glob import glob

MONTHS = {
	1: 'January',
	2: 'Febuary',
	3: 'March',
	4: 'April',
	5: 'May',
	6: 'June',
	7: 'July',
	8: 'August',
	9: 'September',
	10: 'October',
	11: 'November',
	12: 'December',
}
MONTHS_SHORT = { key: value[:3] for key, value in MONTHS.items() }
DEFAULT_INPUT_DIR = './in/'
DEFAULT_OUTPUT_DIR = './out/'
# unix glob format
INPUT_FILENAME_FORMAT = 'Transactions {0} 1, {1} - {0} ??, {1} *.csv'
YEAR = 2024
# TODO: change if budget gets adjusted
BUDGET_PER_MONTH = { i: 1000.00 for i in range(1, 13) }

class Writer:
	def __init__(self, filename: str):
		self.excelWriter = pd.ExcelWriter(filename, engine='xlsxwriter')
		self.workbook = self.excelWriter.book
		currency = { 'num_format': '$#,##0.00' }
		self.formats = {
			'currency': self.workbook.add_format(currency),
			'border_currency': self.workbook.add_format({ 'border': True, **currency }),
		}

	@staticmethod
	def get_csv_filename_from_month(month: str) -> str:
		glob_pattern = INPUT_FILENAME_FORMAT.format(month, YEAR)
		files = sorted(glob(
			glob_pattern,
			root_dir = DEFAULT_INPUT_DIR
		))
		if len(files) < 1:
			raise FileNotFoundError(f'Could not find any matches for {glob_pattern}')
		return os.path.join(DEFAULT_INPUT_DIR, files[-1])

	def read_month(self, month: int):
		filename = self.get_csv_filename_from_month(MONTHS_SHORT[month])
		data = pd.read_csv(
			filename,
			sep =', |,',
			# get rid of warning
			engine='python'
		)
		return data

	def write_month(self, month: int, data: pd.DataFrame):
		sheet_name = MONTHS[month]
		data.to_excel(
			self.excelWriter,
			sheet_name = sheet_name,
			index = False
		)
		rows, cols = data.shape
		cols -= 1
		sheet = self.excelWriter.sheets[sheet_name]
		sheet.add_table(0, 0, rows, cols, {
			'columns': [
				{ 'header': 'Date' },
				{ 'header': 'Category' },
				{ 'header': 'Amount', 'format': self.formats['currency'] },
				{ 'header': 'Note' },
			],
			'name': sheet_name + 'Table',
		})
		start_col = cols + 1
		pivot = data.pivot_table(
			values = 'Amount',
			index = 'Category',
			aggfunc = 'sum'
		).reset_index()
		rows, cols = pivot.shape
		cols -= 1
		pivot.to_excel(
			self.excelWriter,
			sheet_name = sheet_name,
			index = False,
			startcol = start_col
		)
		sheet.add_table(0, start_col, rows + 1, start_col + cols, {
			'columns': [
				{ 'header': 'Category', 'total_string': 'Total' },
				{
					'header': 'Sum of Amount',
					'format': self.formats['currency'],
					'total_function': 'sum'
				},
			],
			'name': sheet_name + 'Pivot',
			'total_row': 1
		})
		sheet.write(rows + 2, start_col, 'Budget', self.formats['border_currency'])
		sheet.write(rows + 2, start_col + 1, BUDGET_PER_MONTH[month], self.formats['border_currency'])
		sheet.write(rows + 3, start_col, 'Over/Under', self.formats['border_currency'])
		sheet.write(rows + 3, start_col + 1, f'={BUDGET_PER_MONTH[month]}-{xl_rowcol_to_cell(rows + 1, start_col + 1)}', self.formats['border_currency'])
		chart = self.workbook.add_chart({ 'type': 'pie' })
		chart.add_series({
			'categories': [sheet_name, 1, start_col, rows, start_col],
			'values': [sheet_name, 1, start_col + 1, rows, start_col + 1],
			'data_labels': { 'value': True, 'percentage': True, 'position': 'best_fit' },
		})
		sheet.insert_chart(rows + 4, start_col, chart)
		# Stupid hack because format in add_table isn't work
		for cells in ('C:C', 'F:F'):
			sheet.set_column(cells, None, self.formats['currency'])
		sheet.autofit()

	def handle_month(self, month: int):
		data = self.read_month(month)
		self.write_month(month, data)

	def save(self):
		self.workbook.close()

def main():
	'''Driver Code'''
	writer = Writer(os.path.join(DEFAULT_OUTPUT_DIR, f'transactions {YEAR}.xlsx'))
	for month in MONTHS.keys():
		writer.handle_month(month)
		break
	writer.save()

if __name__ == '__main__':
	main()
