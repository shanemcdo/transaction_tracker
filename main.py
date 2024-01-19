#!/usr/bin/env python3

import os
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from datetime import datetime
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
INPUT_FILENAME_FORMAT = 'Transactions {0} 1, {1} - {0} ??, {1}*.csv'
YEAR = 2024
# TODO: change if budget gets adjusted
BUDGET_PER_MONTH = { i: 1000.00 for i in range(1, 13) }
BUDGET_PER_MONTH[13] = sum(BUDGET_PER_MONTH.values())
EMPTY = pd.DataFrame({
	'Date': [],
	'Category': [],
	'Amount': [],
	'Note': [],
})

class Writer:
	def __init__(self, filename: str):
		self.excelWriter = pd.ExcelWriter(filename, engine='xlsxwriter')
		self.workbook = self.excelWriter.book
		currency = { 'num_format': '$#,##0.00' }
		border = { 'border': True }
		self.formats = {
			'currency': self.workbook.add_format(currency),
			'border': self.workbook.add_format(border),
			'border_currency': self.workbook.add_format({ **border, **currency }),
		}
		self.data = {}

	@staticmethod
	def get_csv_filename_from_month(month: str) -> str:
		glob_pattern = INPUT_FILENAME_FORMAT.format(month, YEAR)
		files = sorted(glob(
				glob_pattern,
				root_dir = DEFAULT_INPUT_DIR
			),
			# make the order accurate
			key = lambda x: x if '(' in x else x.replace('.csv', ' (0).csv')
		)
		if len(files) < 1:
			raise FileNotFoundError(f'Could not find any matches for {glob_pattern}')
		return os.path.join(DEFAULT_INPUT_DIR, files[-1])

	def read_month(self, month: int) -> (pd.DataFrame, float):
		filename = self.get_csv_filename_from_month(MONTHS_SHORT[month])
		data = pd.read_csv(
			filename,
			sep =', |,',
			# get rid of warning
			engine='python'
		).sort_values(by = 'Date')
		carry_over = data.loc[data.Category == 'Carry Over', 'Amount'].sum()
		data = data[data.Category != 'Carry Over']
		self.data[month] = data.copy()
		return data, carry_over

	def write_month(self, month: int, data: pd.DataFrame, carry_over: float, sheet_name: str = None):
		sheet_name = MONTHS[month] if sheet_name is None else sheet_name
		data.to_excel(
			self.excelWriter,
			sheet_name = sheet_name,
			index = False
		)
		rows, cols = data.shape
		cols -= 1
		sheet = self.excelWriter.sheets[sheet_name]
		sheet.add_table(0, 0, 1000, cols, {
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
		if not pivot.empty:
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
				'total_row': 1,
				'style': 'Table Style Medium 10'
			})
		start_row = rows + 2
		budget_info = pd.DataFrame([
			['Budget', BUDGET_PER_MONTH[month]],
			['Carry Over', carry_over],
			['New Budget', f'={xl_rowcol_to_cell(start_row, start_col + 1)}+{xl_rowcol_to_cell(start_row + 1, start_col + 1)}'],
			['Over/Under', f'={xl_rowcol_to_cell(start_row + 2, start_col + 1)}-{xl_rowcol_to_cell(start_row - 1, start_col + 1)}'],
		])
		budget_info.to_excel(
			self.excelWriter,
			sheet_name,
			index = False,
			header = False,
			startrow = start_row,
			startcol = start_col
		)
		chart = self.workbook.add_chart({ 'type': 'pie' })
		chart.set_title({ 'name': 'By Category' })
		chart.set_legend({ 'position': 'bottom' })
		chart.add_series({
			'categories': [sheet_name, 1, start_col, rows, start_col],
			'values': [sheet_name, 1, start_col + 1, rows, start_col + 1],
			'data_labels': { 'value': True, 'percentage': True, 'position': 'best_fit' },
		})
		start_row += budget_info.shape[0]
		sheet.insert_chart(start_row, start_col, chart, {'y_scale': 2})
		start_col += cols + 1
		data['Day'] = data['Date'].apply(lambda x: datetime.strptime(x, '%m/%d/%Y').strftime('%w%a'))
		pivot = data.pivot_table(
			values = 'Amount',
			index = 'Day',
			aggfunc = 'sum'
		).reset_index()
		pivot['Day'] = pivot['Day'].apply(lambda x: x[1:])
		pivot.to_excel(
			self.excelWriter,
			sheet_name = sheet_name,
			index = False,
			startcol = start_col
		)
		rows, cols = pivot.shape
		cols -= 1
		if not pivot.empty:
			sheet.add_table(0, start_col, rows + 1, start_col + cols, {
				'columns': [
					{ 'header': 'Day', 'total_string': 'Total' },
					{
						'header': 'Sum of Amount',
						'format': self.formats['currency'],
						'total_function': 'sum'
					},
				],
				'name': sheet_name + 'Pivot2',
				'total_row': 1,
				'style': 'Table Style Medium 11'
			})
		chart = self.workbook.add_chart({ 'type': 'pie' })
		chart.set_title({ 'name': 'By Day' })
		chart.set_legend({ 'position': 'bottom' })
		chart.add_series({
			'Name': 'By Day',
			'categories': [sheet_name, 1, start_col, rows, start_col],
			'values': [sheet_name, 1, start_col + 1, rows, start_col + 1],
			'data_labels': { 'value': True, 'percentage': True, 'position': 'best_fit' },
		})
		sheet.insert_chart(start_row, start_col + 4, chart, {'y_scale': 2})
		# Stupid hack because format in add_table isn't work
		for cells in ('C:C', 'F:F'):
			sheet.set_column(cells, None, self.formats['currency'])
		sheet.autofit()

	def handle_month(self, month: int):
		try:
			data, carry_over = self.read_month(month)
			if not data.empty:
				self.write_month(month, data, carry_over)
		except FileNotFoundError:
			pass

	def write_summary(self):
		self.write_month(
			13,
			pd.concat(self.data.values()) if len(self.data) > 0 else EMPTY.copy(),
			0,
			'Summary'
		)

	def save(self):
		self.workbook.close()

def main():
	'''Driver Code'''
	writer = Writer(os.path.join(DEFAULT_OUTPUT_DIR, f'transactions {YEAR}.xlsx'))
	for month in MONTHS.keys():
		writer.handle_month(month)
	writer.write_summary()
	writer.save()

if __name__ == '__main__':
	main()
