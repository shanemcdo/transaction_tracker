#!/usr/bin/env python3

import os
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from datetime import datetime
import calendar
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
STARTING_STYLE_COUNT = 9

def parse_date(date: str) -> datetime:
	return datetime.strptime(date, '%m/%d/%Y')

def stringify_date(day: int) -> str:
	if day < 1:
		return ''
	match day % 10:
		case 1 if day != 11:
			return f'{day}st'
		case 2 if day != 12:
			return f'{day}nd'
		case 3 if day != 13:
			return f'{day}rd'
		case _:
			return f'{day}th'

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
			'percent': self.workbook.add_format({ 'num_format': '0.00%' }),
			'date': self.workbook.add_format({ 'num_format': 'mm/dd/yyyy' }),
		}
		self.data = {}
		self.reset_style_count()

	def reset_style_count(self):
		self.style_count = STARTING_STYLE_COUNT

	def get_style(self) -> dict[str, str]:
		result = { 'style': f'Table Style Medium {self.style_count}' }
		self.style_count += 1
		return result

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

	@staticmethod
	def parse_note(note: str, sep: str = '|') -> float:
		if sep in note:
			try:
				note, cashback = map(lambda x: x.strip('%\n\r\t '), note.split(sep, 1))
				return note, float(cashback) / 100
			except ValueError:
				pass
		return note, 0.0

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
		data.Date = data.Date.apply(parse_date)
		tuple_col = data.Note.apply(self.parse_note)
		data.Note = tuple_col.apply(lambda x: x[0])
		data['CashBack %'] = tuple_col.apply(lambda x: x[1])
		data['CashBack Reward'] = data.Amount * data['CashBack %']
		self.data[month] = data.copy()
		return data, carry_over

	@staticmethod
	def columns(df: pd.DataFrame, *column_kwargs_list: dict) -> dict[str, list[dict]]:
		return { 'columns': [
			{ 'header': column, **column_kwargs }
			for column, column_kwargs in zip(df.columns, column_kwargs_list)
		] }

	# TODO: clean this up in separate funcs
	def write_month(self, month: int, data: pd.DataFrame, carry_over: float, sheet_name: str = None):
		self.reset_style_count()
		chart_series_kwargs = { 'data_labels': { 'category': True, 'value': True, 'percentage': True, 'position': 'best_fit' } }
		chart_legend_kwargs = { 'position': 'none' }
		chart_insert_kwargs = { 'y_scale': 2 }
		pivot_kwargs = { 'values': 'Amount', 'aggfunc': 'sum' }
		column_currency_kwargs = { 'format': self.formats['currency'], 'total_function': 'sum' }
		column_total_kwargs = { 'total_string': 'Total' }
		column_date_kwargs = { 'format': self.formats['date'] }
		sheet_name = MONTHS[month] if sheet_name is None else sheet_name
		sheet = self.workbook.add_worksheet(sheet_name)
		rows, cols = data.shape
		cols -= 1
		sheet.add_table(0, 0, rows + 1, cols, {
			**self.columns(
				data,
				{ **column_total_kwargs, **column_date_kwargs },
				{},
				column_currency_kwargs,
				{},
				{ 'format': self.formats['percent'] },
				column_currency_kwargs
			),
			'name': sheet_name + 'Table',
			'total_row': True,
			'data': data.values.tolist(),
			**self.get_style()
		})
		start_col = cols + 1
		pivot = data.pivot_table(
			index = 'Category',
			**pivot_kwargs
		).reset_index()
		rows, cols = pivot.shape
		cols -= 1
		table_name = sheet_name + 'CatPivot'
		sheet.add_table(0, start_col, rows + 1, start_col + cols, {
			**self.columns(
				pivot,
				column_total_kwargs,
				column_currency_kwargs,
			),
			'name': table_name,
			'total_row': True,
			'data': pivot.values.tolist(),
			**self.get_style()
		})
		start_row = rows + 2
		budget_info = pd.DataFrame([
			['Budget', BUDGET_PER_MONTH[month]],
			['Carry Over', carry_over],
			['New Budget', f'={xl_rowcol_to_cell(start_row, start_col + 1)}+{xl_rowcol_to_cell(start_row + 1, start_col + 1)}'],
			['Remaining', f'={xl_rowcol_to_cell(start_row + 2, start_col + 1)}-{xl_rowcol_to_cell(start_row - 1, start_col + 1)}'],
		])
		sheet.add_table(start_row, start_col, start_row + budget_info.shape[0] - 1, start_col + budget_info.shape[1] - 1, {
			'columns': [{}, { 'format': self.formats['currency'] }],
			'header_row': False,
			'data': budget_info.values.tolist(),
			'name': sheet_name + 'BudgetTable',
			**self.get_style()
		})
		start_row += budget_info.shape[0]
		chart = self.workbook.add_chart({ 'type': 'pie' })
		chart.set_title({ 'name': 'By Category' })
		chart.set_legend(chart_legend_kwargs)
		chart.add_series({
			'categories': f'={sheet_name}!{table_name}[Category]',
			'values': f'={sheet_name}!{table_name}[Amount]',
			**chart_series_kwargs
		})
		sheet.insert_chart(max(start_row, 11), start_col, chart, chart_insert_kwargs)
		start_col += cols + 1
		data['Day'] = data['Date'].apply(lambda x: x.strftime('%w%a'))
		pivot = data.pivot_table(
			index = 'Day',
			**pivot_kwargs
		).reset_index()
		pivot['Day'] = pivot['Day'].apply(lambda x: x[1:])
		rows, cols = pivot.shape
		cols -= 1
		table_name = sheet_name + 'DayPivot'
		sheet.add_table(0, start_col, rows + 1, start_col + cols, {
			**self.columns(
				pivot,
				column_total_kwargs,
				column_currency_kwargs,
			),
			'name': table_name,
			'total_row': True,
			'data': pivot.values.tolist(),
			**self.get_style()
		})
		chart = self.workbook.add_chart({ 'type': 'pie' })
		chart.set_title({ 'name': 'By Day' })
		chart.set_legend(chart_legend_kwargs)
		chart.add_series({
			'categories': f'={sheet_name}!{table_name}[Day]',
			'values': f'={sheet_name}!{table_name}[Amount]',
			**chart_series_kwargs
		})
		sheet.insert_chart(max(start_row, 11), start_col + 4, chart, chart_insert_kwargs)
		start_col += pivot.shape[1]
		if 1 <= month <= 12:
			cal = []
			for row in calendar.monthcalendar(YEAR, month):
				cal.append(map(stringify_date, row))
				cal.append((
					data.loc[data.Date == f'{month:02d}/{cell:02d}/{YEAR:04d}', 'Amount'].sum()
					if cell != 0 else ''
					for cell in row
				))
			cal = pd.DataFrame(cal)
			rows, cols = cal.shape
			cols -= 1
			bounds = 0, start_col, rows, start_col + cols
			sheet.add_table(*bounds, {
				'columns': [ { 'header': day, 'format': self.formats['currency'] } for day in (
					'Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'
				) ],
				'data': cal.values.tolist(),
				**self.get_style()
			})
			sheet.conditional_format(*bounds, {
				'type': '3_color_scale',
				'min_color': '#63be7b',
				'mid_color': '#ffeb84',
				'max_color': '#f8696b',
			})
			sheet.set_column(start_col, start_col + cols, 10)
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
	calendar.setfirstweekday(calendar.SUNDAY)
	writer = Writer(os.path.join(DEFAULT_OUTPUT_DIR, f'transactions {YEAR}.xlsx'))
	for month in MONTHS.keys():
		writer.handle_month(month)
	writer.write_summary()
	writer.save()

if __name__ == '__main__':
	main()
