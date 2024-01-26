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
			'merged': self.workbook.add_format({
				'bold': True,
				'align': 'center',
				'bg_color': '#4e81bd',
				'font_color': 'white',
				'font_size': 15,
				'border_color': 'white',
				'border': 1
			}),
		}
		self.data = {}
		self.reset_style_count()

	def reset_style_count(self):
		self.style_count = STARTING_STYLE_COUNT

	def get_style(self, override: int | None = None) -> dict[str, str]:
		result = { 'style': f'Table Style medium {self.style_count if override is None else override}' }
		if override is None:
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
	def columns(df: pd.DataFrame, *column_kwargs_list: dict) -> list[dict]:
		return [
			{ 'header': column, **column_kwargs }
			for column, column_kwargs in zip(df.columns, column_kwargs_list)
		]

	def write_table(self, data: pd.DataFrame, table_name: str, sheet, start_row: int, start_col: int, columns: list[dict], total: bool = False, headers: bool = True) -> (int ,int):
		'''
		:return: (start_row, start_col) the new start row and col after the space taken up by the table
		'''
		rows, cols = data.shape
		if total and headers:
			rows += 1
		if not total and not headers:
			rows -= 1
		cols -= 1 # no index
		sheet.add_table(start_row, start_col, start_row + rows, start_col + cols, {
			'columns': columns,
			'name': table_name,
			'header_row': headers,
			'total_row': total,
			'data': data.values.tolist(),
			**self.get_style()
		})
		return start_row + rows + 1, start_col + cols + 1

	def write_pie_chart(self, name: str, table_name: str, sheet, start_row: int, start_col: int, categories_field: str, values_field: str, i: int = 0, j: int = 0):
		size = 480
		chart = self.workbook.add_chart({ 'type': 'pie' })
		chart.set_title({ 'name': name })
		chart.set_legend({ 'position': 'none' })
		chart.add_series({
			'categories': f'={table_name}[{categories_field}]',
			'values': f'={table_name}[{values_field}]',
			'data_labels': { 'category': True, 'value': True, 'percentage': True, 'position': 'best_fit' }
		})
		chart.set_size({
			'width': size,
			'height': size,
			'x_offset': j * size,
			'y_offset': i * size,
		})
		sheet.insert_chart(start_row, start_col, chart)

	def write_month_table(self, data: pd.DataFrame, sheet, month: int, start_row: int, start_col: int) -> (int, int):
		'''
		:return: (start_row, start_col) the new start row and col after the space taken up by the table
		'''
		before = start_row, start_col
		if 1 <= month <= 12:
			start_row, start_col = self.write_month_table_helper(
				data,
				sheet,
				month,
				start_row,
				start_col,
			)
		else:
			col = start_col
			for month in range(1, 13):
				start_row, start_col = self.write_month_table_helper(
					data,
					sheet,
					month,
					start_row,
					col,
					header = True
				)
		sheet.conditional_format(*before, start_row, start_col -1, {
			'type': '3_color_scale',
			'min_color': '#63be7b',
			'mid_color': '#ffeb84',
			'max_color': '#f8696b',
		})
		return start_row, start_col

	def write_month_table_helper(self, data: pd.DataFrame, sheet, month: int, start_row: int, start_col: int, header: bool = False) -> (int, int):
		'''
		:return: (start_row, start_col) the new start row and col after the space taken up by the table
		'''
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
		if header:
			sheet.merge_range(start_row, start_col, start_row, start_col + cols, MONTHS[month], self.formats['merged'])
			start_row += 1
		bounds = start_row, start_col, start_row + rows, start_col + cols
		sheet.add_table(*bounds, {
			'columns': [ { 'header': day, 'format': self.formats['currency'] } for day in (
				'Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'
			) ],
			'data': cal.values.tolist(),
			**self.get_style(STARTING_STYLE_COUNT)
		})
		sheet.set_column(start_col, start_col + cols, 10)
		return start_row + rows + 1, start_col + cols + 1

	def write_month(self, month: int, data: pd.DataFrame, carry_over: float, sheet_name: str = None):
		self.reset_style_count()
		column_currency_kwargs = { 'format': self.formats['currency'], 'total_function': 'sum' }
		column_total_kwargs = { 'total_string': 'Total' }
		column_date_kwargs = { 'format': self.formats['date'] }
		column_percent_kwargs = { 'format': self.formats['percent'] }
		pivot_kwargs = { 'values': [ 'Amount', 'CashBack Reward'], 'aggfunc': 'sum' }
		pivot_columns_args = column_percent_kwargs, column_currency_kwargs, column_currency_kwargs
		sheet_name = MONTHS[month] if sheet_name is None else sheet_name
		sheet = self.workbook.add_worksheet(sheet_name)
		start_row, start_col = 0, 0
		_, start_col = self.write_table(
			data,
			sheet_name + 'Table',
			sheet,
			start_row,
			start_col,
			self.columns(
				data,
				{ **column_total_kwargs, **column_date_kwargs },
				{},
				column_currency_kwargs,
				{},
				column_percent_kwargs,
				column_currency_kwargs
			),
			total=True
		)
		pivot = data.pivot_table(
			index = 'Category',
			**pivot_kwargs
		).reset_index()
		cat_table_name = sheet_name + 'CatPivot'
		start_row, max_col = self.write_table(
			pivot,
			cat_table_name,
			sheet,
			start_row,
			start_col,
			self.columns(pivot, *pivot_columns_args),
		)
		data['Day'] = data['Date'].apply(lambda x: x.strftime('%w%a'))
		pivot = data.pivot_table(
			index = 'Day',
			**pivot_kwargs
		).reset_index()
		pivot['Day'] = pivot['Day'].apply(lambda x: x[1:])
		day_table_name = sheet_name + 'DayPivot'
		start_row, col = self.write_table(
			pivot,
			day_table_name,
			sheet,
			start_row,
			start_col,
			self.columns(pivot, *pivot_columns_args),
		)
		max_col = max(max_col, col)
		pivot = data.pivot_table(
			index = 'CashBack %',
			**pivot_kwargs
		).reset_index()
		cash_back_table_name = sheet_name + 'CashBackPivot'
		start_row, col = self.write_table(
			pivot,
			cash_back_table_name,
			sheet,
			start_row,
			start_col,
			self.columns(pivot, *pivot_columns_args),
		)
		max_col = max(max_col, col)
		budget_info = pd.DataFrame([
			['Budget', BUDGET_PER_MONTH[month]],
			['Carry Over', carry_over],
			['New Budget', BUDGET_PER_MONTH[month] + carry_over],
			['Remaining', BUDGET_PER_MONTH[month] + carry_over - data.Amount.sum()],
		])
		start_row, col = self.write_table(
			budget_info,
			sheet_name + 'BudgetTable',
			sheet,
			start_row,
			start_col,
			[{}, { 'format': self.formats['currency'] }],
			headers = False
		)
		sheet.autofit()
		start_col = max_col = max(max_col, col)
		start_row = 0
		start_row, _ = self.write_month_table(
			data,
			sheet,
			month,
			start_row,
			start_col,
		)
		for i, value_field in enumerate(('Amount', 'CashBack Reward')):
			for j, (category_field, table_name) in enumerate((
				('Category', cat_table_name),
				('Day', day_table_name),
				('CashBack %', cash_back_table_name)
			)):
				self.write_pie_chart(
					f'{value_field} By {category_field}',
					table_name,
					sheet,
					start_row,
					start_col,
					category_field,
					value_field,
					i,
					j
				)

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
