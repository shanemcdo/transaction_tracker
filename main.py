#!/usr/bin/env python3

import os
import xlsxwriter
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
	13: 'Whole Year'
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
	'CashBack %': [],
	'CashBack Reward': [],
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
	def parse_note(note: str, sep: str = '|') -> (str, float):
		'''
		Split note on seperator string and return parsed note and cashback
		:note: original note containing note message and cashback %
		:sep: the string to split the note on
		:return: note message and cashback percent in a tuple
		'''
		if sep in note:
			try:
				note, cashback = map(lambda x: x.strip('%\n\r\t '), note.split(sep, 1))
				return note, float(cashback) / 100
			except ValueError:
				pass
		return note, 0.0

	def read_month(self, month: int) -> (pd.DataFrame, float):
		'''
		parse csv and modify data for given month
		:month: int 1-12, its the month to read in
		:return: a tuple of the csv data and the carry_over from the previous month
		'''
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
		'''
		:df: data where column names are taken from
		:column_kwargs_list: additional kwargs for each respective column
		:return: a list of inputted kwargs combined with column names
		'''
		return [
			{ 'header': column, **column_kwargs }
			for column, column_kwargs in zip(df.columns, column_kwargs_list)
		]

	def write_table(self, data: pd.DataFrame, table_name: str, sheet, start_row: int, start_col: int, columns: list[dict], total: bool = False, headers: bool = True) -> (int ,int):
		'''
		write pandas data to an excel table
		:data: data to write to excel table
		:table_name: name of the table in excel
		:sheet: sheet to write to
		:start_row: row in sheet to start writing table
		:start_col: column in sheet to start writing table
		:columns: column config data that contains info about columns
			https://xlsxwriter.readthedocs.io/working_with_tables.html#columns
		:total: whether or not to include the total row
		:headers: whether or not to include the header row
		:return: (start_row, start_col) the new start row and col after the space taken up by the table
		'''
		if data.shape[0] < 1:
			data = data.copy()
			data.loc[-1] = ''
			total = False
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

	def write_pie_chart(self, name: str, chart_type: str, table_name: str, sheet, start_row: int, start_col: int, categories_field: str, values_field: str, i: int = 0, j: int = 0):
		'''
		write pandas data to an excel pie chart
		:name: title for the pie chart
		:chart_type: the kind of chart to use e.g. pie or column
		:table_name: name of the table in excel to get the data from
		:sheet: sheet to write to
		:start_row: row in sheet to start writing pie chart
		:start_col: column in sheet to start writing pie chart
		:categories_field: the name of the field in the table where the categories come from
		:values_field: the name of the field in the table where the values come from
		:i: the y coordinate that offsets the chart
		:j: the x coordinate that offsets the chart
			the i and j values are used to create multiple charts right next to each other
			i.e.
				self.write_pie_chart(..., i = 0, j = 0)
				self.write_pie_chart(..., i = 1, j = 0)
				self.write_pie_chart(..., i = 0, j = 1)
				self.write_pie_chart(..., i = 1, j = 1)
			this will create 4 charts all right next to eachother in a square
		'''
		chart = self.workbook.add_chart({ 'type': chart_type })
		chart.set_title({ 'name': name })
		chart.set_legend({ 'position': 'none' })
		chart.add_series({
			'categories': f'={table_name}[{categories_field}]',
			'values': f'={table_name}[{values_field}]',
			'data_labels': { 'category': True, 'value': True, 'percentage': True, 'position': 'best_fit' }
		})
		size = 480
		chart.set_size({
			'width': size,
			'height': size,
			'x_offset': j * size,
			'y_offset': i * size,
		})
		sheet.insert_chart(start_row, start_col, chart)

	def write_month_table(self, data: pd.DataFrame, sheet, month: int, start_row: int, start_col: int) -> (int, int):
		'''
		writes a table that shows the sum of all transactions on each day of the month
		uses conditional formatting
		:data: the pandas dataframe containing the transaction data for the given month
		:sheet: sheet to write to
		:month: int 1-13, 1-12 represent the months of the year 13 represents all of the months
		:start_row: row in sheet to start writing table
		:start_col: column in sheet to start writing table
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
				if month not in self.data:
					continue
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
		a helper function that
		writes a table that shows the sum of all transactions on each day of the month
		:data: the pandas dataframe containing the transaction data for the given month
		:sheet: sheet to write to
		:month: int 1-12 represent the months of the year
		:start_row: row in sheet to start writing table
		:start_col: column in sheet to start writing table
		:header: whether or not to include the month name header
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
		'''
		Create and write the sheet for a given month
		:month: int 1-13, 1-12 for the months of the year and 13 for all of them
		:data: the dataframe contianing the transactions for the month
		:carry_over: the money leftover from last month (negative means overspent)
		:sheet_name: optional, name to give the sheet created, if left None will be the month name
		'''
		self.reset_style_count()
		column_currency_kwargs = { 'format': self.formats['currency'], 'total_function': 'sum' }
		column_total_kwargs = { 'total_string': 'Total' }
		column_date_kwargs = { 'format': self.formats['date'] }
		column_percent_kwargs = { 'format': self.formats['percent'] }
		pivot_kwargs = { 'values': [ 'Amount', 'CashBack Reward'], 'aggfunc': 'sum' }
		pivot_columns_args = column_currency_kwargs, column_currency_kwargs
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
			self.columns(pivot, {}, *pivot_columns_args),
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
			self.columns(pivot, {}, *pivot_columns_args),
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
			self.columns(pivot, column_percent_kwargs, *pivot_columns_args),
		)
		max_col = max(max_col, col)
		data_copy = data.copy()
		data_copy['Day Number'] = data.Date.apply(lambda x: int(x.strftime('%-d')))
		pivot = data_copy.pivot_table(
			index = 'Day Number',
			**pivot_kwargs
		).reset_index()
		day_number_table_name = sheet_name + 'DayNumberPivot'
		start_row, col = self.write_table(
			pivot,
			day_number_table_name,
			sheet,
			start_row,
			start_col,
			self.columns(pivot, {}, *pivot_columns_args),
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
		start_col = max_col = max(max_col, col)
		sheet.autofit()
		start_row = 0
		_, start_col = self.write_month_table(
			data,
			sheet,
			month,
			start_row,
			start_col,
		)
		for i, value_field in enumerate(('Amount', 'CashBack Reward')):
			for j, (category_field, table_name, chart_type) in enumerate((
				('Category', cat_table_name, 'pie'),
				('Day', day_table_name, 'pie'),
				('CashBack %', cash_back_table_name, 'pie'),
				('Day Number', day_number_table_name, 'column')
			)):
				self.write_pie_chart(
					f'{value_field} By {category_field}',
					chart_type,
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
		'''
		read and write data for the month
		:month: int, 1-12 number representing the months
		'''
		try:
			data, carry_over = self.read_month(month)
			if not data.empty:
				self.write_month(month, data, carry_over)
		except FileNotFoundError:
			pass

	def write_summary(self):
		'''
		write a sheet for a summary of the whole year
		'''
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
	datestring = datetime.now().strftime('%Y%m%d %H%M%S')
	calendar.setfirstweekday(calendar.SUNDAY)
	writer = Writer(os.path.join(DEFAULT_OUTPUT_DIR, f'transactions {datestring}.xlsx'))
	for month in MONTHS.keys():
		writer.handle_month(month)
	writer.write_summary()
	writer.save()

if __name__ == '__main__':
	main()
