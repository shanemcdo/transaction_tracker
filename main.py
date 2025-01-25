#!/usr/bin/env python3

import os
import xlsxwriter
from datetime import datetime
import calendar
import pandas as pd
from glob import glob
import json
from functools import reduce

INCOME_CATEGORIES = [
	'Cashback',
	'Salary',
	'Fatherly Support',
	'Check',
	'Reward',
	'Sale',
	'Carry Over',
]
MONTHS = {
	1: 'January',
	2: 'February',
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
BUDGETS_DIR = './budgets/'
BALANCES_DIR = './balances/'
RAW_TRANSACTIONS_DIR = './raw_transactions/'
TRANSACTION_REPORTS_DIR = './transaction_reports/'
# unix glob format
RAW_TRANSACTION_FILENAME_FORMAT = 'Transactions {0} 1, {1} - {0} ??, {1}*.csv'
EMPTY = pd.DataFrame({
	'Date': [],
	'Category': [],
	'Amount': [],
	'Note': [],
	'CashBack %': [],
	'CashBack Reward': [],
	'Account': [],
})
STARTING_STYLE_COUNT = 9
ENDING_STYLE_COUNT = 14
DEFAULT_ACCOUNT = 'Monthly'
STARTING_YEAR = 2024

def get_year() -> int:
	return datetime.now().year

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
				'bg_color': '#222222',
				'font_color': '#eeeeee',
				'font_size': 15,
				'border_color': 'white',
				'border': 1
			}),
		}
		# {
		#   year: {
		#     month: DF [ Date, Category, Amount, Note, Cashback %, Cashback reward ],
		#     ...
		#   }, ...
		# }
		self.data = {}
		# {
		#   year: {
		#     month: DF [ Category, Expected ],
		#     ...
		#   }, ...
		# }
		self.monthly_budget = {}
		# {
		#   year: {
		#     balance category: starting value,
		#     ...
		#   }, ...
		# }
		self.starting_balances = {}
		# {
		#   balance category: starting value,
		#   ...
		# }
		self.balances = {}
		self.reset_style_count()
		self.set_year(get_year())

	def reset_position(self):
		'''
		reset all position back to 0, 0
		'''
		self.row = 0
		self.column = 0
		self.next_row = 0 # always zero
		self.next_column = 0

	def go_to_next(self):
		'''
		go to top of next available column
		'''
		self.column = self.next_column
		self.row = self.next_row

	def reset_style_count(self):
		self.style_count = STARTING_STYLE_COUNT

	def get_style(self, override: int | None = None) -> dict[str, str]:
		result = { 'style': f'Table Style medium {self.style_count if override is None else override}' }
		if override is None:
			self.style_count += 1
			if self.style_count > ENDING_STYLE_COUNT:
				self.reset_style_count()
		return result

	def set_starting_balances(self):
		'''
		set the starting balances of the year
		'''
		self.starting_balances[self.year + 1] = self.balances.copy()

	def get_balances(self):
		'''
		get starting balances from json file

		example file:
		{
			"Bigger purchases": 0,
			"Emergency": 1000
		}
		'''
		filename = f'starting_balances{self.year}.json'
		filepath = os.path.join(BALANCES_DIR, filename)
		try:
			with open(filepath) as f:
				self.starting_balances[self.year] = json.load(f)
				self.reset_balances()
		except FileNotFoundError:
			pass

	def reset_balances(self):
		'''
		set balances back to starting balances
		'''
		self.balances = self.starting_balances[self.year].copy()

	def get_budget_df(self, month: int) -> str:
		'''
		read the budget from the file

		example file:
		Category,Expected
		Rent & Utilities,2490.0
		Investing,500.0
		Fuel,150.0
		Groceries,500.0
		Eating Out,300.0
		Other,200.0
		'''
		df = pd.read_csv(os.path.join(BUDGETS_DIR, f'{self.year}{month:02d}budget.csv'))
		return df

	def get_csv_filename_from_month(self, month: str) -> str:
		# e.g. 'Transactions Nov 1, 2024 - Nov 30, 2024 (7).csv'
		glob_pattern = RAW_TRANSACTION_FILENAME_FORMAT.format(month, self.year)
		files = glob(
			glob_pattern,
			root_dir = RAW_TRANSACTIONS_DIR
		)
		if len(files) < 1:
			raise FileNotFoundError(f'Could not find any matches for {glob_pattern}')
		elif len(files) == 1:
			file = files[0]
		else:
			filename = files[-1]
			biggest = -1, None
			for file in files:
				if '(' not in file:
					continue
				number = int(file[file.find('(') + 1 : file.find(')')])
				if number > biggest[0]:
					biggest = number, file
			file = biggest[1]
		return os.path.join(RAW_TRANSACTIONS_DIR, file)

	@staticmethod
	def parse_note(note: str, sep: str = '|') -> (str, float):
		'''
		Split note on seperator string and return parsed note and cashback
		:note: original note containing note message and cashback %
		:sep: the string to split the note on
		:return: note message and cashback percent in a tuple
		'''
		note = str(note)
		if note == 'nan':
			note = ''
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
			sep ='\s*,\s*',
			# get rid of warning
			engine='python'
		).sort_values(by = 'Date')
		data.Amount *= -1
		data = data[data.Category != 'Carry Over']
		data.Date = data.Date.apply(parse_date)
		tuple_col = data.Note.apply(self.parse_note)
		data.Note = tuple_col.apply(lambda x: x[0])
		data['CashBack %'] = tuple_col.apply(lambda x: x[1])
		data['CashBack Reward'] = data.Amount * data['CashBack %']
		self.data[self.year][month] = data.copy()
		self.monthly_budget[self.year][month] = self.get_budget_df(month)
		return data

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

	def write_table(self, data: pd.DataFrame, table_name: str, sheet, columns: list[dict], total: bool = False, headers: bool = True):
		'''
		write pandas data to an excel table at the location saved in the class
		:data: data to write to excel table
		:table_name: name of the table in excel
		:sheet: sheet to write to
		:columns: column config data that contains info about columns
			https://xlsxwriter.readthedocs.io/working_with_tables.html#columns
		:total: whether or not to include the total row
		:headers: whether or not to include the header row
		'''
		self.row, col = self.write_table_at(
			data,
			table_name,
			sheet,
			self.row,
			self.column,
			columns,
			total,
			headers
		)
		if col > self.next_column:
			self.next_column = col

	def write_table_at(self, data: pd.DataFrame, table_name: str, sheet, start_row: int, start_col: int, columns: list[dict], total: bool = False, headers: bool = True) -> (int ,int):
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

	def write_pie_chart_at(self, name: str, chart_type: str, table_name: str, sheet, start_row: int, start_col: int, categories_field: str, values_field: str, i: int = 0, j: int = 0, show_value: bool = True):
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
				self.write_pie_chart_at(..., i = 0, j = 0)
				self.write_pie_chart_at(..., i = 1, j = 0)
				self.write_pie_chart_at(..., i = 0, j = 1)
				self.write_pie_chart_at(..., i = 1, j = 1)
			this will create 4 charts all right next to eachother in a square
		'''
		chart = self.workbook.add_chart({ 'type': chart_type })
		chart.set_title({ 'name': name })
		chart.set_legend({ 'position': 'none' })
		chart.add_series({
			'categories': f'={table_name}[{categories_field}]',
			'values': f'={table_name}[{values_field}]',
			'data_labels': {
				'category': chart_type == 'pie',
				'value': show_value,
				'percentage': True,
				'position': 'best_fit' if chart_type == 'pie' else 'outside_end'
			}
		})
		size = 520
		chart.set_size({
			'width': size,
			'height': size,
			'x_offset': j * size,
			'y_offset': i * size,
		})
		sheet.insert_chart(start_row, start_col, chart)

	def write_month_table(self, data: pd.DataFrame, sheet, month: int):
		'''
		writes a table that shows the sum of all transactions on each day of the month
		uses conditional formatting
		:data: the pandas dataframe containing the transaction data for the given month
		:sheet: sheet to write to
		:month: int 1-13, 1-12 represent the months of the year 13 represents all of the months
		:return: (start_row, start_col) the new start row and col after the space taken up by the table
		'''
		self.row, col = self.write_month_table_at(
			data,
			sheet,
			month,
			self.row,
			self.column
		)
		if col > self.next_column:
			self.next_column = col

	def write_month_table_at(self, data: pd.DataFrame, sheet, month: int, start_row: int, start_col: int) -> (int, int):
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
				if month not in self.data[self.year]:
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
			'min_type': 'num',
			'min_value': 0,
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
		for row in calendar.monthcalendar(self.year, month):
			cal.append(map(stringify_date, row))
			cal.append((
				data.loc[data.Date == f'{month:02d}/{cell:02d}/{self.year:04d}', 'Amount'].sum()
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

	def write_title(self, sheet, title: str, width: int):
		self.row, col = self.write_title_at(sheet, title, width, self.row, self.column)
		if col > self.next_column:
			self.next_column = col

	def write_title_at(self, sheet, title: str, width: int, start_row: int, start_col: int) -> (int, int):
		sheet.merge_range(self.row, self.column, self.row, self.column + width - 1, title, self.formats['merged'])
		return self.row + 1, self.column + width

	def write_month(self, month: int, data: pd.DataFrame, sheet_name: str = None, budget: dict = None):
		'''
		Create and write the sheet for a given month
		:month: int 1-14, 1-12 for the months of the year
				13 for all of them
				14 is a total summary regardless of year
		:data: the dataframe contianing the transactions for the month
		:sheet_name: optional, name to give the sheet created, if left None will be the month name
		:budget: optional, budget to use instead of reading from self.monthly_budget
		'''
		self.reset_style_count()
		self.reset_position()
		data_headers = data.columns
		column_total_sum_kwargs = { 'total_function': 'sum' }
		column_currency_kwargs = { 'format': self.formats['currency'], **column_total_sum_kwargs }
		column_total_kwargs = { 'total_string': 'Total' }
		column_date_kwargs = { 'format': self.formats['date'] }
		column_percent_kwargs = { 'format': self.formats['percent'] }
		pivot_kwargs = { 'values': [ 'Amount', 'CashBack Reward'], 'aggfunc': 'sum' }
		pivot_columns_args = column_currency_kwargs, column_currency_kwargs, {}
		sheet_name = self.get_sheetname(month) if sheet_name is None else sheet_name
		sheet = self.workbook.add_worksheet(sheet_name)
		# table of default transactions
		def write_transaction_table(data: pd.DataFrame, table_name: str, include_cashback: bool = True):
			if data.shape[0] == 0: return
			self.write_title(sheet, table_name, len(data.columns))
			self.write_table(
				data if include_cashback else data[['Date', 'Category', 'Amount', 'Note']],
				sheet_name + table_name.replace(' ', '_').replace('&', '') + 'Table',
				sheet,
				self.columns(
					default_transactions,
					{ **column_total_kwargs, **column_date_kwargs },
					{},
					column_currency_kwargs,
					{},
					{},
					column_percent_kwargs,
					column_currency_kwargs,
				),
				total=True
			)
		default_transactions = data.loc[data.Account == DEFAULT_ACCOUNT, data_headers]
		income_condition = default_transactions.Category.map(lambda x: x in INCOME_CATEGORIES) & (default_transactions.Amount < 0)
		default_income_transactions = default_transactions[income_condition]
		default_income_transactions.Amount *= -1
		default_transactions = default_transactions[~income_condition]
		positive_default_transactions = default_transactions[default_transactions.Amount > 0]
		all_expenses = data.loc[data.Category.map(lambda x: x not in INCOME_CATEGORIES), data_headers]
		eligible_expenses = all_expenses[all_expenses.Category.map(lambda x: x not in [ 'Investing', 'Transfer' ])]
		write_transaction_table(default_transactions, DEFAULT_ACCOUNT)
		write_transaction_table(default_income_transactions, DEFAULT_ACCOUNT + ' Income', False)
		accounts = data.loc[data.Account != DEFAULT_ACCOUNT, 'Account'].sort_values().unique()
		pre_balances_sum = sum(self.balances.values())
		for account in accounts:
			transactions = data.loc[data.Account == account, data_headers]
			transactions.Amount *= -1
			write_transaction_table(transactions, account)
			self.balances[account] = self.balances.get(account, 0) + transactions.Amount.sum()
		self.set_starting_balances()
		write_transaction_table(all_expenses, 'All Expenses')
		self.go_to_next()
		self.reset_style_count()
		# Total budget / carryover / remaining
		income_sum = default_income_transactions.Amount.sum()
		expenses_sum = default_transactions.Amount.sum()
		income_and_balances_sum = income_sum + pre_balances_sum
		all_expenses_sum = all_expenses.Amount.sum()
		budget_info = pd.DataFrame([
			['Monthly Income', income_sum],
			['Monthly Expenses', expenses_sum],
			['Monthly Expenses - transfers', expenses_sum - default_transactions[default_transactions.Category == 'Transfer'].Amount.sum()],
			['Monthly Income - Monthly Expenses', income_sum - expenses_sum],
			['Monthly Income - All Expenses', income_sum - all_expenses_sum],
			['Income + Balances', income_and_balances_sum],
			['All Expenses', all_expenses_sum],
			['Income + Balances - All Expenses', income_and_balances_sum - all_expenses_sum],
		])
		self.write_title(sheet, 'Overall Budget', len(budget_info.columns))
		self.write_table(
			budget_info,
			sheet_name + 'BudgetTable',
			sheet,
			[{}, column_currency_kwargs],
			headers = False
		)
		# balances
		balances_df = pd.DataFrame(
			([
				f'{account}',
				self.balances.get(account, 0),
				-data[data.Account == account].Amount.sum(),
				data[(data.Account == account) & (data.Amount > 0)].Amount.sum(),
				-data[(data.Account == account) & (data.Amount < 0)].Amount.sum(),
				len(data[data.Account == account].Amount),
			] for account in sorted(set((*accounts, *self.balances.keys())))),
			columns = ['Account', 'New Balance', 'Net Change', 'Spent', 'Saved', 'Transaction Count']
		)
		balances_df = balances_df[
			(abs(balances_df['New Balance']) >= 0.001) |
			(abs(balances_df['Net Change']) >= 0.001) |
			(abs(balances_df['Spent']) >= 0.001) |
			(abs(balances_df['Saved']) >= 0.001)
		]
		self.write_title(sheet, 'Balances', len(balances_df.columns))
		self.write_table(
			balances_df,
			sheet_name + 'BalancesTable',
			sheet,
			self.columns(
				balances_df,
				{},
				column_currency_kwargs,
				column_currency_kwargs,
				column_currency_kwargs,
				column_currency_kwargs,
				column_total_sum_kwargs,
			),
			total = True
		)
		# Budget Categories Table
		pivot = default_transactions.pivot_table(
			index = 'Category',
			**pivot_kwargs
		).reset_index().join(
			all_expenses.Category.value_counts(),
			on='Category'
		).rename(columns={'count': 'Transaction Count'})
		budget = self.monthly_budget[self.year][month] if budget is None else budget
		budget_categories_df = budget.join(
			pivot[['Category', 'Amount', 'Transaction Count']].set_index('Category'),
			on='Category',
		)
		all_cats = set()
		for cat in budget_categories_df.Category:
			if '&' not in cat:
				all_cats.add(cat)
				continue
			cats = set(map(lambda x: x.strip(), cat.split('&')))
			all_cats.update(cats)
			budget_categories_df.loc[budget_categories_df.Category == cat, 'Amount'] = pivot[pivot.Category.map(lambda x: x in cats)].Amount.sum()
			budget_categories_df.loc[budget_categories_df.Category == cat, 'Transaction Count'] = pivot[pivot.Category.map(lambda x: x in cats)]['Transaction Count'].sum()
		budget_categories_df.Amount = budget_categories_df.Amount.fillna(0)
		budget_categories_df['Transaction Count'] = budget_categories_df['Transaction Count'].fillna(0)
		budget_categories_df.loc[budget_categories_df.Category == 'Other', 'Amount'] = pivot[pivot.Category.map(lambda x: (x not in all_cats or x == 'Other') and x != 'Transfer')].Amount.sum()
		budget_categories_df.loc[budget_categories_df.Category == 'Other', 'Transaction Count'] = len(pivot[pivot.Category.map(lambda x: (x not in all_cats or x == 'Other') and x != 'Transfer')]['Transaction Count'])
		budget_categories_df['Remaining'] = budget_categories_df.Expected - budget_categories_df.Amount
		budget_categories_df['Usage %'] = budget_categories_df['Amount'] / budget_categories_df['Expected']
		transaction_count_col = budget_categories_df.pop('Transaction Count')
		budget_categories_df.insert(len(budget_categories_df.columns), 'Transaction Count', transaction_count_col)
		budget_categories_table_name = sheet_name + 'BudgetCategoriesTable'
		self.write_title(sheet, 'Budget Categories', len(budget_categories_df.columns))
		before_row = self.row
		self.write_table(
			budget_categories_df,
			budget_categories_table_name,
			sheet,
			self.columns(
				budget_categories_df,
				{},
				column_currency_kwargs,
				column_currency_kwargs,
				column_currency_kwargs,
				column_percent_kwargs,
				column_total_sum_kwargs,
			),
			True
		)
		sheet.conditional_format(before_row + 1, self.column + 4, self.row - 1, self.column + 4, {
			'type': '3_color_scale',
			'min_color': '#63be7b',
			'min_type': 'num',
			'min_value': 0,
			'mid_color': '#ffeb84',
			'mid_type': 'num',
			'mid_value': 0.5,
			'max_color': '#f8696b',
			'max_type': 'num',
			'max_value': 1,
		})
		# category pivot & reimbursement/refund table
		pivot = all_expenses.pivot_table(
			index = 'Category',
			**pivot_kwargs
		).reset_index()
		categories_list = sorted(all_expenses.Category.unique())
		spent_list =      [ (all_expenses[(all_expenses.Category == cat) & (all_expenses.Amount > 0)]).Amount.sum() for cat in categories_list ]
		reimbursed_list = [ (all_expenses[(all_expenses.Category == cat) & (all_expenses.Amount < 0)]).Amount.sum() for cat in categories_list ]
		reimbursement_df = pd.DataFrame({
			'Category': categories_list,
			'Spent': spent_list,
			'Reimbursed/Refunded': reimbursed_list,
		}).join(
			pivot[['Category', 'Amount', 'CashBack Reward']].set_index('Category'),
			on='Category'
		).join(
			all_expenses.Category.value_counts(),
			on='Category'
		).rename(columns={'count': 'Transaction Count'})
		reimbursement_df['Reimbursed/Refunded'] *= -1
		cat_table_name = sheet_name + 'CatPivot'
		self.write_title(sheet, 'Categories Pivot', len(reimbursement_df.columns))
		self.write_table(
			reimbursement_df,
			cat_table_name,
			sheet,
			self.columns(
				reimbursement_df,
				{},
				column_currency_kwargs,
				column_currency_kwargs,
				column_currency_kwargs,
				column_currency_kwargs,
				column_total_sum_kwargs,
			),
			True
		)
		# Account Pivot
		pivot = all_expenses.pivot_table(
			index = 'Account',
			**pivot_kwargs
		).reset_index()
		account_list = sorted(all_expenses.Account.unique())
		spent_list =      [ (all_expenses[(all_expenses.Account == account) & (all_expenses.Amount > 0)]).Amount.sum() for account in account_list ]
		reimbursed_list = [ (all_expenses[(all_expenses.Account == account) & (all_expenses.Amount < 0)]).Amount.sum() for account in account_list ]
		reimbursement_df = pd.DataFrame({
			'Account': account_list,
			'Spent': spent_list,
			'Reimbursed/Refunded': reimbursed_list,
		}).join(
			pivot[['Account', 'Amount', 'CashBack Reward']].set_index('Account'),
			on='Account'
		).join(
			all_expenses.Account.value_counts(),
			on='Account'
		).rename(columns={
			'count': 'Transaction Count',
			'Spent': 'Amount',
			'Amount': 'Net Change',
			'Reimbursed/Refunded': 'Saved'
		})
		reimbursement_df['Net Change'] = reimbursement_df.Amount + reimbursement_df.Saved
		reimbursement_df.Saved *= -1
		# reimbursement_df.loc[reimbursement_df.Account == DEFAULT_ACCOUNT, 'Account'] = DEFAULT_ACCOUNT + ' (excluding transfers)'
		self.write_title(sheet, 'Account Pivot', len(reimbursement_df.columns))
		account_table_name = sheet_name + 'AccountPivot'
		self.write_table(
			reimbursement_df,
			account_table_name,
			sheet,
			self.columns(
				reimbursement_df,
				{},
				column_currency_kwargs,
				column_currency_kwargs,
				column_currency_kwargs,
				column_currency_kwargs,
				column_total_sum_kwargs,
			),
			True
		)
		# day pivot
		all_expenses['Day'] = all_expenses['Date'].apply(lambda x: x.strftime('%w%a'))
		pivot = all_expenses.pivot_table(
			index = 'Day',
			**pivot_kwargs
		).reset_index()
		for day in [ '0Sun', '1Mon', '2Tue', '3Wed', '4Thu', '5Fri', '6Sat']:
			if day not in pivot['Day'].values:
				pivot = pd.concat([pivot, pd.DataFrame([[day, 0, 0]], columns=pivot.columns)])
		pivot = pivot.sort_values('Day').join(
			all_expenses.Day.value_counts(),
			on='Day'
		).rename(columns={'count': 'Transaction Count'}).fillna(0)
		pivot['Day'] = pivot['Day'].apply(lambda x: x[1:])
		day_table_name = sheet_name + 'DayPivot'
		self.write_title(sheet, 'Day Pivot', len(pivot.columns))
		print(pivot)
		# replace nan with zero here
		self.write_table(
			pivot,
			day_table_name,
			sheet,
			self.columns(pivot, {}, *pivot_columns_args),
		)
		# cashback pivot
		pivot = all_expenses.pivot_table(
			index = 'CashBack %',
			**pivot_kwargs
		).reset_index().join(
			all_expenses['CashBack %'].value_counts(),
			on='CashBack %'
		).rename(columns={'count': 'Transaction Count'})
		cash_back_table_name = sheet_name + 'CashBackPivot'
		self.write_title(sheet, 'Cashback Pivot', len(pivot.columns))
		self.write_table(
			pivot,
			cash_back_table_name,
			sheet,
			self.columns(pivot, column_percent_kwargs, *pivot_columns_args),
		)
		# avg cashback 
		cashback_sum = all_expenses['CashBack Reward'].sum()
		eligible_expenses_sum = eligible_expenses.Amount.sum()
		cashback_info = pd.DataFrame({
			'Eligible Spending Sum': [ eligible_expenses_sum ],
			'Cashback Sum': [ cashback_sum ],
			'Average cashback yield': [ cashback_sum / eligible_expenses_sum ],
			'Average cashback yield excluding 0% cashback': [ cashback_sum / pivot[pivot['CashBack %'] != 0].Amount.sum() ],
		})
		self.write_title(sheet, 'Cashback Info', len(cashback_info.columns))
		self.write_table(
			cashback_info,
			sheet_name + 'CashBackInfoTable',
			sheet,
			self.columns(
				cashback_info,
				column_currency_kwargs,
				column_currency_kwargs,
				column_percent_kwargs,
				column_percent_kwargs,
			)
		)
		# day number pivot
		all_expenses_copy = all_expenses.copy()
		all_expenses_copy['Day Number'] = all_expenses.Date.apply(lambda x: int(x.strftime('%-d')))
		pivot = all_expenses_copy.pivot_table(
			index = 'Day Number',
			**pivot_kwargs
		).reset_index()
		for i in range(1, 32):
			if any(pivot['Day Number'] == i):
				continue
			pivot = pd.concat([pivot, pd.DataFrame([[i, 0, 0]], columns = pivot.columns)])
		pivot = pivot.sort_values(by='Day Number').join(
			all_expenses_copy['Day Number'].value_counts(),
			on='Day Number'
		).rename(columns={'count': 'Transaction Count'}).fillna(0)

		day_number_table_name = sheet_name + 'DayNumberPivot'
		self.write_title(sheet, 'Day Number Pivot', len(pivot.columns))
		self.write_table(
			pivot,
			day_number_table_name,
			sheet,
			self.columns(pivot, {}, *pivot_columns_args),
		)
		self.go_to_next()
		sheet.autofit()
		if month != 14:
			# month table
			self.write_month_table(
				all_expenses,
				sheet,
				month
			)
			self.go_to_next()
		# charts
		for i, value_field in enumerate(('Amount', 'Transaction Count', 'CashBack Reward' )):
			for j, (category_field, table_name, chart_type, show_value) in enumerate((
				('Category', cat_table_name, 'pie', True),
				('Account', account_table_name, 'pie', True),
				('Day', day_table_name, 'column', True),
				('CashBack %', cash_back_table_name, 'pie', True),
				('Day Number', day_number_table_name, 'column', False)
			)):
				self.write_pie_chart_at(
					f'{value_field} By {category_field}',
					chart_type,
					table_name,
					sheet,
					self.row,
					self.column,
					category_field,
					value_field,
					i,
					j,
					show_value
				)

	def handle_month(self, month: int) -> bool:
		'''
		read and write data for the month
		:month: int, 1-12 number representing the months
		:returns: true if it succeeds and false if it fails
		'''
		if month < 1 or month > 12:
			raise ValueError(f'month must be between 1-12 inclusive. actual = {month}')
		try:
			data = self.read_month(month)
			if not data.empty:
				self.write_month(month, data)
		except FileNotFoundError as e:
			print(f'Couldn \'t find file for month {month}. Continuing')
			return False
		return True

	def write_summary(self):
		'''
		write a sheet for a summary of the whole year
		'''
		if len(self.data[self.year]) <= 0:
			return
		self.monthly_budget[self.year][13] = pd.concat(self.monthly_budget[self.year].values()).groupby('Category', sort=False).sum().reset_index()
		self.write_month(
			13,
			pd.concat(self.data[self.year].values()),
			f'Summary{self.year}'
		)

	def write_summary_all(self):
		'''
		write a sheet for a summary of all recorded history
		write_summary must be called for this to work correctly
		'''
		budget = pd.concat(map(lambda x: x.get(13, pd.DataFrame()), self.monthly_budget.values())).groupby('Category', sort=False).sum().reset_index()
		data = pd.concat(reduce(lambda x, y: x + list(y.values()), self.data.values(), []))
		self.write_month(14, data, 'SummaryAll', budget)

	def focus(self, month: int):
		'''
		Focus on a specific sheet when the workbook opens
		:month: the sheet to focus on
		'''
		print(self.get_sheetname(month))
		sheet = self.workbook.get_worksheet_by_name(self.get_sheetname(month))
		if sheet:
			sheet.activate()

	def hide(self, month: int):
		'''
		Hide a sheet based on the month number
		:month: the month of the sheet to focus on 1-12
		'''
		sheet = self.workbook.get_worksheet_by_name(self.get_sheetname(month))
		if sheet:
			sheet.hide()

	def get_sheetname(self, month: int):
		'''
		generate a sheet name based on month and year
		:month: the month 1-12 to focus on
		'''
		if month < 1 or month > 12:
			raise ValueError(f'month must be between 1-12 inclusive. actual = {month}')
		return f'{MONTHS[month]}{self.year}'

	def full_screen(self):
		'''Make the window full screen'''
		# just make it big enough to fill any screen
		self.workbook.set_size(1000000, 1000000)

	def save(self):
		self.workbook.close()

	def set_year(self, year: int):
		self.year = year
		if year not in self.data:
			self.data[year] = {}
		if year not in self.monthly_budget:
			self.monthly_budget[year] = {}

def main():
	'''Driver Code'''
	now = datetime.now()
	datestring = now.strftime('%Y%m%d %H%M%S')
	calendar.setfirstweekday(calendar.SUNDAY)
	current_year = get_year()
	writer = Writer(os.path.join(TRANSACTION_REPORTS_DIR, f'transactions {datestring}.xlsx'))
	for year in range(STARTING_YEAR, current_year + 1):
		writer.set_year(year)
		writer.get_balances()
		any_success = False
		for month in range(1,13):
			any_success |= writer.handle_month(month)
			if current_year != year:
				writer.hide(month)
		if any_success:
			writer.reset_balances()
			writer.write_summary()
	writer.set_year(STARTING_YEAR)
	writer.reset_balances()
	writer.write_summary_all()
	writer.set_year(current_year)
	writer.focus(now.month)
	writer.full_screen()
	writer.save()

if __name__ == '__main__':
	main()
