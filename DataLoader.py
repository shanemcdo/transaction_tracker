from utils import *
from glob import glob
import datetime
import json
import pandas as pd
from functools import reduce

# unix glob format
RAW_TRANSACTION_FILENAME_FORMAT = getenv('RAW_TRANSACTION_FILENAME_FORMAT')
BUDGETS_DIR = getenv('BUDGETS_DIR')
BALANCES_DIR = getenv('BALANCES_DIR')
RAW_TRANSACTIONS_DIR = getenv('RAW_TRANSACTIONS_DIR')

def parse_date(date: str) -> datetime.date:
	return datetime.datetime.strptime(date, '%m/%d/%Y').date()

class DataLoader:
	def __init__(self):
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
		#   balance category: starting value,
		#   ...
		# }
		self.starting_balances = {}

	def get_all_data(self) -> pd.DataFrame:
		'''
		concat all data into one pandas object
		'''
		return pd.concat(reduce(lambda x, y: x + list(y.values()), self.data.values(), [])).sort_values('Date')

	def read_month(self, month: int, year: int) -> pd.DataFrame | None:
		'''
		parse csv and modify data for given month
		:month: int 1-12, its the month to read in
		:year: int, the year to read in
		:return: a tuple of the csv data and the carry_over from the previous month
		'''
		try:
			filename = self.get_csv_filename_from_month(MONTHS_SHORT[month], year)
			data = pd.read_csv(
				filename,
				sep =r'\s*,\s*',
				# get rid of warning
				engine='python'
			).sort_values(by = 'Date')
		except FileNotFoundError:
			return None
		data.Amount *= -1
		data = data.loc[data.Category != 'Carry Over']
		data.Date = data.Date.apply(parse_date)
		tuple_col = data.Note.apply(self.parse_note)
		data.Note = tuple_col.apply(lambda x: x[0])
		data['CashBack %'] = tuple_col.apply(lambda x: x[1])
		data['CashBack Reward'] = data.Amount * data['CashBack %']
		if year not in self.data:
			self.data[year] = {}
		self.data[year][month] = data.copy()
		return data

	def get_csv_filename_from_month(self, month: str, year: int) -> str:
		# e.g. 'Transactions Nov 1, 2024 - Nov 30, 2024 (7).csv'
		glob_pattern = RAW_TRANSACTION_FILENAME_FORMAT.format(month, year)
		files = glob(
			glob_pattern,
			root_dir = RAW_TRANSACTIONS_DIR
		)
		if len(files) < 1:
			raise FileNotFoundError(f'Could not find any matches for {glob_pattern}')
		elif len(files) == 1:
			file = files[0]
		else:
			biggest = -1, None
			for file in files:
				if '(' not in file:
					continue
				number = int(file[file.find('(') + 1 : file.find(')')])
				if number > biggest[0]:
					biggest = number, file
			file = biggest[1]
			assert file is not None
		return os.path.join(RAW_TRANSACTIONS_DIR, file)

	@staticmethod
	def parse_note(note: str, sep: str = '|') -> tuple[str, float]:
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

	def read_starting_balances(self, year: int):
		'''
		get starting balances from json file
		:year: the year of the starting balances to get

		example file:
		{
			"Bigger purchases": 0,
			"Emergency": 1000
		}
		'''
		filename = f'starting_balances{year}.json'
		filepath = os.path.join(BALANCES_DIR, filename)
		try:
			with open(filepath) as f:
				self.starting_balances = json.load(f)
		except FileNotFoundError:
			pass

	def read_budget(self, month: int, year: int, max_recursions: int = 100) -> pd.DataFrame:
		'''
		read the budget from the file
		if it cannot find one it will create a new one copying the last month's
		if it cannot find any within the lasst {max_recursion} months it will raise an FileNotFoundError

		example file:
		Category,Expected
		Rent & Utilities,2490.0
		Investing,500.0
		Fuel,150.0
		Groceries,500.0
		Eating Out,300.0
		Other,200.0
		'''
		filename = os.path.join(BUDGETS_DIR, f'{year}{month:02d}budget.csv')
		try:
			df = pd.read_csv(filename)
		except FileNotFoundError as e:
			if max_recursions < 1:
				raise e
			new_month = month - 1
			new_year = year
			if new_month < 1:
				new_month = 12
				new_year -= 1
			df = self.read_budget(new_month, new_year, max_recursions - 1)
			df.to_csv(filename, index = False)
		if year not in self.monthly_budget:
			self.monthly_budget[year] = {}
		self.monthly_budget[year][month] = df
		return df

	def load(self, starting_year: int):
		'''
		Read all transaction data into pd.DataFrames 1 month at a time
		Read all budget csv files into data 1 month at a time
		Read the 1 starting balances file for the given starting_year
		'''
		current_year = get_year()
		for year in range(starting_year, current_year + 1):
			for month in range(1, 13):
				if self.read_month(month, year) is not None:
					self.read_budget(month, year)
		self.read_starting_balances(starting_year)
