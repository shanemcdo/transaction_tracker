from Loader import Loader
from utils import *
from glob import glob
import datetime
import pandas as pd
from utils import parse_note

# unix glob format
RAW_TRANSACTION_FILENAME_FORMAT = getenv('RAW_TRANSACTION_FILENAME_FORMAT')
RAW_TRANSACTIONS_DIR = getenv('RAW_TRANSACTIONS_DIR')

def parse_date(date: str) -> datetime.date:
	return datetime.datetime.strptime(date, '%m/%d/%Y').date()

class CsvLoader(Loader):
	'''
	Loads Export CSVs from "Spending Tracker App"
	'''

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
		tuple_col = data.Note.apply(parse_note)
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
