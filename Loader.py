from utils import *
import json
import pandas as pd
from functools import reduce

from abc import ABC, abstractmethod

# unix glob format
BUDGETS_DIR = getenv('BUDGETS_DIR')
BALANCES_DIR = getenv('BALANCES_DIR')

class Loader(ABC):
	'''
	Abstract and must be inherited
	Loads budgets and Balances CSVs
	Acts as a parent class for CsvLoader and SqliteLoader
	'''

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

	@abstractmethod
	def load(self, starting_year: int):
		raise NotImplemented
