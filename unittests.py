#!/usr/bin/env python3

import unittest
from loaders import CSVLoader, SQLiteLoader
from utils import getenv
from pandas.testing import assert_frame_equal

class TransactionTrackerTests(unittest.TestCase):

	starting_year = int(getenv('STARTING_YEAR'))

	def test_compare_loads(self):
		csv_loader = CSVLoader()
		csv_loader.load(self.starting_year)
		csv_df = csv_loader.get_all_data()
		sqlite_loader = SQLiteLoader()
		sqlite_loader.load(self.starting_year)
		sqlite_df = sqlite_loader.get_all_data()
		assert_frame_equal(
			csv_df,
			sqlite_df
		)

		for year, value in csv_loader.monthly_budget.items():
			for month, df in value.items():
				assert_frame_equal(df, sqlite_loader.monthly_budget[year][month])
		for year, value in sqlite_loader.monthly_budget.items():
			for month, df in value.items():
				assert_frame_equal(df, csv_loader.monthly_budget[year][month])

		self.assertDictEqual(csv_loader.starting_balances, sqlite_loader.starting_balances)

	def test_csv_transaction_columns(self):
		loader = CSVLoader()
		loader.load(self.starting_year)
		df = loader.get_all_data()
		self.assertTrue((df.columns == [ 'Date', 'Category', 'Amount', 'Note', 'Account', 'CashBack %', 'CashBack Reward']).all(), 'Incorrect Data Headers')

	def test_csv_budget_columns(self):
		loader = CSVLoader()
		loader.load(self.starting_year)
		for year, val in loader.monthly_budget.items():
			for month, df in val.items():
				self.assertTrue(
					(df.columns == ['Category', 'Expected']).all(),
					f'budget for {year}/{month} has incorrect columns: "{df.columns}"'
				)

	def test_sqlite_transaction_columns(self):
		loader = SQLiteLoader()
		loader.load(self.starting_year)
		df = loader.get_all_data()
		self.assertTrue((df.columns == [ 'Date', 'Category', 'Amount', 'Note', 'Account', 'CashBack %', 'CashBack Reward']).all(), 'Incorrect Data Headers')

	def test_sqlite_budget_columns(self):
		loader = SQLiteLoader()
		loader.load(self.starting_year)
		for year, val in loader.monthly_budget.items():
			for month, df in val.items():
				self.assertTrue(
					(df.columns == ['Category', 'Expected']).all(),
					f'budget for {year}/{month} has incorrect columns: "{df.columns}"'
				)

if __name__ == '__main__':
	unittest.main()
