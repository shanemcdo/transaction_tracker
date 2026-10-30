#!/usr/bin/env python3

import unittest
from loaders import CSVLoader, SQLiteLoader
from pandas.testing import assert_frame_equal

class TransactionTrackerTests(unittest.TestCase):

	def test_loads(self):
		csv_loader = CSVLoader()
		csv_loader.load(2024)
		csv_df = csv_loader.get_all_data()
		sqlite_loader = SQLiteLoader()
		sqlite_loader.load(2024)
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

if __name__ == '__main__':
	unittest.main()
