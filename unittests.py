#!/usr/bin/env python3

import unittest
from CSVLoader import CSVLoader
from SQLiteLoader import SQLiteLoader

class TransactionTrackerTests(unittest.TestCase):

	def test_loads(self):
		csv_loader = CSVLoader()
		csv_loader.load(2024)
		csv_df = csv_loader.get_all_data()
		sqlite_loader = SQLiteLoader()
		sqlite_loader.load(2024)
		sqlite_df = sqlite_loader.get_all_data()
		print(csv_df)
		print(sqlite_df)
		self.assertEqual(
			csv_df,
			sqlite_df,
			'Two load methods got different results'
		)

if __name__ == '__main__':
	unittest.main()
