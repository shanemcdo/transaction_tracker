#!/usr/bin/env python3

import unittest
from CSVLoader import CSVLoader
from SQLiteLoader import SQLiteLoader
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

if __name__ == '__main__':
	unittest.main()
