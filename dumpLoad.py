#!/usr/bin/env python3

from loaders import SQLiteLoader, CSVLoader
from sys import argv


def main() -> int:
	if len(argv) < 2:
		print('Not enough arguments: Must pass load type')
		return 1
	match argv[1]:
		case 'SQLite':
			loader = SQLiteLoader()
		case 'CSV':
			loader = CSVLoader()
		case load_type:
			print(f'invalid load type: "{load_type}" please pass "CSV" or "SQLite"')
			return 2
	loader.load(2024)
	data = loader.get_all_data()
	for row in data.values:
		print(*row, sep = ',')
	return 0

if __name__ == '__main__':
	exit(main())
