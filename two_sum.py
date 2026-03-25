#!/usr/bin/env python3

import numpy as np
import pandas as pd
from sys import stdin

def get_data() -> np.ndarray:
	print('Enter data and hit ctrl-d to stop:')
	csv = pd.read_csv(
		stdin,
		sep = '\t',
		names = ('price', 'label')
	)
	# csv.info()
	csv['price'] = csv['price'].str.replace(',', '')
	csv['price'] = csv['price'].str.replace('$', '')
	csv['price'] = csv['price'].str.replace(')', '')
	csv['price'] = csv['price'].str.replace('(', '-')
	csv['price'] = csv['price'].astype(float)
	csv = csv.sort_values('price')
	#csv = csv[csv['price'] > 0]
	arr = csv.to_numpy()
	return arr

def get_target() -> float:
	while True:
		try:
			val = input('Enter target value> ')
			return float(val)
		except ValueError:
			pass

if __name__ == '__main__':
	target = get_target()
	data = get_data()
	size = len(data)
	start = 0
	end = size - 1

	while start < end:
		check = round(data[start][0] + data[end][0], 2)
		if check < target:
			start += 1
		elif check > target:
			end -= 1
		else:
			print(data[start])
			print(data[end])
			break


