from Loader import Loader
import sqlite3
import datetime
import calendar
import pandas as pd
import glob
import os
from utils import *

#              magic chatgpt number
#              |           9 hours (for weird japan honeymoon dates)
#              v           v
EPOCH_OFFSET = 978307200 + 32400
AUTO_BACKUP_DB_DIR = getenv('AUTO_BACKUP_DB_DIR')

class SqliteLoader(Loader):
	'''
	Loads from sqlite db files from "Spending Tracker App"
	'''

	def get_db_filename(self) -> str:
		files = glob.glob(os.path.join(AUTO_BACKUP_DB_DIR, '*'))
		return max(files, key = os.path.getctime)

	def load(self, starting_year: int):
		filename = self.get_db_filename()
		con = sqlite3.connect(filename)
		cur = con.cursor()

		default_acnt_pk,  = cur.execute('SELECT z_pk FROM zprofile WHERE zisdefault = 1').fetchone()

		now = datetime.datetime.now()
		_month_end_weekday, month_end = calendar.monthrange(now.year, now.month)
		query = f'''SELECT
			datetime(tran.zdate + {EPOCH_OFFSET}, 'unixepoch') AS date_utc,
			cat.zname,
			tran.zamount,
			IFNULL(tran.znote, ''),
			acnt.zname,
			tran.ztype
			FROM ztransaction tran,
			zcategory cat,
			zprofile acnt
			WHERE date_utc < datetime('{now.year}-{now.month}-{month_end} 23:59:59')
			AND date_utc > datetime('{starting_year}-01-01 00:00:00')
			AND tran.ztype <> 'Carry'
			AND tran.zcategory = cat.z_pk
			AND IFNULL(tran.zprofile, {default_acnt_pk}) = acnt.z_pk
			ORDER BY date_utc
			'''
		for row in cur.execute(query):
			date, name, amnt, note, acnt, type = row
			date = datetime.datetime.fromisoformat(date).date()
			if type != 'Expense': # Income or Transfer
				amnt *= -1
			note, cb = parse_note(note)
			note = note.replace(',', '')
			cb_reward = amnt * cb
			df_row = pd.DataFrame(
				data = [[date, name, amnt, note, acnt, cb, cb_reward]],
				columns= ['Date', 'Category', 'Amount', 'Note', 'Account', 'Cashback %', 'Cashback Reward']
			)
			# TODO: do this in a more efficient way
			if date.year not in self.data:
				self.data[date.year] = {}
			if date.month not in self.data[date.year]:
				self.data[date.year][date.month] = df_row
			else:
				self.data[date.year][date.month] = pd.concat([self.data[date.year][date.month], df_row])
		for year, values in self.data.items():
			for month in values.keys():
				self.read_budget(month, year)
		self.read_starting_balances(starting_year)
