#!/usr/bin/env python3

import sqlite3
import datetime
import calendar
from utils import parse_note

#              magic chatgpt number
#              |           9 hours (for weird japan honeymoon dates)
#              v           v
EPOCH_OFFSET = 978307200 + 32400

if __name__ == '__main__':
	filename = '/Users/shane/temp/trans.db'
	starting_year = 2024
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
		# print(row)
		date, name, amnt, note, acnt, type = row
		date = datetime.datetime.fromisoformat(date).date()
		if type != 'Expense': # Income or Transfer
			amnt *= -1
		note, cb = parse_note(note)
		note = note.replace(',', '')
		cb_reward = amnt * cb
		print(date, name, amnt, note, acnt, cb, cb_reward, sep=',')
