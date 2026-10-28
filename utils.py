from dotenv import load_dotenv
import datetime
import os

MONTHS = {
	1: 'January',
	2: 'February',
	3: 'March',
	4: 'April',
	5: 'May',
	6: 'June',
	7: 'July',
	8: 'August',
	9: 'September',
	10: 'October',
	11: 'November',
	12: 'December',
	13: 'Whole Year'
}
getenv = lambda x: os.getenv(x) or ''
MONTHS_SHORT = { key: value[:3] for key, value in MONTHS.items() }
load_dotenv()
# names of accounts in balances that are stored in savings accounts
SAVINGS_ACCOUNTS = getenv('SAVINGS_ACCOUNTS').split(',')
INCOME_CATEGORIES = getenv('INCOME_CATEGORIES').split(',')
DEFAULT_ACCOUNT = getenv('DEFAULT_ACCOUNT')

def get_year() -> int:
	return datetime.datetime.now().year

def parse_note(note: str, sep: str = '|') -> tuple[str, float]:
	'''
	Split note on seperator string and return parsed note and cashback
	:note: original note containing note message and cashback %
	:sep: the string to split the note on
	:return: note message and cashback percent in a tuple
	'''
	note = str(note)
	if note == 'nan':
		note = ''
	if sep in note:
		try:
			note, cashback = map(lambda x: x.strip('%\n\r\t '), note.split(sep, 1))
			return note, float(cashback) / 100
		except ValueError:
			pass
	return note, 0.0
