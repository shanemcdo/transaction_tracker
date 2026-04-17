from dotenv import load_dotenv
import datetime
import os
import pandas as pd

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
TRANSACTION_REPORTS_DIR = getenv('TRANSACTION_REPORTS_DIR')
FINANCE_PATH = getenv('FINANCE_PATH')
BUDGET_BALANCES_SHEET = getenv('BUDGET_BALANCES_SHEET')
STARTING_STYLE_COUNT = int(getenv('STARTING_STYLE_COUNT'))
ENDING_STYLE_COUNT = int(getenv('ENDING_STYLE_COUNT'))
DEFAULT_ACCOUNT = getenv('DEFAULT_ACCOUNT')
EMPTY = pd.DataFrame({
	'Date': [],
	'Category': [],
	'Amount': [],
	'Note': [],
	'CashBack %': [],
	'CashBack Reward': [],
	'Account': [],
})

def get_year() -> int:
	return datetime.datetime.now().year

