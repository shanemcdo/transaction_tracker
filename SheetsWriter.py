from utils import *
from DataLoader import DataLoader
import gspread

SHEET_URL = getenv('SHEET_URL')

class SheetsWriter:

	def __init__(self, loader: DataLoader):
		self.data_loader = loader
		gc = gspread.oauth() # pyright: ignore
		self.sheets = gc.open_by_url(SHEET_URL)

	def write_raw_transactions(self):
		data = self.data_loader.get_all_data()
		data.Date = data.Date.map(lambda x: x.strftime('%Y/%m/%d'))
		datasheet = self.sheets.worksheet('Raw Transactions')
		datasheet.clear()
		datasheet.update(data.values.tolist(), value_input_option = 'USER_ENTERED') # pyright: ignore

	def write_budgets(self):
		budgetsheet = self.sheets.worksheet('Budgets')
		budgetsheet.clear()
		budgets = pd.DataFrame(
			data = [
				(year, month, df.loc[df.Category == 'Rent & Utilities', 'Expected'].sum(), df.loc[df.Category == 'Fuel', 'Expected'].sum(), df.loc[df.Category == 'Groceries', 'Expected'].sum(), df.loc[df.Category == 'Eating Out', 'Expected'].sum(), df.loc[df.Category == 'Other', 'Expected'].sum())
				for year, months in self.data_loader.monthly_budget.items()
				for month, df in months.items()
				if month < 13
			],
			columns = ['year', 'month', 'Rent & Utilities', 'Fuel', 'Groceries', 'Eating Out', 'Other'],
		)
		budgetsheet.update([budgets.columns.values.tolist()] + budgets.values.tolist(), value_input_option = 'USER_ENTERED') # pyright: ignore

	def write_date_last_updated(self):
		datesheet = self.sheets.worksheet('Date Last Updated')
		datesheet.update([[datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')]], value_input_option = 'USER_ENTERED') # pyright: ignore

	def write_google_sheets(self):
		'''
		Update google sheets data sheet with fresh data
		'''
		self.write_raw_transactions()
		self.write_budgets()
		self.write_date_last_updated()
