# Transaction Tracker

This creates an excel file using the transactions from ["Spending Tracker"](https://apps.apple.com/us/app/spending-tracker/id548615579)

## Input

### .env

The `.env` file contains settings such as various directory paths, style counts, starting year, etc.

### Transactions

Looks for raw CSV data in the `$RAW_TRANSACTION_DIR` directory.
The file name should match the `$RAW_TRANSACTION_FILENAME_FORMAT`.

expects a csv file with 4 columns representing date, category, ammount, and note

```
Date,Category,Amount,Note,Account
01/01/2024,Eating Out,123.45,Wood Ranch,Default
```
optionally if note contains a `|` then it will be split and the right side will be read
as cashback percentage.

the account `$DEFAULT_ACCOUNT` is a monthly budget and other accounts are for earmarked categories such as "Emergency".

### Budgets

Looks in the `$BUDGETS_DIR` folder.
The file name should match `YYYYMMbudget.csv`.

This program expects a 2 column format with Category and Expected spend.

example:
```
Category,Expected
Rent & Utilities,2000.0
Fuel,150.0
Groceries,500.0
Eating Out,300.0
Other,200.0
```

### Balances

Looks in the `$BALANCES_DIR` folder.
The file name should match `starting_balancesYYYY.json`.
A new one of these does not need to be created for every year.
Only when you start using this application you have already existing balances that need to be accounted for does this need to be created.

example:
```
{
    "Emergency": 1000,
    "Wedding": 2000,
    "Savings": 3000
}
```

## output

writes files in the `$TRANSACTIONS_REPORTS_DIR` folder.

The output created is of the name `transactions YYYYmmdd HHMMSS.xlsx`.

An excel file is created.

## Running

- Using [virtualenv](https://pypi.org/project/virtualenv/) run `virtualenv venv` to create a virtual environment.
- Then use `source /venv/bin/activate` to activate it.
- then use `pip3 install -r requirements.txt` to install the requirements.
- This is only required once.
- Modify the .env to customize paths and names.
- use `./main.py` or `./run` in order to run the program.
