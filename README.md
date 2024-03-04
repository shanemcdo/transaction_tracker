# Transaction Tracker

This creates an excel file using the transactions from "Spending Tracker"

## Input

Looks for files in the `./in` folder. It is recommended to use symbolic links to 
point to the input file folder.

### File Name

the input expects `Transactions Mmm 1, YYYY - Mmm ??, YYYY *.csv` format.

e.g. `Transactions Jan 1, 2024 - Jan 31, 2024 (1).csv`

### File Format

expects a csv file with 4 columns representing date, category, ammount, and note

```
Date,Category,Amount,Note,Account
01/01/2024,Eating Out,123.45,Wood Ranch,Default
```
optionally if note contains a `|` then it will be split and the right side will be read
as cashback percentage.

the account `Default` is a monthly budget and other accounts are for earmarked categories such as "Emergency".

## output

writes files in the `./out` folder.

### File Name

The output created is of the name `transactions YYYYmmdd HHMMSS.xlsx`.

### File format

An excel file is created

## Running

- Using [virtualenv](https://pypi.org/project/virtualenv/) run `virtualenv venv` to create a virtual environment.
- Then use `source /venv/bin/activate` to activate it.
- then use `pip3 install -r requirements.txt` to install the requirements.
- This is only required once.
- use `ln -s /old/path in` to link raw transaction data directory.
- use `ln -s /old/path old` to link output directory.
- use `./main.py` or `./run` in order to run the program.
