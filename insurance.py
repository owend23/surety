import pandas as pd
import datetime as dt

def parse_general_ledger(filename):
	df = pd.read_excel(filename, sheet_name='General Ledger Entries', 
			   usecols=['Posting Date','G/L Account No.','Description','Debit Amount','Credit Amount',
				   'State','Dept Code','BRANCH ID CODE','BRANCH ID VALUE'])
	
	branches = set(df['BRANCH ID VALUE'].tolist())
	accounts = set(df['G/L Account No.'].tolist())
	states = set(df['State'].tolist())
	full_accounts = []
	
	for account in accounts:
		acct = account
		for branch in branches:
			acct += '-' + branch
		for state in states:
			acct += '-' + state
		full_accounts.append(acct)
	
	for i, cell in enumerate(df['Posting Date']):
		date = str(df['Posting Date'][i].split('/'))
		if not int(date[1]) == 1:
			date[1] = '01'
		df['Posting Date'][i] = '/'.join(date)
