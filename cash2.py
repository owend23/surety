import pandas as pd
import numpy as np
import pickle

class Data:
	
	def __init__(self):
		self.accts = self.loadDatabase('/data/workspace_files/databases/accounts.pickle')
		self.branches = self.loadDatabase('/data/workspace_files/databases/branches.pickle')
		self.accounts = self.loadDatabase('/data/workspace_files/databases/closing.pickle')
		
	def loadDatabase(self, filename):
		with open(filename, 'rb') as f:
			return pickle.load(f)
		
class Cash(Data):
	def __init__(self, escrow, df):
		super().__init__()
		self.escrow = escrow
		self.sheet = self.accts[escrow]['sheet']
		self.date = self.getDate(df)
		self.filename = '/data/workspace_files/journals/cash_receipts/{} cash_receipts.xlsx'.format(self.date.replace('/', '_'))
		self.df = self.groupDataframe(df)
		self.total = self.total_row(df)
		self.report = pd.concat([self.df, self.total], ignore_index=True)
		
	def groupDataframe(self, df):
		shorts = ['66300', '66302']
		shortages = df[df.AcctCode.isin(shorts)]
		cash = df[~df.AcctCode.isin(shorts)]
		cash = cash.groupby(['TitleCoNum','Branch','state_code','AcctCode']).agg({'Invoice Line Total':'sum', 'PaymentDate':'first',
																				  'File Number':'first', 'CloseAgent':'first', 'File':'first'}).reset_index()
		cash = pd.concat([cash, shortages], ignore_index=True)
		
		cash['Type'] = ['G/L Account' for _ in range(len(cash))]
		cash['Account Desr'] = ['{} {}'.format(cash['File Number'][i], cash['CloseAgent'][i]) if cash.AcctCode[i] in shorts else np.nan for i in range(len(cash))]
		cash['Description Reference'] = ['{} RQ DEP'.format(cash['PaymentDate'][i]) for i in range(len(cash))]
		cash['Debits'] = [abs(round(x,2)) if x < 0 else np.nan for x in cash['Invoice Line Total']]
		cash['Credits'] = [round(x,2) if x >= 0 else np.nan for x in cash['Invoice Line Total']]
		
		cash = cash[['PaymentDate','Type','AcctCode','state_code','Branch','dept','Account Desr','Description Reference','Debits','Credits']]
		cash.rename(columns={'PaymentDate':'Date', 'AcctCode':'Account', 'state_code':'St', 'dept':'Dept'}, inplace=True)
		cash.sort_values(by=['Branch', 'Account'], inplace=True)
		return cash
	
	def getDate(self, df):
		return df['PaymentDate'].tolist()[0]
	
	def total_row(self, df):
		total = round(df['Invoice Line Total'].sum(), 2)
		description = '{} RQ DEP'.format(self.date)
		return pd.DataFrame({
			'Date': [self.date],
			'Type': ['Bank Account'],
			'Account': [self.accts[self.escrow]['bank']],
			'St': ['00'],
			'Branch': ['000'],
			'Dept': ['00'],
			'Account Desr': [np.nan],
			'Description Reference': ['{} RQ DEP'.format(self.date)],
			'Debits': [round(df['Invoice Line Total'].sum(), 2)],
			'Credits': [np.nan],
		})
	
class Count(Data):
	
		def __init__(self, escrow, df):
			super().__init__()
			self.frame = df
			self.escrow = escrow
			self.date = self.frame['PaymentDate'].tolist()[0]
			self.df = self.groupFrame(df)
			self.totals = list(self.totals())
			self.counts = list(self.counts())
			self.debits = list(self.debits())
			self.report = self.create_report()
			
		def groupFrame(self, df):
			return df.groupby(['TitleCoNum', 'state_code', 'OrderCategory']).agg({'Invoice Line Total':'sum', 'File':'first', 'Branch':'first', 'dept':'first'}).reset_index()
		
		def totals(self):
			for i, cell in enumerate(self.df['Invoice Line Total']):
				yield round(cell, 2)
				
		def counts(self):
			for (tco, st, oc), df in self.frame.groupby(['TitleCoNum', 'state_code', 'OrderCategory']):
				if oc in [1, 4]:
					yield len(set(df[df.AcctCode == '40000']['File Number']))
				elif oc in [2, 5]:
					yield len(set(df[df.AcctCode == '40002']['File Number']))
				else:
					yield len(set(df['File Number']))
		
		def acct_codes(self):
			for i in range(len(self.df)):
				oc = self.df['OrderCategory'][i]
				yield self.accounts[oc]['revenue']
				yield self.accounts[oc]['count']
			yield '99998'
			
		def states(self):
			for i in range(len(self.df)):
				yield self.df['state_code'][i]
				yield self.df['state_code'][i]
			yield '00'
		
		def co_branches(self):
			for i in range(len(self.df)):
				if self.df['File'][i] in self.branches.keys():
					yield self.branches[self.df['File'][i]]
					yield self.branches[self.df['File'][i]]
				else:
					yield '000'
					yield '000'
			yield '000'
			
		def create_report(self):
			return pd.DataFrame({
				'Date': [self.date for _ in range(len(self.debits))],
				'Type': ['G/L Account' for _ in range(len(self.debits))],
				'Account': list(self.acct_codes()),
				'St': list(self.states()),
				'Branch': list(self.co_branches()),
				'Dept': ['00' for _ in range(len(self.debits))]
				'Account Desr': [np.nan for _ in range(len(self.debits))],
				'Description Reference': ['{} RQ DEP'.format(self.date) for _ in range(len(self.debits))],
				'Debits': self.debits,
				'Credits': [np.nan for _ in range(len(self.debits))]
			})
		
def load_worksheet(filename):
	# replace order categories if not in [1,2,4,5] but acctcodes [40000,40002] in file
	df = pd.read_excel(filename, converters={'AcctCode':str, 'Invoice Line Total':float})
	
	df.State.replace(np.nan, 'NJ', inplace=True)
	df.AcctCode.replace('40003', '40000', inplace=True)
	df.OrderCategory.replace(25, 8, inplace=True)
	
	df['File'] = df['File Number'].str[-5:]
	
	for i in range(len(df['File'])):
		if not df['File'][i][-1].isnumeric():
			df.loc[i, 'File'] = df['File'][i].split('-')[0][1:] + '-01'
		elif df['File'][i].startswith('CS'):
			df.loc[i, 'File'] = 'ST-01'
	
	with open('/data/workspace_files/databases/branches.pickle', 'rb') as f:
		branches = pickle.load(f)
		
	df['Branch'] = df['File'].str[-5:].map(lambda x: branches[x] if x in branches.keys() else '000')
	df['state_code'] = df['File'].str[-2:]
	df['dept'] = ['02' if x.startswith('6') else '00' for x in df.AcctCode]
	df['PaymentDate'] = pd.to_datetime(df['PaymentDate']).dt.strftime('%m/%d/%Y')
	return df

def create_worksheet(filename):
	df = load_worksheet(filename)
	filename = ''
	arr = []
	
	for escrow, frame in df.groupby('EscrowBank'):
		cash = Cash(escrow, frame)
		count = Count(escrow, frame)
		count.fix_report()
		report = pd.concat([cash.report, count.report], ignore_index=True)
		filename = cash.filename
		s = report[report.Type == 'Bank Account'].index.values[0] + 3
		e = report[report.Account == '99998'].index.values[0] + 1
		f = '=SUM(I{}:I{})'.format(s, e)
		arr.append((report, cash.sheet, f, e))
	
	arr = sorted(arr, key=lambda x: x[1])
	
	with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
		print('Creating', filename.split('/')[-1])
		for i in range(len(arr)):
			print('Adding sheet', arr[i][1])
			arr[i][0].to_excel(writer, sheet_name=arr[i][1], index=False)
			workbook = writer.book
			worksheet = writer.sheets[arr[i][1]]
			num_format = workbook.add_format({'num_format':'##0.00'})
			worksheet.set_column(0, 2, 12)
			worksheet.set_column(3, 3, 3)
			worksheet.set_column(4, 5, 7)
			worksheet.set_column(6, 6, 15)
			worksheet.set_column(7, 7, 21)
			worksheet.set_column('I:J', 9, cell_format=num_format)
			worksheet.write_formula('J{}'.format(arr[i][-1] + 1), arr[i][-2])
		print('Worksheet finished')

if __name__ == "__main__":
	create_worksheet('cash.xls')
