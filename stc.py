import pandas as pd
import numpy as np
import pickle
import datetime
import os

class Cash:
    def __init__(self, path):
        self.path = path
        self.df = pd.read_excel(path)
        self.df2 = self.df2()
        self._branches = self.get_branches()
        self._companies = self.get_companies()
        self.co = self.get_co()
        self.co2 = self.get_co2()
        self.date = self.get_date()
        self.filename = self.construct_filename()
        self.report = pd.DataFrame({
                                    'Date': list(self.dates()),
                                    'Type': list(self.types()),
                                    'Account': list(self.accounts()),
                                    'St': list(self.states()),
                                    'Branch': list(self.branches()),
                                    'Dept': list(self.departments()),
                                    'Account Desr': list(self.acct_descriptions()),
                                    'Description Reference': list(self.description_references()),
                                    'Debits': list(self.debits()),
                                    'Credits': list(self.credits()),
        })

    def check_report(self):
        if '66302' in self.report['Account'].tolist():
            print('66302', self.path.split('/')[-1].split('.')[0])
            print()
        elif '66300' in self.report['Account'].tolist():
            print('66300', self.path.split('/')[-1].split('.')[0])
            print()
        deposit_total = self.report[self.report['Type'] == 'Bank Account']['Debits'].sum()
        acct_total = 0
        for i, cell in enumerate(self.report['Account']):
            if str(cell).startswith('96') or str(cell).startswith('?'): acct_total += self.report['Debits'][i]
        deposit_total = round(deposit_total, 2)
        acct_total = round(acct_total, 2)
        if deposit_total == acct_total: pass
        else:
            print(self.path.split('/')[-1].split('.')[0], 'is off balance')
            print('Deposit Total:', deposit_total)
            print('Account Total:', acct_total)
            print('Difference:', round(deposit_total - acct_total, 2))
            print('\n')

    def construct_filename(self):
        date = self.date.replace('/', '_')
        return f'{date} cash_receipts.xlsx'

    def get_date(self):
        for i, cell in enumerate(self.df['Description']):
            if str(cell) == 'Accounting Code Total': return datetime.date.strftime(self.df['PaymentDate'][i - 1], "%m/%d/%Y")
    
    def df2(self):
        df = pd.read_excel(self.path.split('.')[0] + '_closing.xls')
        df.columns = ['col' + str(i) for i in range(len(df.columns))]
        return df

    def get_co2(self):
        return self.df2.iloc[0, 0]

    def get_branches(self):
        with open('branches.pickle', 'rb') as f:
            return pickle.load(f)

    def get_companies(self):
        with open('companies.pickle', 'rb') as f:
            return pickle.load(f)

    def get_co(self):
        return self.df.iloc[0, 0]

    def accounts(self):
        for i, cell in enumerate(self.df['Description']):
            if not type(cell) == float:
                if len(cell) == 5 and str(cell).isnumeric(): yield cell
                elif str(cell) == 'Deposit Total':
                    if self.co in self._companies.keys(): yield self._companies[self.co]
                    else: yield np.nan
        for i, cell in enumerate(self.df2['col0']):
            if not type(cell) == float:
                if str(cell).startswith('9'):
                    if cell == '96021-' and self.co2 == 'ESL Title and Settlement Services ':
                        yield '96024'
                        yield str(self.df2['col5'][i]).strip('-')
                    else:
                        yield str(cell).strip('-')
                        yield str(self.df2['col5'][i]).strip('-')
        yield '99998'

    def credits(self):
        total = 0
        for i, cell in enumerate(self.df['Description']):
            if str(cell) == 'Accounting Code Total':
                if self.df['Invoice Line Total'][i] >= 0: yield round(self.df['Invoice Line Total'][i], 2)
                else: yield np.nan
            elif str(cell) == 'Deposit Total': yield np.nan
        for i, cell in enumerate(self.df2['col0']):
            if not type(cell) == float:
                if str(cell).startswith('9'):
                    total += round(self.df2['col4'][i], 2)
                    total += round(self.df2['col10'][i], 2)
                    yield np.nan
                    yield np.nan
        yield total

    def debits(self):
        for i, cell in enumerate(self.df['Description']):
            if str(cell) == 'Accounting Code Total':
                if self.df['Invoice Line Total'][i] < 0: yield abs(round(self.df['Invoice Line Total'][i], 2))
                else: yield np.nan
            elif str(cell) == 'Deposit Total': yield round(self.df['Invoice Line Total'][i], 2)
        for i, cell in enumerate(self.df2['col0']):
            if not type(cell) == float:
                if str(cell).startswith('9'):
                    yield round(self.df2['col4'][i], 2)
                    yield round(self.df2['col10'][i], 2)
        yield np.nan
    
    def dates(self):
        for i, cell in enumerate(self.df['Description']):
            if str(cell) == 'Accounting Code Total': yield self.date
            elif str(cell) == 'Deposit Total': yield self.date 
        for i, cell in enumerate(self.df2['col0']):
            if not type(cell) == float:
                if str(cell).startswith('9'):
                    yield self.date
                    yield self.date
        yield self.date

    def types(self):
        for i, cell in enumerate(self.df['Description']):
            if str(cell) == 'Accounting Code Total': yield 'G/L Account'
            elif str(cell) == 'Deposit Total': yield 'Bank Account'
        for i, cell in enumerate(self.df2['col0']):
            if not type(cell) == float:
                if str(cell).startswith('9'):
                    yield 'G/L Account'
                    yield 'G/L Account'
        yield 'G/L Account'

    def states(self):
        for i, cell in enumerate(self.df['Description']):
            if str(cell) == 'Accounting Code Total':
                if not self.df['File Number'][i - 1].startswith('CS'):
                    if self.df['File Number'][i - 1].split('-')[-1] == 'R': yield '01'
                    else: yield self.df['File Number'][i - 1].split('-')[-1]
                else: yield '01'
            elif str(cell) == 'Deposit Total': yield '00'
        for i, cell in enumerate(self.df2['col0']):
            if not type(cell) == float:
                if str(cell).startswith('9'):
                    if not str(self.df2['col1'][i]).strip('-') == '???????':
                        yield str(self.df2['col1'][i]).strip('-')
                        yield str(self.df2['col1'][i]).strip('-')
                    else:
                        yield '01'
                        yield '01'
        yield '00'

    def branches(self):
        for i, cell in enumerate(self.df['Description']):
            if str(cell) == 'Accounting Code Total':
                k = self.df['File Number'][i - 1][-5:]
                if self.df['File Number'][i - 1][-5:] in self._branches.keys(): yield self._branches[self.df['File Number'][i - 1][-5:]]
                else: yield '000'
            elif str(cell) == 'Deposit Total': yield '000'
        for i, cell in enumerate(self.df2['col0']):
            if not type(cell) == float:
                if str(cell).startswith('9'):
                    yield str(self.df2['col2'][i]).strip('-')
                    yield str(self.df2['col2'][i]).strip('-')
        yield '000'

    def departments(self):
        for i, cell in enumerate(self.df['Description']):
            if len(str(cell)) == 5 and str(cell).isnumeric(): acct = cell
            if str(cell) == 'Accounting Code Total':
                if str(acct).startswith('6'): yield '02'
                else: yield '00'
            elif str(cell) == 'Deposit Total': yield '00'
        for i, cell in enumerate(self.df2['col0']):
            if not type(cell) == float:
                if str(cell).startswith('9'):
                    yield '00'
                    yield '00'
        yield '00'

    def acct_descriptions(self):
        for i, cell in enumerate(self.df['Description']):
            if str(cell) == 'Accounting Code Total': yield np.nan
            elif str(cell) == 'Deposit Total': yield np.nan
        for i, cell in enumerate(self.df2['col0']):
            if not type(cell) == float:
                if str(cell).startswith('9'):
                    yield np.nan
                    yield np.nan
        yield np.nan
    
    def description_references(self):
        for i, cell in enumerate(self.df['Description']):
            if str(cell) == 'Accounting Code Total': yield self.date + ' RQ DEP'
            elif str(cell) == 'Deposit Total': yield self.date + ' RQ DEP'
        for i, cell in enumerate(self.df2['col0']):
            if not type(cell) == float:
                if str(cell).startswith('9'):
                    yield self.date + ' RQ DEP'
                    yield self.date + ' RQ DEP'
        yield self.date + ' RQ DEP'

def cash_receipts(dir):
    filename = ''
    frames = []
    for file in os.listdir(dir):
        if not file.startswith('.'):
            if not file.endswith('_closing.xls'):
                if not dir.endswith('/'):
                    Cash(dir + '/' + file).check_report()
                    frames.append((Cash(dir + '/' + file).report, file.split('.')[0]))
                    filename = Cash(dir + '/' + file).filename
                else:
                    Cash(dir + file).check_report()
                    frames.append((Cash(dir + file).report, file.split('.')[0]))
                    filename = Cash(dir + file).filename
        
    with pd.ExcelWriter(filename) as writer:
        for i in range(len(frames)):
            frames[i][0].to_excel(writer, sheet_name=frames[i][1], index=False)

    
