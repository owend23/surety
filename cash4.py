import pandas as pd
import numpy as np
import datetime
import pickle

class Cash:

    def __init__(self, escrow, df):
        self.escrow = escrow
        self.frame = df
        self.df = self.split_shortages()
        self.date = self.get_posting_date()
        self.accts = self.load_accts()
        self.branch = self.load_branches()
        self.total = self.get_total()
        self.acct = self.get_acct()
        self.sheet = self.get_sheetname()
        self.report = pd.DataFrame({
            'Date': list(self.dates()),
            'Type': list(self.types()),
            'Account': list(self.accounts()),
            'St': list(self.states()),
            'Branch': list(self.branches()),
            'Dept': list(self.depts()),
            'Account Desr': list(self.acct_desrs()),
            'Description Reference': list(self.descr_refs()),
            'Debits': list(self.debits()),
            'Credits': list(self.credits()),
            })

    def load_accts(self):
        with open('accounts.pickle', 'rb') as f: return pickle.load(f)

    def load_branches(self):
        with open('branches.pickle', 'rb') as f: return pickle.load(f)

    def split_shortages(self):
        shorts = ['66300','66302']
        if '66300' or '66302' in self.frame['AcctCode']:
            report = self.frame[~self.frame['AcctCode'].isin(shorts)]
            shortages = self.frame[self.frame['AcctCode'].isin(shorts)]
            report = report.groupby(['TitleCoNum','State','AcctCode']).agg({'Invoice Line Total':'sum','File Number':'first',
                                                                           'PaymentDate':'first','CloseAgent':'first'}).reset_index()
            shortages = shortages[['TitleCoNum','State','AcctCode','Invoice Line Total','File Number','PaymentDate','CloseAgent']]
            return pd.concat([report, shortages], ignore_index=True)
        else:
            return self.frame.groupby(['TitleCoNum','State','AcctCode']).agg({'Invoice Line Total':'sum','PaymentDate':'last',
                                                                              'File Number':'last','EscrowBank':'last','CloseAgent':'last'})

    def get_posting_date(self):
        for i, cell in enumerate(self.df['PaymentDate']):
            if type(cell) == pd._libs.tslibs.timestamps.Timestamp:
                return datetime.date.strftime(cell, '%m/%d/%Y')

    def get_total(self):
        return round(self.df['Invoice Line Total'].sum(), 2)

    def get_acct(self):
        return self.accts[self.escrow]['bank']

    def get_sheetname(self):
        return self.accts[self.escrow]['sheet']

    def states(self):
        for i, cell in enumerate(self.df['File Number']):
            if not cell.startswith('CS'):
                state = cell.split('-')[-1]
                if cell.split('-')[-1] == 'R':
                    yield '01'
                else:
                    yield cell.split('-')[-1]
            else:
                yield '01'
        yield '00'

    def branches(self):
        for i, cell in enumerate(self.df['File Number']):
            if cell[-5:] in self.branch.keys(): yield self.branch[cell[-5:]]
            else: yield '000'
        yield '000'

    def accounts(self):
        for i, cell in enumerate(self.df['AcctCode']): yield str(cell)
        yield self.acct

    def debits(self):
        for i, cell in enumerate(self.df['Invoice Line Total']):
            if not cell >= 0:
                yield abs(round(cell, 2))
            else:
                yield np.nan
        yield self.total

    def credits(self):
        for i, cell in enumerate(self.df['Invoice Line Total']):
            if not cell >= 0:
                yield np.nan
            else:
                yield round(cell, 2)
        yield np.nan

    def dates(self):
        for _ in range(len(self.df)):
            yield self.date
        yield self.date

    def types(self):
        for i in range(len(self.df)):
            yield 'G/L Account'
        yield 'Bank Account'

    def depts(self):
        for i, cell in enumerate(self.df['AcctCode']):
            if not str(cell).startswith('6'):
                yield '00'
            else:
                yield '02'
        yield '00'

    def acct_desrs(self):
        accounts = self.df['AcctCode'].tolist()
        files = self.df['File Number'].tolist()
        agents = self.df['CloseAgent'].tolist()
        for i in range(len(accounts)):
            if accounts[i] in ['66300','66302']:
                yield files[i] + ' ' + agents[i]
            else:
                yield np.nan
        yield np.nan

    def descr_refs(self):
        for _ in range(len(self.df)):
            yield '{} RQ DEP'.format(self.date)
        yield '{} RQ DEP'.format(self.date)

    def format(self):
        self.report.loc["Debits":"Credits"].style.format("{:.2%}")

class Count:

    def __init__(self, df, escrow):
        self.df = df
        self.escrow = escrow
        self.totals = list(self.totals())
        self.accts = self.accounts()
        self.branches = self.branches()
        self.g = self.group()
        self.counts = list(self.counts())
        self.debits = list(self.debits())
        self.date = self.get_date()
        self.report = pd.DataFrame({
                       'Date': list(self.dates()),
                       'Type': list(self.types()),
                       'Account': list(self.acct_codes()),
                       'St': list(self.states()),
                       'Branch': list(self.co_branches()),
                       'Dept': list(self.depts()),
                       'Account Desr': list(self.acct_desrs()),
                       'Description Reference': list(self.references()),
                       'Debits': self.debits,
                       'Credits': list(self.credits()),
        })

    def get_date(self):
        for i, cell in enumerate(self.df['PaymentDate']):
            if type(cell) == pd._libs.tslibs.timestamps.Timestamp:
                return datetime.date.strftime(cell, '%m/%d/%Y')

    def accounts(self):
        with open('accounts.pickle', 'rb') as f:
            return pickle.load(f)

    def branches(self):
        with open('branches.pickle', 'rb') as f:
            return pickle.load(f)

    def group(self):
        return list(self.df.groupby(['TitleCoNum','State','OrderCategory']))

    def totals(self):
        arr = self.df.groupby(['EscrowBank','TitleCoNum','State','OrderCategory']).agg({'Invoice Line Total':'sum','File Number':'first'}).reset_index()['Invoice Line Total'].tolist()
        for cost in arr:
            yield round(cost, 2)

    def co_branches(self):
        df = self.df.groupby(['EscrowBank','TitleCoNum','State','OrderCategory']).agg({'File Number':'first'}).reset_index()
        for i, cell in enumerate(df['File Number']):
            k = cell[-5:]
            if k in self.branches.keys():
                yield self.branches[k]
                yield self.branches[k]
            else:
                yield '000'
                yield '000'
        yield '000'

    def states(self):
        df = self.df.groupby(['EscrowBank','TitleCoNum','State','OrderCategory']).agg({'File Number':'first'}).reset_index()
        for i, cell in enumerate(df['File Number']):
            if cell.startswith('CS'):
                yield '01'
                yield '01'
            elif cell.endswith('R'):
                yield '01'
                yield '01'
            else:
                yield cell.split('-')[-1]
                yield cell.split('-')[-1]
        yield '00'
    
    def counts(self):
        for i in range(len(self.g)):
            tco,state,oc = self.g[i][0]
            df2 = self.g[i][1]
            if oc in [1, 4]:
                yield len(set(df2[df2.AcctCode == '40000']['File Number']))
            elif oc in [2, 5]:
                yield len(set(df2[df2.AcctCode == '40002']['File Number']))
            else:
                yield len(set(df2['File Number']))

    def debits(self):
        total = 0
        for i in range(len(self.totals)):
            yield self.totals[i]
            yield self.counts[i]
            total += self.totals[i]
            total += self.counts[i]
        self.deposit_total = round(total, 2)
        yield np.nan

    def credits(self):
        for i in range(len(self.debits)):
            yield np.nan

    def dates(self):
        for _ in range(len(self.debits)):
            yield self.date 

    def types(self):
        for _ in range(len(self.debits)):
            yield 'G/L Account'

    def acct_codes(self):
        D = {2: {'revenue':'96002','count':'90502'},
             1: {'revenue': '96000', 'count':'90500'},
             4: {'revenue': '96004', 'count': '90504'},
             5: {'revenue': '96005', 'count': '90505'},
             7: {'revenue': '96023', 'count': '90511'},
             11: {'revenue': '96021', 'count': '90605'},
             8: {'revenue': '96023', 'count': '?????'},
             10: {'revenue': '96021', 'count': '90601'},
             16: {'revenue': '96003', 'count': '90503'},
             9: {'revenue': '96020', 'count': '90600'},
             12: {'revenue': '90600', 'count': '90512'},
             29: {'revenue': '90600', 'count': '90512'},
             }
        df = self.df.groupby(['TitleCoNum','State','OrderCategory']).agg({'Invoice Line Total':'sum','File Number':'first'}).reset_index()
        for i in range(len(df)):
            oc = df['OrderCategory'][i]
            yield D[oc]['revenue']
            yield D[oc]['count']
        yield '99998'

    def depts(self):
        for _ in range(len(self.debits)):
            yield '00'

    def acct_desrs(self):
        for _ in range(len(self.debits)):
            yield np.nan

    def references(self):
        for _ in range(len(self.debits)):
            yield '{} RQ DEP'.format(self.date)

    def fix_report(self):
        for i, cell in enumerate(self.report['Account']):
            if cell.startswith('???'):
                self.report.drop([i], inplace=True)
            elif cell == '96021' and self.accts[self.escrow]['sheet'] == 'ESL NJ':
                self.report.loc[i, 'Account'] = '96024'
        for i, cell in enumerate(self.report['Debits']):
            if cell == 0:
                self.report.drop([i], inplace=True)
        
    def format(self):
        self.report.loc[:, "Debits":"Credits"].style.format("{:.2%}")

def create_worksheet(filename):
    df = pd.read_excel(filename, converters={'AcctCode':str, 'Invoice Line Total':float})

    ending = 'cash_receipts.xlsx'

    arr = []

    for escrow, df in df.groupby('EscrowBank'):
        frames = []
        cash = Cash(escrow, df)
        cash.split_shortages()
        cash.format()
        count = Count(df, escrow)
        count.format()
        count.fix_report()
        wb_date = cash.date
        frames.append(cash.report)
        frames.append(count.report)
        frame = pd.concat(frames, ignore_index=True)
        s = frame[frame['Type'] == 'Bank Account'].index.values[0] + 3
        e = frame[frame['Account'] == '99998'].index.values[0] + 1
        f = '=SUM(I{}:I{})'.format(s,e)
        arr.append((frame, cash.sheet, f, e))

    wb_date = wb_date.replace('/', '_')
    filename = 'journals/{} {}'.format(wb_date, ending)

    arr = sorted(arr, key=lambda x: x[1])

    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        print('Creating', filename.split('/')[-1])
        for i in range(len(arr)):
            print('Adding Sheet', arr[i][1])
            arr[i][0].to_excel(writer, sheet_name=arr[i][1], index=False)
            workbook = writer.book
            worksheet = writer.sheets[arr[i][1]]
            format1 = workbook.add_format({'num_format': '##0.00'})
            worksheet.set_column(0, 2, 12)
            worksheet.set_column(3, 3, 3)
            worksheet.set_column(4, 5, 7)
            worksheet.set_column(6, 6, 15)
            worksheet.set_column(7, 7, 21)
            worksheet.set_column('I:J', 9, cell_format=format1)
            worksheet.write_formula('J{}'.format(arr[i][-1] + 1), arr[i][-2])
        print('Worksheet finished')

if __name__ == '__main__':
    create_worksheet('cash.xls')
