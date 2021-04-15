import pandas as pd
import numpy as np
import datetime
import pickle
import openpyxl

class Cash:

    def __init__(self, df):
        self.df = df
        self.accts = self.load_accts()
        self.branch = self.load_branches()
        self.date = self.get_date()
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
    
    def load_branches(self):
        with open('branches.pickle', 'rb') as f:
            return pickle.load(f)

    def load_accts(self):
        with open('accounts.pickle', 'rb') as f:
            return pickle.load(f)

    def get_sheetname(self):
        for i, cell in enumerate(self.df['EscrowBank']):
            if cell in self.accts.keys():
                return self.accts[cell]['sheet']

    def get_acct(self):
        for i, cell in enumerate(self.df['EscrowBank']):
            if cell in self.accts.keys():
                return self.accts[cell]['bank']

    def show_dataframe(self):
        print(self.df)

    def get_date(self):
        for i, cell in enumerate(self.df['PaymentDate']):
            if type(cell) == pd._libs.tslibs.timestamps.Timestamp:
                return datetime.date.strftime(cell, "%m/%d/%Y")

    def get_total(self):
        return round(self.df['Invoice Line Total'].sum(), 2)

    def states(self):
        for i, cell in enumerate(self.df['File Number']):
            if not cell.startswith('CS'):
                state = cell.split('-')[-1]
                if state == 'R':
                    yield '01'
                else:
                    yield state
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
        yield self.accts[self.df['EscrowBank'][i]]['bank']

    def debits(self):
        for i, cell in enumerate(self.df['Invoice Line Total']):
            if not cell >= 0:
                yield abs(round(cell, 2))
            else:
                yield np.nan
        yield round(self.total, 2)
    
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
        for _ in range(len(self.df)):
            yield 'G/L Account'
        yield 'G/L Account'

    def depts(self):
        for i, cell in enumerate(self.df['AcctCode']):
            if not str(cell).startswith('6'):
                yield '00'
            else:
                yield '02'
        yield '00'

    def acct_desrs(self):
        for _ in range(len(self.df)):
            yield np.nan
        yield np.nan
    
    def descr_refs(self):
        for _ in range(len(self.df)):
            yield '{} {}'.format(self.date, 'RQ DEP')
        yield '{} {}'.format(self.date, 'RQ DEP')

    def format(self):
        self.report.loc["Debits":"Credits"].style.format("{:.2%}")

class Count:

    def __init__(self, df):
        self.df = df
        self.totals = list(self.totals())
        self.accts = self.accounts()
        self.branches = self.branches()
        self.g = self.group()
        self.counts = list(self.counts())
        self.debits = list(self.debits())
        self.date = self.get_date()
        self.sheet = self.get_sheetname()
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
        self.ws = self.fix_report()


    def get_sheetname(self):
        for i, cell in enumerate(self.df['EscrowBank']):
            if cell in self.accts.keys():
                return self.accts[cell]['sheet']

    def total_report(self):
        total = 0
        for i, cell in enumerate(self.report['Account']):
            if cell.startswith('9'):
                total += self.report['Debits'][i]
        self.report.iloc[-1, -1] = round(total, 2)
        self.report = self.report


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
        L = self.df.groupby(['EscrowBank','TitleCoNum','State','OrderCategory']).agg({'Invoice Line Total':'sum','File Number':'first'}).reset_index()['Invoice Line Total'].tolist()
        for cost in L:
            yield round(cost, 2)

    def grouped_summary(self):
        df = self.df.groupby(['EscrowBank','TitleCoNum','OrderCategory']).agg({'Invoice Line Total':'sum','File Number':'first'}).reset_index()
        return df

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
                yield len(set(df2[df2.AcctCode == 40000]['File Number']))
            elif oc in [2, 5]:
                yield len(set(df2[df2.AcctCode == 40002]['File Number']))
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
        r = len(self.debits) - 1
        for i in range(r):
            yield np.nan
        yield self.deposit_total
        

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
        s,e,n = 0,0,0
        for i, cell in enumerate(self.report['Account']):
            if cell.startswith('???'):
                n = self.report['Debits'][i]
                self.report.drop([i], inplace=True)
            elif cell == '99998':
                e = i
        total = self.report['Credits'][e]
        self.report.loc[e, "Credits"] = round(total, 2) - n
        if round(self.report['Debits'].sum(), 2) != round(self.report['Credits'].sum(), 2):
            print(self.sheet, 'off balance by', round(self.report['Credits'].sum() - self.report['Debits'].sum(), 2))
            
        
    def format(self):
        self.report.loc[:, "Debits":"Credits"].style.format("{:.2%}")

def create_frames(filename):
    wb_date = ''
    ending = 'cash_receipts.xlsx'
    df = pd.read_excel(filename)
    for i, cell in enumerate(df['TitleCoNum']):
        if cell == 138:
            df.loc[i, 'TitleCoNum'] = 139

    L = list(df.groupby('EscrowBank'))

    arr = []
    for i in range(len(L)):
        frames = []
        df2 = L[i][1]
        count = Count(df2)
        count.fix_report()
        count.format()
        wb_date = count.date
        df2 = df2.groupby(['TitleCoNum','State','AcctCode']).agg({"Invoice Line Total":'sum','PaymentDate':'last','File Number':'last','EscrowBank':'last'}).reset_index()
        c = Cash(df2)
        c.format()
        sheetname = c.sheet
        frames.append(c.report)
        frames.append(count.report)
        frame = pd.concat(frames)
        arr.append((frame, sheetname))

    wb_date = wb_date.replace('/', '_')
    filename = '{} {}'.format(wb_date, ending)

    arr = sorted(arr, key=lambda x: x[1])

    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        print('Creating', filename)
        for i in range(len(arr)):
            print('Adding Sheet', arr[i][1])
            arr[i][0].to_excel(writer, sheet_name=arr[i][1], index=False)
            workbook = writer.book
            worksheet = writer.sheets[arr[i][1]]
            format1 = workbook.add_format({'num_format': '##0.00'})
            worksheet.set_column(0, 2, 11.29)
            worksheet.set_column(3, 3, 2.29)
            worksheet.set_column(4, 5, 6.29)
            worksheet.set_column(6, 6, 11.86)
            worksheet.set_column(7, 7, 20.29)
            worksheet.set_column('I:J', 8.29, cell_format=format1)
        print('Worksheet finished')