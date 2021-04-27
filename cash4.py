import pandas as pd
import numpy as np
import pickle
import datetime

class Data:

    def __init__(self, frame, escrow):
        self.frame = self.load_dataframe(frame)
        self.escrow = escrow
        self.accts = self.load_accts()
        self.branch = self.load_branches()
        self.date = self.get_posting_date()
        self.filename = self.get_filename()
        self.total = self.get_total()
        self.acct = self.get_acct()
        self.sheet = self.get_sheetname()

    def load_dataframe(self, frame):
        for i, cell in enumerate(frame['State']):
            if str(cell) == 'nan':
                frame.loc[i, 'State'] = 'NJ'
        return frame

    def load_accts(self):
        with open('accounts.pickle', 'rb') as f: return pickle.load(f)

    def load_branches(self):
        with open('branches.pickle','rb') as f: return pickle.load(f)

    def get_posting_date(self):
        for i, cell in enumerate(self.frame['PaymentDate']):
            if type(cell) == pd._libs.tslibs.timestamps.Timestamp:
                return datetime.date.strftime(cell, '%m/%d/%Y')

    def get_filename(self):
        return 'journals/{} cash_receipts.xlsx'.format(self.date.replace('/','_'))

    def get_total(self):
        return round(self.frame['Invoice Line Total'].sum(), 2)

    def get_acct(self):
        return self.accts[self.escrow]['bank']

    def get_sheetname(self):
        return self.accts[self.escrow]['sheet']

class Cash(Data):
    
    def __init__(self, frame, escrow):
        super().__init__(frame, escrow)
        self.df = self.split_shortages()
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

    def split_shortages(self):
        shorts = ['66300','66302']
        if '66300' or '66302' in self.frame['AcctCode']:
            report = self.frame[~self.frame['AcctCode'].isin(shorts)]
            shortages = self.frame[self.frame['AcctCode'].isin(shorts)]
            report = report.groupby(['TitleCoNum','State','AcctCode']).agg({'Invoice Line Total':'sum','File Number':'first',
                                                                            'PaymentDate':'first','CloseAgent':'first'}).reset_index()
            return pd.concat([report, shortages], ignore_index=True)
        return self.frame.groupby(['TitleCoNum','State','AcctCode']).agg({'Invoice Line Total':'sum','File Number':'first',
                                                                       'PaymentDate':'first','CloseAgent':'first'}).reset_index()

    def states(self):
        for i, cell in enumerate(self.df['File Number']):
            if not cell.startswith('CS'):
                if cell.split('-')[-1] == 'R':
                    yield '01'
                else:
                    yield cell.split('-')[-1]
            else:
                yield '01'
        yield '00'

    def branches(self):
        for i, cell in enumerate(self.df['File Number']):
            if cell[-5:] in self.branch.keys():
                yield self.branch[cell[-5:]]
            else:
                yield '000'
        yield '000'

    def accounts(self):
        for i, cell in enumerate(self.df['AcctCode']):
            if cell == '43502' and self.sheet == 'G21':
                yield '43501'
            else:
                yield cell
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

    def fix_report(self):
        for i in range(len(self.report)):
            if self.report['Credits'][i] == 0:
                self.report.drop([i], inplace=True)

    def final_report(self):
        shortages = self.report[self.report.Account.str.startswith('663')]
        report = self.report[~self.report.Account.str.startswith('663')]
        report = report[report['Type'] == 'G/L Account']
        total = self.report[self.report['Type'] == 'Bank Account']

        report = report.groupby(['Account','Branch','St']).agg({'Date':'first','Type':'first','Dept':'first',
                                                                         'Account Desr':'first','Description Reference':'first',
                                                                         'Debits':'sum','Credits':'sum'},as_index=False).reset_index()
        report = pd.DataFrame({'Date':report.Date.tolist(),
                       'Type':report.Type.tolist(),
                       'Account':report.Account.tolist(),
                       'St':report.St.tolist(),
                       'Branch':report['Branch'].tolist(),
                       'Dept':report.Dept.tolist(),
                       'Account Desr':report['Account Desr'].tolist(),
                       'Description Reference':report['Description Reference'].tolist(),
                       'Debits': report.Debits.tolist(),
                       'Credits': report.Credits.tolist()})

        for i, cell in enumerate(report['Debits']):
            if cell == 0:
                report.loc[i, 'Debits'] = np.nan
            elif report['Credits'][i] == 0:
                report.loc[i, 'Credits'] = np.nan
        
        return pd.concat([report, shortages, total], ignore_index=True)

class Count(Data):

    def __init__(self, frame, escrow):
        super().__init__(frame, escrow)
        self.df = self.group_frame()
        self.totals = list(self.totals())
        self.counts = list(self.counts())
        self.debits = list(self.debits())
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
        self.final = self.final_report()

    def group_frame(self):
        return self.frame.groupby(['TitleCoNum','State','OrderCategory']).agg({'Invoice Line Total':'sum','File Number':'first'}).reset_index()

    def totals(self):
        for i, cell in enumerate(self.df['Invoice Line Total']):
            yield round(cell, 2)

    def co_branches(self):
        for i, cell in enumerate(self.df['File Number']):
            if cell[-5:] in self.branch.keys():
                yield self.branch[cell[-5:]]
                yield self.branch[cell[-5:]]
            else:
                yield '000'
                yield '000'
        yield '000'

    def states(self):
        for i, cell in enumerate(self.df['File Number']):
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
        L = list(self.frame.groupby(['TitleCoNum','State','OrderCategory']))
        for i in range(len(L)):
            tco,state,oc = L[i][0]
            df2 = L[i][1]
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
        for i in range(len(self.df)):
            oc = self.df['OrderCategory'][i]
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

    def final_report(self):
        total = self.report[self.report.Account == '99998']
        report = self.report[self.report.Account != '99998']
        report = report.groupby(['Account','Branch','St']).agg({'Date':'first','Type':'first',
                                                                         'Dept':'first','Account Desr':'first',
                                                                         'Description Reference':'first','Debits':'sum',
                                                                         'Credits':'sum'},as_index=False).reset_index()
        report = report.sort_values('Account', ascending=False)


        report = pd.DataFrame({'Date':report.Date.tolist(),
                       'Type':report.Type.tolist(),
                       'Account':report.Account.tolist(),
                       'St':report.St.tolist(),
                       'Branch':report['Branch'].tolist(),
                       'Dept':report.Dept.tolist(),
                       'Account Desr':report['Account Desr'].tolist(),
                       'Description Reference':report['Description Reference'].tolist(),
                       'Debits': report.Debits.tolist(),
                       'Credits': report.Credits.tolist()})

        report = pd.concat([report, total], ignore_index=True)

        for i, cell in enumerate(report['Debits']):
            if cell == 0:
                report.loc[i, 'Debits'] = np.nan
            elif report['Credits'][i] == 0:
                report.loc[i, 'Credits'] = np.nan
        
        return report

def update_account_database():
    with open('accounts.pickle', 'rb') as f:
        D = pickle.load(f)
    escrow = int(input('Escrow Bank: '))
    acct = str(input('Account: '))
    sheet = str(input('Sheetname: '))
    D[escrow] = {'bank':acct, 'sheet':sheet}
    with open('accounts.pickle', 'wb') as f:
        pickle.dump(D, f, pickle.HIGHEST_PROTOCOL)

def create_worksheet(filename):
    sheet = pd.read_excel(filename, converters={'AcctCode':str,'Invoice Line Total':float})
    filename = ''
    arr = []

    for escrow, df in sheet.groupby('EscrowBank'):
        data = Data(df, escrow)
        cash = Cash(df, escrow)
        cash.fix_report()
        count = Count(df, escrow)
        count.fix_report()
        df = pd.concat([cash.report, count.report], ignore_index=True)
        filename = data.filename 
        s = df[df.Type == 'Bank Account'].index.values[0] + 3
        e = df[df.Account == '99998'].index.values[0] + 1
        f = '=SUM(I{}:I{})'.format(s,e)
        arr.append((df, data.sheet, f, e))
        
    arr = sorted(arr, key=lambda x: x[1])

    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        print('Creating', filename.split('/')[-1])
        for i in range(len(arr)):
            print('Adding Sheet', arr[i][1])
            arr[i][0].to_excel(writer, sheet_name=arr[i][1], index=False)
            workbook = writer.book
            worksheet = writer.sheets[arr[i][1]]
            format1 = workbook.add_format({'num_format':'##0.00'})
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
