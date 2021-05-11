import pandas as pd
import numpy as np
import pickle
import calendar
import os

class TD:
    def __init__(self, path):
        self.path = path
        self._accts = self.load_accts()
        self._cos = self.load_cos()
        self._employee = self.load_employee()
        self._keywords = self.load_descriptions()
        self.df = self.load_df()
        self.date = self.get_posting_date()
        self.report = pd.DataFrame({
                                    'Type': list(self.types()),
                                    'No': list(self.accounts()),
                                    'State': list(self.states()),
                                    'Branch Code': list(self.branches()),
                                    'Dept Code': list(self.depts()),
                                    'Description/Comment': list(self.descriptions()),
                                    'Quantity': list(self.quantity()),
                                    'Direct Unit Cost': list(self.costs()),
                                    'IC Partner Ref Type': list(self.ref_type()),
                                    'IC Partner Code': list(self.ic_partner_codes()),
                                    'IC Partner Reference': list(self.ic_partner_refs()),
        })

    def load_descriptions(self):
        with open('/data/workspace_files/databases/keywords.pickle', 'rb') as f:
            return pickle.load(f)

    def load_employee(self):
        with open('/data/workspace_files/databases/employee.pickle', 'rb') as f:
            return pickle.load(f)

    def load_accts(self):
        with open('/data/workspace_files/databases/accts.pickle', 'rb') as f: 
            return pickle.load(f)
    
    def load_cos(self):
        with open('/data/workspace_files/databases/cos.pickle', 'rb') as f: 
            return pickle.load(f)

    def load_df(self):
        df = pd.read_excel(self.path, converters={'MCC/SIC Code':str, 'Originating Account Number':str})
        df = df[df['Merchant Name'] != 'AUTO PAYMENT DEDUCTION']
        return df
    
    def accounts(self):
        for i, cell in enumerate(self.df['MCC/SIC Code']):
            if cell in self._accts.keys(): 
                yield self._accts[cell]
            else: 
                yield np.nan

    def types(self):
        for _ in range(len(self.df)): 
            yield 'G/L Account'

    def states(self):
        for i, cell in enumerate(self.df['Originating Account Number']):
            k = cell[-4:]
            if k in self._employee.keys(): 
                yield self._employee[k]['state']
            else: 
                yield np.nan

    def ref_type(self):
        for i, cell in enumerate(self.df['Account Number']):
            yield 'G/L Account'

    def branches(self):
        for i, cell in enumerate(self.df['Originating Account Number']):
            k = cell[-4:]
            if k in self._employee.keys(): 
                yield self._employee[k]['branch']
            else: 
                yield np.nan

    def descriptions(self):
        L = [cell for i, cell in enumerate(self.df['Merchant Name'])]
        for i, cell in enumerate(self.df['Originating Account Name']):
            if str(cell) == 'nan': 
                continue
            else:
                if len(cell.split()) == 2 and cell != 'COMMERCIAL DEPARTMENT':
                    name = cell.split()
                    initials = name[0][0] + name[1][0]
                    L[i] = '{}-{}'.format(initials, L[i])
                else: 
                    continue
        for descr in L: 
            yield descr

    def quantity(self):
        for _ in range(len(self.df)): 
            yield 1

    def costs(self):
        for i, cell in enumerate(self.df['Original Amount']): 
            yield round(cell, 2)

    def ic_partner_codes(self):
        for i, cell in enumerate(self.df['Originating Account Number']):
            k = cell[-4:]
            if k in self._employee.keys():
                yield self._employee[k]['ic_code']
            else:
                yield np.nan

    def ic_partner_refs(self):
        for _ in range(len(self.df)):
            yield np.nan

    def depts(self):
        # create conversion table based on accounting code / employee
        for i, cell in enumerate(self.df['Originating Account Number']): 
            k = cell[-4:]
            if k in self._employee.keys():
                yield self._employee[k]['dept']
            else:
                yield np.nan

    def get_posting_date(self):
        date = self.df.iloc[0, 0]
        return '{}/{}/{}'.format(date.month, calendar.monthrange(date.year, date.month)[1], date.year)

    def fix_sheet(self):
        import string
        schars = list(string.punctuation)

        for i, cell in enumerate(self.report['Description/Comment']):
            for char in schars:
                if char in cell:
                    cell = cell.replace(char, ' ')
            for word in cell.split():
                for k, v in self._keywords.items():
                    if word in v:
                        self.report.loc[i, 'No'] = k
            if self.report['Direct Unit Cost'][i] < 0:
                self.report.loc[i, 'No'] = '19999'
            if 'CONDOCERTS' in cell.split() and self.report['Direct Unit Cost'] > 300:
                self.report.loc[i, 'No'] = '63400'


def generate_td(dir):
    frames = []
    filename = ''
    for file in os.listdir(dir):
        sheet = file.split('TD CARD')[0].strip()
        sheet = sheet.split()[0]
        if not dir.endswith('/'): 
            dir += '/'
        td = TD(dir + file)
        td.fix_sheet()
        filename = '/data/workspace_files/journals/td/' + TD(dir + file).date.replace('/','_') + ' TD_Statements.xlsx'
        frames.append((td.report, sheet))

    frames = sorted(frames, key=lambda x: x[1])
    
    with pd.ExcelWriter(filename) as writer:
        for i in range(len(frames)): 
            frames[i][0].to_excel(writer, sheet_name=frames[i][1], index=False)

if __name__ == '__main__':
    generate_td('td statements')
