class Data:

    def __init__(self):
        self.accts = self.load_database('/data/workspace_files/accounts.pickle')
        self.branches = self.load_database('/data/workspace_files/branches.pickle')
        self.accounts = {2: {'revenue':'96002','count':'90502'},
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


    def load_database(self, filename):
        with open(filename, 'rb') as f:
            return pickle.load(f)

class Cash(Data):

    def __init__(self, escrow, df):
        super().__init__()
        self.escrow = escrow
        self.date = self.get_date(df)
        self.df = self.group_dataframe(df)
        self.total = self.total_row(df)
        self.report = pd.concat([self.df,self.total], ignore_index=True)

    def group_dataframe(self, df):
        shorts = ['66300','66302']
        shortages = df[df.AcctCode.isin(shorts)]
        cash = df[~df.AcctCode.isin(shorts)]
        cash = cash.groupby(['TitleCoNum','Branch','state_code','AcctCode']).agg({'Invoice Line Total':'sum','PaymentDate':'first','dept':'first',
                                                                        'File Number':'first','CloseAgent':'first','File':'first'}).reset_index()
        cash = pd.concat([cash, shortages], ignore_index=True)

        cash['Type'] = ['G/L Account' for i in range(len(cash))]    
        cash['Account Desr'] = ['{} {}'.format(cash['File Number'][i],cash['CloseAgent'][i]) if cash.AcctCode[i] in shorts else np.nan for i in range(len(cash))]
        cash['Description Reference'] = ['{} RQ DEP'.format(cash['PaymentDate'][i]) for i in range(len(cash))]
        cash['Debits'] = [abs(round(x, 2)) if x < 0 else np.nan for x in cash['Invoice Line Total']]
        cash['Credits'] = [round(x, 2) if x >= 0 else np.nan for x in cash['Invoice Line Total']]

        cash = cash[['PaymentDate','Type','AcctCode','state_code','Branch','dept','Account Desr','Description Reference','Debits','Credits']]
        cash.rename(columns = {'PaymentDate':'Date','AcctCode':'Account','state_code':'St'}, inplace=True)

        return cash
    
    def get_date(self, df):
        return df['PaymentDate'].tolist()[0]

    def total_row(self, df):
        total = round(df['Invoice Line Total'].sum(), 2)
        description = '{} RQ DEP'.format(self.date)
        D = pd.DataFrame({
            'Date': [self.date],
            'Type': ['Bank Account'],
            'Account': [self.accts[self.escrow]['bank']],
            'St': ['00'],
            'Branch': ['00'],
            'Dept': ['00'],
            'Account Desr': [np.nan],
            'Description Reference': [description],
            'Debits': [total],
            'Credits': [np.nan]
        })
        return D
        

class Count(Data):

    def __init__(self, escrow, df):
        super().__init__()
        self.frame = df
        self.escrow = escrow
        self.date = self.frame['PaymentDate'].tolist()[0]
        self.df = self.group_frame(df)
        self.totals = list(self.totals())
        self.counts = list(self.counts())
        self.debits = list(self.debits())
        self.report = self.create_report()

    def group_frame(self, df):
        return df.groupby(['TitleCoNum','state_code','OrderCategory']).agg({'Invoice Line Total':'sum', 'File':'first'}).reset_index()

    def totals(self):
        for i, cell in enumerate(self.df['Invoice Line Total']):
            yield round(cell, 2)

    def counts(self):
        for (tco, st, oc), df in self.frame.groupby(['TitleCoNum','state_code','OrderCategory']):
            if oc in [1, 4]:
                yield len(set(df[df.AcctCode == '40000']['File Number']))
            elif oc in [2, 5]:
                yield len(set(df[df.AcctCode == '40002']['File Number']))
            else:
                yield len(set(df['File Number']))

    def debits(self):
        for i in range(len(self.totals)):
            yield self.totals[i]
            yield self.counts[i]
        yield np.nan

    def create_report(self):
        df = pd.DataFrame({
            'Date': [self.date for i in range(len(self.debits))],
            'Type': ...,
            'St': ...,
            'Branch': ...,
            'Dept': ...,
            'Account Desr': ...,
            'Description Reference': ...,
            'Debits': ...,
            'Credits': ...
        })
        return df

def load_worksheet(filename):
    df = pd.read_excel(filename, converters={'AcctCode':str,'Invoice Line Total':np.float64})
    df.State.replace(np.nan, 'NJ', inplace=True)

    df['File'] = df['File Number'].str[-5:]
    for i in range(len(df['File'])):
        if not df['File'][i][-1].isnumeric():
            df.loc[i, 'File'] = df['File'][i].split('-')[0][1:] + '-01'
        elif df['File'][i].startswith('CS'):
            df.loc[i, 'File'] = 'ST-01'
    
    def load_database(path):
        with open(path, 'rb') as f:
            return pickle.load(f)

    branches = load_database('/data/workspace_files/branches.pickle')

    df['Branch'] = df['File'].str[-5:].map(lambda x: branches[x] if x in branches.keys() else '000')
    df['state_code'] = df['File'].str[-2:]
    df['dept'] = ['02' if x.startswith('6') else '00' for x in df.AcctCode]
    df['PaymentDate'] = pd.to_datetime(df['PaymentDate']).dt.strftime("%m/%d/%Y")

    return df
        

df = load_worksheet('/data/workspace_files/cash.xls')

for escrow, frame in df.groupby('EscrowBank'):
    count = Count(escrow, frame)
    print(count.report)
    
