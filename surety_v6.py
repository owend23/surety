import pandas as pd

def parse_worksheet(df, date):
    accounts, states, branches, depts, debits, credits = [], [], [], [], [], []
    dates, types, acct_descr, descr_ref = [], [], [], []

    acct = ''
    dr = date + ' RQ DEP'
    at = 'G/L Account'

    for i, cell in enumerate(df['Description'][:-3]):
        if not type(cell) == float:
            if str(cell).isnumeric():
                acct = cell
        else:
            if str(cell) == 'nan':
                try:
                    file = df['File Number'][i - 1]
                    state = df['File Number'][i - 1].split('-')[-1]
                    if df['Invoice Line Total'][i] >= 0:
                        credits.append(df['Invoice Line Total'][i])
                        debits.append('')
                    else:
                        debits.append(abs(df['Invoice Line Total'][i]))
                        credits.append('')
                except AttributeError:
                    file = df['File Number'][i - 3]
                    state = df['File Number'][i - 3].split('-')[-1]
                    if df['Invoice Line Total'][i - 1] >= 0:
                        credits.append(df['Invoice Line Total'][i - 2])
                        debits.append('')
                    else:
                        debits.append(abs(df['Invoice Line Total'][i - 2]))
                        credits.append('')
    
                if not state == 'R':
                    states.append(state)
                else:
                    states.append('01')
    
                if file.endswith('CD-02') or file.endswith('CD-01'):
                    branches.append('007')
                elif file.endswith('WB-01'):
                    branches.append('009')
                elif file.endswith('LW-01'):
                    branches.append('006')
                elif file.endswith('OC-01'):
                    branches.append('010')
                elif file.endswith('WW-01'):
                    branches.append('011')
                elif file.endswith('MM-01'):
                    branches.append('016')
                elif file.endswith('NF-01'):
                    branches.append('012')
                elif file.endswith('TR-01'):
                    branches.append('004')
                elif file.endswith('PN-02'):
                    branches.append('013')
                elif file.endswith('SF-01') or file.endswith('SF-02'):
                    branches.append('018')
                else:
                    branches.append('000')
    
                if acct.startswith('6'):
                    depts.append('02')
                else:
                    depts.append('00')

            dates.append(date)
            types.append(at)
            acct_descr.append('')
            descr_ref.append(dr)
            accounts.append(acct)

    return pd.DataFrame({
        'Date': dates,
        'Type': types,
        'Accounts': accounts,
        'St': states,
        'Branch': branches,
        'Dept': depts,
        'Account Desr': acct_descr,
        'Description Reference': descr_ref,
        'Debits': debits,
        'Credits': credits
    })

def parse_counts(df, date):
    accounts, states, branches, depts, debits, credits = [], [], [], [], [], []
    dates, types, acct_descr, descr_ref = [], [], [], []
    df.columns = ['col' + str(i) for i in range(len(df.columns))]    

    for i, cell in enumerate(df['col0']):
        if not type(cell) == float:
            if cell.startswith('9'):
                accounts.append(cell.strip('-'))
                states.append(df['col1'][i].strip('-'))
                branches.append(df['col2'][i].strip('-'))
                depts.append(str(df['col3'][i]))
                debits.append(float(df['col4'][i]))
                credits.append('')
                dates.append(date)
                types.append('G/L Account')
                acct_descr.append('')
                descr_ref.append(date + ' RQ DEP')
                
    for i, cell in enumerate(df['col5']):
        if not type(cell) == float:
            if cell.startswith('9'):
                accounts.append(cell.strip('-'))
                states.append(df['col6'][i].strip('-'))
                branches.append(df['col7'][i].strip('-'))
                depts.append(str(df['col8'][i]))
                debits.append(df['col10'][i])
                credits.append('')
                dates.append(date)
                types.append('G/L Account')
                acct_descr.append('')
                descr_ref.append(date + ' RQ DEP')
    
    unit_offset_sum = sum(debits)
    dates.append(date)
    types.append('G/L Account')
    accounts.append('99998')
    states.append('00')
    branches.append('000')
    depts.append('00')
    acct_descr.append('')
    descr_ref.append(date + ' RQ DEP')
    debits.append('')
    credits.append(float(unit_offset_sum))
    
    return pd.DataFrame({
        'Date': dates,
        'Type': types,
        'Accounts': accounts,
        'St': states,
        'Branch': branches,
        'Dept': depts,
        'Account Desr': acct_descr,
        'Description Reference': descr_ref,
        'Debits': debits,
        'Credits': credits
    })
    
def deposit_total(df, date):
    accounts, states, branches, depts, debits, credits = [], [], [], [], [], []
    dates, types, acct_descr, descr_ref = [], [], [], []
    for i, cell in enumerate(df['Description']):
        if cell == 'Deposit Total':
            total = df['Invoice Line Total'][i]
    
    co = df.iloc[0, 0]
    bank_account = ''

    if co == 'Atlantic Shore Title, LLC':
        bank_account = 'OPERTD8038'
    elif co == 'Agents Title, LLC':
        bank_account = 'OPERTD5749'
    elif co == 'ESL Abstract, LLC':
        bank_account = 'OPERTD4040'
    elif co == 'ESL Title and Settlement Services, LLC':
        bank_account = 'OPERTD4040'
    elif co == 'GDV Abstract, LLC':
        bank_account = 'OPERUNI6481'
    elif co == 'Group 21 Title Agency, LLC':
        bank_account = 'OPERTD4819'
    elif co == 'InDeed Abstract, LLC':
        bank_account = 'OPERTD7451'
    elif co == 'Northern New Jersey Title Services, LLC':
        bank_account = 'OPERTD3694'
    elif co == 'Surety Lender Services, LLC':
        bank_account = 'OPERTD3694'
    elif co == 'Surety Title Agency Coastal Region, LLC':
        bank_account = 'OPERTD6791'
    elif co == 'Surety Title Agency of Haddonfield, LLC':
        bank_account = 'OPERTD2176'
    elif co == 'Surety Title Company, LLC':
        bank_account = 'OPERTD9831'
        
    accounts.append(bank_account)
    states.append('00')
    branches.append('000')
    depts.append('00')
    debits.append(total)
    credits.append('')
    
    dates.append(date)
    types.append('Bank Account')
    acct_descr.append('')
    descr_ref.append(date + ' RQ DEP')
    
    return pd.DataFrame({
        'Date': dates,
        'Type': types,
        'Accounts': accounts,
        'St': states,
        'Branch': branches,
        'Dept': depts,
        'Account Desr': acct_descr,
        'Description Reference': descr_ref,
        'Debits': debits,
        'Credits': credits
    })
