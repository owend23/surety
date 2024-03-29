import os
import pandas as pd

def parse_worksheet(df, date):
    accounts, states, branches, depts, debits, credits = [], [], [], [], [], []
    dates, types, acct_descr, descr_ref = [], [], [], []

    acct = ''
    dr = date + ' RQ DEP'
    at = 'G/L Account'

    for i, cell in enumerate(df['Description'][:-3]):
        if not type(cell) == float:
            if str(cell).isnumeric() and len(str(cell)) == 5:
                acct = cell
        else:
            if str(cell) == 'nan':
                if df['Invoice Line Total'][i] >= 0:
                    credits.append(round(df['Invoice Line Total'][i], 2))
                    debits.append('')
                else:
                    debits.append(abs(round(df['Invoice Line Total'][i], 2)))
                    credits.append('')
                
                file = df['File Number'][i - 1]
                
                if file.startswith('CS'):
                    state = '01'
                else:
                    state = df['File Number'][i - 1].split('-')[-1]
    
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
                elif file.endswith('SG-01'):
                    branches.append('014')
                elif file.endswith('JC-01'):
                    branches.append('019')
                elif file.endswith('TU-01'):
                    branches.append('003')
                else:
                    branches.append('000')
    
                if str(acct).startswith('6'):
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
        'Account': accounts,
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
    
    co = df.iloc[0, 0]
    for i, cell in enumerate(df['col0']):
        if not type(cell) == float:
            if str(cell).startswith('9'):
                if cell == '96021-' and co == 'ESL Title and Settlement Services ':
                    accounts.append('96024')
                    accounts.append(df['col5'][i].strip('-'))
                else:
                    accounts.append(str(cell).strip('-'))
                    accounts.append(df['col5'][i].strip('-'))
                
                if not df['col1'][i].strip('-') == '???????':
                    states.append(df['col1'][i].strip('-'))
                else:
                    states.append('')
                    
                if not df['col6'][i].strip('-') == '???????':
                    states.append(df['col6'][i].strip('-'))
                else:
                    states.append('')
                    
                branches.append(df['col2'][i].strip('-'))
                branches.append(df['col7'][i].strip('-'))
                depts.append('00')
                depts.append('00')
                debits.append(round(df['col4'][i], 2))
                debits.append(round(df['col10'][i], 2))
                credits.append('')
                credits.append('')
                dates.append(date)
                dates.append(date)
                types.append('G/L Account')
                types.append('G/L Account')
                acct_descr.append('')
                acct_descr.append('')
                descr_ref.append(date + ' RQ DEP')
                descr_ref.append(date + ' RQ DEP')
                
    offset_sum = sum(debits)
    dates.append(date)
    types.append('G/L Account')
    accounts.append('99998')
    states.append('00')
    branches.append('000')
    depts.append('00')
    acct_descr.append('')
    descr_ref.append(date + ' RQ DEP')
    debits.append('')
    credits.append(float(offset_sum))
    
    return pd.DataFrame({
        'Date': dates,
        'Type': types,
        'Account': accounts,
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
            total = round(df['Invoice Line Total'][i], 2)
    
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
        bank_account = 'OPERTD1263'
    elif co == 'Surety Title Agency Coastal Region, LLC':
        bank_account = 'OPERTD6791'
    elif co == 'Surety Title Agency of Haddonfield, LLC' or co == 'Surety Title Abstract Group, LLC':
        bank_account = 'OPERTD2176'
    elif co == 'Surety Title Company, LLC':
        bank_account = 'OPERTD9831'
    elif co == 'Surety Abstract Services, LLC':
        bank_account = 'OPERTD9831'
    elif co == 'RealSafe Title, LLC':
        bank_account = 'OPERBU9313'
    elif co == 'One Abstract Services, LLC':
        bank_account = 'OPERTD4550'
    elif co == 'Surety Title Services North Region':
        bank_account = 'OPERTD3694'
    elif co == 'Your Hometown Title, LLC' or co == 'YHT Settlement Track, LLC':
        bank_account = 'OPERTD3093'
    elif co == 'Surety Abstract Capital, LLC':
        bank_account = 'OPERTD6631'
    elif co == 'Facilities Abstract, LLC':
        bank_account = 'OPERBU0331'
    elif co == 'Turton Signature Title Agency':
        bank_account = 'OPERTD6791'
    elif co == 'Jackson Title Agency, LLC':
        bank_account = 'OPERBU0595'
    elif co == 'Surety Abstract Ventures, LLC':
        bank_account = 'OPERTD8201'
        
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
        'Account': accounts,
        'St': states,
        'Branch': branches,
        'Dept': depts,
        'Account Desr': acct_descr,
        'Description Reference': descr_ref,
        'Debits': debits,
        'Credits': credits
    })

def check_workbook(file):
    wb = pd.ExcelFile(file)
    sheets = wb.sheet_names
    
    for sheet in sheets:
        ws = wb.parse(sheet)
        deposit_total = ws[ws['Type'] == 'Bank Account']['Debits'].sum()
        acct_total = 0
        
        for i, cell in enumerate(ws['Account']):
            if str(cell).startswith('96'):
                acct_total += ws['Debits'][i]
                
        deposit_total = round(deposit_total, ndigits=2)
        acct_total = round(acct_total, ndigits=2)
                
        if deposit_total == acct_total:
            continue
        else:
            print(sheet + ' is off balance')
            print('Deposit Total:', deposit_total)
            print('Account Total:', acct_total)
            print('Difference:', round(deposit_total - acct_total, ndigits=2))
            print('\n')

def create_sheet(dir):
    date = input('Date: ')
    wb_date = date.replace('/', '_')
    ending = ' cash_receipts.xlsx'
    
    sheets = []
    deposits,counts = [],[]
    for file in os.listdir(dir):
        if not file.endswith('_closing.xls'):
            if not file.startswith('.'):
                deposits.append(file)
        else:
            if not file.startswith('.'):
                counts.append(file)
    
    deposits = sorted(deposits)
    counts = sorted(counts)
                
    for i, file in enumerate(deposits):
        sheet = file.split('.')[0]
        
        df = pd.read_excel('docs/cash_receipts/' + file)
        df2 = pd.read_excel('docs/cash_receipts/' + counts[i])
        
        try:
            report = parse_worksheet(df, date)
        except AttributeError:
            for i, cell in enumerate(df.iloc):
                if not df.iloc[i, :].any():
                    print('Row ' + str(i + 2) + ' of ' + sheet + '.xls is blank please delete')
            
        total = deposit_total(df, date)
        count = parse_counts(df2, date)
        frames = [report, total, count]
        df = pd.concat(frames)
        sheets.append((df, sheet))
        
    file = open(wb_date + ending, 'wb')
    file.close()

    file = wb_date + ending

    with pd.ExcelWriter(file) as writer:
        for i in range(len(sheets)):
            sheets[i][0].to_excel(writer, sheet_name=sheets[i][1], index=False)

    check_workbook(file)
    
create_sheet('docs/cash_receipts/')
        
