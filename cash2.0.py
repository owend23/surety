import pandas as pd
import numpy as np 
import datetime
import os

def find_shortages(filename):
    xl = pd.ExcelFile(filename)
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        if '66302' in df['Account'].tolist():
            print('66302')
            print(sheet)
            print()
        elif '66300' in df['Account'].tolist():
            print('66300')
            print(sheet)
            print()

def get_accounts(df):
    for i, cell in enumerate(df['Description']):
        if not type(cell) == float:
            if len(cell) == 5 and str(cell).isnumeric():
                yield cell
            elif str(cell) == 'Deposit Total':
                co = df.iloc[0, 0]

                if co == 'Atlantic Shore Title, LLC':
                    yield 'OPERTD8038'
                elif co == 'Agents Title, LLC':
                    yield 'OPERTD5749'
                elif co == 'ESL Abstract, LLC':
                    yield 'OPERTD4040'
                elif co == 'ESL Title and Settlement Services, LLC':
                    yield 'OPERTD4040'
                elif co == 'GDV Abstract, LLC':
                    yield 'OPERUNI6481'
                elif co == 'Group 21 Title Agency, LLC':
                    yield 'OPERTD4819'
                elif co == 'InDeed Abstract, LLC':
                    yield 'OPERTD7451'
                elif co == 'Northern New Jersey Title Services, LLC':
                    yield 'OPERTD3694'
                elif co == 'Surety Lender Services, LLC':
                    yield 'OPERTD1263'
                elif co == 'Surety Title Agency Coastal Region, LLC':
                    yield 'OPERTD6791'
                elif co == 'Surety Title Agency of Haddonfield, LLC' or co == 'Surety Title Abstract Group, LLC':
                    yield 'OPERTD2176'
                elif co == 'Surety Title Company, LLC':
                    yield 'OPERTD9831'
                elif co == 'Surety Abstract Services, LLC':
                    yield 'OPERTD9831'
                elif co == 'RealSafe Title, LLC':
                    yield 'OPERBU9313'
                elif co == 'One Abstract Services, LLC':
                    yield 'OPERTD4550'
                elif co == 'Surety Title Services North Region':
                    yield 'OPERTD3694'
                elif co == 'Your Hometown Title, LLC' or co == 'YHT Settlement Track, LLC':
                    yield 'OPERTD3093'
                elif co == 'Surety Abstract Capital, LLC':
                    yield 'OPERTD6631'
                elif co == 'Facilities Abstract, LLC':
                    yield 'OPERBU0331'
                elif co == 'Turton Signature Title Agency':
                    yield 'OPERTD6791'
                elif co == 'Jackson Title Agency, LLC':
                    yield 'OPERBU0595'
                elif co == 'Surety Abstract Ventures, LLC':
                    yield 'OPERTD8201'
                else:
                    yield ''

def get_credits(df):
    for i, cell in enumerate(df['Description']):
        if str(cell) == 'Accounting Code Total':
            if df['Invoice Line Total'][i] >= 0:
                yield round(df['Invoice Line Total'][i], ndigits=2)
            else:
                yield np.nan
        elif str(cell) == 'Deposit Total':
            yield np.nan

def get_debits(df):
    for i, cell in enumerate(df['Description']):
        if str(cell) == 'Accounting Code Total':
            if df['Invoice Line Total'][i] < 0:
                yield abs(round(df['Invoice Line Total'][i], ndigits=2))
            else:
                yield np.nan
        elif str(cell) == 'Deposit Total':
            yield round(df['Invoice Line Total'][i], ndigits=2)

def get_dates(df):
    for i, cell in enumerate(df['Description']):
        if str(cell) == 'Accounting Code Total':
            yield datetime.date.strftime(df['PaymentDate'][i - 1], "%m/%d/%Y")
        elif str(cell) == 'Deposit Total':
            yield datetime.date.strftime(df['PaymentDate'][i - 2], "%m/%d/%Y")

def get_types(df):
    for i, cell in enumerate(df['Description']):
        if str(cell) == 'Accounting Code Total':
            yield 'G/L Account'
        elif str(cell) == 'Deposit Total':
            yield 'Bank Account'

def get_state_codes(df):
    for i, cell in enumerate(df['Description']):
        if str(cell) == 'Accounting Code Total':
            file = df['File Number'][i - 1]
            if not df['File Number'][i - 1].startswith('CS'):
                state = df['File Number'][i - 1].split('-')[-1]
                if state == 'R':
                    yield '01'
                else:
                    yield state
            else:
                yield '01'
        elif str(cell) == 'Deposit Total':
            yield '00'

def get_branches(df):
    for i, cell in enumerate(df['Description']):
        if str(cell) == 'Accounting Code Total':
            file = df['File Number'][i - 1]

            if file.endswith('CD-02') or file.endswith('CD-01'):
                yield '007'
            elif file.endswith('WB-01'):
                yield '009'
            elif file.endswith('LW-01'):
                yield '006'
            elif file.endswith('OC-01'):
                yield '010'
            elif file.endswith('WW-01'):
                yield '011'
            elif file.endswith('MM-01'):
                yield '016'
            elif file.endswith('NF-01'):
                yield '012'
            elif file.endswith('TR-01'):
                yield '004'
            elif file.endswith('PN-02'):
                yield '013'
            elif file.endswith('SF-01') or file.endswith('SF-02'):
                yield '018'
            elif file.endswith('SG-01'):
                yield '014'
            elif file.endswith('JC-01'):
                yield '019'
            elif file.endswith('TU-01'):
                yield '003'
            else:
                yield '000'
        elif str(cell) == 'Deposit Total':
            yield '000'

def get_department_codes(df):
    for i, cell in enumerate(df['Description']):
        if len(str(cell)) == 5 and str(cell).isnumeric():
            acct = cell
        if str(cell) == 'Accounting Code Total':
            if str(acct).startswith('6'):
                yield '02'
            else:
                yield '00'
        elif str(cell) == 'Deposit Total':
            yield '00'

def get_account_descriptions(df):
    for i, cell in enumerate(df['Description']):
        if str(cell) == 'Accounting Code Total':
            yield ''
        elif str(cell) == 'Deposit Total':
            yield ''

def get_description_references(df):
    for i, cell in enumerate(df['Description']):
        if str(cell) == 'Accounting Code Total':
            yield datetime.date.strftime(df['PaymentDate'][i - 1], "%m/%d/%Y") + ' RQ DEP'
        elif str(cell) == 'Deposit Total':
            yield datetime.date.strftime(df['PaymentDate'][i - 2], "%m/%d/%Y") + ' RQ DEP'

def create_report(filename):
    df = pd.read_excel(filename)
    
    df = pd.DataFrame({
        'Date': list(get_dates(df)),
        'Type': list(get_types(df)),
        'Account': list(get_accounts(df)),
        'St': list(get_state_codes(df)),
        'Branch': list(get_branches(df)),
        'Dept': list(get_department_codes(df)),
        'Account Desr': list(get_account_descriptions(df)),
        'Description Reference': list(get_description_references(df)),
        'Debits': list(get_debits(df)),
        'Credits': list(get_credits(df)),
        })

    df['Debits'].astype(float)
    df['Credits'].astype(float)
    df['Account'].astype(str)
    df['St'].astype(str)
    df['Branch'].astype(str)
    df['Dept'].astype(str)

    debits = round(df['Debits'].sum(), ndigits=2)
    credits = round(df['Credits'].sum(), ndigits=2)

    if debits == credits:
        return df

def get_date(filename):
    df = pd.read_excel(filename)
    for i, cell in enumerate(df['Description']):
        if str(cell) == 'Accounting Code Total':
            return datetime.date.strftime(df['PaymentDate'][i - 1], "%m/%d/%Y")

def parse_counts(filename, date):
    df = pd.read_excel(filename)

    accounts, states, branches, depts, debits, credits = [], [], [], [], [], []
    dates, types, acct_descr, descr_ref = [], [], [], []
    df.columns = ['col' + str(i) for i in range(len(df.columns))]

    co = df.iloc[0, 0]
    for i, cell in enumerate(df['col0']):
        if not type(cell) == float:
            if str(cell).startswith('9'):
                if cell == '96021-' and co == 'ESL Title and Settlement Services ':
                    accounts.append('96024')
                    accounts.append(str(df['col5'][i]).strip('-'))
                else:
                    accounts.append(str(cell).strip('-'))
                    accounts.append(str(df['col5'][i]).strip('-'))

                if not str(df['col1'][i]).strip('-') == '???????':
                    states.append(str(df['col1'][i]).strip('-'))
                else:
                    states.append(np.nan)

                if not str(df['col6'][i]).strip('-') == '???????':
                    states.append(str(df['col6'][i]).strip('-'))
                else:
                    states.append(np.nan)

                branches.append(str(df['col2'][i]).strip('-'))
                branches.append(str(df['col7'][i]).strip('-'))
                depts.append('00')
                depts.append('00')
                debits.append(round(df['col4'][i], 2))
                debits.append(round(df['col10'][i], 2))
                credits.append(np.nan)
                credits.append(np.nan)
                dates.append(date)
                dates.append(date)
                types.append('G/L Account')
                types.append('G/L Account')
                acct_descr.append(np.nan)
                acct_descr.append(np.nan)
                descr_ref.append(date + ' RQ DEP')
                descr_ref.append(date + ' RQ DEP')

    offset_sum = sum(debits)
    dates.append(date)
    types.append('G/L Account')
    accounts.append('99998')
    states.append('00')
    branches.append('000')
    depts.append('00')
    acct_descr.append(np.nan)
    descr_ref.append(date + ' RQ DEP')
    debits.append(np.nan)
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

def check_workbook(file):
    wb = pd.ExcelFile(file)
    for sheet in wb.sheet_names:
        ws = wb.parse(sheet)
        deposit_total = ws[ws['Type'] == 'Bank Account']['Debits'].sum()
        acct_total = 0

        for i, cell in enumerate(ws['Account']):
            if str(cell).startswith('96') or str(cell).startswith('?'):
                acct_total += ws['Debits'][i]

        deposit_total = round(deposit_total, 2)
        acct_total = round(acct_total, 2)

        if deposit_total == acct_total:
            continue
        else:
            print(sheet, ' is off balance')
            print('Deposit Total:', deposit_total)
            print('Account Total:', acct_total)
            print('Difference:', round(deposit_total - acct_total, 2))
            print('\n')

def get_count_accounts(df):
    df.columns = ['col' + str(i) for i in range(len(df.columns))]
    co = df.iloc[0, 0]
    for i, cell in enumerate(df['col0']):
        if not type(cell) == float:
            if str(cell).startswith('9'):
                if cell == '96021-' and co == 'ESL Title and Settlement Services ':
                    yield '96024'
                    yield str(df['col5'][i]).strip('-')
                else:
                    yield str(cell).strip('-')
                    yield str(df['col5'][i]).strip('-')
    yield '99998'

def get_count_states(df):
    df.columns = ['col' + str(i) for i in range(len(df.columns))]
    for i, cell in enumerate(df['col0']):
        if not type(cell) == float:
            if str(cell).startswith('9'):
                if not str(df['col1'][i]).strip('-') == '???????':
                    yield str(df['col1'][i]).strip('-')
                    yield str(df['col1'][i]).strip('-')
                else:
                    yield '01'
                    yield '01'
    yield '00'

def get_count_branches(df):
    df.columns = ['col' + str(i) for i in range(len(df.columns))]
    for i, cell in enumerate(df['col0']):
        if not type(cell) == float:
            if str(cell).startswith('9'):
                yield str(df['col2'][i]).strip('-')
                yield str(df['col2'][i]).strip('-')
    yield '000'

def get_count_depts(df):
    df.columns = ['col' + str(i) for i in range(len(df.columns))]
    for i, cell in enumerate(df['col0']):
        if not type(cell) == float:
            if str(cell).startswith('9'):
                yield '00'
                yield '00'
    yield '00'

def get_count_debits(df):
    total = 0
    df.columns = ['col' + str(i) for i in range(len(df.columns))]
    for i, cell in enumerate(df['col0']):
        if not type(cell) == float:
            if str(cell).startswith('9'):
                yield round(df['col4'][i], 2)
                yield round(df['col10'][i], 2)
                total += round(df['col4'][i], 2)
                total += round(df['col10'][i], 2)
    yield total


def get_count_credits(df):
    df.columns = ['col' + str(i) for i in range(len(df.columns))]
    for i, cell in enumerate(df['col0']):
        if not type(cell) == float:
            if str(cell).startswith('9'):
                yield np.nan
                yield np.nan
    yield np.nan

def get_count_dates(df, date):
    df.columns = ['col' + str(i) for i in range(len(df.columns))]
    for i, cell in enumerate(df['col0']):
        if not type(cell) == float:
            if str(cell).startswith('9'):
                yield date
                yield date
    yield date

def get_count_types(df):
    df.columns = ['col' + str(i) for i in range(len(df.columns))]
    for i, cell in enumerate(df['col0']):
        if not type(cell) == float:
            if str(cell).startswith('9'):
                yield 'G/L Account'
                yield 'G/L Account'
    yield 'G/L Account'

def get_count_acct_descr(df):
    df.columns = ['col' + str(i) for i in range(len(df.columns))]
    for i, cell in enumerate(df['col0']):
        if not type(cell) == float:
            if str(cell).startswith('9'):
                yield np.nan
                yield np.nan
    yield np.nan

def get_count_descr_refs(df, date):
    df.columns = ['col' + str(i) for i in range(len(df.columns))]
    for i, cell in enumerate(df['col0']):
        if not type(cell) == float:
            if str(cell).startswith('9'):
                yield date + ' RQ DEP'
                yield date + ' RQ DEP'
    yield date + ' RQ DEP'

def create_count_report(filename, date):
    df = pd.read_excel(filename)
    return pd.DataFrame({
        'Date': list(get_count_dates(df, date)),
        'Type': list(get_count_types(df)),
        'Account': list(get_count_accounts(df)),
        'St': list(get_count_states(df)),
        'Branch': list(get_count_branches(df)),
        'Dept': list(get_count_depts(df)),
        'Account Desr': list(get_count_acct_descr(df)),
        'Description Reference': list(get_count_descr_refs(df, date)),
        'Debits': list(get_count_debits(df)),
        'Credits': list(get_count_credits(df)),
        })
    

def create_sheet(dir):
    wb_date = ''
    ending = ' cash_receipts.xlsx'
    sheets = []
    deposits, counts = [], []
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
        date = get_date('docs/cash_receipts/' + file)
        wb_date = date.replace('/', '_')
        sheet = file.split('.')[0]

        report = create_report('docs/cash_receipts/' + file)
        count = create_count_report('docs/cash_receipts/' + counts[i], date)
        frames = [report, count]
        df = pd.concat(frames)
        sheets.append((df, sheet))

    
    file = open(f'{wb_date} {ending}', 'wb')
    file.close()

    file = wb_date + ending

    with pd.ExcelWriter(file) as writer:
        for i in range(len(sheets)):
            sheets[i][0].to_excel(writer, sheet_name=sheets[i][1], index=False)

    check_workbook(file)


create_sheet('docs/cash_receipts')
