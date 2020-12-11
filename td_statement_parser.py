import pandas as pd


def parse_td_statement(df):
    types,accounts,states,branches,depts,descriptions,qty,direct_unit_cost,ic_partner_ref_type,ic_partner_code,ic_partner_ref = [],[],[],[],[],[],[],[],[],[],[]
    account = ''
    
    for i, cell in enumerate(df['Original Amount']):
        if cell > 0:
            if not df['Merchant Name'][i] == 'STANDARD VCF 4.4 100':
                account = '60300'
            else:
                account = '63004'
            types.append('G/L Account')
            accounts.append(account)
            states.append('01')
            branches.append('000')
            depts.append('00')
            descriptions.append(df['Merchant Name'][i])
            qty.append(1)
            direct_unit_cost.append(df['Original Amount'][i])
            ic_partner_ref_type.append('G/L Account')
            ic_partner_code.append('')
            ic_partner_ref.append('')
    
    return pd.DataFrame({
        'Type': types,
        'NO': accounts,
        'STATE': states,
        'Branch Code': branches,
        'Dept Code': depts,
        'Description/Comment': descriptions,
        'Quantity': qty,
        'Direct Unit Cost': direct_unit_cost,
        'IC Partner Ref Type': ic_partner_ref_type,
        'IC Partner Code': ic_partner_code,
        'IC Partner Reference': ic_partner_ref,
        })


while True:
    file = input('Filename: ')
    if file == '':
        break
    df = pd.read_excel('docs/td_statements/' + file)
    df = parse_td_statement(df)

    df.to_excel(file, index=False)

