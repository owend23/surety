import pandas as pd
import numpy as np

xl = pd.ExcelFile('journals/04_19_2021 cash_receipts.xlsx')
shorts = ['66300','66302']
frames = []

for sheet in xl.sheet_names:
    df = xl.parse(sheet, converters={'Account':str,'St':str,'Branch':str,'Dept':str})
    if '66300' or '66302' in df['Account']:
        df1 = df[~df['Account'].isin(shorts)]
        df1 = df1[df1['Type'] != 'Bank Account']
        df1 = df1[~df1.Account.str.startswith('9')]
        total = df[df['Type'] == 'Bank Account']
        df2 = df[df['Account'].isin(shorts)]
        df3 = df[df.Account.str.startswith('9')]


        df1 = df1.groupby(['Account','St','Branch']).agg({'Date':'first','Type':'first','Dept':'first','Account Desr':'first',
                                                          'Description Reference':'first','Debits':'sum','Credits':'sum'}).reset_index()
        if len(df2) > 0:
            df = pd.concat([df1,df2,total,df3], ignore_index=True)
        else:
            df = pd.concat([df1,total,df3])
        for i, cell in enumerate(df['Debits']):
            if cell == 0:
                df.loc[i, 'Debits'] = np.nan
        for i, cell in enumerate(df['Credits']):
            if cell == 0:
                df.loc[i, 'Credits'] = np.nan
        df = pd.DataFrame({'Date':df.Date,'Type':df.Type,'Account':df.Account,'St':df.St,'Branch':df.Branch,
        'Dept':df.Dept,'Account Desr':df['Account Desr'],'Description Reference':df['Description Reference'],
        'Debits':df.Debits,'Credits':df.Credits})
        frames.append((df, sheet))
