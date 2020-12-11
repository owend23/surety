#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
main loop for program.
"""
import pandas as pd
import surety_v4 as st4

date = input('Date: ')

wb_date = date.replace('/', '_')

ending = ' cash_receipts.xlsx'

sheets = []

while True:
    
    filename = input('Enter Filename: ')
    if filename == '':
        break
    
    sheet = filename.split('.')[0]
    
    filename2 = sheet + '_closing.xls'
    
    df = pd.read_excel('docs/cash_receipts/' + filename)
    df2 = pd.read_excel('docs/cash_receipts/' + filename2)

    journal = st4.parse_worksheet(df, date)
    deposit = st4.deposit_total(df, date)
    counts = st4.parse_counts(df2, date)
    
    frames = [journal, deposit, counts]
    
    df = pd.concat(frames)
    
    sheets.append((df, sheet))
    
file = open(wb_date + ending, 'wb')
file.close()

file = wb_date + ending

with pd.ExcelWriter(file) as writer:
    for i in range(len(sheets)):
        sheets[i][0].to_excel(writer, sheet_name=sheets[i][1], index=False)
