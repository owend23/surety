#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Main loop creates a new spreadsheet based off the input from the user giving the filename.
"""

import surety_v4 as st4
import pandas as pd

date = input('Date: ')

ending = '_cash_receipts.xlsx'

while True:

    filename = input('Enter Filename: ')
    if filename == '':
        break
    
    filename2 = input('Enter Count File: ')
    if filename2 == '':
        break
    
    sheet_name = filename.split('.')[0]
    
    df = pd.read_excel('docs/' + filename)
    df2 = pd.read_excel('docs/' + filename2)
    
    
    journal = st4.parse_worksheet(df, date)
    deposit = st4.deposit_total(df, date)
    counts = st4.parse_counts(df2, date)
    
    frames = [journal, deposit, counts]
    
    df = pd.concat(frames)
    
    df.to_excel(sheet_name + ending)
