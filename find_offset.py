#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Jan 14 10:51:40 2021

@author: owen
"""

import pandas as pd

file = 'MARLTON SUSPENSE.xlsx'

df = pd.read_excel(file)

debits = df['Debit Amount'].tolist()
credits = df['Credit Amount'].tolist()

debits = sorted(debits)
credits = sorted(credits)

debits = debits[121:]
credits = credits[126:]

missing = []
for c in credits:
    if credits.count(c) != debits.count(c):
        missing.append(c)
    
missing2 = []

for d in debits:
    if debits.count(d) != credits.count(d):
        missing2.append(d)
        
