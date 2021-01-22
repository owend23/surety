#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Jan  7 14:29:17 2021

@author: owen
"""
import pandas as pd

def find_file(df1, df2):
    files1 = set(df1['File'].tolist())
    files2 = set(df2['File Number'].tolist())
    
    for file in files1:
        if not file in files2:
            print(file)
            

    