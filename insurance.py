import pandas as pd
import datetime

def parse_general_ledger(filename):
	df = pd.read_excel(filename)
