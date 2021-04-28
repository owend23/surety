import pandas as pd
import numpy as np
import pickle

class Data:
	
	def __init__(self, df, escrow):
		self.df = df
		self.escrow = escrow
		self.accts = self.load_database('/data/workspace_files/accounts.pickle')
		self.branches = self.load_database('/data/workspace_files/branches.pickle')
		
	def load_database(self, filename):
		with open(filename, 'rb') as f:
			return pickle.load(f)
