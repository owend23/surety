import numpy as np
import pandas as pd
from sklearn.tree import DecisionTreeClassifier
from sklearn.preprocessing import LabelEncoder
from sklearn.model_selection import train_test_split

df = pd.read_excel('cash_data.xls')
df = df[['Description','AcctCode','Amount']]

df = df[df.AcctCode != 19999]
df.dropna(how='any', inplace=True)

refis = df[df.AcctCode == 40002].index.tolist()
data.at[refis, 'Description'] = 'ReFinance Loan'

le = LabelEncoder()

acctCode = le.fit_transform(df['AcctCode'].tolist())
description = le.fit_transform(df['Description'].tolist())
amounts = df['Amount'].tolist()

acct_dict = dict(zip(acctCode, df['AcctCode'].tolist()))
description_dict = dict(zip(description, df['Description'].tolist()))

X = list(zip(description, amounts))
y = acctCode

X_train, X_test, y_train, y_test = train_test_split(X, y, stratify=y, random_state=42)

tree = DecisionTreeClassifier(random_state=0)

tree.fit(X_train, y_train)

print("Accuracy on training set: {:.3f}".format(tree.score(X_train, y_train)))
print("Accuracy on test set: {:.3f}".format(tree.score(X_test, y_test)))
print()

predicted = tree.predict(X_test)

correct, wrong = [], []
for x in range(len(predicted)):
	test_data = list(X_test[x])
	test_data[0] = description_dict[test_data[0]]
	if predicted[x] == y_test[x]:
		correct.append(f"Predicted: {acct_dict[predicted[x]]} Actual: {acct_dict[y_test[x]]} Data: {test_data}")
	else:
		wrong.append(f"Predicted: {acct_dict[predicted[x]]} Actual: {acct_dict[y_test[x]]} Data: {test_data}")
		
print("WRONG PREDICTIONS")
for item in sorted(wrong):
	print(item)
	
print()

print("CORRECT PREDICTIONS")
for item in sorted(correct):
	print(item)
