import sklearn
from sklearn.neighbors import KNeighborsClassifier
import pandas as pd
from sklearn import preprocessing
import pickle

data = pd.read_excel('cash (1).xls')
data = data[['Description','AcctCode','Invoice Line Total','OrderCategory']]
data.dropna(how='any', inplace=True)

print(data.head())

le = preprocessing.LabelEncoder()
description = le.fit_transform(list(data['Description']))
acctCode = le.fit_transform(list(data['AcctCode']))
orderCategory = le.fit_transform(list(data['OrderCategory']))
invoice_total = list(data['Invoice Line Total'])

accts = set(list(data['AcctCode']))
acctCodeValues = set(list(acctCode))

D = dict(zip(acctCodeValues,accts))

X = list(zip(description, orderCategory, invoice_total))
y = list(acctCode)


x_train, x_test, y_train, y_test = sklearn.model_selection.train_test_split(X, y, test_size=0.1)

model = KNeighborsClassifier(n_neighbors=1)

model.fit(x_train, y_train)
acc = model.score(x_test, y_test)
print(acc)

predicted = model.predict(x_test)

correct, wrong = [], []

for x in range(len(predicted)):
    if predicted[x] == y_test[x]:
        correct.append(f"Predicted: {D[predicted[x]]} Data: {x_test[x]}, Actual: {D[y_test[x]]}")
    else:
        wrong.append(f"Predicted: {D[predicted[x]]} Data: {x_test[x]} Actual:  {D[y_test[x]]}")

print('CORRECT')
[print(x) for x in correct]
print()
print('WRONG')
[print(x) for x in wrong]
