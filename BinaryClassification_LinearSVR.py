import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.svm import LinearSVR
from sklearn.metrics import roc_auc_score

import numpy as np
import json

train = pd.read_csv("training.csv")

def get_data_info(data):
    print("Data types:")
    print(data.dtypes)
    print("Rows and columns:")
    print(data.shape)
    print("Column info:")
    print(data.columns)
    print("Null value:")
    print(data.apply(lambda x: sum(x.isnull()) / len(data)))

#get_data_info(train)
means = train.mean()
train = train.fillna(means)

maxes = train.max()
train /= maxes

cols = train.columns[2:]
data = train[cols]
target = train["target"]

data_train, data_test, target_train, target_test = train_test_split(data, target, test_size = 0.2)
svr_model = LinearSVR(max_iter = 100000, verbose = 1)
model = svr_model.fit(data_train, target_train)
pred_train = model.predict(data_test)

print("AUC:", roc_auc_score(target_test, pred_train))

test = pd.read_csv("testing.csv")
test = test.fillna(means)
test /= maxes
pred_test = model.predict(test[test.columns[1:]])

pred_test = np.minimum(np.maximum(pred_test, 0), 1)
d = {}

for i in range(len(pred_test)):
    d[i] = pred_test[i]

f = open("output.json", "w")
json.dump(d, f)
f.close()
