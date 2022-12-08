import pandas as pd

oldFile = 'testDifferenceOld.xlsx'
newFile = 'testDifferenceNew.xlsx'

dfOld = pd.read_excel(oldFile)
dfNew = pd.read_excel(newFile)
dfDifference = pd.merge([dfOld, dfNew], on=['character', 'race'], join="inner").drop_duplicates()
print(dfDifference)
# print('Old')
# print(dfOld)
# print('New')
# print(dfNew)
# print('difference')
# print(dfDifference)
# dfCommon =

# df = df_one[~df_one.index.isin(df_two.index)]