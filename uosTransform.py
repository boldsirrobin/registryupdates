# Convert Registry Excel to Moodle upload file - only for UoS weekly update files!




import pandas as pd
import numpy as np
import re

# Set sheet-specific variables
# -----------------------------
## Filenames
partner = 'UoS'
intake = 'weekly'
dateReceived = '20221202'
inputSuffix = '.xlsx'
outputSuffix = '.csv'
inputFile = partner + intake + '_' + dateReceived + inputSuffix
outputFile = partner + intake + '_' + dateReceived + outputSuffix
print('File from Registry is ', inputFile)
print('File to upload to Moodle is ', outputFile)


##Operations
newAccounts = False  # Set to True if there may be new students
middleNameColumn = False  # Set to true if there is a middle name column
statusChange = True  # Set to True if there is a column for status
cohortChange = True # Set to true if there is a column for courses
groupChange = True # Set to True if there is campus/group information
levelDataNeeded = False  # Set to True if there is a column with level data etc. that is needed to determine cohort.

programmes = ['2019-10-01', '2020-02-01', '2020-10-01', '2021-02-01', '2021-10-01', '2022-02-01', '2022-06-01', '2022-10-01']
cohorts = ['UOSOCT2019', 'UOSFEB2020', 'UOSOCT2020', 'UOSFEB2021', 'UOSOCT2021', 'UOSFEB2022', 'UOSJUN22', 'UOSSEP22']

# Should be no need to change anything below this line
# =====================================================


# Create workbook object df
df = pd.read_excel(inputFile)

print("\nThe column headers :")
oldColumnHeaders = df.columns  # Get the column headers from the file
columnHeaders = [i.lower() for i in oldColumnHeaders]  # Make them all lowercase

# Search for keywords and replace with the headers we actually want
columnHeaders = ['username' if 'lead' in i or 'zoho' in i else i for i in columnHeaders]
columnHeaders = ['firstname' if 'first' in i else i for i in columnHeaders]
columnHeaders = ['middlename' if 'middle' in i else i for i in columnHeaders]
columnHeaders = ['lastname' if 'last' in i else i for i in columnHeaders]
columnHeaders = ['cohort1' if 'course' in i or 'cohort' in i and 'tt' not in i else i for i in columnHeaders]
columnHeaders = ['level' if 'study' in i or 'tt' in i else i for i in columnHeaders]
columnHeaders = ['campus' if 'campus' in i and 'change' not in i else i for i in columnHeaders]
columnHeaders = ['group' if 'group' in i and 'campus' not in i else i for i in columnHeaders]
columnHeaders = ['suspended' if 'status' in i and 'visa' not in i else i for i in columnHeaders]
print(columnHeaders)
# Write the new column headers to the dataframe
df.columns = columnHeaders
# Make a list of the columns we want
columns = ['username', 'cohort1', 'campus', 'group']
if newAccounts:
    columns.append('firstname')
    columns.append('lastname')
if middleNameColumn:
    columns.append('middlename')
if levelDataNeeded:
    columns.append('level')
if statusChange:
    columns.append('suspended')
# Rstrict the dataframe to these columns
df = df[columns]

# Make it all strings (or you get NaN for some usernames)
df = df.astype(str)
# Strip whitespace from cell contents
for col in df.columns:
    try:
        df[col] = df[col].str.strip()
    except AttributeError:
        pass

# Make usernames lowercase
df['username'] = df['username'].str.lower()

# Make new cohorts replace old cohorts where not blank (NaT).
# df.loc[df['level'] != 'NaT', 'cohort1'] = df['level']

# Search-replace
# ---------------

## Account suspension
df['suspended'] = df['suspended'].replace(
    '(?i).*withdraw.*|.*cancel.*|.*declined.*|.*IOS.*|.*intercalated.*|.*suspend.*|.*defer.*|.*complete.*', '1', regex=True)
df['suspended'] = df['suspended'].replace('(?i).*enrol.*|.*attended.*|.*pending.*|.*progress.*', '0', regex=True)

## Cohort names

for i in range(len(programmes)):
    df['cohort1'] = df['cohort1'].replace(programmes[i], cohorts[i], regex=True)

## Campus names
df['campus'] = df['campus'].replace('(?i).*Global.*|.*Duncan.*|.*GedU.*', 'Global_Edu_', regex=True)
df['campus'] = df['campus'].replace('(?i).*Greenford.*', 'Greenford_', regex=True)
df['campus'] = df['campus'].replace('(?i).*Manchester.*', 'Manchester_', regex=True)
df['campus'] = df['campus'].replace('(?i).*Cam.*|.*Stratford.*', 'Cam_Road_', regex=True)
df['campus'] = df['campus'].replace('(?i).*Bow.*', 'Bow_Road_', regex=True)
df['campus'] = df['campus'].replace('(?i).*Leeds.*', 'Leeds_', regex=True)
df['campus'] = df['campus'].replace('(?i).*Republic.*|.*Poplar.*', 'Republic_', regex=True)
df['campus'] = df['campus'].replace('(?i).*Birmingham.*', 'Birmingham_', regex=True)


# Add extra fields for new users
if newAccounts:
    df['email'] = df['username'] + '@globalbanking.ac.uk'
    df['password'] = 'Gbs1234@'

# Clean up the group column and merge with campus column
df['group'] = df['group'].str[:2]
df['profile_field_GroupAllocation'] = df['campus'] + df['group']

# Delete unwanted columns
df = df.drop(['campus', 'group'], axis=1)

df.to_csv(outputFile, index=False)
print(df)


#==============




