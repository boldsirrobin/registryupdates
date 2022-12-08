# Convert Registry Excel to Moodle upload file
# This script currently works for GBS, BSU and normal CCCU files.
# It does NOT currently work on CCCU withdrawal files or UoS Weekly Updates.
# Edit the sections "Filenames" and "Operations" according to the file to be processed.

# TODO
# Remove "nan"


import pandas as pd
import sys

# Set sheet-specific variables
# -----------------------------
## Filenames
partner = 'CCCU' # This will be CCCU, GBS or BSU
intake = 'Sep22' # For updates, this might be something like  "weekly".
dateReceived = '20221207' # This is a date in Unix format and should be part of the file name.
inputSuffix = '.xlsx' # Set to .xlsx or .csv.
outputSuffix = '.csv'
inputFile = partner + intake + '_' + dateReceived + inputSuffix
outputFile = partner + intake + '_' + dateReceived + outputSuffix
print('File from Registry is ', inputFile)

##Operations
newAccounts = False  # Set to True if there may be new students
middleNameColumn = False  # Set to true if there is a middle name column (only important for new acocunts)
statusChange = False  # Set to True if there is a column for status
cohortChange = True  # Set to true if there is a column for courses
groupChange = True # Set to True if there is campus/group information
levelDataNeeded = False  # Set to True if there is a column with level data etc. that is needed to determine cohort.
maxFileLength = 0 # Maximum number of rows for CSV files. Set to 0 if you don't want to split the file.

## Programme/Cohort names
if partner == 'GBS' and cohortChange:
    programmes = ['.*Healthcare.*', '.*Digital Technologies.*', '.*Construction.*', '.*Business.*', '.*Education.*']
    if intake == 'Jan22':
        cohorts = ['GBSJAN22HC_L4', 'GBSJAN22DT_L4', 'GBSJAN22CM_L4', 'GBSJAN22BUS_L4', 'GBSJAN22ET_L5']
    elif intake == 'Jun22':
        cohorts = ['GBSJUN22HC_L4', 'GBSJUN22DT_L4', 'GBSJUN22CM_L4', 'GBSJUN22BUS_L4', 'GBSJUN22ET_L5']
    elif intake == 'Oct22' or intake == 'Sep22':
        cohorts = ['GBSSEP22HC_L4', 'GBSSEP22DT_L4', 'GBSSEP22CM_L4', 'GBSSEP22BUS_L4', 'GBSSEP22ET_L5']
    else:
        sys.exit("Invalid intake.")

if partner == 'BSU' and cohortChange:
    programmes = ['.*Construction.*', '.*Business.*']
    if intake == 'Feb22':
        cohorts = ['BSUCM_FEB2022_L3', 'BSUTOPUP_FEB2022']
    elif intake == 'Jun22':
        cohorts = ['BSUCM_JUNE2022_L3', 'BSUTOPUP_JUNE2022']
    elif intake == 'Oct22':
        cohorts = ['BSUCM_SEP2022_L3', 'BSUOCT22_TOPUP']
    elif intake == 'Oct22':
        cohorts = ['BSUCM_FEB2023_L3', 'BSUTOPUP_FEB2023']
    else:
        sys.exit("Invalid intake.")

if partner == 'CCCU' and cohortChange:
    if intake == 'Jun21':
        programmes = ['.*Tourism.*']
        cohorts = ['CCCUJUN21BTM_L4_01']
    elif intake == 'Sep21':
        programmes = ['.*Accounting.*', '.*Tourism.*']
        cohorts = ['CCCUSEP21AFM_02', 'CCCUSEP21BTM_L4']
    elif intake == 'Jan22':
        programmes = ['.*Accounting.*', '.*Tourism.*']
        cohorts = ['CCCUJAN22AFM_01', 'CCCUJAN22BTM_01']
    elif intake == 'Jun22':
        programmes = ['.*Accounting.*', '.*Tourism.*']
        cohorts = ['CCCUJUN22AF_L3_01', 'CCCUJUN22BTM_L3_01']
    elif intake == 'Sep22':
        programmes = ['.*Accounting.*', '.*Tourism.*']
        cohorts = ['CCCUSEP22AF_L3_01', 'CCCUSEP22BTM_L3_01']
    else:
        sys.exit("Invalid intake.")

# Should be no need to change anything below this line
# ====================================================

# Create workbook object df
if inputSuffix == '.xlsx':
    df = pd.read_excel(inputFile)
elif inputSuffix == '.csv':
    df = pd.read_csv(inputFile)
else:
    sys.exit("Invalid file type.")

# Create the new column headers
oldColumnHeaders = df.columns  # Get the column headers from the file
columnHeaders = [i.lower() for i in oldColumnHeaders]  # Make them all lowercase
## Search for keywords and replace with the headers we actually want
columnHeaders = ['username' if 'lead' in i or 'gbs id' in i or 'student id' in i else i for i in columnHeaders]
columnHeaders = ['firstname' if 'first name' in i else i for i in columnHeaders]
columnHeaders = ['middlename' if 'middle' in i and newAccounts else i for i in columnHeaders]
columnHeaders = ['lastname' if 'last name' in i else i for i in columnHeaders]
columnHeaders = ['cohort1' if 'course' in i or 'program' in i else i for i in columnHeaders]
columnHeaders = ['level' if 'study' in i else i for i in columnHeaders]
columnHeaders = ['campus' if ('campus' in i and not 'change' in i) or 'final campus' in i else i for i in columnHeaders]
columnHeaders = ['group' if 'group' in i and not 'campus' in i else i for i in columnHeaders]
columnHeaders = ['suspended' if 'status' in i and not 'visa' in i and not 'date' in i and not 'progression' in i else i
                 for i in columnHeaders]
print("\nThe column headers :")
print(columnHeaders)

df.columns = columnHeaders

# Pull out the columns we want
columns = ['username']
if newAccounts:
    columns.append('firstname')
    columns.append('lastname')
if middleNameColumn:
    columns.append('middlename')
if levelDataNeeded:
    columns.append('level')
if cohortChange:
    columns.append('cohort1')
if statusChange:
    columns.append('suspended')
if groupChange:
    columns.append('campus')
    columns.append('group')

df = df[columns]

# Make it all strings
df = df.astype(str)
# Strip whitespace from cell contents
for col in df.columns:
    try:
        df[col] = df[col].str.strip()
    except AttributeError:
        pass

# Make usernames lowercase
df['username'] = df['username'].str.lower()

# Merge columns for programme/level
if levelDataNeeded:
    df['cohort1'] = df['cohort1'] + ' ' + df['level']

# Search-replace
# ---------------

## Account suspension
if statusChange:
    df['suspended'] = df['suspended'].replace(
        '(?i).*withdraw.*|.*cancel.*|.*declined.*|.*IOS.*|.*intercalated.*|.*suspend.*|.*complete.*|.*defer.*', '1',
        regex=True)
    df['suspended'] = df['suspended'].replace('(?i).*enrol.*|.*attended.*|.*pending.*|.*progress.*', '0', regex=True)

## Cohort names
if cohortChange:
    for i in range(len(programmes)):
        df['cohort1'] = df['cohort1'].replace(programmes[i], cohorts[i], regex=True)

## Campus names
if groupChange:
    df['campus'] = df['campus'].replace('(?i).*Greenford.*', 'Greenford_', regex=True)
    df['campus'] = df['campus'].replace('(?i).*Global.*|.*Duncan.*|.*GedU.*', 'Global_Edu_', regex=True)
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
# Clean up group names (take first two characters of string)
if groupChange:
    df['group'] = df['group'].str[:2]
# Merge columns for groups and delete unwanted columns
if groupChange:
    df['profile_field_GroupAllocation'] = df['campus'] + df['group']
    df = df.drop(['campus', 'group'], axis=1)
    # df.loc[df['suspended'] == '1', 'profile_field_GroupAllocation'] = "suspended"
    # This is an option in case we want a separate group for suspended accounts - not in use ATM

if levelDataNeeded:
    df = df.drop(['level'], axis=1)

# Get rid of NaN cells
df.dropna() # This doesn't seem to be working at the moment.

# Split dataframes for ease of uploading if desired
if maxFileLength > 0:
    rowsToGo = len(df.index)
    print(rowsToGo)
    frameNumber = 0
    print('Files to upload to Moodle are:')
    while rowsToGo > 0:
        frameNumber = frameNumber + 1
        frameName = f"df{frameNumber}"
        if rowsToGo < maxFileLength:
            startPoint = 0
        else:
            startPoint = rowsToGo - maxFileLength
        frameName = df.iloc[startPoint:rowsToGo]
        outputFile = f"{partner}{intake}_{frameNumber}{outputSuffix}"
        frameName.to_csv(outputFile, index=False)
        print(f"Rows remaining: {rowsToGo}")
        print(f"Start point: {startPoint}")
        rowsToGo = rowsToGo - maxFileLength
        print('\n', outputFile)
        print(frameName)

else:
    print('File to upload to Moodle is ', outputFile)
    df.to_csv(outputFile, index=False)
    print(df)
