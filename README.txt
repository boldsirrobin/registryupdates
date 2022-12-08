The files in this folder are simple Python utilities. Obviously you need Python installed to run them; for the registry
file converters you also need pandas and numpy.

You need to have a folder with the structure
- Main folder
  |- python
  |- excelfiles
  |- csvfiles

All the python files go in ./python, the files from registry go in ./excelfiles, and the CSV output will end up in
./csvfiles (if you prefer to have them all in the same folder, just edit the program to remove all refereces to the
inputPath and outputAth variables). Excel filenames must be in the form partnerintake_date.xlsx e.g.,
CCCUSep22_20221123.xls. The wording of the partner part is
important (GBS, CCCU, BSU or UoS); the intake part can be anything you like so long as it's the same as what you enter
into the script.

You need to set the variables at the beginning of the script for partner (e.g., CCCU), intake (whatever
you put in the file name) and dateRecieved (the date, preferably in Unix format for consistency, e.g., 20221219). There
are also Booleans (e.g., newAccounts = False). Set these according to what you want to do with the file. For example,
some of the files we get are only for withdrawals and reinstatements, so the only variable that needs to be True is
status Change. After setting the variables, you can run the script and it should do its magic regardless of the
format of the file (fingers crossed).

SPECIFIC FILES

registryTransform.py

This is the main script. Currently it works for most GBS files, CCCU and BSU but NOT for UoS Weekly updates.

uosTransform.py

This is just for UoS weekly updates.

fileDiff.py

Experimental - ignore this