# VBA

Some basic functions I put together from stackoverflow and googling mainly.

- **backup.bas** creates a copy of the file and adds a timestamp to the filename of the copy. I used it when I ran a huge macro, that took 3 minutes to process thousands of csv files, and was constantly afraid I'd loose the file. This was my quick and dirty attempt at version control for that one instance.

- **clearSheets.bas** deletes every row of the sheet starting from the 5th.

- **email.bas** An email opens in Outlook, fully composed, with the excel file attached. You'll just have to glance at it, and press 'Send'. When I used this, I mapped it to "Ctrl+E" for more elegant magic.

- **importCSV.bas** imports a CSV file to a sheet.

- **sort.bas** sorts specified rows by two different keys in the appropriate order of priorities, and then renumbers the index column.
