# GoogleSheetsNotification

2020 year
After my customer migrated to GoogleTables, he asked me to automatize some processes in GoogleTables.
He needed to read definite cells and analyze them.
Then color cells and show at the desktop.

There are 4 programs in the package:
- Notificater for the 1st manager
- Notificater for the 2nd manager (different countries)
- Setup wizard
- Uninstaller (kills all processes)

I used DataGridView to draw tables.

I doubled two similar projects in one package.
They are made for two named people, who work with different spreadsheets.


About mistakes.

1) MainForm.cs is too long. It would be better to divide it into several abstracts.