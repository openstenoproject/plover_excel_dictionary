# Plover Excel dictionary

Add support for LibreOffice (`.ods`) and Excel (`.xlsx`) dictionaries.

## Usage

How this (should) work: all the sheets of the spreadsheet are merged on load;
for each sheet, the first column should be the steno, and the second column
the translation. The implementation tries to keep extraneous columns' contents
when possible (see *Notes* below). New/modified entries are added to the `NEW`
spreadsheet. The order of the entries is kept, invalid entries are pruned.

Note:
 - changing an entry through Plover is equivalent to deleting it and creating
   a new one, so any previous extra data is lost, and it's moved to the `NEW`
   sheet.
 - formulas are expanded on load, and the result of those expansions is saved
   back if the dictionary is modified in Plover, so you should avoid editing
   your dictionary in Plover if you want to keep those formulas.


## Release history

### 1.0.0

* ensure the best available plugin is used for each format save/load operation
* switch to a better maintained `pyexcel` plugin for our default `.ods` support
* drop support for keeping extraneous data when changing an entry's translation

### 0.2.4

* drop Python 2 support
