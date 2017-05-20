Plover Excel dictionary
=======================

Add support for LibreOffice (`.ods`) and Excel (`.xlsx`) dictionaries.

Usage
-----

How this (should) work: all the sheets of the spreadsheet are merged on load;
for each sheet, the first column should be the steno, and the second column the
translation. The implementation tries to keep extraneous columns contents, so
if you do modify the dictionary through Plover, you should not loose those
(except for deleted entries of course). New entries are added to the `NEW`
spreadsheet. The order of the entries is kept, invalid entries are prunned.

Note:
 - changing an entry's strokes (not its translation) through Plover's editor is
   equivalent to deleting it and and adding a new one, so any extra data is
   lost, and it's moved to the `NEW` sheet.
 - formulas are expanded on load, and the result of those expansions is saved
   back if the dictionary is modified in Plover, so you should avoid editing
   your dictionary in Plover if you want to keep those formulas.
