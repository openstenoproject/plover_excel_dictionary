# vim: set fileencoding=utf-8 :

from collections import defaultdict, OrderedDict
import os
import shutil

# Python 2/3 compatibility.
from six import iteritems

import pyexcel

from plover.steno_dictionary import StenoDictionary
from plover.steno import normalize_steno


NEW_SHEET_NAME = 'NEW'


class ExcelDictionary(StenoDictionary):

    def __init__(self):
        super(ExcelDictionary, self).__init__()
        self._dict = OrderedDict()
        self._sheets = []
        self._extras = {}

    def __delitem__(self, key):
        super(ExcelDictionary, self).__delitem__(key)
        del self._extras[key]

    def _load(self, filename):
        book = pyexcel.get_book_dict(file_name=filename)
        def load():
            for sheet, entries in book.items():
                self._sheets.append(sheet)
                for row in entries:
                    if not row or not row[0]:
                        continue
                    translation = row[1] if len(row) > 1 else ''
                    steno = normalize_steno(row[0])
                    yield steno, translation
                    self._extras[steno] = (sheet, row[2:])
        self.update(load())

    def _save(self, filename):
        book = OrderedDict()
        for sheet in self._sheets:
            book[sheet] = []
        book[NEW_SHEET_NAME] = []
        default_extras = (NEW_SHEET_NAME, [])
        for k, v in iteritems(self._dict):
            sheet, extras = self._extras.get(k, default_extras)
            book[sheet].append(['/'.join(k), v] + extras)
        # pyexcel needs the correct extension to detect the file type...
        ext = os.path.splitext(self.path)[1]
        assert ext in filename
        pyexcel.save_book_as(bookdict=book, dest_file_name=filename + ext)
        shutil.move(filename + ext, filename)
