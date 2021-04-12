# vim: set fileencoding=utf-8 :

from collections import OrderedDict
import os

import pyexcel
import importlib_metadata

from plover import log
from plover.steno_dictionary import StenoDictionary
from plover.steno import normalize_steno


def _first_available_package(*package_list):
    for package in package_list:
        try:
            importlib_metadata.distribution(package)
        except importlib_metadata.PackageNotFoundError:
            continue
        return package
    return None


# Preferred reader/writers for each formats.
PREFERRED_READER = {
    '.ods': _first_available_package('pyexcel-ods', 'pyexcel-ods3'),
    '.xlsx': _first_available_package('pyexcel-xlsx'),
}
PREFERRED_WRITER = {
    '.ods': PREFERRED_READER['.ods'],
    '.xlsx': _first_available_package('pyexcel-libxlsxw', 'pyexcel-xlsx'),
}

# Sheet name for modified entries.
NEW_SHEET_NAME = 'NEW'


class ExcelDictionary(StenoDictionary):

    def __init__(self):
        super().__init__()
        self._dict = OrderedDict()
        self._sheets = []
        self._extras = {}

    def __delitem__(self, key):
        super().__delitem__(key)
        del self._extras[key]

    def _load(self, filename):
        ext = os.path.splitext(filename)[1]
        reader = PREFERRED_READER[ext]
        log.info('reading %r using %s reader', filename,
                 repr(reader) if reader else 'default')
        book = pyexcel.get_book_dict(file_name=filename, library=reader)
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
        ext = os.path.splitext(self.path)[1]
        assert filename.endswith(ext)
        writer = PREFERRED_WRITER[ext]
        log.info('writing %r using %s writer', filename,
                 repr(writer) if writer else 'default')
        book = OrderedDict()
        for sheet in self._sheets:
            book[sheet] = []
        book[NEW_SHEET_NAME] = []
        default_extras = (NEW_SHEET_NAME, [])
        for k, v in self._dict.items():
            sheet, extras = self._extras.get(k, default_extras)
            book[sheet].append(['/'.join(k), v] + extras)
        pyexcel.save_book_as(bookdict=book, dest_file_name=filename, dest_library=writer)
